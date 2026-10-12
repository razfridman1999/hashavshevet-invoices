"""Where invoice PDFs come from: classic Outlook on this PC, or a plain folder (for testing).

Outlook is driven through COM, which works only with *classic* Outlook for Windows.
The "new Outlook" app has no COM interface; if it is the one in use, switch back with
the toggle in its top-right corner.
"""
from __future__ import annotations

import hashlib
import re
import shutil
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path

INVOICE_HINT = re.compile(r"חשבונית|חשבון\s*עסקה|קבלה|invoice|receipt|tax", re.I)


@dataclass
class Fetched:
    pdf: Path
    mail_id: str        # stable id, used to never process the same attachment twice
    sender: str
    subject: str
    received: datetime
    entry_id: str = ""  # Outlook's own id, to move the mail to "טופל" once it is handled
    original: str = ""  # the whole PDF, when this is one invoice split out of it
    note: str = ""      # e.g. why a PDF that may hold several invoices was not split


def file_hash(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()[:16]


def from_folder(folder: Path, inbox: Path) -> list[Fetched]:
    out = []
    for pdf in sorted(folder.glob("*.pdf")):
        dest = inbox / pdf.name
        if not dest.exists():
            shutil.copy2(pdf, dest)
        out.append(Fetched(dest, "file:" + file_hash(pdf), "", pdf.stem,
                           datetime.fromtimestamp(pdf.stat().st_mtime)))
    return out


def _namespace():
    import win32com.client  # only needed on the Outlook machine
    return win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")


def _parts(folder_path: str) -> list[str]:
    """"a/b" or "a\\b" below the Inbox; the mapping table she keeps uses backslashes."""
    return [p.strip() for p in re.split(r"[\\/]", folder_path or "") if p.strip()]


def _folder(ns, folder_path: str):
    folder = ns.GetDefaultFolder(6)  # 6 = olFolderInbox
    for part in _parts(folder_path):
        folder = folder.Folders[part]
    return folder


def from_outlook(inbox: Path, days_back: int = 0, folder_path: str = "") -> list[Fetched]:
    """Every invoice PDF in the folder. A client's own folder is read from start to end
    (handled mail leaves it for "טופל"); only the plain Inbox keeps a time limit and the
    check that the mail looks like an invoice at all."""
    folder = _folder(_namespace(), folder_path)
    own_folder = bool(_parts(folder_path))
    # The cut-off is applied here in Python. Outlook's Restrict() takes the date as a
    # string in the machine's regional format, so a different Windows setting would
    # silently move the cut-off.
    since = datetime.now() - timedelta(days=days_back) if days_back and not own_folder else None
    items = folder.Items
    items.Sort("[ReceivedTime]", True)  # newest first, so we can stop at the first old one
    out = []
    for item in items:
        if getattr(item, "Class", None) != 43:  # 43 = olMail; skip meeting requests etc.
            continue
        try:
            received = item.ReceivedTime.replace(tzinfo=None)  # pywin32 labels local time as UTC
        except Exception:
            continue
        if since and received < since:
            break
        for i in range(1, item.Attachments.Count + 1):
            att = item.Attachments.Item(i)
            name = att.FileName or ""
            if not name.lower().endswith(".pdf"):
                continue
            if not own_folder and not (INVOICE_HINT.search(item.Subject or "") or INVOICE_HINT.search(name)
                                       or INVOICE_HINT.search(item.Body[:2000] if item.Body else "")):
                continue
            # never put ':' in a Windows file name: NTFS reads it as a hidden data stream
            tag = mail_tag(item.EntryID)
            mail_id = f"outlook:{tag}-{i}"
            safe = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", name)[-80:]
            dest = inbox / f"{received:%Y%m%d}_{tag[:8]}-{i}_{safe}"
            if not dest.exists():
                att.SaveAsFile(str(dest))
            out.append(Fetched(dest, mail_id, item.SenderEmailAddress or "", item.Subject or "", received,
                               entry_id=item.EntryID))
    return out


def mail_tag(entry_id: str) -> str:
    return hashlib.sha1(entry_id.encode()).hexdigest()[:12]


def tag_of(key: str) -> str:
    """"outlook:<tag>-<n>[#...]" -> "<tag>"; "" for anything that is not an Outlook mail."""
    m = re.match(r"outlook:([0-9a-f]{12})-", key or "")
    return m.group(1) if m else ""


def move_mails(entry_ids: list[str], folder_path: str) -> tuple[int, list[str]]:
    """Move handled mail to the client's "טופל" folder. (moved, problems). A problem never
    undoes anything: the invoices are already in the import file."""
    if not entry_ids:
        return 0, []
    ns = _namespace()
    try:
        target = _folder(ns, folder_path)
    except Exception:
        return 0, [f'תיקיית "{folder_path}" לא נמצאה באאוטלוק, המיילים לא הועברו']
    moved, problems = 0, []
    for eid in entry_ids:
        try:
            ns.GetItemFromID(eid).Move(target)
            moved += 1
        except Exception:
            problems.append("מייל אחד לא הועבר (אולי כבר הועבר או נמחק)")
    return moved, problems


def folder_paths(depth: int = 3) -> list[str]:
    """Every folder below the Inbox as "a\\b", for the settings page to pick from."""
    out: list[str] = []

    def walk(folder, prefix: str, level: int) -> None:
        if level > depth:
            return
        for i in range(1, folder.Folders.Count + 1):
            sub = folder.Folders.Item(i)
            path = f"{prefix}\\{sub.Name}" if prefix else sub.Name
            out.append(path)
            walk(sub, path, level + 1)

    walk(_namespace().GetDefaultFolder(6), "", 1)
    return out
