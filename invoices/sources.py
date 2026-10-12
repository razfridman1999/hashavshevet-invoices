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


def from_outlook(inbox: Path, days_back: int, folder_path: str = "") -> list[Fetched]:
    import win32com.client  # only needed on the Outlook machine

    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    folder = ns.GetDefaultFolder(6)  # 6 = olFolderInbox
    for part in filter(None, folder_path.split("/")):
        folder = folder.Folders[part]

    # The cut-off is applied here in Python. Outlook's Restrict() takes the date as a
    # string in the machine's regional format, so a different Windows setting would
    # silently move the cut-off.
    since = datetime.now() - timedelta(days=days_back)
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
        if received < since:
            break
        for i in range(1, item.Attachments.Count + 1):
            att = item.Attachments.Item(i)
            name = att.FileName or ""
            if not name.lower().endswith(".pdf"):
                continue
            if not (INVOICE_HINT.search(item.Subject or "") or INVOICE_HINT.search(name)
                    or INVOICE_HINT.search(item.Body[:2000] if item.Body else "")):
                continue
            # never put ':' in a Windows file name: NTFS reads it as a hidden data stream
            tag = hashlib.sha1(item.EntryID.encode()).hexdigest()[:12]
            mail_id = f"outlook:{tag}-{i}"
            safe = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", name)[-80:]
            dest = inbox / f"{received:%Y%m%d}_{tag[:8]}-{i}_{safe}"
            if not dest.exists():
                att.SaveAsFile(str(dest))
            out.append(Fetched(dest, mail_id, item.SenderEmailAddress or "", item.Subject or "", received))
    return out
