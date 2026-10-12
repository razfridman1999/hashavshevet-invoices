"""A report she can send Raz when something reads wrong. It is built only from codes and
yes/no facts: no supplier names, ID numbers, amounts, invoice numbers, e-mails or file
names - it is her employer's data. She sees the whole text before she sends anything."""
from __future__ import annotations

import platform
import re
from collections import Counter
from datetime import datetime
from pathlib import Path

from . import journal

FRAME = re.compile(r'^\s*File "(?P<path>[^"]+)", line (?P<line>\d+), in (?P<func>\w+)')
EXC = re.compile(r"^(?P<name>[A-Za-z_][\w.]*(Error|Exception|Exit|Interrupt))\b")


def _error_outline(home: Path) -> list[str]:
    """Where the last crash happened (our file, line, function) and its type. The
    exception's message is left out: it can quote a path or a value from an invoice."""
    log = home / "שגיאה אחרונה.txt"
    if not log.exists():
        return []
    lines = [f"עודכן: {datetime.fromtimestamp(log.stat().st_mtime):%Y-%m-%d %H:%M}"]
    for raw in log.read_text(encoding="utf-8", errors="replace").splitlines():
        if m := FRAME.match(raw):
            lines.append(f'  {Path(m["path"]).name}:{m["line"]} {m["func"]}')
        elif m := EXC.match(raw.strip()):
            lines.append(f'  {m["name"]}')
    return lines


def build(home: Path, version: str, her_note: str = "") -> str:
    entries = list(journal.load(home).values())
    statuses = Counter(e.get("status", "") for e in entries)
    codes = Counter(c for e in entries for c in e.get("codes", []))
    out = [f"דוח לרז - חשבוניות לחשבשבת {version}",
           f"נוצר: {datetime.now():%Y-%m-%d %H:%M} | {platform.platform()} | Python {platform.python_version()}",
           "", "מה כתבה:", her_note.strip() or "(לא נכתב כלום)", "",
           f"חשבוניות ביומן: {len(entries)}"]
    out += [f"  {status}: {n}" for status, n in statuses.most_common()]
    out += ["", "סוגי הערות (כמה פעמים):"] + ([f"  {c}: {n}" for c, n in codes.most_common()] or ["  אין"])
    out += ["", "חשבוניות עם הערות (בלי תוכן, רק מה נמצא):"]
    n = 0
    for e in entries:
        if not e.get("codes"):
            continue
        n += 1
        found = " ".join(f"{label}={'כן' if e.get(k) not in (None, '') else 'לא'}"
                         for k, label in (("supplier_id", "חפ"), ("invoice_number", "מספר"),
                                          ("invoice_date", "תאריך"), ("total", "סכום"), ("vat", "מעמ")))
        out.append(f"  #{n} {e.get('status', '')} | {','.join(e['codes'])} | {found}")
    if not n:
        out.append("  אין")
    error = _error_outline(home)
    out += ["", "שגיאה אחרונה:"] + (error or ["  אין"])
    return "\n".join(out) + "\n"


def write(home: Path, version: str, her_note: str = "") -> Path:
    path = home / f"דוח לרז {datetime.now():%Y-%m-%d %H%M%S}.txt"
    path.write_text(build(home, version, her_note), encoding="utf-8-sig")   # BOM: Notepad shows Hebrew
    return path
