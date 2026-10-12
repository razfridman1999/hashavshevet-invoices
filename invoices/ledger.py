"""The suppliers' ledger (כרטסת ספקים) exported from Hashavshevet to Excel. It answers two
questions before anything is imported: was this invoice already entered (by anyone, not
only by this tool), and which expense card does this supplier usually go to.

Layout of the export: a few title lines, then a header row ("ח-ן נגדי", "אסמ'", "אסמ'2",
"פרטים", "... חובה", "... זכות"). Each supplier is a block: a line with its name (A) and
card (B), an opening balance, its movements (A empty), and three 'סה"כ' lines.
An invoice is a credit line; its counter account (ח-ן נגדי) is the expense card. A payment
is a debit line against the bank and says nothing about the expense card. The invoice
number may be in אסמ' or in אסמ'2 (some clients keep אסמ' at 0)."""
from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path

from .extract import journal_reference
from .supplier_index import _text, read_rows

NAME_HINT = "כרטסת"


@dataclass
class Line:
    account: str
    refs: set[str]
    amount: Decimal        # always positive
    credit: bool           # True = an invoice line (זכות)
    counter: str           # ח-ן נגדי


@dataclass
class Ledger:
    lines: dict[str, list[Line]] = field(default_factory=dict)
    source: str = ""

    def has(self, account: str, invoice_number: str, total) -> bool:
        """Same supplier card + same invoice number (in either reference) + same amount."""
        ref = journal_reference(invoice_number or "")
        amount = _amount(total)
        if not account or not ref or amount is None:
            return False
        return any(ref in ln.refs and ln.amount == abs(amount) for ln in self.lines.get(str(account), []))

    def expense_accounts(self, account: str) -> list[str]:
        """The expense cards this supplier's invoices went to, most used first."""
        used = Counter(ln.counter for ln in self.lines.get(str(account), []) if ln.credit and ln.counter)
        return [card for card, _ in used.most_common()]


def _amount(v) -> Decimal | None:
    if v is None or v == "":
        return None
    try:
        return Decimal(str(v).replace(",", "").strip()).quantize(Decimal("0.01"))
    except InvalidOperation:
        return None


def parse(rows: list[list], source: str = "") -> Ledger:
    led = Ledger(source=source)
    cols = None
    account = ""
    for row in rows:
        cells = [_text(c) for c in row]
        if cols is None:
            if "ח-ן נגדי" in cells:
                find = lambda pred: next((i for i, c in enumerate(cells) if pred(c)), None)
                cols = {"counter": find(lambda c: c == "ח-ן נגדי"), "ref1": find(lambda c: c == "אסמ'"),
                        "ref2": find(lambda c: c == "אסמ'2"),
                        "debit": find(lambda c: "/" in c and c.endswith("חובה")),
                        "credit": find(lambda c: "/" in c and c.endswith("זכות"))}
            continue
        get = lambda k: cells[cols[k]] if cols[k] is not None and cols[k] < len(cells) else ""
        first = cells[0] if cells else ""
        if first:                       # a supplier's own line, or its totals
            if not first.startswith('סה"כ') and len(cells) > 1 and cells[1]:
                account = cells[1]
            continue
        debit, credit = _amount(get("debit")), _amount(get("credit"))
        if not account or not get("counter") or not (debit or credit):
            continue                    # opening balance, totals, blank lines
        refs = {journal_reference(get(k)) for k in ("ref1", "ref2")} - {"", "0"}
        led.lines.setdefault(account, []).append(
            Line(account, refs, abs(credit or debit), bool(credit), get("counter")))
    if cols is None:
        raise ValueError('לא נמצאה בקובץ שורת כותרות עם "ח-ן נגדי". לייצא את הכרטסת מחשבשבת לאקסל שוב.')
    return led


def find(cfg: dict) -> Path | None:
    """The ledger she set in the client's settings; otherwise the newest file with
    "כרטסת" in its name in the client's invoice folders (she drops it in the day's folder)."""
    if cfg.get("ledger"):
        return Path(cfg["ledger"])
    base = Path(cfg["files_folder"]) if cfg.get("files_folder") else None
    if not base or not base.is_dir():
        return None
    found = [p for pattern in (f"*{NAME_HINT}*.xls*", f"*/*{NAME_HINT}*.xls*") for p in base.glob(pattern)
             if not p.name.startswith("~$")]
    return max(found, key=lambda p: p.stat().st_mtime) if found else None


def load(path: Path) -> Ledger:
    return parse(read_rows(path), str(path))
