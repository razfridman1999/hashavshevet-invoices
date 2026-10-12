"""What the log can tell her without anyone typing anything: a monthly VAT summary of
what went into import files, and suppliers who bill every month but not this time."""
from __future__ import annotations

from datetime import date
from decimal import Decimal

from . import journal


def _month(v) -> str:
    """'2026-10' from a date, a datetime or their text form; '' when unknown."""
    s = str(v or "")
    return s[:7] if len(s) >= 7 and s[4] == "-" else ""


def _shift(month: str, n: int) -> str:
    y, m = int(month[:4]), int(month[5:7]) - 1 + n
    return f"{y + m // 12:04d}-{m % 12 + 1:02d}"


def monthly_summary(entries: dict, months: int = 12) -> list[dict]:
    """Per invoice month, only what is in an import file. A month with an entry that
    lacks its amounts is marked partial instead of showing a sum that looks complete."""
    by_month: dict[str, dict] = {}
    for e in entries.values():
        month = _month(e.get("invoice_date"))
        if e.get("status") != journal.EXPORTED or not month:
            continue
        m = by_month.setdefault(month, {"month": month, "count": 0, "net": Decimal(0), "vat": Decimal(0),
                                        "total": Decimal(0), "partial": False})
        m["count"] += 1
        for k in ("net", "vat", "total"):
            if e.get(k) is None or e.get(k) == "":
                m["partial"] = True
            else:
                m[k] += Decimal(str(e[k]))
    rows = sorted(by_month.values(), key=lambda m: m["month"], reverse=True)[:months]
    return [{**m, **{k: float(m[k].quantize(Decimal("0.01"))) for k in ("net", "vat", "total")}} for m in rows]


def missing_invoices(entries: dict, today: date | None = None) -> list[dict]:
    """A supplier with an invoice in each of the three months before last month, and
    none for last month. Bi-monthly bills (electricity) never trigger it: no false alarm."""
    today = today or date.today()
    last = _shift(f"{today:%Y-%m}", -1)
    before = {_shift(last, -n) for n in (1, 2, 3)}
    months: dict[str, set[str]] = {}
    names: dict[str, str] = {}
    for e in entries.values():
        sid, month = e.get("supplier_id"), _month(e.get("invoice_date"))
        if not sid or not month or e.get("status") == journal.REJECTED:
            continue
        months.setdefault(sid, set()).add(month)
        names[sid] = e.get("supplier_name") or sid
    return sorted(({"supplier_id": sid, "supplier_name": names[sid], "month": last}
                   for sid, seen in months.items() if before <= seen and last not in seen),
                  key=lambda m: m["supplier_name"])
