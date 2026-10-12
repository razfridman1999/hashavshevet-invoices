"""One invoice over several expense cards ("חשבונית מפוצלת"). The building manager stamps
the invoice and writes which lines go to which card ("8110 - שורה 1, 8120 - שורות 5-15").
It is entered as one journal line per card, all with the same reference; the rounding
goes on the largest line, as she does it by hand.

On her sheet it is one row with a "פיצול לסעיפים" cell: "8110=1000; 8120=4000" - the
net of each card's lines. She can type it herself for any invoice."""
from __future__ import annotations

import re
from decimal import Decimal, InvalidOperation

PIECE = re.compile(r"(\d{2,12})\s*[=:]\s*(-?[\d,]*\.?\d+)")
TOLERANCE = Decimal("1.00")   # line totals vs. the invoice's net: rounding lines, never more


def parse(text) -> list[tuple[str, Decimal]]:
    out = []
    for card, amount in PIECE.findall(str(text or "")):
        try:
            out.append((card, Decimal(amount.replace(",", "")).quantize(Decimal("0.01"))))
        except InvalidOperation:
            continue
    return out


def format_split(pieces) -> str:
    return "; ".join(f"{card}={Decimal(str(net)):.2f}" for card, net in pieces)


def _d(v) -> Decimal:
    return Decimal(str(v or 0)).quantize(Decimal("0.01"))


def parts(row: dict) -> tuple[list[dict], str]:
    """The row as one journal line per card, or (the row alone, why it cannot be split).
    Net, VAT and total of the lines always add up to the invoice's own, to the agora."""
    pieces = parse(row.get("split"))
    if len(pieces) < 2:
        return [row], ""
    net, vat, total = _d(row.get("net")), _d(row.get("vat")), _d(row.get("total"))
    sign = Decimal(-1) if total < 0 else Decimal(1)
    net, vat = abs(net), abs(vat)
    nets = [abs(n) for _, n in pieces]
    diff = net - sum(nets)
    if abs(diff) > TOLERANCE or not net:
        return [row], (f"סכומי הסעיפים ({sum(nets)}) לא מסתכמים לסכום לפני מע\"מ ({net}): "
                       'לתקן את "פיצול לסעיפים"')
    big = max(range(len(nets)), key=lambda i: nets[i])
    nets[big] += diff                                  # the rounding goes on the largest line
    vats = [(n * vat / net).quantize(Decimal("0.01")) for n in nets]
    vats[big] = vat - (sum(vats) - vats[big])
    return [{**row, "expense_account": card, "net": float(sign * n), "vat": float(sign * v),
             "total": float(sign * (n + v))}
            for (card, _), n, v in zip(pieces, nets, vats)], ""
