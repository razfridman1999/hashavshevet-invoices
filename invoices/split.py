"""One PDF, several invoices. Some suppliers send a month of invoices in one file, so each
page is read on its own and the file is cut where a new invoice starts. An invoice that
runs over several pages stays whole. Where it is not clear where one invoice ends and
the next begins, nothing is cut: the file stays one row, marked for her to check."""
from __future__ import annotations

from pathlib import Path

import fitz  # PyMuPDF

from .extract import SPACES, extract_from_text

UNCLEAR = "ייתכן שבקובץ יש כמה חשבוניות ולא ברור איפה כל אחת מתחילה: לבדוק, ולהוסיף ידנית את מה שחסר"

Key = tuple[str, str]   # (ח.פ of the supplier, invoice number); "" where a page shows none


def page_keys(path: Path, own_ids: set[str] = frozenset()) -> list[Key]:
    with fitz.open(path) as doc:
        texts = [page.get_text("text").translate(SPACES) for page in doc]
    keys = []
    for text in texts:
        inv = extract_from_text(text, str(path), own_ids)
        keys.append((inv.supplier_id or "", inv.invoice_number or ""))
    return keys


def invoice_groups(keys: list[Key]) -> tuple[list[list[int]], bool]:
    """Pages grouped into invoices, and whether the cut was unclear.

    A page with both a ח.פ and a number different from the invoice before it starts a
    new invoice. A page with neither continues the one before. Anything in between - a
    page with only one of the two that does not match, or a first page with no number
    followed by real invoices - is unclear, and then the file is never cut."""
    groups: list[list[int]] = []
    current: Key | None = None          # the invoice being collected, when fully known
    for i, (sid, num) in enumerate(keys):
        full = bool(sid and num)
        if not groups:
            groups.append([i])
            current = (sid, num) if full else None
        elif not sid and not num:
            groups[-1].append(i)
        elif full and current is None:
            return [list(range(len(keys)))], True
        elif full:
            if (sid, num) == current:
                groups[-1].append(i)
            else:
                groups.append([i])
                current = (sid, num)
        elif current and ((sid and sid == current[0]) or (num and num == current[1])):
            groups[-1].append(i)
        else:
            return [list(range(len(keys)))], True
    return groups, False


def split_pdf(path: Path, own_ids: set[str] = frozenset()) -> tuple[list[tuple[Path, str]], str]:
    """([(part file, "first-last" pages)], note). No parts when the file is one invoice
    or when the cut is unclear; the note says why it was left whole."""
    with fitz.open(path) as doc:
        if doc.page_count < 2:
            return [], ""
    groups, unclear = invoice_groups(page_keys(path, own_ids))
    if unclear:
        return [], UNCLEAR
    if len(groups) < 2:
        return [], ""
    parts = []
    with fitz.open(path) as doc:
        for g in groups:
            pages = f"{g[0] + 1}-{g[-1] + 1}"
            out = path.with_name(f"{path.stem}__עמ{pages}.pdf")
            if not out.exists():
                part = fitz.open()
                part.insert_pdf(doc, from_page=g[0], to_page=g[-1])
                part.save(out)
                part.close()
            parts.append((out, pages))
    return parts, ""
