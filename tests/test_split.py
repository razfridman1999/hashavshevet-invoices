"""One PDF with several invoices is cut into one invoice per file; an invoice that runs
over several pages stays whole; an unclear file is never cut, only marked."""
from __future__ import annotations

import sys
from pathlib import Path

import fitz
import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

from invoices import journal, main, split  # noqa: E402
from test_webapp import home  # noqa: E402,F401  (fixture)

A, B = "513100008", "513100016"


@pytest.mark.parametrize("keys, groups", [
    ([(A, "1"), (B, "2")], [[0], [1]]),                       # two suppliers
    ([(A, "1"), (A, "2")], [[0], [1]]),                       # two invoices of one supplier
    ([(A, "1"), ("", ""), (B, "2")], [[0, 1], [2]]),          # a second page without a header
    ([(A, "1"), (A, "1"), (B, "2")], [[0, 1], [2]]),          # the header repeats on page 2
    ([(A, "1"), (A, ""), ("", "1")], [[0, 1, 2]]),            # only the ח.פ, only the number
])
def test_pages_are_grouped_into_invoices(keys, groups):
    assert split.invoice_groups(keys) == (groups, False)


@pytest.mark.parametrize("keys", [
    [("", ""), (A, "1")],            # a cover letter, then an invoice: where does it belong?
    [(A, "1"), ("", "9")],           # a new number with no ח.פ: new invoice or a delivery note?
    [(A, "1"), (B, "")],             # another company's ח.פ with no number
])
def test_unclear_files_are_never_cut(keys):
    assert split.invoice_groups(keys) == ([list(range(len(keys)))], True)


def merge(out: Path, *pdfs: Path) -> Path:
    doc = fitz.open()
    for p in pdfs:
        with fitz.open(p) as src:
            doc.insert_pdf(src)
    doc.save(out)
    return out


def test_real_pdf_with_two_invoices_is_cut(home):
    from make_demo_invoices import OUT
    both = merge(home / "both.pdf", OUT / "01_electric.pdf", OUT / "02_office_table.pdf")
    parts, note = split.split_pdf(both)
    assert note == "" and [pages for _, pages in parts] == ["1-1", "2-2"]
    for part, _ in parts:
        assert fitz.open(part).page_count == 1
    assert split.split_pdf(OUT / "01_electric.pdf") == ([], "")


def test_fetch_gives_each_invoice_its_own_row(home):
    from make_demo_invoices import OUT
    merge(OUT / "monthly.pdf", OUT / "01_electric.pdf", OUT / "02_office_table.pdf")
    (OUT / "01_electric.pdf").unlink()
    (OUT / "02_office_table.pdf").unlink()
    main.fetch()
    entries = journal.load(home)
    cut = {k: e for k, e in entries.items() if "#" in k}
    assert sorted(e["invoice_number"] for e in cut.values()) == ["20261045", "58213"]
    assert {k.rsplit("#", 1)[1] for k in cut} == {"1-1", "2-2"}
    main.fetch()                                          # a second run: no new rows
    assert len(journal.load(home)) == len(entries)


def test_file_handled_whole_before_is_not_cut_into_new_rows(home):
    """An older version handled it as one invoice: cutting it now would bring both back."""
    from make_demo_invoices import OUT
    from invoices import sources
    monthly = merge(OUT / "monthly.pdf", OUT / "01_electric.pdf", OUT / "02_office_table.pdf")
    main.save(main.STATE, {"mail_ids": ["file:" + sources.file_hash(monthly)], "invoice_keys": []})
    main.fetch()
    assert not any("#" in k for k in journal.load(home))


def test_unclear_file_stays_whole_and_says_why(home):
    from make_demo_invoices import OUT
    cover = fitz.open()
    cover.new_page().insert_text((72, 72), "Cover letter for the attached documents, see next pages please")
    cover.save(home / "cover.pdf")
    merge(OUT / "with_cover.pdf", home / "cover.pdf", OUT / "01_electric.pdf")
    main.fetch()
    row = next(e for k, e in journal.load(home).items() if k.endswith("with_cover.pdf") or "with_cover" in str(e.get("file")))
    assert "multi_unclear" in journal.issue_codes({"notes": row["todo"]})
