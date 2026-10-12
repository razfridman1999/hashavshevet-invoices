"""Every layout the demo set covers, read end to end from real PDF files."""
from __future__ import annotations

import sys
from datetime import date
from decimal import Decimal
from pathlib import Path

import fitz
import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "demo"))

import make_demo_invoices as demo  # noqa: E402
from invoices.extract import extract, unreverse_numbers, valid_israeli_id  # noqa: E402

OWN = {demo.OWN_ID}
S = demo.S


@pytest.fixture(scope="module")
def pdfs(tmp_path_factory):
    demo.OUT = tmp_path_factory.mktemp("invoices")
    demo.main()
    return demo.OUT


def read(pdfs, name):
    return extract(pdfs / name, OWN)


def test_israeli_id_checksum():
    assert valid_israeli_id("513100008")
    assert not valid_israeli_id("513100009")
    assert not valid_israeli_id("12345678a")


def test_plain_invoice(pdfs):
    inv = read(pdfs, "01_electric.pdf")
    assert (inv.supplier_name, inv.supplier_id, inv.invoice_number) == ('חשמל השרון בע"מ', S[0], "20261045")
    assert inv.invoice_date == date(2026, 10, 5)
    assert (inv.net, inv.vat, inv.total) == (Decimal("420.00"), Decimal("75.60"), Decimal("495.60"))
    assert inv.allocation_number == "300112233"
    assert inv.warnings == []


def test_table_layout_with_two_digit_year(pdfs):
    inv = read(pdfs, "02_office_table.pdf")
    assert inv.invoice_number == "58213"
    assert inv.invoice_date == date(2026, 10, 7)
    assert (inv.vat, inv.total) == (Decimal("209.66"), Decimal("1374.46"))
    assert inv.warnings == []


def test_exempt_dealer_has_no_vat(pdfs):
    inv = read(pdfs, "03_exempt_dealer.pdf")
    assert inv.invoice_number == "112"
    assert (inv.vat, inv.total) == (Decimal("0"), Decimal("1500.00"))
    assert inv.warnings == []


def test_credit_note_is_negative(pdfs):
    inv = read(pdfs, "04_credit_note.pdf")
    assert inv.invoice_number == "7731"
    assert (inv.net, inv.vat, inv.total) == (Decimal("-200.00"), Decimal("-36.00"), Decimal("-236.00"))
    assert inv.warnings == []


def test_buyer_id_printed_first_is_not_taken_for_supplier(pdfs):
    inv = read(pdfs, "05_software_large_buyer_first.pdf")
    assert inv.supplier_id == S[4]
    assert inv.supplier_name == 'תוכנות מרום בע"מ'
    assert inv.total == Decimal("8024.00")


def test_wrong_vat_is_flagged(pdfs):
    inv = read(pdfs, "06_wrong_vat.pdf")
    assert any("לא תואם" in w for w in inv.warnings)


def test_scan_goes_to_manual(pdfs):
    inv = read(pdfs, "07_scanned.pdf")
    assert not inv.has_text and inv.warnings


def visual_order_pdf(path: Path, vat: str, total: str) -> Path:
    """Some invoicing systems store each line backwards; digits then come out reversed."""
    lines = ['שירותי מחשוב צפון בע"מ', f"ח.פ {S[6]}", "חשבונית מס 4521", "תאריך: 06/10/2026",
             'סה"כ לפני מע"מ: 500.00', f'מע"מ 18%: {vat}', f'סה"כ לתשלום: {total}']
    doc = fitz.open()
    pg = doc.new_page()
    for i, line in enumerate(lines):
        pg.insert_text((60, 60 + 22 * i), line[::-1], fontname="arial",
                       fontfile=r"C:\Windows\Fonts\arial.ttf", fontsize=12)
    doc.save(path)
    return path


def test_visual_order_wrong_vat_is_still_flagged(tmp_path):
    """Picking the text orientation must never 'repair' an invoice whose VAT is wrong."""
    inv = extract(visual_order_pdf(tmp_path / "v.pdf", "100.00", "600.00"), OWN)
    assert inv.total == Decimal("600.00")
    assert any("לא תואם" in w for w in inv.warnings)


def test_visual_order_pdf_is_read_correctly(tmp_path):
    inv = extract(visual_order_pdf(tmp_path / "visual.pdf", "90.00", "590.00"), OWN)
    assert (inv.supplier_id, inv.invoice_number) == (S[6], "4521")
    assert inv.invoice_date == date(2026, 10, 6)
    assert (inv.vat, inv.total) == (Decimal("90.00"), Decimal("590.00"))
    assert inv.warnings == []


def test_unreverse_numbers():
    assert unreverse_numbers("6202/01/60 :תאריך 00.095") == "06/10/2026 :תאריך 590.00"
