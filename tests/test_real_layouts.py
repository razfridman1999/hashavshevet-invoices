"""Layouts from the first real run at her office (125 invoices). The text is as the PDFs
gave it; every company number is replaced with a fictional one and the names are made up.
Each case failed before: VAT read as the rate, as the total or as the buyer's VAT file;
numbers glued to their labels or printed before them."""
from __future__ import annotations

import sys
from datetime import date
from decimal import Decimal as D
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "demo"))

from invoices import journal  # noqa: E402
from invoices.extract import extract_from_text  # noqa: E402
from make_demo_invoices import OWN_ID as OWN, S  # noqa: E402

SUP = S[3]


def read(text: str):
    return extract_from_text(text, "x.pdf", {OWN})


def amounts(inv):
    return inv.total, inv.vat, inv.net


def test_vat_rate_on_its_own_line_then_the_amount():          # a mobile operator's bill
    inv = read(f'תקשורת לדוגמה בע"מ\nח.פ {SUP}\nחשבונית מס 3344556677\nתאריך: 28/09/2026\n'
               'סה"כ חשבון נוכחי כולל מע"מ₪ 59.00\nסכומים חייבים במע"מ50.00\n מע"מ18%\n9.00')
    assert amounts(inv) == (D("59.00"), D("9.00"), D("50.00")) and not inv.warnings


def test_net_vat_and_percent_split_over_lines_and_a_glued_number():   # a phone company
    inv = read(f'בינלאומי לדוגמה בע"מ\nח.פ {SUP}\nחשבונית מס מקור2026999123456\nתאריך: 01/11/2026\n'
               'תאריך עריכת חשבון28/09/26\nתקופת חשבון09/2026\n'
               'סה"כ לפני מע"מ300.00\n% מע"מ18.00\n54.00\nסך הכל לתשלום354.00ש"ח')
    assert amounts(inv) == (D("354.00"), D("54.00"), D("300.00"))
    assert inv.invoice_number == "2026999123456" and inv.invoice_date == date(2026, 9, 28)


def test_total_taken_as_vat_in_a_table():                       # a control company
    inv = read(f'בקרה לדוגמה בע"מ\nח.פ {SUP}\nחשבונית מס 5512\nתאריך: 03/09/2026\n'
               'סה"כ1,500.00\n270.00\n18.0 %מע"מ\n1,770.00סה"כ כולל מע"מ')
    assert amounts(inv) == (D("1770.00"), D("270.00"), D("1500.00"))


def test_buyers_vat_file_number_is_never_an_amount():          # an elevator company
    inv = read(f'מעליות לדוגמה בע"מ\nח.פ {SUP}\n334455667מספר חשבונית\n13579246מספר לקוח\n24681357מספר חוזה\n'
               f'{OWN}מספר מע"מ\nתאריך: 01/09/2026\n2,360.00\n360.00\n18.00\n2,000.00')
    assert amounts(inv) == (D("2360.00"), D("360.00"), D("2000.00"))
    assert inv.invoice_number == "334455667"


def test_vat_above_its_label_and_allocation_before_its_label():
    inv = read(f'שירותים לדוגמה בע"מ\nח.פ {SUP}\nחשבונית מס 7001\nתאריך: 31/08/2026\n'
               f"{OWN} :מס' תיק במע\"מ\n"
               '50,000.00סה"כ\n9,000.00\n(18.00%) מע"מ\n ש"ח59,000.00סה"כ כולל מע"מ\n300112299:מספר הקצאה')
    assert amounts(inv) == (D("59000.00"), D("9000.00"), D("50000.00"))
    assert inv.allocation_number == "300112299"


def test_rate_after_the_vat_label_and_amounts_far_below():     # a document system's layout
    inv = read(f'ניהול לדוגמה בע"מ\nח.פ {SUP}\nחשבונית מס 880\nתאריך: 10/09/2026\n'
               ':סה"כ לפני מע"מ\n:מע"מ18.00\n:סה"כ לתשלום11,800.00\nהערות\n-0.42\n10,000.00\n1,800.00')
    assert amounts(inv) == (D("11800.00"), D("1800.00"), D("10000.00"))


def test_an_amount_that_merely_fits_is_never_taken():
    """Only a pair whose both sides are printed: otherwise the mismatch stays visible."""
    inv = read(f'ספק לדוגמה בע"מ\nח.פ {SUP}\nחשבונית מס 9\nתאריך: 05/10/2026\n'
               'מע"מ: 5.00\nפריט 17.97\nסה"כ לתשלום: 118.00')
    assert inv.vat == D("5.00") and any(w.startswith('מע"מ ') for w in inv.warnings)


def test_a_report_that_is_not_an_invoice_says_so_and_names_no_supplier():
    inv = read("דוח בדיקה תקופתי למעלית\nמתאריך: 01/12/2011\nמספר נותן שרות: 40001\nהבדיקה הבאה עד 01/12/2012")
    assert inv.supplier_name == ""
    assert journal.issue_codes({"notes": " | ".join(inv.warnings)})[0] == "not_invoice"


def test_two_pairs_that_fit_means_no_choice_is_made():
    inv = read(f'ספק לדוגמה בע"מ\nח.פ {SUP}\nחשבונית מס 9\nתאריך: 05/10/2026\n'
               'מע"מ: 7.00\n200.00\n36.00\n199.96\n36.04\nסה"כ לתשלום: 236.00')
    assert inv.vat == D("7.00")                     # left as read, and the mismatch is shown
    assert any(w.startswith('מע"מ ') for w in inv.warnings)


def test_a_rate_written_before_the_number_is_not_an_amount():
    inv = read(f'ספק לדוגמה בע"מ\nח.פ {SUP}\nחשבונית מס 9\nתאריך: 05/10/2026\n'
               'סה"כ לפני מע"מ 100.00\nמע"מ %18.00\nסה"כ לתשלום: 118.00')
    assert inv.vat != D("18.00")


def test_a_contract_number_under_the_total_label_is_not_the_total():
    inv = read(f'ספק לדוגמה בע"מ\nח.פ {SUP}\nחשבונית מס 9\nתאריך: 05/10/2026\n'
               'סה"כ לתשלום\n24681357 מספר חוזה\n1,180.00\nמע"מ 18%: 180.00')
    assert inv.total == D("1180.00")


def test_the_vat_file_line_is_not_the_vat_amount():
    inv = read(f'ספק לדוגמה בע"מ\nח.פ {SUP}\nחשבונית מס 95\nתאריך: 05/10/2026\n'
               "מס' תיק במע\"מ 942\nסה\"כ לתשלום: 118.00\nמע\"מ 18%: 18.00")
    assert inv.vat == D("18.00") and not inv.warnings


def test_a_page_full_of_figures_is_read_quickly():
    """The pair search is n squared: a crafted PDF with thousands of amounts must not hang her fetch."""
    import time
    figures = "\n".join(f"{i}.{i % 100:02d}" for i in range(1, 6000))
    start = time.monotonic()
    read(f'ספק לדוגמה בע"מ\nח.פ {SUP}\nחשבונית מס 95\nתאריך: 05/10/2026\n{figures}')   # no total: the pair search
    assert time.monotonic() - start < 3
