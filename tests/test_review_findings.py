"""Regression tests for the independent review (one test per finding)."""
from __future__ import annotations

import sys
from datetime import date
from decimal import Decimal
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from invoices.extract import _extract_text  # noqa: E402

SUP, SUP2, OWN = "513100008", "513100016", "515550002"


def read(text: str):
    return _extract_text(text, "x.pdf", {OWN})


def has(inv, fragment: str) -> bool:
    return any(fragment in w for w in inv.warnings)


# --- finding 1: VAT not found must never pass silently --------------------------------
def test_english_layout_subtotal_vat_total():
    inv = read(f"Example Ltd\nVAT No {SUP}\nTax Invoice 5521\nDate: 05/10/2026\n"
               "Sub Total 100.00\nVAT 18% 18.00\nTotal 118.00")
    assert (inv.total, inv.vat) == (Decimal("118.00"), Decimal("18.00"))
    assert inv.warnings == []


def test_missing_vat_is_flagged():
    inv = read(f'ספק בע"מ\nח.פ {SUP}\nחשבונית מס 77\nתאריך: 05/10/2026\nסה"כ לתשלום: 118.00')
    assert has(inv, 'לא נמצא מע"מ')


def test_zero_vat_without_exempt_wording_is_flagged():
    inv = read(f'ספק בע"מ\nח.פ {SUP}\nחשבונית מס 77\nתאריך: 05/10/2026\nמע"מ: 0.00\nסה"כ לתשלום: 118.00')
    assert has(inv, 'מע"מ 0')


# --- finding 2: credit notes under other titles, and trailing minus ---------------------
def test_credit_note_title_variants():
    for title in ("חשבונית מס זיכוי 55", "חשבונית מס / זיכוי 55", "תעודת זיכוי 55"):
        inv = read(f'ספק בע"מ\nח.פ {SUP}\n{title}\nתאריך: 05/10/2026\n'
                   'סה"כ לפני מע"מ: 100.00\nמע"מ 18%: 18.00\nסה"כ לתשלום: 118.00')
        assert inv.total == Decimal("-118.00"), title
        assert inv.invoice_number == "55", title


def test_trailing_minus_is_negative():
    inv = read(f'ספק בע"מ\nח.פ {SUP}\nחשבונית מס 9\nתאריך: 05/10/2026\n'
               'סה"כ לפני מע"מ: 100.00-\nמע"מ 18%: 18.00-\nסה"כ לתשלום: 118.00-')
    assert (inv.total, inv.vat) == (Decimal("-118.00"), Decimal("-18.00"))


def test_stray_zikui_word_holds_the_invoice():
    inv = read(f'ספק בע"מ\nח.פ {SUP}\nחשבונית מס 9\nתאריך: 05/10/2026\nסה"כ לפני מע"מ: 100.00\n'
               'מע"מ 18%: 18.00\nסה"כ לתשלום: 118.00\nבמקרה של החזרה תונפק הודעת-זיכוי לפי מדיניות')
    assert inv.total == Decimal("118.00") or has(inv, "זיכוי")


# --- finding 10: credit note number is its own, not the original invoice's ---------------
def test_credit_note_number_not_the_referenced_invoice():
    inv = read(f'ספק בע"מ\nח.פ {SUP}\nהודעת זיכוי מס\' 777\nבגין חשבונית 4001\nתאריך: 05/10/2026\n'
               'סה"כ לפני מע"מ: 100.00\nמע"מ 18%: 18.00\nסה"כ לתשלום: 118.00')
    assert inv.invoice_number == "777"


# --- finding 12: hyphen separators and percentages are not amounts ----------------------
def test_hyphen_separator_is_not_minus():
    inv = read(f'ספק בע"מ\nח.פ {SUP}\nחשבונית מס 9\nתאריך: 05/10/2026\n'
               'סה"כ לפני מע"מ - 100.00\nמע"מ - 18.00\nסה"כ לתשלום - 118.00')
    assert (inv.total, inv.vat) == (Decimal("118.00"), Decimal("18.00"))


def test_decimal_percentage_is_not_vat():
    inv = read(f'ספק בע"מ\nח.פ {SUP}\nחשבונית מס 9\nתאריך: 05/10/2026\nמע"מ 18.00%\n'
               'סה"כ לפני מע"מ: 100.00\nסה"כ לתשלום: 118.00')
    assert inv.vat != Decimal("18")


# --- finding 4: supplier ID ---------------------------------------------------------------
def test_id_followed_by_date_on_same_line():
    inv = read(f'ספק בע"מ\nע.מ {SUP} 12/10/2026\nטל 0501234567\nחשבונית מס 9\nתאריך: 05/10/2026\n'
               'סה"כ לפני מע"מ: 100.00\nמע"מ 18%: 18.00\nסה"כ לתשלום: 118.00')
    assert inv.supplier_id == SUP


def test_two_candidate_supplier_ids_warns():
    inv = read(f'ספק בע"מ\nח.פ {SUP}\nלכבוד: לקוח אחר בע"מ ח.פ {SUP2}\nחשבונית מס 9\nתאריך: 05/10/2026\n'
               'סה"כ לפני מע"מ: 100.00\nמע"מ 18%: 18.00\nסה"כ לתשלום: 118.00')
    assert has(inv, "שני מספרי")


# --- finding 5: the invoice date, not another date ---------------------------------------
def test_prefers_invoice_date_over_order_date():
    inv = read(f'ספק בע"מ\nח.פ {SUP}\nחשבונית מס 9\nתאריך הזמנה: 01/08/2026\nתאריך חשבונית: 05/10/2026\n'
               'תאריך פירעון: 30/11/2026\nסה"כ לפני מע"מ: 100.00\nמע"מ 18%: 18.00\nסה"כ לתשלום: 118.00')
    assert inv.invoice_date == date(2026, 10, 5)
