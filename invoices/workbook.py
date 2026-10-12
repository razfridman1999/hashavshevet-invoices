"""The two Excel files the bookkeeper actually touches.

1. review.xlsx  - one row per invoice. Green rows are ready; yellow rows need her
                  (new supplier, missing field, VAT mismatch). She fills the gaps.
2. import.xlsx  - built from the approved rows, in the shape Hashavshevet ERP's
                  "יבוא תנועות הנה"ח > יבוא מטבלת ODBC" expects: one header row,
                  account keys and references stored as TEXT, and a named range
                  over the data (without a name the import window shows NONAME).
"""
from __future__ import annotations

from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

REVIEW_COLS = [
    ("approve", "לאשר? (כן/לא)"),
    ("supplier_name", "ספק"),
    ("supplier_id", 'ח.פ / ע.מ'),
    ("supplier_account", "כרטיס ספק"),
    ("expense_account", "כרטיס הוצאה"),
    ("invoice_number", "מס' חשבונית"),
    ("invoice_date", "תאריך"),
    ("net", 'לפני מע"מ'),
    ("vat", 'מע"מ'),
    ("total", 'סה"כ'),
    ("allocation_number", "מספר הקצאה"),
    ("notes", "הערות"),
    ("file", "קובץ"),
    ("key", "מזהה"),
]

# Column titles in import.xlsx. They only have to be recognisable: the first import
# maps each one to a Hashavshevet field once ("שידוך"), and that mapping is saved.
IMPORT_COLS = ["אסמכתא", "אסמכתא2", "תאריך אסמכתא", "תאריך ערך", "פרטים",
               "חשבון חובה1", "סכום חובה1", "חשבון חובה2", "סכום חובה2",
               "חשבון זכות1", "סכום זכות1", "חשבון זכות2", "סכום זכות2"]

GREEN = PatternFill("solid", fgColor="E2F0D9")
YELLOW = PatternFill("solid", fgColor="FFF2CC")
HEAD = PatternFill("solid", fgColor="1F4E78")


def write_review(rows: list[dict], path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "לבדיקה"
    ws.sheet_view.rightToLeft = True
    ws.append([title for _, title in REVIEW_COLS])
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = HEAD
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    for r in rows:
        ws.append([r.get(k, "") for k, _ in REVIEW_COLS])
        fill = YELLOW if r.get("notes") else GREEN
        for cell in ws[ws.max_row]:
            cell.fill = fill
    widths = [12, 26, 13, 12, 12, 14, 12, 12, 10, 12, 14, 45, 30, 8]
    for col, w in zip("ABCDEFGHIJKLMN", widths):
        ws.column_dimensions[col].width = w
    for row in ws.iter_rows(min_row=2):
        row[6].number_format = "DD/MM/YYYY"
        for c in row[7:10]:
            c.number_format = "#,##0.00"
    ws.column_dimensions["N"].hidden = True
    ws.freeze_panes = "B2"
    wb.save(path)


def read_review(path: Path) -> list[dict]:
    ws = load_workbook(path, data_only=True).active
    keys = [k for k, _ in REVIEW_COLS]
    rows = []
    for values in ws.iter_rows(min_row=2, values_only=True):
        if not any(values):
            continue
        row = dict(zip(keys, values))
        for k in ("supplier_id", "supplier_account", "expense_account", "invoice_number",
                  "allocation_number", "supplier_name", "key"):
            v = row.get(k)
            # Excel turns typed account numbers into floats (40012 -> 40012.0)
            row[k] = "" if v is None else (str(int(v)) if isinstance(v, float) and v.is_integer() else str(v).strip())
        if not row["key"]:  # a row she typed in herself: give it a stable identity
            row["key"] = f'manual:{row["supplier_id"]}:{row["invoice_number"]}:{row.get("total")}'
        rows.append(row)
    return rows


def _d(v) -> Decimal:
    return Decimal(str(v or 0)).quantize(Decimal("0.01"))


def write_import(rows: list[dict], vat_account: str, path: Path) -> None:
    """One journal line per invoice: expense + input VAT debited, supplier credited.

    A credit note (negative total) is the mirror image: supplier debited, expense and
    VAT credited, all amounts positive - safer than handing the import negative sums.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "תנועות"
    ws.append(IMPORT_COLS)
    for r in rows:
        net, vat, total = _d(r["net"]), _d(r["vat"]), _d(r["total"])
        ref = r["invoice_number"]
        ref_digits = "".join(ch for ch in ref if ch.isdigit())[-9:]  # אסמכתא: numeric, max 9
        d = r["invoice_date"]
        d = d.date() if hasattr(d, "date") else d
        head = [ref_digits, (r.get("allocation_number") or "")[-9:], d, d]
        vat_acc = vat_account if vat else ""
        # Unused amount cells hold 0, never blank: ODBC guesses each column's type from its
        # first rows, and a column that starts out blank reads a later value as NULL.
        name = r.get("supplier_name") or ""
        if total >= 0:
            ws.append(head + [f'{name} חש\' {ref}'[:50],
                              r["expense_account"], float(net), vat_acc, float(vat),
                              r["supplier_account"], float(total), "", 0.0])
        else:
            ws.append(head + [f'{name} זיכוי {ref}'[:50],
                              r["supplier_account"], float(-total), "", 0.0,
                              r["expense_account"], float(-net), vat_acc, float(-vat)])
    last = ws.max_row
    for row in ws.iter_rows(min_row=2):
        for idx in (0, 1, 5, 7, 9, 11):      # keys & references as text, per the H-ERP guide
            row[idx].number_format = "@"
            if row[idx].value is not None:
                row[idx].value = str(row[idx].value)
        row[2].number_format = row[3].number_format = "DD/MM/YYYY"
    ref = f"'{ws.title}'!$A$1:$M${last}"
    wb.defined_names["invoices"] = DefinedName("invoices", attr_text=ref)
    wb.save(path)


def save_xls_copy(xlsx: Path) -> Path | None:
    """Also save an Excel 97-2003 .xls next to the .xlsx, through the installed Excel.

    A 32-bit Hashavshevet often reads only .xls through ODBC (the H-ERP guide's own
    example is an .xls). Returns None when Excel is not available.
    """
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        return None
    xls = xlsx.with_suffix(".xls")
    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")  # private instance, never her open Excel
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(str(xlsx.resolve()))
        wb.SaveAs(str(xls.resolve()), FileFormat=56)  # 56 = xlExcel8
        wb.Close(False)
        return xls
    except Exception:
        return None
    finally:
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()
