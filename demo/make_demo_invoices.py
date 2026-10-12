"""Fictional invoices for demonstrating and testing the tool.

Every supplier, number and amount here is invented, and every page is stamped
"דוגמה". The set covers the layouts a bookkeeper actually meets: line-item tables,
VAT-exempt dealers, credit notes, a scan with no text layer, the buyer's own company
number printed before the supplier's, and the same invoice arriving twice.
"""
from __future__ import annotations

import sys
from pathlib import Path

import fitz

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from invoices.extract import valid_israeli_id  # noqa: E402

OUT = Path(__file__).parent / "invoices"


def ids(start: int, n: int) -> list[str]:
    return [str(x) for x in range(start, start + 2000) if valid_israeli_id(str(x))][:n]


OWN_ID = ids(515550000, 1)[0]          # the employer ("חברת הלקוח")
S = ids(513100000, 7)                  # seven fictional suppliers
STAMP = '<p style="color:#c00;font-size:10px">דוגמה בלבד - מסמך פיקטיבי לבדיקת מערכת</p>'
BUYER = f'<p>לכבוד: גלבוע ייעוץ וניהול בע"מ, ח.פ {OWN_ID}</p>'


def page(html: str, name: str) -> None:
    doc = fitz.open()
    pg = doc.new_page()
    pg.insert_htmlbox(fitz.Rect(36, 36, 560, 806),
                      f'<div dir="rtl" style="font-family:Arial;font-size:12px">{STAMP}{html}</div>')
    doc.save(OUT / name)


def simple(supplier, sid, title, d, net, vat, total, extra="", buyer_first=False):
    head = f"<h2>{supplier}</h2><p>עוסק מורשה {sid}</p>"
    return (BUYER + head if buyer_first else head + BUYER) + f"""
<h3>{title}</h3><p>תאריך: {d}</p>
<p>סה"כ לפני מע"מ: {net}</p><p>מע"מ 18%: {vat}</p><p>סה"כ לתשלום: {total}</p>{extra}"""


def table(supplier, sid, num, d, lines):
    net = sum(q * p for _, q, p in lines)
    vat = round(net * 0.18, 2)
    rows = "".join(f"<tr><td>{n}</td><td>{q}</td><td>{p:,.2f}</td><td>{q * p:,.2f}</td></tr>"
                   for n, q, p in lines)
    return f"""<h2>{supplier}</h2><p>ח.פ: {sid} | טל' 03-5550000</p>{BUYER}
<h3>חשבונית מס מקור מספר {num}</h3><p>תאריך הפקה: {d}</p>
<table border="1" cellpadding="4" style="border-collapse:collapse;width:100%">
<tr><th>פריט</th><th>כמות</th><th>מחיר</th><th>סה"כ</th></tr>{rows}</table>
<p>סה"כ לפני מע"מ: {net:,.2f}</p><p>מע"מ 18%: {vat:,.2f}</p>
<p><b>סה"כ כולל מע"מ: {net + vat:,.2f}</b></p><p>מספר הקצאה: 3{num.zfill(8)}</p>"""


def scanned(name: str) -> None:
    """A 'scan': the invoice rendered to an image, so the PDF has no text at all."""
    src = fitz.open()
    pg = src.new_page()
    pg.insert_htmlbox(fitz.Rect(36, 36, 560, 806), f'<div dir="rtl" style="font-family:Arial">{STAMP}'
                      + simple("מוסך הגליל", S[6], "חשבונית מס 4410", "02/10/2026", "600.00", "108.00", "708.00")
                      + "</div>")
    pix = pg.get_pixmap(dpi=110)
    doc = fitz.open()
    doc.new_page().insert_image(fitz.Rect(0, 0, 595, 842), pixmap=pix)
    doc.save(OUT / name)


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    for old in OUT.glob("*.pdf"):
        old.unlink()
    page(simple('חשמל השרון בע"מ', S[0], "חשבונית מס 20261045", "05/10/2026", "420.00", "75.60", "495.60",
                "<p>מספר הקצאה: 300112233</p>"), "01_electric.pdf")
    page(table('אופיס סנטר ישיר בע"מ', S[1], "58213", "07.10.26",
               [("נייר A4 קרטון", 4, 115.0), ("טונר למדפסת", 2, 289.9), ("קלסרים", 10, 12.5)]), "02_office_table.pdf")
    page(f"""<h2>יעל ברק - עיצוב גרפי</h2><p>עוסק פטור {S[2]}</p>{BUYER}
<h3>קבלה מס' 112</h3><p>תאריך: 30/09/2026</p><p>עבור: עיצוב לוגו</p>
<p>סה"כ לתשלום: 1,500.00</p><p>עוסק פטור - לא נגבה מע"מ</p>""", "03_exempt_dealer.pdf")
    page(simple('ניקיון הכרמל בע"מ', S[3], "חשבונית זיכוי 7731", "08/10/2026", "-200.00", "-36.00", "-236.00",
                "<p>זיכוי עבור שירות שלא בוצע</p>"), "04_credit_note.pdf")
    page(simple('תוכנות מרום בע"מ', S[4], "חשבונית מס 9902", "01/10/2026", "6,800.00", "1,224.00", "8,024.00",
                buyer_first=True), "05_software_large_buyer_first.pdf")
    page(simple('דפוס הנגב', S[5], "חשבונית מס 315", "03/10/2026", "100.00", "20.00", "120.00"),
         "06_wrong_vat.pdf")
    scanned("07_scanned.pdf")
    # the electricity bill again, forwarded a second time
    page(simple('חשמל השרון בע"מ', S[0], "חשבונית מס 20261045", "05/10/2026", "420.00", "75.60", "495.60",
                "<p>מספר הקצאה: 300112233</p>"), "08_electric_forwarded_again.pdf")
    print("own company id:", OWN_ID)
    print("suppliers:", S)


if __name__ == "__main__":
    main()
