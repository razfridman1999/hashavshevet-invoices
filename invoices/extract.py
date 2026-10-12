"""Read an Israeli supplier invoice PDF and pull out the fields a journal entry needs.

Everything here runs locally. A PDF with a real text layer (most invoices sent by
email today) is parsed with rules; a scanned PDF has no text, comes back empty,
and is sent to manual review.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal, InvalidOperation
from pathlib import Path

import fitz  # PyMuPDF

# Israeli VAT history, newest first: (effective from, rate)
VAT_RATES = [(date(2025, 1, 1), Decimal("0.18")), (date(2015, 10, 1), Decimal("0.17"))]

ID_LABELS = r'(?:ח\.?\s?פ\.?|ח"פ|ע\.?\s?מ\.?|ע"מ|עוסק\s+(?:מורשה|פטור)|מס[\'׳]?\s*עוסק|מספר\s+עוסק|ח\.צ|ע\.ר|VAT\s*(?:No|ID)?|Reg\.?\s*No)'
DATE_RE = re.compile(r"\b(\d{1,2})[./-](\d{1,2})[./-](\d{4}|\d{2})\b")
# 1,234.56 | 1234.56 | 1234 - never a percentage ("18%", "18.00%") or part of a longer
# number. A minus counts when it touches the digits, sits alone at the end of the line
# before them (RTL layouts split it off), or trails them ("200.00-"). A hyphen used as a
# separator ("סה"כ - 118.00") is not a minus.
AMOUNT_RE = (r"(?<![\d.,%])(-(?:[ \t]*\n\s*)?)?(\d{1,3}(?:,\d{3})+(?:\.\d{1,2})?|\d+\.\d{1,2}|\d+)"
             r"(?!\d|[.,]\d|\s*%)(-(?![ \t]*\d))?")
TOTAL_STRONG = (r'סה"?כ\s*(?:לתשלום|כולל\s*מע"?מ)|סך\s*הכל\s*לתשלום|Grand\s+Total|Total\s+(?:Due|Amount|incl)'
                r'|Amount\s+Due|(?<![Ss]ub )(?<![Ss]ub)(?<![Ss]ub-)\bTotal\b')
VAT_LABEL = r'מע"?מ|מס\s+ערך\s+מוסף|\bVAT\b(?!\s*(?:No|ID|Reg|#))'
CREDIT_TITLE = r"(?:הודעת[ \t-]+זיכוי|תעודת[ \t]+זיכוי|חשבונית(?:[ \t]+מס)?[ \t]*/?[ \t]*זיכוי|Credit[ \t]+Note)"
# date labels: which date on the page is the invoice's own
DATE_PREFER = re.compile(r"חשבונית|הפקה|ההפקה|המסמך|מסמך|הוצאה|Invoice", re.I)
DATE_EXCLUDE = re.compile(r"פירעון|תשלום|אספקה|הזמנה|הדפסה|ערך|לידה|תפוגה|Due|Order|Print|Delivery", re.I)
MONEY_TOKEN = re.compile(r"(?<![\d.,])(?:[1-9]\d{0,2}(?:,\d{3})+|[1-9]\d*|0)\.\d{2}(?![\d.])")
# PDFs are full of no-break spaces and invisible direction marks between words
SPACES = str.maketrans({"\xa0": " ", " ": " ", " ": " ", "‏": None, "‎": None,
                        "‪": None, "‫": None, "‬": None, "‭": None, "‮": None,
                        "⁦": None, "⁧": None, "⁨": None, "⁩": None, "״": '"'})
# words that turn "מע"מ" into a different line: "לפני מע"מ", "כולל מע"מ", "ללא מע"מ"
NOT_VAT_LINE = re.compile(r'(?:לפני|כולל|ללא|פטור\s+מ|חייב\s*ב?|נגבה|[Ss]ub\s*[Tt]otal\s*(?:excl\.?|before)?)\s*$')
# a document title, then optional qualifiers, then the number - all on one line,
# so a title never borrows a number (a date, an amount) from the line below it
DOC_TITLE = (r"(?:" + CREDIT_TITLE + r"|חשבונית(?:[ \t]+מס)?(?:[ \t]*/[ \t]*קבלה)?(?:[ \t]+(?:מקור|העתק|זיכוי|עסקה|נאמן[ \t]+למקור))?"
             r"|חשבון[ \t]+עסקה|קבלה|Tax[ \t]+Invoice|Invoice)")
NUMBER_RE = re.compile(DOC_TITLE + r"[ \t]*(?:מס[\'׳]?|מספר|No\.?|#)?[ \t]*[:#]?[ \t]*([A-Za-z]{0,3}-?\d{2,12})(?![\d/.])")
NUMBER_BEFORE_RE = re.compile(r"(?<![\d/.])(\d{2,12})[ \t:]*" + DOC_TITLE)   # RTL visual order


@dataclass
class Invoice:
    source: str                      # file path of the PDF
    supplier_name: str = ""
    supplier_id: str = ""            # ח.פ / ע.מ, 9 digits, checksum-valid
    invoice_number: str = ""
    invoice_date: date | None = None
    total: Decimal | None = None     # incl. VAT
    vat: Decimal | None = None
    allocation_number: str = ""      # מספר הקצאה (חשבונית ישראל)
    has_text: bool = True
    warnings: list[str] = field(default_factory=list)

    @property
    def net(self) -> Decimal | None:
        if self.total is None:
            return None
        return self.total - (self.vat or Decimal("0"))


def valid_israeli_id(num: str) -> bool:
    """Checksum used by ת.ז, ח.פ and ע.מ alike."""
    if not num.isdigit() or len(num) > 9:
        return False
    num = num.zfill(9)
    total = 0
    for i, ch in enumerate(num):
        d = int(ch) * (1 if i % 2 == 0 else 2)
        total += d - 9 if d > 9 else d
    return total % 10 == 0


def vat_rate_on(day: date | None) -> Decimal:
    day = day or date.today()
    for since, rate in VAT_RATES:
        if day >= since:
            return rate
    return VAT_RATES[-1][1]


def pdf_text(path: Path) -> str:
    with fitz.open(path) as doc:
        return "\n".join(page.get_text("text") for page in doc).translate(SPACES)


def _num(s: str) -> Decimal | None:
    try:
        return Decimal(s.replace(",", ""))
    except InvalidOperation:
        return None


def _parse_date(m: re.Match) -> date | None:
    d, mth, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if y < 100:
        y += 2000
    try:
        return date(y, mth, d)
    except ValueError:
        return None


def _amount_after(text: str, label: str, skip_before: re.Pattern | None = None) -> Decimal | None:
    """Amount on the same line as a label, or on the next line (PDF layouts split them)."""
    for m in re.finditer(label, text):
        if skip_before and skip_before.search(text[max(0, m.start() - 12): m.start()]):
            continue
        line_start = text.rfind("\n", 0, m.start()) + 1
        line_end = text.find("\n", m.end())
        line_end = len(text) if line_end == -1 else line_end
        # 1. same line, after the label  2. same line, before it (RTL visual order)
        # 3. the next few lines (layouts that split "label" / ":" / "amount")
        a = re.search(AMOUNT_RE, text[m.end(): line_end])
        if a:
            return _signed(a)
        b = list(re.finditer(AMOUNT_RE, text[line_start: m.start()]))
        if b:
            return _signed(b[-1])
        c = re.search(AMOUNT_RE, text[line_end: line_end + 40])
        if c:
            return _signed(c)
    return None


def _signed(m: re.Match) -> Decimal | None:
    n = _num(m.group(2))
    return -n if n is not None and (m.group(1) or m.group(3)) else n


def _invoice_date(text: str) -> tuple[date | None, bool]:
    """(date, sure). Prefer a date labelled as the invoice's own; skip due/order/print
    dates. Not sure when it had to fall back to an unlabelled date among several."""
    preferred, neutral, excluded_spans = [], [], []
    for line_m in re.finditer(r"[^\n]+", text):
        line = line_m.group(0)
        label = re.search(r"תאריך|Date", line, re.I)
        if not label:
            continue
        for dm in DATE_RE.finditer(line):
            d = _parse_date(dm)
            if not d:
                continue
            # the words between the label and the date (either order, RTL or LTR)
            between = line[min(label.end(), dm.end()): max(label.start(), dm.start())]
            between += line[label.end(): label.end() + 15]
            if DATE_EXCLUDE.search(between):
                excluded_spans.append((line_m.start() + dm.start()))
            elif DATE_PREFER.search(between):
                preferred.append(d)
            else:
                neutral.append(d)
    if preferred:
        return preferred[0], True
    if neutral:
        return neutral[0], True
    loose = [d for m in DATE_RE.finditer(text) if m.start() not in excluded_spans and (d := _parse_date(m))]
    return (loose[0] if loose else None), len(set(loose)) <= 1


def _supplier_name(lines: list[str], supplier_id: str) -> str:
    """The issuer's name sits on, or just above, the line with its own ח.פ / ע.מ."""
    def is_name(s: str) -> bool:
        return (len(s) >= 2 and not re.fullmatch(r"[\d\W]+", s)
                and not re.match(r"\s*(?:לכבוד|תאריך|טל|דוגמה)", s))

    for i, line in enumerate(lines):
        if supplier_id and supplier_id in line:
            rest = re.sub(ID_LABELS + r"|[\d:.\-|]", " ", line).strip()
            if is_name(rest):
                return rest[:60]
            for prev in reversed(lines[max(0, i - 3): i]):
                if is_name(prev.strip()):
                    return prev.strip()[:60]
    return next((l.strip()[:60] for l in lines if is_name(l.strip())), "")


def unreverse_numbers(text: str) -> str:
    """Undo 'visual order' PDFs, where every run of digits comes out backwards
    (513100073 -> 370001315, 06/10/2026 -> 6202/01/60) while the Hebrew is fine."""
    return re.sub(r"\d[\d.,/:%-]*\d", lambda m: m.group(0)[::-1], text)


def orientation_score(text: str) -> int:
    """How much of the text looks like correctly ordered numbers.

    Deliberately blind to the invoice's content (VAT, totals): if it rewarded a
    reading whose VAT "adds up", a genuinely wrong invoice could be silently
    "fixed" by picking a garbled reading. Only shape-level evidence counts.
    """
    ids = sum(valid_israeli_id(n) for n in re.findall(r"(?<!\d)\d{9}(?!\d)", text))
    dates = 0
    for m in DATE_RE.finditer(text):
        d = _parse_date(m)
        dates += bool(d and abs((d - date.today()).days) < 3 * 366)
    return 3 * ids + 2 * dates + len(MONEY_TOKEN.findall(text))


def extract(path: Path, own_ids: set[str] = frozenset()) -> Invoice:
    text = pdf_text(path)
    if len(text.strip()) < 30:
        inv = Invoice(source=str(path), has_text=False)
        inv.warnings.append("PDF סרוק / בלי טקסט, צריך הזנה ידנית")
        return inv
    flipped = unreverse_numbers(text)
    if orientation_score(flipped) > orientation_score(text):   # ties keep the text as printed
        text = flipped
    return _extract_text(text, str(path), own_ids)


def _extract_text(text: str, source: str, own_ids: set[str]) -> Invoice:
    inv = Invoice(source=source)

    # Supplier ID: a checksum-valid 9-digit number near an ID label that is not the
    # buyer's own company number.
    candidates = []
    id_num = r"(\d{9}|\d{2,3}-\d{6,7}|\d{3} \d{3} \d{3})"
    labelled = sorted(list(re.finditer(ID_LABELS + r"[ \t:#.]*" + id_num + r"(?!\d)", text))
                      + list(re.finditer(r"(?<!\d)" + id_num + r"[ \t:]*" + ID_LABELS, text)),  # RTL order
                      key=lambda m: m.start())
    for m in labelled:
        n = re.sub(r"\D", "", m.group(1))
        if len(n) == 9 and valid_israeli_id(n) and n not in own_ids and n not in candidates:
            candidates.append(n)
    if not candidates:  # fall back to any valid 9-digit number in the text
        candidates = [n for n in re.findall(r"(?<!\d)(\d{9})(?!\d)", text)
                      if valid_israeli_id(n) and n not in own_ids]
        if candidates:
            inv.warnings.append("הח.פ של הספק נמצא בלי כותרת, לוודא שזה הספק הנכון")
    if candidates:
        inv.supplier_id = candidates[0]
        if len(set(candidates)) > 1 and not inv.warnings:
            inv.warnings.append(f"נמצאו שני מספרי ח.פ ({candidates[0]}, {candidates[1]}): לוודא מי הספק")
    else:
        inv.warnings.append("לא נמצא ח.פ / ע.מ של הספק")

    alloc = re.search(r"(?:מספר|מס[\'׳]?)\s*הקצאה[\s:]*(\d{9,})", text)
    if alloc:
        inv.allocation_number = alloc.group(1)

    num = NUMBER_RE.search(text) or NUMBER_BEFORE_RE.search(text)
    if num:
        inv.invoice_number = num.group(1)
    else:
        inv.warnings.append("לא נמצא מספר חשבונית")

    inv.invoice_date, sure = _invoice_date(text)
    if not inv.invoice_date:
        inv.warnings.append("לא נמצא תאריך")
    elif not sure:
        inv.warnings.append("יש כמה תאריכים במסמך: לוודא שזה תאריך החשבונית")

    inv.total = _amount_after(text, TOTAL_STRONG)
    if inv.total is None:
        inv.total = _amount_after(text, r'סה"?כ|סך\s*הכל')
    if inv.total is None:
        inv.warnings.append("לא נמצא סכום כולל")

    vat = _amount_after(text, VAT_LABEL, skip_before=NOT_VAT_LINE)
    if re.search(r"עוסק\s+פטור", text):
        # an exempt dealer charges no VAT; a VAT amount on the same page is a contradiction
        if vat:
            inv.warnings.append('כתוב "עוסק פטור" אבל מופיע סכום מע"מ: לבדוק')
        inv.vat = Decimal("0")
    else:
        inv.vat = vat
        if vat is None and inv.total is not None:
            inv.warnings.append('לא נמצא מע"מ במסמך: לבדוק את הסכומים')
        elif vat == 0:
            inv.warnings.append('מע"מ 0 בחשבונית שלא כתוב בה "עוסק פטור": לבדוק')

    # A credit note is a negative invoice even when the page prints the amounts unsigned.
    # It counts only as a title (start of a line, or followed by its number); a passing
    # mention of a credit note elsewhere on the page only holds the invoice for review.
    if re.search(r"(?m)^[ \t]*" + CREDIT_TITLE + r"|" + CREDIT_TITLE + r"[ \t]*(?:מס[\'׳]?|מספר|No\.?)?[ \t]*\d", text, re.I):
        inv.total = -abs(inv.total) if inv.total is not None else None
        inv.vat = -abs(inv.vat) if inv.vat is not None else None
    elif "זיכוי" in text and inv.total is not None and inv.total > 0:
        inv.warnings.append('המילה "זיכוי" מופיעה במסמך: לוודא שזו לא חשבונית זיכוי')

    inv.supplier_name = _supplier_name(text.splitlines(), inv.supplier_id)
    check_amounts(inv)
    return inv


def check_amounts(inv: Invoice) -> None:
    """VAT must match the rate in force on the invoice date (±1 agora per 100 NIS)."""
    if inv.total is None or inv.vat is None or inv.vat == 0:
        return
    expected = (inv.net * vat_rate_on(inv.invoice_date)).quantize(Decimal("0.01"))
    tolerance = max(Decimal("0.05"), abs(inv.total) * Decimal("0.0005"))
    if abs(expected - inv.vat) > tolerance:
        inv.warnings.append(f'מע"מ {inv.vat} לא תואם {vat_rate_on(inv.invoice_date):.0%} מהסכום ({expected})')
