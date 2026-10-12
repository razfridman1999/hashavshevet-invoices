"""Two commands, one per desktop shortcut:

  fetch   pull new invoice PDFs from Outlook, read them, write review.xlsx
  export  read the approved review.xlsx, learn new suppliers, write import.xlsx
  journal open the log of every invoice and its status
"""
from __future__ import annotations

import json
import os
import sys
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

from . import journal, sources, supplier_index, ui
from .extract import Invoice, extract
from .workbook import description_of, details_text, read_review, save_xls_copy, write_import, write_review

HOME = Path(os.environ.get("INVOICES_HOME", Path.home() / "Documents" / "חשבוניות לחשבשבת"))
CONFIG = HOME / "config.json"
SUPPLIERS = HOME / "suppliers.json"   # ח.פ -> {name, supplier_account, expense_account, description, senders}
STATE = HOME / "state.json"          # what was already processed, so nothing is entered twice

DEFAULT_CONFIG = {
    "source": "outlook",              # "outlook" or "folder"
    "test_folder": "",                # used when source == "folder"
    "outlook_folder": "",             # e.g. "חשבוניות" for a sub-folder of the Inbox; empty = Inbox
    "days_back": 45,
    "own_company_ids": [],            # the employer's own ח.פ, so it is never taken for the supplier
    "vat_account": "",                # כרטיס מע"מ תשומות
    "ttype": "",                      # ס.ת for supplier invoices - differs between companies
    "supplier_index": "",             # Excel export of Hashavshevet's supplier index (optional)
    "large_amount": 5000,             # invoices above this always go to review
}
DEAL_NOTE = 'חשבון עסקה: לא נקלט (נקלטת רק חשבונית מס). אם זו בכל זאת חשבונית מס, למחוק "לא" ולכתוב "כן"'


def load(path: Path, default):
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    return default


def save(path: Path, data) -> None:
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    tmp.replace(path)


def from_index(index: dict[str, list[dict]], supplier_id: str) -> tuple[dict | None, str]:
    """(supplier, note) for a supplier not learned yet, from Hashavshevet's own index."""
    cards = index.get(supplier_id, []) if supplier_id else []
    if len(cards) == 1:
        return {**cards[0], "from_index": True}, ""
    if len(cards) > 1:
        listed = ", ".join(f'{c["supplier_account"]} {c["name"]}' for c in cards)
        return None, f"לח.פ הזה יש כמה כרטיסי ספק באינדקס ({listed}): לבחור כרטיס ספק"
    return None, ""


def review_notes(inv: Invoice, supplier: dict | None, entered: set[str], in_batch: set[str],
                 cfg: dict, sender: str = "") -> list[str]:
    """Why this invoice needs a human look. An empty list means the row is ready (green)."""
    notes = list(inv.warnings)
    if inv.deal_only:
        notes.insert(0, DEAL_NOTE)
    if not supplier:
        notes.append("ספק חדש: למלא כרטיס ספק וכרטיס הוצאה (פעם אחת, מכאן זה אוטומטי)")
    elif supplier.get("from_index"):
        notes.append("ספק ראשון בכלי: כרטיס הספק מולא מהאינדקס, למלא כרטיס הוצאה ולבדוק את הפרטים")
    elif sender and sender.lower() not in supplier.get("senders", []):
        # Anyone can email a PDF that carries a known supplier's ח.פ - classic invoice
        # fraud. A known supplier writing from an address it never used is held back.
        notes.append(f"שולח לא מוכר לספק הזה ({sender}): לוודא שהחשבונית אמיתית")
    key = f"{inv.supplier_id}:{inv.invoice_number}"
    if inv.invoice_number and key in entered:
        notes.append("כפילות: החשבונית הזו כבר נקלטה בעבר")
    elif inv.invoice_number and key in in_batch:
        notes.append("כפילות: אותה חשבונית הגיעה פעמיים (כנראה הועברה שוב במייל)")
    if inv.total is not None and inv.total > cfg["large_amount"]:
        notes.append(f"סכום גבוה מ-{cfg['large_amount']:,}")
    if inv.total is not None and inv.total < 0:
        notes.append("סכום שלילי: כנראה חשבונית זיכוי")
    return notes


def _latest_review_rows() -> dict[str, dict]:
    """Rows of the newest review sheet. Errors propagate: silently starting from an
    empty sheet would throw away the accounts and approvals she already typed."""
    reviews = sorted(HOME.glob("לבדיקה *.xlsx"), key=lambda p: p.stat().st_mtime)
    if not reviews:
        return {}
    return {r["key"]: r for r in read_review(reviews[-1])}


def _balance_problem(r: dict) -> str:
    """net + VAT must equal the total, also after she edited a number by hand."""
    try:
        net, vat, total = (Decimal(str(r.get(k) or 0)).quantize(Decimal("0.01")) for k in ("net", "vat", "total"))
    except InvalidOperation:
        return "סכום לא תקין"
    return "" if net + vat == total else f'לפני מע"מ {net} + מע"מ {vat} לא שווה לסה"כ {total}'


def _log(changes: list[tuple[dict, str, str]], export_file: str = "") -> str:
    """Update the log of statuses. It only reports: a failure here must never undo or
    block what was already done, so it turns into a line in her message instead."""
    try:
        return journal.record(HOME, changes, export_file)
    except Exception as e:
        return f"\n\nיומן החשבוניות לא התעדכן ({e}). החשבוניות עצמן טופלו כרגיל."


def fetch() -> Path | None:
    cfg = {**DEFAULT_CONFIG, **load(CONFIG, {})}
    suppliers = load(SUPPLIERS, {})
    state = load(STATE, {"mail_ids": [], "invoice_keys": []})
    inbox = HOME / "קבצים"
    inbox.mkdir(parents=True, exist_ok=True)

    if not cfg["own_company_ids"]:  # without it the buyer's own ח.פ can be taken for the supplier
        ui.fail('עוד לא הוגדר הח.פ של החברה. להריץ קודם את "0 - התקנה ראשונה".')
    if cfg["source"] == "folder":
        fetched = sources.from_folder(Path(cfg["test_folder"]), inbox)
    else:
        try:
            fetched = sources.from_outlook(inbox, cfg["days_back"], cfg["outlook_folder"])
        except Exception as e:  # COM errors are cryptic; say what to check instead
            ui.fail("לא הצלחתי לקרוא מאאוטלוק.\n\n"
                    "לבדוק שאאוטלוק הרגיל פתוח (לא 'אאוטלוק החדש'), ושתיקיית החשבוניות בהגדרות קיימת.\n\n"
                    f"פרטים טכניים: {e}")
    try:  # edits she already made in a sheet she has not exported yet
        previous = _latest_review_rows()
    except Exception as e:
        ui.fail("לא הצלחתי לקרוא את טבלת הבדיקה הקודמת. אם היא פתוחה באקסל, לסגור אותה ולנסות שוב.\n\n"
                f"פרטים טכניים: {e}")
    try:
        index = supplier_index.load(cfg["supplier_index"])
    except Exception as e:
        ui.fail("לא הצלחתי לקרוא את קובץ אינדקס הספקים. אם הוא פתוח באקסל, לסגור אותו ולנסות שוב, "
                f'או לבחור אותו מחדש ב"0 - התקנה ראשונה".\n\n{cfg["supplier_index"]}\n\nפרטים טכניים: {e}')

    done = set(state["mail_ids"])
    seen_keys = set(state["invoice_keys"])
    rows, batch_keys = [], set()
    for f in fetched:
        if f.mail_id in done:
            continue
        if f.mail_id in previous:
            r = previous.pop(f.mail_id)
            rows.append(r)
            batch_keys.add(f'{r["supplier_id"]}:{r["invoice_number"]}')
            continue
        try:
            inv = extract(f.pdf, set(cfg["own_company_ids"]))
        except Exception:  # one broken attachment must not block all the others
            inv = Invoice(source=str(f.pdf), has_text=False)
            inv.warnings.append("לא הצלחתי לפתוח את הקובץ (פגום או מוגן בסיסמה): הזנה ידנית")
        if inv.invoice_date and abs((inv.invoice_date - f.received.date()).days) > 120:
            inv.warnings.append(f"תאריך החשבונית רחוק מהיום שבו הגיע המייל ({f.received:%d/%m/%Y}): לבדוק")
        sup = suppliers.get(inv.supplier_id)
        index_note = ""
        if not sup:
            sup, index_note = from_index(index, inv.supplier_id)
        key = f"{inv.supplier_id}:{inv.invoice_number}"
        notes = review_notes(inv, sup, seen_keys, batch_keys, cfg, f.sender)
        if index_note:
            notes.append(index_note)
        batch_keys.add(key)
        name = (sup or {}).get("name") or inv.supplier_name
        rows.append({
            "approve": "לא" if inv.deal_only else ("" if notes else "כן"),
            "supplier_name": name,
            "supplier_id": inv.supplier_id,
            "supplier_account": (sup or {}).get("supplier_account", ""),
            "expense_account": (sup or {}).get("expense_account", ""),
            "invoice_number": inv.invoice_number,
            "invoice_date": inv.invoice_date,
            "net": float(inv.net) if inv.net is not None else None,
            "vat": float(inv.vat) if inv.vat is not None else None,
            "total": float(inv.total) if inv.total is not None else None,
            "allocation_number": inv.allocation_number,
            "details": details_text(inv.invoice_date, "זיכוי" if inv.total is not None and inv.total < 0
                                    else (sup or {}).get("description", ""), name),
            "notes": " | ".join(notes),
            "file": str(f.pdf),
            "key": f.mail_id,
            "sender": f.sender.lower(),
        })

    # Unfinished rows whose mail was not fetched this time (aged out, moved, deleted, or
    # typed in by hand) stay on the sheet: they disappear only by export or by "לא".
    rows += [r for key, r in previous.items() if key not in done]

    if not rows:
        ui.message("אין חשבוניות חדשות.")
        return None
    # time in the name: a second run the same day must not overwrite a sheet she is filling in
    out = HOME / f"לבדיקה {datetime.now():%Y-%m-%d %H%M%S}.xlsx"
    write_review(rows, out)
    ready = sum(1 for r in rows if not r["notes"])
    logged = _log([(r, *journal.status_after_fetch(r)) for r in rows])
    ui.message(f"נמצאו {len(rows)} חשבוניות: {ready} מוכנות, {len(rows) - ready} לבדיקה.\n\nהטבלה תיפתח עכשיו."
               + logged)
    return out


def export(review_path: Path | None = None) -> Path | None:
    cfg = {**DEFAULT_CONFIG, **load(CONFIG, {})}
    if not cfg["vat_account"]:
        ui.fail('חסר מספר כרטיס מע"מ תשומות. להריץ את "0 - התקנה ראשונה".')
    if not cfg["ttype"]:
        ui.fail('חסר ס.ת (סוג תנועה) לחשבוניות ספקים. להריץ את "0 - התקנה ראשונה".')
    suppliers = load(SUPPLIERS, {})
    state = load(STATE, {"mail_ids": [], "invoice_keys": []})
    if review_path is None:
        reviews = sorted(HOME.glob("לבדיקה *.xlsx"), key=lambda p: p.stat().st_mtime)
        if not reviews:
            ui.fail('לא נמצאה טבלת בדיקה. קודם להריץ "1 - משיכת חשבוניות".')
        review_path = reviews[-1]
    try:
        review_rows = read_review(review_path)
    except PermissionError:
        ui.fail("הטבלה פתוחה באקסל. לשמור, לסגור אותה ולנסות שוב.")

    approved, rejected, held_rows = [], [], []
    done_ids = set(state["mail_ids"])
    entered, approved_keys = set(state["invoice_keys"]), set()
    for r in review_rows:
        if r["key"] in done_ids:     # already exported from this very sheet: never twice
            continue
        answer = str(r.get("approve") or "").strip()
        if answer == "לא":           # explicitly rejected: never show it again
            rejected.append(r)
            continue
        if answer not in ("כן", "v", "V", "1", "yes"):
            continue                 # left empty: stays pending for the next run
        missing = [label for k, label in (("supplier_account", "כרטיס ספק"), ("expense_account", "כרטיס הוצאה"),
                                          ("invoice_number", "מספר חשבונית"), ("invoice_date", "תאריך"),
                                          ("total", 'סה"כ')) if not r.get(k)]
        label = f'{r["supplier_name"]} {r["invoice_number"]}'
        if missing:
            held_rows.append((r, f'{label}: חסר {", ".join(missing)}'))
            continue
        if bad := _balance_problem(r):
            held_rows.append((r, f"{label}: {bad}"))
            continue
        key = f'{r["supplier_id"]}:{r["invoice_number"]}'
        if r["invoice_number"] and (key in entered or key in approved_keys):
            held_rows.append((r, f"{label}: כבר נקלטה או מופיעה פעמיים - לא נכנסה שוב"))
            continue
        approved_keys.add(key)
        approved.append(r)
        if r["supplier_id"]:  # learn: next time this supplier is filled in automatically
            before = suppliers.get(r["supplier_id"], {})
            known = before.get("senders", [])
            senders = sorted(set(known) | ({r["sender"]} if r.get("sender") else set()))
            description = (description_of(r.get("details"), r["supplier_name"])
                           if float(r["total"]) >= 0 else before.get("description", ""))  # a credit note teaches nothing
            suppliers[r["supplier_id"]] = {"name": r["supplier_name"],
                                           "supplier_account": r["supplier_account"],
                                           "expense_account": r["expense_account"],
                                           # "<מה נעשה>" from her פרטים, reused next month
                                           "description": description,
                                           "senders": senders}

    problems = [text for _, text in held_rows]
    held = ("\n\nהשורות האלה סומנו לאישור אבל לא נכנסו:\n• " + "\n• ".join(problems)) if problems else ""
    state["mail_ids"] = sorted(done_ids | {r["key"] for r in rejected})
    not_exported = ([(r, journal.REJECTED, 'סומנה "לא"') for r in rejected]
                    + [(r, journal.PENDING, text) for r, text in held_rows])
    if not approved:
        save(STATE, state)
        ui.message("אין שורות מאושרות." + held + _log(not_exported))
        return None

    out = HOME / f"לקליטה בחשבשבת {datetime.now():%Y-%m-%d %H%M%S}.xlsx"
    write_import(approved, cfg["vat_account"], out, cfg["ttype"])
    xls = save_xls_copy(out)
    # state right after the file exists: if anything later fails, a rerun must not
    # export these rows a second time
    state["mail_ids"] = sorted(set(state["mail_ids"]) | {r["key"] for r in approved})
    state["invoice_keys"] = sorted(entered | approved_keys)
    save(STATE, state)
    save(SUPPLIERS, suppliers)
    logged = _log([(r, journal.EXPORTED, "") for r in approved] + not_exported, out.name)
    ui.message(f"נוצר קובץ קליטה עם {len(approved)} חשבוניות.\n\n"
               f"בחשבשבת: כללי > ממשקים > יבוא תנועות הנה\"ח > ODBC, ולבחור את הקובץ:\n{out.name}"
               + (f"\n\nאם חשבשבת לא רואה אותו, לבחור את:\n{xls.name}" if xls else "") + held + logged)
    return out


def show_journal() -> None:
    if not (HOME / journal.JSON_NAME).exists():
        ui.message('היומן עוד ריק. הוא יתמלא אחרי "1 - משיכת חשבוניות".')
        return
    view = HOME / journal.XLSX_NAME
    try:
        journal.write_view(journal.load(HOME), view)
    except PermissionError:
        pass                     # already open in Excel: showing it is all she asked for
    if not os.environ.get("INVOICES_NO_OPEN"):
        os.startfile(view)


APP_DIR = Path(__file__).resolve().parents[1]
JOURNAL_BAT = "3 - יומן חשבוניות.bat"
APP_BAT = "חשבוניות לחשבשבת.bat"
LAUNCHERS = {name: ('@echo off\r\ncd /d "%~dp0"\r\n'
                    'rem py/pyw first: on some PCs "python" is only the Microsoft Store shortcut and pythonw does not exist\r\n'
                    f'where pyw >nul 2>nul && (start "" pyw run.py {command} & exit /b)\r\n'
                    f'start "" pythonw run.py {command}\r\n')
             for name, command in ((JOURNAL_BAT, "journal"), (APP_BAT, "app"))}


def ensure_launchers(app: Path = APP_DIR) -> None:
    """An updater only copies the files it knew about when it was installed, so a button
    added later never arrives through it. The new code puts its own buttons in place."""
    if not (app / "run.py").exists():
        return
    for name, text in LAUNCHERS.items():
        path = app / name
        try:
            if not path.exists() or b"pyw run.py" not in path.read_bytes():   # missing, or pythonw-only
                path.write_bytes(text.encode("ascii"))
        except OSError:
            pass                 # a missing shortcut must not stop her real work


def ensure_desktop_icon(cfg: dict) -> bool:
    """Once: an icon on her desktop that opens the home page. Never again after that,
    so an icon she chose to delete stays deleted. True when it was just made."""
    if cfg.get("desktop_icon_done") or os.environ.get("INVOICES_NO_OPEN") or sys.platform != "win32":
        return False
    cfg["desktop_icon_done"] = True
    save(CONFIG, cfg)
    try:
        import win32com.client
        shell = win32com.client.Dispatch("WScript.Shell")
        link = shell.CreateShortcut(str(Path(shell.SpecialFolders("Desktop")) / "חשבוניות לחשבשבת.lnk"))
        exe = Path(sys.executable)
        pythonw = exe.with_name("pythonw.exe")
        link.TargetPath = str(pythonw if pythonw.exists() else exe)
        link.Arguments = "run.py app"
        link.WorkingDirectory = str(APP_DIR)
        link.IconLocation = r"%SystemRoot%\System32\imageres.dll,112"
        link.Description = "חשבוניות לחשבשבת"
        link.Save()
        return True
    except Exception:
        return False


def run(command: str = "fetch") -> None:
    ensure_launchers()
    HOME.mkdir(parents=True, exist_ok=True)
    if not CONFIG.exists():
        save(CONFIG, DEFAULT_CONFIG)
    if command == "setup":
        from . import setup_check
        setup_check.run()
        return
    if command == "journal":
        show_journal()
        return
    if ensure_desktop_icon({**DEFAULT_CONFIG, **load(CONFIG, {})}) and command != "app":
        ui.message('חדש: על שולחן העבודה יש עכשיו אייקון "חשבוניות לחשבשבת". '
                   "הוא פותח מסך אחד עם כל הכפתורים ועם המצב של כל החשבוניות.")
    if command == "app":
        from . import webapp
        webapp.serve()
        return
    try:
        result = {"fetch": fetch, "export": export}[command]()
    except SystemExit:
        raise
    except Exception:  # run without a console: an unexpected error must still be seen
        import traceback
        log = HOME / "שגיאה אחרונה.txt"
        log.write_text(traceback.format_exc(), encoding="utf-8")
        ui.fail(f"משהו השתבש. שום דבר לא נקלט.\n\nלשלוח לרז את הקובץ:\n{log}")
    if result and not os.environ.get("INVOICES_NO_OPEN"):
        os.startfile(result if result.name.startswith("לבדיקה") else result.parent)


if __name__ == "__main__":
    run(sys.argv[1] if len(sys.argv) > 1 else "fetch")
