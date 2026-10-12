"""First-run setup: ask the three settings, then check Outlook and Excel are reachable."""
from __future__ import annotations

from . import main, ui
from .extract import valid_israeli_id


def outlook_ok() -> str:
    try:
        import win32com.client
        inbox = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI").GetDefaultFolder(6)
        return f"✔ אאוטלוק מחובר ({inbox.Items.Count} הודעות בדואר הנכנס)"
    except Exception:
        return ("✘ אין גישה לאאוטלוק. אם מופעל 'אאוטלוק החדש', צריך לכבות את המתג שלו "
                "בפינה העליונה של החלון ולחזור לגרסה הרגילה.")


def excel_ok() -> str:
    try:
        import win32com.client
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Quit()
        return "✔ אקסל מותקן (ייווצר גם קובץ בפורמט הישן)"
    except Exception:
        return "• אקסל לא נמצא: ייווצר קובץ בפורמט החדש בלבד"


def updates_ok() -> str:
    try:
        import updater
        updater.remote_version(timeout=8)
        return "✔ עדכונים אוטומטיים פעילים"
    except Exception:
        return ("⚠ המחשב לא מצליח להגיע לשרת העדכונים (כנראה חסימה ברשת של העבודה). "
                "התוכנה תעבוד, אבל עדכונים לא יגיעו לבד. לספר לרז.")


def run() -> None:
    main.HOME.mkdir(parents=True, exist_ok=True)
    cfg = {**main.DEFAULT_CONFIG, **main.load(main.CONFIG, {})}
    answers = ui.ask_settings(cfg)
    if answers is None:
        ui.fail("ההגדרה בוטלה. אפשר להריץ אותה שוב בכל זמן.")
    bad = [n for n in answers["own_company_ids"] if not valid_israeli_id(n)]
    cfg.update(answers)
    main.save(main.CONFIG, cfg)
    lines = [outlook_ok(), excel_ok(), updates_ok(), f"✔ ההגדרות נשמרו ב:\n{main.CONFIG}"]
    if bad:
        lines.append(f"⚠ המספר {bad[0]} לא נראה כמו ח.פ תקין, כדאי לבדוק אותו.")
    ui.message("\n\n".join(lines), error=bool(bad) or "✘" in lines[0])


if __name__ == "__main__":
    run()
