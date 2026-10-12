"""Monthly VAT summary, the missing-invoice alert, opening a PDF from the page, and the
report to Raz - which must never carry her employer's data."""
from __future__ import annotations

import json
import re
import sys
import threading
from datetime import date
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

from invoices import insights, journal, main, report, split, webapp  # noqa: E402
from test_webapp import app, call, home  # noqa: E402,F401  (fixtures)

E, P, R = journal.EXPORTED, journal.PENDING, journal.REJECTED


def entry(status, sid="513100008", day="2026-09-05", net=100.0, vat=18.0, total=118.0, name="ספק", **kw):
    return {"status": status, "supplier_id": sid, "supplier_name": name, "invoice_date": day,
            "net": net, "vat": vat, "total": total, "updated": "2026-10-01 10:00", **kw}


# --- monthly summary --------------------------------------------------------------------
def test_summary_counts_only_what_went_into_import_files():
    entries = {"a": entry(E), "b": entry(E, day="2026-09-20 00:00:00", net=0.1, vat=0.2, total=0.3),
               "c": entry(P), "d": entry(R), "e": entry(E, day="2026-08-01", net=-200, vat=-36, total=-236)}
    sep, aug = insights.monthly_summary(entries)
    assert sep == {"month": "2026-09", "count": 2, "net": 100.1, "vat": 18.2, "total": 118.3, "partial": False}
    assert aug["count"] == 1 and aug["total"] == -236.0          # a credit note lowers the month


def test_summary_month_with_missing_amounts_says_partial():
    (m,) = insights.monthly_summary({"a": entry(E), "b": entry(E, vat=None)})
    assert m["partial"] and m["count"] == 2


# --- missing invoices ---------------------------------------------------------------------
def months(*ms, sid="513100008", status=E):
    return {f"{sid}{m}{status}": entry(status, sid=sid, day=f"{m}-10", name=f"ספק {sid[-2:]}") for m in ms}


def test_monthly_supplier_without_last_month_is_flagged():
    found = insights.missing_invoices(months("2026-06", "2026-07", "2026-08"), today=date(2026, 10, 12))
    assert [(m["supplier_id"], m["month"]) for m in found] == [("513100008", "2026-09")]


@pytest.mark.parametrize("history", [
    ("2026-06", "2026-07", "2026-08", "2026-09"),       # it did arrive
    ("2026-05", "2026-07", "2026-09"),                  # bi-monthly bill: never a false alarm
    ("2026-07", "2026-08"),                             # too little history to call it monthly
])
def test_no_alert_when_nothing_is_missing(history):
    assert insights.missing_invoices(months(*history), today=date(2026, 10, 12)) == []


def test_rejected_invoices_do_not_make_a_supplier_monthly():
    entries = {**months("2026-06", "2026-07"), **months("2026-08", status=R)}
    assert insights.missing_invoices(entries, today=date(2026, 10, 12)) == []


def test_january_looks_back_across_the_year():
    found = insights.missing_invoices(months("2026-09", "2026-10", "2026-11"), today=date(2027, 1, 3))
    assert found[0]["month"] == "2026-12"


# --- the report to Raz ----------------------------------------------------------------------
SECRETS = ["מאפיית השקד בע\"מ", "513100008", "98765", "4,321.00", "4321", "billing@shaked.example",
           "חשבונית-שקד-אוקטובר.pdf", "נתיב-סודי"]


def test_report_carries_no_invoice_data(home):
    notes = (f"שולח לא מוכר לספק הזה (billing@shaked.example): לוודא | "
             f"נמצאו שני מספרי ח.פ (513100008, 515550002): לוודא מי הספק | משהו חדש לגמרי 98765")
    row = {"key": "k1", "supplier_name": SECRETS[0], "supplier_id": "513100008", "invoice_number": "98765",
           "invoice_date": "2026-10-05", "total": 4321.0, "net": None, "vat": None, "notes": notes,
           "file": str(home / "קבצים" / "חשבונית-שקד-אוקטובר.pdf")}
    journal.record(home, [(row, P, notes)])
    (home / "שגיאה אחרונה.txt").write_text(
        'Traceback (most recent call last):\n  File "C:\\Users\\x\\app\\invoices\\main.py", line 151, in fetch\n'
        "    inv = extract(f.pdf)\nValueError: could not read נתיב-סודי 4,321.00\n", encoding="utf-8")
    text = report.build(home, "2026.10.12.7", her_note="")
    for secret in SECRETS + ["Users", "\\x\\"]:      # a full path carries her Windows user name
        assert secret not in text, secret
    assert "unknown_sender" in text and "two_ids" in text and "other" in text      # codes do travel
    assert "מספר=כן" in text and "מעמ=לא" in text
    assert "main.py:151 fetch" in text and "ValueError" in text                    # where, not what


def test_every_note_the_tool_writes_has_a_code():
    """A new kind of note would show up as 'other' and tell Raz nothing."""
    src = "".join((ROOT / "invoices" / f).read_text(encoding="utf-8") for f in ("extract.py", "main.py"))
    texts = re.findall(r'(?:warnings|notes)\.(?:append|insert)\((?:0, )?f?([\'"])(.+?)\1', src)
    texts = [re.sub(r"\{[^}]*\}", "20.00", t) for _, t in texts] + [main.DEAL_NOTE, split.UNCLEAR]
    assert len(texts) > 20
    for t in texts:
        assert journal.issue_codes({"notes": t}) != ["other"], t


def test_report_file_opens_in_notepad_with_hebrew(home):
    path = report.write(home, "v", "בדיקה")
    assert path.read_bytes().startswith(b"\xef\xbb\xbf") and "בדיקה" in path.read_text(encoding="utf-8-sig")


# --- through the page ----------------------------------------------------------------------
def post(a, path, body=None, token="ours"):
    import http.client
    c = http.client.HTTPConnection("127.0.0.1", a.port, timeout=30)
    data = json.dumps(body).encode() if body is not None else b""
    headers = {"Host": f"127.0.0.1:{a.port}", "Content-Type": "application/json"}
    if token:
        headers["X-Token"] = a.token if token == "ours" else token
    c.request("POST", path, body=data, headers=headers)
    r = c.getresponse()
    out = r.status, r.read()
    c.close()
    return out


def test_pdf_opens_only_files_the_tool_saved(app, home):
    inbox = home / "קבצים"
    inbox.mkdir(exist_ok=True)
    (inbox / "a.pdf").write_bytes(b"%PDF-1.4")
    outside = home / "elsewhere.pdf"
    outside.write_bytes(b"%PDF-1.4")
    (inbox / "run_me.bat").write_bytes(b"@echo off")
    journal.record(home, [({"key": "good", "file": str(inbox / "a.pdf")}, P, ""),
                          ({"key": "not_pdf", "file": str(inbox / "run_me.bat")}, P, ""),
                          ({"key": "outside", "file": str(outside)}, P, ""),
                          ({"key": "climb", "file": str(inbox / ".." / "elsewhere.pdf")}, P, ""),
                          ({"key": "gone", "file": str(inbox / "deleted.pdf")}, P, "")])
    assert json.loads(post(app, "/api/pdf?key=good")[1])["ok"]
    for key in ("not_pdf", "outside", "climb", "gone", "no-such-key", "..%2F..%2Fconfig.json"):
        assert not json.loads(post(app, f"/api/pdf?key={key}")[1])["ok"], key
    assert post(app, "/api/pdf?key=good", token=None)[0] == 403


def test_report_from_the_page(app, home):
    status, body = post(app, "/api/report", {"note": "הסכום לא נמצא"})
    r = json.loads(body)
    assert status == 200 and r["ok"] and "דוח לרז" in r["text"]
    (path,) = home.glob("דוח לרז *.txt")
    assert "הסכום לא נמצא" in path.read_text(encoding="utf-8-sig")
    assert post(app, "/api/report", {"note": "x"}, token=None)[0] == 403
    assert post(app, "/api/report", {"note": "x" * 70_000})[0] == 413


def test_status_carries_summary_and_alerts(app, home):
    journal.record(home, [({"key": "k", "supplier_id": "513100008", "invoice_date": "2026-09-01",
                            "net": 1.0, "vat": 0.18, "total": 1.18}, E, "")])
    s = json.loads(call(app, "GET", "/api/status")[1])
    assert s["summary"][0]["count"] == 1 and s["missing"] == [] and s["items"][0]["has_file"] is False
