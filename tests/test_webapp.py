"""The home page server: does the same work as the old buttons, and answers nobody but
her own page (any website can send requests to 127.0.0.1)."""
from __future__ import annotations

import http.client
import json
import sys
import threading
from pathlib import Path

import openpyxl
import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "demo"))

import make_demo_invoices as demo  # noqa: E402
from invoices import journal, main, ui, webapp  # noqa: E402
from invoices.workbook import COL  # noqa: E402


@pytest.fixture
def home(tmp_path, monkeypatch):
    monkeypatch.setenv("INVOICES_NO_OPEN", "1")
    demo.OUT = tmp_path / "mail"
    demo.main()
    h = tmp_path / "home"
    h.mkdir()
    for name, value in (("HOME", h), ("CONFIG", h / "config.json"),
                        ("SUPPLIERS", h / "suppliers.json"), ("STATE", h / "state.json")):
        monkeypatch.setattr(main, name, value)
    main.save(main.CONFIG, {**main.DEFAULT_CONFIG, "source": "folder", "test_folder": str(demo.OUT),
                            "own_company_ids": [demo.OWN_ID], "vat_account": "77100", "ttype": "ח"})
    return h


@pytest.fixture
def app(home):
    a = webapp.App()
    t = threading.Thread(target=a.server.serve_forever, daemon=True)
    t.start()
    yield a
    a.server.shutdown()
    a.server.server_close()


def call(a, method, path, token="ours", host=None, header=True):
    token = a.token if token == "ours" else token
    c = http.client.HTTPConnection("127.0.0.1", a.port, timeout=60)
    headers = {"Host": host or f"127.0.0.1:{a.port}"}
    if header and token is not None:
        headers["X-Token"] = token
    c.request(method, path, headers=headers)
    r = c.getresponse()
    body = r.read()
    c.close()
    return r.status, body, r


def run(a, action):
    status, body, _ = call(a, "POST", f"/api/run/{action}")
    assert status == 200
    return json.loads(body)


def status_of(a):
    return json.loads(call(a, "GET", "/api/status")[1])


# --- who may talk to it -------------------------------------------------------------------
def test_page_needs_the_token(app):
    assert call(app, "GET", "/", token=None)[0] == 403
    assert call(app, "GET", "/?t=wrong", token=None)[0] == 403
    status, body, r = call(app, "GET", f"/?t={app.token}", token=None)
    assert status == 200 and "חשבוניות לחשבשבת".encode() in body
    assert "default-src 'none'" in r.getheader("Content-Security-Policy")


def test_actions_need_the_token_header(app):
    assert call(app, "POST", "/api/run/fetch", token=None)[0] == 403
    assert call(app, "POST", "/api/run/fetch", token="guess")[0] == 403
    # a token in the address is not enough for an action
    assert call(app, "POST", f"/api/run/fetch?t={app.token}", header=False)[0] == 403
    assert not list(main.HOME.glob("לבדיקה *.xlsx"))


def test_foreign_host_is_refused_even_with_the_token(app):
    """DNS rebinding: an outside name pointed at 127.0.0.1 arrives with its own Host."""
    assert call(app, "GET", "/api/status", host=f"evil.example:{app.port}")[0] == 403
    assert call(app, "POST", "/api/run/fetch", host="evil.example")[0] == 403


# --- doing the work ---------------------------------------------------------------------
def test_fetch_from_the_page_reports_back_and_updates_the_counts(app):
    assert status_of(app)["counts"] == {"pending": 0, "ready": 0, "exported_month": 0}
    r = run(app, "fetch")
    assert r["ok"] and "נמצאו 8 חשבוניות" in r["text"]
    s = status_of(app)
    assert s["counts"]["pending"] == 8 and s["review"].startswith("לבדיקה")
    assert len(s["items"]) == 8 and s["configured"]


def test_export_from_the_page(app):
    run(app, "fetch")
    review = sorted(main.HOME.glob("לבדיקה *.xlsx"))[-1]
    wb = openpyxl.load_workbook(review)
    for row in wb.active.iter_rows(min_row=2):
        if Path(row[COL["file"]].value).name == "01_electric.pdf":
            row[COL["approve"]].value, row[COL["supplier_account"]].value, row[COL["expense_account"]].value \
                = "כן", "40010", "81100"
    wb.save(review)
    r = run(app, "export")
    assert r["ok"] and "נוצר קובץ קליטה עם 1 חשבוניות" in r["text"]
    assert status_of(app)["counts"]["exported_month"] == 1


def test_a_refusal_is_shown_and_the_server_lives_on(app):
    cfg = main.load(main.CONFIG, {})
    main.save(main.CONFIG, {**cfg, "vat_account": ""})
    r = run(app, "export")                      # ui.fail -> sys.exit inside the action
    assert not r["ok"] and 'כרטיס מע"מ' in r["text"]
    assert status_of(app)["configured"] is False
    assert ui.captured is None                  # later messages go to boxes again


def test_unexpected_error_is_logged_and_reported(app, monkeypatch):
    def boom():
        raise RuntimeError("disk on fire")
    monkeypatch.setattr(main, "fetch", boom)
    r = run(app, "fetch")
    assert not r["ok"] and "שגיאה אחרונה.txt" in r["text"]
    assert "disk on fire" in (main.HOME / "שגיאה אחרונה.txt").read_text(encoding="utf-8")
    assert ui.captured is None


def test_two_actions_never_run_at_once(app):
    with app.busy:
        status, body, _ = call(app, "POST", "/api/run/fetch")
    assert status == 409 and "עוד רצה" in json.loads(body)["text"]
    assert not list(main.HOME.glob("לבדיקה *.xlsx"))


def test_unknown_action_does_nothing(app):
    r = run(app, "format_disk")
    assert not r["ok"]


def test_status_items_carry_pdf_text_as_plain_strings(app):
    journal.record(main.HOME, [({"key": "k", "supplier_name": "<img src=x onerror=alert(1)>"},
                                journal.PENDING, "")])
    item = status_of(app)["items"][0]
    assert item["supplier_name"] == "<img src=x onerror=alert(1)>"   # data; the page sets textContent


def test_page_never_builds_html_from_data():
    page = (ROOT / "invoices" / "app.html").read_text(encoding="utf-8")
    assert "innerHTML" not in page and "insertAdjacentHTML" not in page and "document.write" not in page


CHILD = """
import sys, threading
sys.path.insert(0, sys.argv[1])
from invoices import main, ui, webapp
def fetch():
    import win32com.client
    win32com.client.Dispatch("Scripting.Dictionary")   # any real COM object
    ui.message("com ok")
main.fetch = fetch
results = []
for _ in range(2):    # the first thread to import pythoncom is set up by the import itself
    t = threading.Thread(target=lambda: results.append(webapp.run_action("fetch")))
    t.start(); t.join()
sys.stdout = sys.__stdout__
print("RESULT", [r["text"][:30] for r in results])
"""


@pytest.mark.skipif(sys.platform != "win32", reason="COM is Windows only")
def test_actions_can_use_com_from_a_server_thread(tmp_path):
    """Each request runs in a new thread; Outlook and Excel are COM, which fails in a
    thread that did not call CoInitialize. A fake Outlook in the other tests hid this.
    A fresh process: anything else in this test run may already have set COM up."""
    import os
    import subprocess
    env = {**os.environ, "INVOICES_HOME": str(tmp_path), "INVOICES_NO_OPEN": "1", "PYTHONIOENCODING": "utf-8"}
    out = subprocess.run([sys.executable, "-c", CHILD, str(ROOT)], capture_output=True, text=True,
                         encoding="utf-8", env=env, timeout=60)
    assert "RESULT ['com ok', 'com ok']" in out.stdout, out.stdout + out.stderr
