"""Several clients from one page, each in its own folder: settings, suppliers, log and
sheets never cross from one client to another."""
from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest
from urllib.parse import quote

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

from invoices import clients, journal, main  # noqa: E402
from test_insights_report import post  # noqa: E402
from test_webapp import app, call, home  # noqa: E402,F401  (fixtures)


def test_main_folder_is_the_first_client_and_others_are_added(home):
    assert clients.list_clients(home) == [{"id": "", "name": clients.MAIN_NAME}]
    cid, problem = clients.add(home, "מגדל לדוגמה")
    assert problem == "" and (home / "לקוחות" / cid / "config.json").exists()
    assert [c["name"] for c in clients.list_clients(home)] == [clients.MAIN_NAME, "מגדל לדוגמה"]
    assert clients.add(home, "מגדל לדוגמה")[0] is None             # no two clients by one name
    assert clients.add(home, '  /:*  ')[0] is None                  # nothing left of the name


@pytest.mark.parametrize("cid", ["..", "../..", "לקוחות", "no-such", "x/../../y", "C:\\Windows"])
def test_only_a_listed_client_has_a_folder(home, cid):
    clients.add(home, "א")
    assert clients.home_of(home, cid) is None


def test_settings_keep_only_known_fields(home):
    cid, _ = clients.add(home, "ב")
    h = clients.home_of(home, cid)
    ok, _ = clients.save_settings(h, {"vat_account": "77100", "ttype": "ח", "hid": "305",
                                      "own_company_ids": "515550002", "source": "folder", "test_folder": "C:\\"})
    cfg = main.load(h / "config.json", {})
    assert ok and cfg["hid"] == "305" and cfg["own_company_ids"] == ["515550002"]
    assert cfg["source"] == "outlook" and cfg["test_folder"] == ""   # the page cannot redirect the source
    assert clients.save_settings(h, {"hid": "jurnal305"})[0] is False
    assert clients.save_settings(h, {"own_company_ids": "123456789"})[0] is False


def test_each_client_works_in_its_own_folder(app, home):
    status, body = post(app, "/api/clients", {"name": "לקוח ב"})
    cid = json.loads(body)["id"]
    q = quote(cid)                                         # the page sends it URL-encoded
    other = home / "לקוחות" / cid
    cfg = {**main.load(home / "config.json", {}), "client_name": "לקוח ב"}
    main.save(other / "config.json", cfg)                 # same demo mail folder, own everything else
    assert json.loads(post(app, f"/api/run/fetch?c={q}")[1])["ok"]
    assert list(other.glob("לבדיקה *.xlsx")) and not list(home.glob("לבדיקה *.xlsx"))
    assert journal.load(other) and not journal.load(home)
    s = json.loads(call(app, "GET", f"/api/status?c={q}")[1])
    assert s["client"] == cid and s["items"] and s["settings"]["client_name"] == "לקוח ב"
    assert [c["name"] for c in s["clients"]] == [clients.MAIN_NAME, "לקוח ב"]     # the page's picker
    assert json.loads(call(app, "GET", "/api/status")[1])["items"] == []


def test_unknown_client_is_refused_for_actions(app, home):
    assert post(app, "/api/run/fetch?c=..")[0] == 404
    assert post(app, "/api/settings?c=../..", {"hid": "1"})[0] == 404
    s = json.loads(call(app, "GET", "/api/status?c=gone")[1])
    assert s["client"] == ""                              # the page falls back to the main client


def test_settings_from_the_page(app, home):
    status, body = post(app, "/api/settings", {"hid": "999", "ttype": "ח"})
    assert status == 200 and json.loads(body)["ok"]
    assert main.load(home / "config.json", {})["hid"] == "999"
    assert post(app, "/api/settings", {"hid": "x"}, token=None)[0] == 403
