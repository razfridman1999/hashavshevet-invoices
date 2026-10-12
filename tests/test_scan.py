"""Scanned invoices read by Claude Code - only when she turned it on for the client.
Claude itself is never called here: a fake stands in and counts the calls."""
from __future__ import annotations

import sys
from decimal import Decimal as D
from pathlib import Path

import openpyxl
import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

from invoices import clients, journal, main, scan, sources  # noqa: E402
from test_workflow import S, home  # noqa: E402,F401  (fixture)
from make_demo_invoices import OWN_ID  # noqa: E402

READ = {"doc_type": "tax_invoice", "supplier_name": "מוסך הגליל", "supplier_id": S[6], "invoice_number": "4410",
        "date": "02/10/2026", "net": 600.0, "vat": 108.0, "total": 708.0, "allocation_number": "", "section": "81900"}


def test_what_claude_read_is_checked_like_any_invoice():
    inv = scan.to_invoice(READ, "s.pdf", {OWN_ID})
    assert (inv.supplier_id, inv.invoice_number, inv.total, inv.vat, inv.section) == (S[6], "4410", D("708.00"), D("108.00"), "81900")
    assert inv.warnings == [scan.NOTE]                          # always: compare it with the invoice


@pytest.mark.parametrize("change, code", [
    ({"supplier_id": OWN_ID}, "no_id"),                          # the buyer is never the supplier
    ({"supplier_id": "123456789"}, "no_id"),                     # checksum
    ({"vat": 50.0}, "vat_mismatch"),                             # 600 + 50 is not 708
    ({"net": 650.0}, "vat_mismatch"),                            # VAT fits the total, the net was misread
    ({"doc_type": "other"}, "not_invoice"),
    ({"total": None}, "no_total"),
])
def test_a_bad_reading_is_flagged(change, code):
    inv = scan.to_invoice({**READ, **change}, "s.pdf", {OWN_ID})
    assert code in journal.issue_codes({"notes": " | ".join(inv.warnings)})


def test_credit_note_and_deal_invoice():
    credit = scan.to_invoice({**READ, "doc_type": "credit_note"}, "s.pdf", set())
    assert credit.total == D("-708.00") and credit.vat == D("-108.00")
    assert scan.to_invoice({**READ, "doc_type": "deal_invoice"}, "s.pdf", set()).deal_only


@pytest.fixture
def fake_claude(monkeypatch):
    calls = []
    monkeypatch.setattr(scan, "claude_path", lambda: "claude.exe")
    monkeypatch.setattr(scan, "read_data", lambda pdf, own, model="sonnet": calls.append(pdf.name) or dict(READ))
    return calls


def turn_on(on: bool = True):
    main.save(main.CONFIG, {**main.load(main.CONFIG, {}), "scan_reader": "claude" if on else ""})


def scanned_row(review):
    rows = list(openpyxl.load_workbook(review).active.iter_rows(values_only=True))
    return next(dict(zip(rows[0], r)) for r in rows[1:] if "07_scanned" in str(r[13]))


def test_off_by_default_nothing_is_sent(home, fake_claude):
    assert main.DEFAULT_CONFIG["scan_reader"] == ""
    row = scanned_row(main.fetch())
    assert fake_claude == [] and "scan" in journal.issue_codes({"notes": row["הערות"]})


def test_on_the_scan_is_read_and_still_waits_for_her(home, fake_claude, capsys):
    turn_on()
    row = scanned_row(main.fetch())
    assert fake_claude == ["07_scanned.pdf"]                     # only the scan, never a text PDF
    assert (row["ח.פ / ע.מ"], row["מס' חשבונית"], row['סה"כ']) == (S[6], "4410", 708)
    assert row["כרטיס הוצאה"] == "81900"                          # her rule: the stamped "סעיף" first
    assert "scan_read" in journal.issue_codes({"notes": row["הערות"]}) and row["לאשר? (כן/לא)"] != "כן"
    assert "נקראו 1 סריקות" in capsys.readouterr().out


def test_a_scan_is_sent_once(home, fake_claude):
    turn_on()
    (home / "קבצים").mkdir(exist_ok=True)
    fetched = sources.from_folder(Path(main.load(main.CONFIG, {})["test_folder"]), home / "קבצים")
    cfg = {**main.DEFAULT_CONFIG, **main.load(main.CONFIG, {})}
    main._read_invoices(fetched, cfg)
    main._read_invoices(fetched, cfg)
    assert fake_claude == ["07_scanned.pdf"]


def test_stamped_section_against_her_usual_card_is_flagged(home, fake_claude):
    main.save(main.SUPPLIERS, {S[6]: {"name": "מוסך", "supplier_account": "40016", "expense_account": "81600",
                                      "senders": [""]}})
    turn_on()
    row = scanned_row(main.fetch())
    assert row["כרטיס הוצאה"] == "81900"
    assert "section_differs" in journal.issue_codes({"notes": row["הערות"]})


def test_no_claude_on_the_pc_says_so(home, monkeypatch, capsys):
    monkeypatch.setattr(scan, "claude_path", lambda: None)
    turn_on()
    main.fetch()
    assert "Claude Code לא נמצא" in capsys.readouterr().out


def test_failed_reading_stays_manual(home, monkeypatch):
    monkeypatch.setattr(scan, "claude_path", lambda: "claude.exe")
    monkeypatch.setattr(scan, "read_data", lambda pdf, own, model="sonnet": None)
    turn_on()
    row = scanned_row(main.fetch())
    assert "scan_failed" in journal.issue_codes({"notes": row["הערות"]})


def test_the_switch_takes_only_on_or_off(home):
    h = main.HOME
    clients.save_settings(h, {"scan_reader": "claude"})
    assert main.load(h / "config.json", {})["scan_reader"] == "claude"
    clients.save_settings(h, {"scan_reader": "everything"})
    assert main.load(h / "config.json", {})["scan_reader"] == ""


def test_claude_runs_with_no_tools_and_the_images_inside_the_message(monkeypatch, tmp_path):
    """A crafted scan can carry text like "read file X": with no tools it cannot open anything."""
    seen = {}

    def fake_run(cmd, input, cwd, **kw):
        seen.update(cmd=cmd, input=input, cwd=cwd, files=list(Path(cwd).iterdir()))
        out = '{"type":"result","is_error":false,"structured_output":{"supplier_id":"1"}}'
        return type("R", (), {"returncode": 0, "stdout": out})()
    monkeypatch.setattr(scan.subprocess, "run", fake_run)
    assert scan.ask_claude([b"png-bytes"], "p", "claude.exe") == {"supplier_id": "1"}
    cmd = seen["cmd"]
    assert cmd[cmd.index("--tools") + 1] == "" and "--allowedTools" not in cmd and "Read" not in cmd
    assert "--strict-mcp-config" in cmd and "--no-session-persistence" in cmd
    assert seen["files"] == []                                  # an empty folder, nothing to find
    msg = __import__("json").loads(seen["input"])
    assert msg["message"]["content"][0]["type"] == "image"


def test_a_failed_or_refused_answer_is_no_answer(monkeypatch):
    def fake_run(cmd, input, cwd, **kw):
        return type("R", (), {"returncode": 0, "stdout": '{"type":"result","is_error":true,"structured_output":{"a":1}}'})()
    monkeypatch.setattr(scan.subprocess, "run", fake_run)
    assert scan.ask_claude([b"x"], "p", "claude.exe") is None
