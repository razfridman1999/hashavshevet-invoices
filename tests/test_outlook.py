"""The Outlook stage against a fake Outlook object model (no Outlook on the dev machine).

The fake mirrors the parts of the COM API sources.from_outlook touches: 1-based
Attachments.Item(i), Class 43 for mail, timezone-aware ReceivedTime, Restrict().
"""
from __future__ import annotations

import sys
import types
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from invoices import sources  # noqa: E402

PDF = b"%PDF-1.4 fake"


class Attachment:
    def __init__(self, name):
        self.FileName = name

    def SaveAsFile(self, path):
        Path(path).write_bytes(PDF)


class Attachments:
    def __init__(self, names):
        self._a = [Attachment(n) for n in names]
        self.Count = len(self._a)

    def Item(self, i):  # COM collections are 1-based
        return self._a[i - 1]


class Mail:
    def __init__(self, entry, subject, files, days_ago=1, cls=43, body=""):
        self.Class = cls
        self.EntryID = f"{entry:0>40}"
        self.Subject = subject
        self.Body = body
        self.SenderEmailAddress = "billing@example.co.il"
        self.ReceivedTime = datetime.now(timezone.utc) - timedelta(days=days_ago)
        self.Attachments = Attachments(files)


class Items(list):
    restricted_with = None

    def Restrict(self, flt):
        Items.restricted_with = flt
        return self

    def Sort(self, prop, descending=False):
        assert prop == "[ReceivedTime]"
        self.sort(key=lambda m: m.ReceivedTime, reverse=descending)


class Folder:
    def __init__(self, items, sub=None):
        self.Items = Items(items)
        self.Folders = sub or {}


def install_fake_outlook(monkeypatch, inbox: Folder):
    ns = types.SimpleNamespace(GetDefaultFolder=lambda n: inbox if n == 6 else None)
    app = types.SimpleNamespace(GetNamespace=lambda name: ns)
    client = types.SimpleNamespace(Dispatch=lambda prog: app)
    win32com = types.ModuleType("win32com")
    win32com.client = client
    monkeypatch.setitem(sys.modules, "win32com", win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", client)


def test_picks_invoice_pdfs_only(tmp_path, monkeypatch):
    inbox = Folder([
        Mail(1, "חשבונית מס 4410 - חשמל השרון", ["inv.pdf", "logo.png"]),
        Mail(2, "סיכום פגישה", ["minutes.pdf"]),                                   # no invoice hint
        Mail(3, "מסמך מצורף", ["Invoice_778.PDF"]),                                # hint in the file name
        Mail(4, "עדכון", ["doc.pdf"], body="מצורפת חשבונית לחודש אוקטובר"),       # hint in the body
        Mail(5, "חשבונית", ["x.pdf"], cls=53),                                    # a meeting request
        Mail(6, "Re: חשבונית מס/קבלה", ["a.pdf", "b.pdf"]),                        # two in one mail
    ])
    install_fake_outlook(monkeypatch, inbox)
    got = sources.from_outlook(tmp_path, days_back=45)
    assert sorted(f.subject for f in got) == sorted(
        ["חשבונית מס 4410 - חשמל השרון", "מסמך מצורף", "עדכון", "Re: חשבונית מס/קבלה", "Re: חשבונית מס/קבלה"])
    assert len({f.mail_id for f in got}) == len(got)            # every attachment its own id
    assert all(f.pdf.exists() and f.pdf.parent == tmp_path for f in got)
    assert all(f.received.tzinfo is None for f in got)


def test_same_mail_twice_gives_same_ids(tmp_path, monkeypatch):
    inbox = Folder([Mail(7, "חשבונית", ["a.pdf"])])
    install_fake_outlook(monkeypatch, inbox)
    first = sources.from_outlook(tmp_path, 45)
    second = sources.from_outlook(tmp_path, 45)
    assert [f.mail_id for f in first] == [f.mail_id for f in second]


def test_saved_file_names_are_plain_ntfs_names(tmp_path, monkeypatch):
    """Review finding 6: a ':' in a Windows file name creates a hidden data stream."""
    install_fake_outlook(monkeypatch, Folder([Mail(9, "חשבונית", ["inv.pdf"])]))
    got = sources.from_outlook(tmp_path, 45)
    assert ":" not in got[0].pdf.name
    assert got[0].pdf.stat().st_size == len(PDF)


def test_old_mail_is_skipped_without_a_locale_dependent_filter(tmp_path, monkeypatch):
    """Review finding 12: the date cut-off must not depend on Windows' date format."""
    install_fake_outlook(monkeypatch, Folder([Mail(10, "חשבונית", ["new.pdf"], days_ago=3),
                                              Mail(11, "חשבונית", ["old.pdf"], days_ago=90)]))
    Items.restricted_with = None
    got = sources.from_outlook(tmp_path, 45)
    assert [f.pdf.name.split("_")[-1] for f in got] == ["new.pdf"]
    assert Items.restricted_with is None


def test_subfolder(tmp_path, monkeypatch):
    inbox = Folder([], sub={"חשבוניות": Folder([Mail(8, "חשבונית", ["a.pdf"])])})
    install_fake_outlook(monkeypatch, inbox)
    assert len(sources.from_outlook(tmp_path, 45, "חשבוניות")) == 1
