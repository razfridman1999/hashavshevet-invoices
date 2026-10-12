"""The saved field mapping ("שידוך") of Hashavshevet's ODBC import, as a .wpr file next to
the import file, so she only loads it (the yellow folder icon in the ODBC window) instead
of matching every field by hand.

Layout, read from a mapping saved at her office and used for a real import (מנה 762):
  02 'V2', then length-prefixed cp1255 strings: the ODBC connection string, the table
  'Journal', the range '[jurnal<hid>]', and then the mapping itself - one slot per
  Hashavshevet field, holding the name of the column that feeds it (empty = unused)."""
from __future__ import annotations

from pathlib import Path

ENC = "cp1255"
# Hashavshevet's journal fields in its own slot order, and the column feeding each one.
SLOTS = (["סוג תנועה", "אסמכתא 1", "אסמכתא 2", "תאריך אסמכתא", "תאריך ערך", "קוד תמחיר", "קוד מטבע",
          "פרטים", "חשבון חובה 1", "חשבון חובה 2", "חשבון זכות 1", "חשבון זכות 2",
          "סכום חובה 1", "סכום חובה 2", "סכום זכות 1", "סכום זכות 2"]
         + [""] * 11 + ["מספר עמ"] + [""] * 5 + ["מספר הקצאה"] + [""] * 3)


def pstr(s: str) -> bytes:
    b = s.encode(ENC)
    if len(b) > 255:
        raise ValueError(f"too long for a .wpr string ({len(b)} bytes)")
    return bytes([len(b)]) + b


def mapping() -> bytes:
    return b"\x00\x03WIN%\x00\x00\x00\x00" + b"".join(pstr(s) for s in SLOTS) + b"\x010" * 4


def build(data_file: Path, hid: str) -> bytes:
    """.xls goes through the Excel 97 driver (790), .xlsx through the 2007 one (1046)."""
    driver = 790 if data_file.suffix.lower() == ".xls" else 1046
    conn = (f"ODBC;DSN=Excel Files;DBQ={data_file};DefaultDir={data_file.parent};"
            f"DriverId={driver};MaxBufferSize=2048;PageTimeout=5;")
    return b"\x02V2" + pstr(conn) + pstr("Journal") + pstr(f"[jurnal{hid}]") + mapping()


def write(data_file: Path, hid: str) -> Path:
    """The name Hashavshevet itself gives a saved mapping: Journal.<file>.jurnal<hid>.wpr"""
    out = data_file.with_name(f"Journal.{data_file.name}.jurnal{hid}.wpr")
    out.write_bytes(build(data_file, hid))
    return out
