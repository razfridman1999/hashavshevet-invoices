"""Several clients, each kept apart: its own settings, suppliers, log and sheets. Supplier
cards and ס.ת differ between clients and must never cross from one to another, so each
client is a folder of its own and the tool simply works inside the chosen one.

The first client is the main folder itself (where everything lived before), so nothing
already there moves. Every other client is a folder under "לקוחות". The real names and
paths are typed in on her PC and stay there."""
from __future__ import annotations

import re
from pathlib import Path

from .extract import valid_israeli_id

FOLDER = "לקוחות"
MAIN_NAME = "לקוח ראשי"
TEXT_FIELDS = ("client_name", "outlook_folder", "done_folder", "vat_account", "ttype", "hid",
               "supplier_index", "files_folder", "rep_folder")


def _cfg(home: Path) -> dict:
    from . import main
    return {**main.DEFAULT_CONFIG, **main.load(home / "config.json", {})}


def list_clients(root: Path) -> list[dict]:
    out = [{"id": "", "name": _cfg(root).get("client_name") or MAIN_NAME}]
    base = root / FOLDER
    if base.is_dir():
        for d in sorted(base.iterdir(), key=lambda p: p.name):
            if (d / "config.json").exists():
                out.append({"id": d.name, "name": _cfg(d).get("client_name") or d.name})
    return out


def home_of(root: Path, client_id: str) -> Path | None:
    """The folder of a client the list knows - never a path built from what a request sent."""
    if not client_id:
        return root
    known = {c["id"] for c in list_clients(root)}
    return root / FOLDER / client_id if client_id in known else None


def folder_name(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|\x00-\x1f]', "", name).strip(" .")[:60]


def add(root: Path, name: str) -> tuple[str | None, str]:
    """(new client's id, problem)."""
    from . import main
    cid = folder_name(name or "")
    if not cid:
        return None, "צריך לכתוב שם ללקוח."
    if any(c["name"] == name.strip() for c in list_clients(root)) or (root / FOLDER / cid).exists():
        return None, "כבר יש לקוח בשם הזה."
    home = root / FOLDER / cid
    home.mkdir(parents=True)
    main.save(home / "config.json", {**main.DEFAULT_CONFIG, "client_name": name.strip()})
    return cid, ""


def settings(home: Path) -> dict:
    cfg = _cfg(home)
    out = {k: str(cfg.get(k) or "") for k in TEXT_FIELDS}
    out["own_company_ids"] = ", ".join(cfg.get("own_company_ids") or [])
    return out


def save_settings(home: Path, data: dict) -> tuple[bool, str]:
    """Only the known fields, as text. Returns (saved, what to tell her)."""
    from . import main
    cfg = _cfg(home)
    for k in TEXT_FIELDS:
        if k in data:
            cfg[k] = str(data[k] or "").strip()[:300]
    if "own_company_ids" in data:
        cfg["own_company_ids"] = [x for x in re.split(r"[\s,]+", str(data["own_company_ids"] or "")) if x][:5]
    problems = []
    if cfg["hid"] and not cfg["hid"].isdigit():
        problems.append('מזהה החברה בחשבשבת הוא מספר בלבד (למשל 305, מתוך "jurnal305").')
    bad = [n for n in cfg["own_company_ids"] if not valid_israeli_id(n)]
    if bad:
        problems.append(f"המספר {bad[0]} לא נראה כמו ח.פ תקין.")
    if problems:
        return False, "\n".join(problems)
    main.save(home / "config.json", cfg)
    missing = [label for k, label in (("own_company_ids", "ח.פ"), ("vat_account", 'כרטיס מע"מ'), ("ttype", "ס.ת"))
               if not cfg[k]]
    return True, "נשמר." + (f" עוד חסר: {', '.join(missing)}." if missing else "")
