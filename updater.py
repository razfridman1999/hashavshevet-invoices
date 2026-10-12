"""Self-update from the public GitHub repo's `release` branch.

Runs before every launch and never blocks work: no network, a timeout or a broken
download all mean "keep the version that is here". Before replacing anything, the
current code is copied to .backup/, so run.py can roll back if a new version fails
to even start.
"""
from __future__ import annotations

import io
import py_compile
import shutil
import tempfile
import urllib.request
import zipfile
from pathlib import Path

REPO = "razfridman1999/hashavshevet-invoices"
BRANCH = "release"
ROOT = Path(__file__).resolve().parent
VERSION = ROOT / "VERSION"
BACKUP = ROOT / ".backup"
# what an update replaces; her data never lives here (it is in Documents)
MANAGED = ["invoices", "run.py", "updater.py", "VERSION", "requirements.txt", "מדריך.html",
           "0 - התקנה ראשונה.bat", "1 - משיכת חשבוניות.bat", "2 - יצירת קובץ לחשבשבת.bat"]


def local_version() -> str:
    return VERSION.read_text(encoding="utf-8").strip() if VERSION.exists() else ""


def _get(url: str, timeout: float) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": "hashavshevet-invoices-updater",
                                               "Cache-Control": "no-cache"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def remote_version(timeout: float = 4) -> str:
    return _get(f"https://raw.githubusercontent.com/{REPO}/{BRANCH}/VERSION", timeout).decode().strip()


def _validate(tree: Path) -> None:
    """Refuse a download that is incomplete or does not even compile."""
    for name in ("invoices/main.py", "invoices/extract.py", "run.py", "updater.py", "VERSION"):
        if not (tree / name).exists():
            raise RuntimeError(f"update is missing {name}")
    for py in list((tree / "invoices").glob("*.py")) + [tree / "run.py", tree / "updater.py"]:
        py_compile.compile(str(py), doraise=True, cfile=str(Path(tempfile.gettempdir()) / "hi_check.pyc"))


def backup() -> None:
    if BACKUP.exists():
        shutil.rmtree(BACKUP)
    BACKUP.mkdir()
    for name in MANAGED:
        src = ROOT / name
        if src.is_dir():
            shutil.copytree(src, BACKUP / name, ignore=shutil.ignore_patterns("__pycache__"))
        elif src.exists():
            shutil.copy2(src, BACKUP / name)


def restore() -> bool:
    """Put the pre-update code back. Returns False when there is nothing to restore."""
    if not (BACKUP / "invoices").exists():
        return False
    for name in MANAGED:
        src, dst = BACKUP / name, ROOT / name
        if not src.exists():
            continue
        if dst.is_dir():
            shutil.rmtree(dst)
        if src.is_dir():
            shutil.copytree(src, dst)
        else:
            shutil.copy2(src, dst)
    return True


def update(timeout: float = 4) -> str | None:
    """Install a newer release if there is one. Returns the new version, or None."""
    if (ROOT / ".git").exists():          # a developer checkout: never overwrite it
        return None
    try:
        new = remote_version(timeout)
        if not new or new == local_version():
            return None
        data = _get(f"https://codeload.github.com/{REPO}/zip/refs/heads/{BRANCH}", timeout * 5)
        with tempfile.TemporaryDirectory() as tmp:
            zipfile.ZipFile(io.BytesIO(data)).extractall(tmp)
            tree = next(Path(tmp).iterdir())          # GitHub wraps it in <repo>-<branch>/
            _validate(tree)
            if (tree / "VERSION").read_text(encoding="utf-8").strip() != new:
                return None                           # raced with a newer push: try next time
            backup()
            for name in MANAGED:
                src, dst = tree / name, ROOT / name
                if not src.exists():
                    continue
                if src.is_dir():
                    if dst.exists():
                        shutil.rmtree(dst)
                    shutil.copytree(src, dst)
                elif name != "VERSION":
                    shutil.copy2(src, dst)
            shutil.copy2(tree / "VERSION", VERSION)   # last: a half-done update retries next time
        return new
    except Exception:
        return None
