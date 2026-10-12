"""Self-update: works when it should, does nothing when it should not, rolls back a
version that cannot start. Runs against a throwaway copy of the project."""
from __future__ import annotations

import importlib.util
import io
import shutil
import sys
import zipfile
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]


def load_module(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


@pytest.fixture
def install(tmp_path, monkeypatch):
    """A copy of the app as it sits on her PC (no .git), with its own updater module."""
    app = tmp_path / "app"
    shutil.copytree(ROOT, app, ignore=shutil.ignore_patterns(".git", "tests", "demo", ".backup",
                                                             "__pycache__", ".pytest_cache"))
    (app / "VERSION").write_text("1.0\n", encoding="utf-8")
    mod = load_module("updater", app / "updater.py")
    monkeypatch.setattr(sys, "path", [str(app)] + [p for p in sys.path if p != str(ROOT)])
    yield app, mod
    for name in [m for m in sys.modules if m in ("updater", "run") or m.startswith("invoices")]:
        del sys.modules[name]


def release_zip(app: Path, version: str, tweak=None) -> bytes:
    """What GitHub's codeload would serve for the release branch."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        for f in app.rglob("*"):
            if f.is_file() and ".backup" not in f.parts and "__pycache__" not in f.parts:
                rel = f.relative_to(app).as_posix()
                data = f.read_bytes()
                if rel == "VERSION":
                    data = f"{version}\n".encode()
                if tweak and rel in tweak:
                    data = tweak[rel].encode()
                z.writestr(f"hashavshevet-invoices-release/{rel}", data)
    return buf.getvalue()


def serve(monkeypatch, mod, version: str, zip_bytes: bytes | None = None, fail=False):
    def fake_get(url, timeout):
        if fail:
            raise OSError("no network")
        return version.encode() if url.endswith("/VERSION") else zip_bytes
    monkeypatch.setattr(mod, "_get", fake_get)


def test_newer_release_is_installed_and_backed_up(install, monkeypatch):
    app, mod = install
    marker = "# release 2.0\n"
    main_src = (app / "invoices" / "main.py").read_text(encoding="utf-8")
    serve(monkeypatch, mod, "2.0", release_zip(app, "2.0", {"invoices/main.py": marker + main_src}))
    assert mod.update() == "2.0"
    assert (app / "VERSION").read_text().strip() == "2.0"
    assert (app / "invoices" / "main.py").read_text(encoding="utf-8").startswith(marker)
    assert (app / ".backup" / "VERSION").read_text().strip() == "1.0"


def test_same_version_does_nothing(install, monkeypatch):
    app, mod = install
    serve(monkeypatch, mod, "1.0", b"never downloaded")
    assert mod.update() is None
    assert not (app / ".backup").exists()


def test_no_network_keeps_working(install, monkeypatch):
    app, mod = install
    serve(monkeypatch, mod, "2.0", fail=True)
    assert mod.update() is None
    assert (app / "VERSION").read_text().strip() == "1.0"


def test_release_that_does_not_compile_is_refused(install, monkeypatch):
    app, mod = install
    before = (app / "invoices" / "extract.py").read_text(encoding="utf-8")
    serve(monkeypatch, mod, "2.0", release_zip(app, "2.0", {"invoices/extract.py": "def broken(:\n"}))
    assert mod.update() is None
    assert (app / "invoices" / "extract.py").read_text(encoding="utf-8") == before
    assert (app / "VERSION").read_text().strip() == "1.0"


def test_release_that_crashes_on_start_is_rolled_back(install, monkeypatch):
    """It compiles, but blows up on import: run.py must restore the old version and run it."""
    app, mod = install
    serve(monkeypatch, mod, "2.0", release_zip(app, "2.0", {"invoices/main.py": "raise RuntimeError('bad release')\n"}))
    ran = []
    run = load_module("run", app / "run.py")
    real_load = run._load_main

    def load_and_stub():
        m = real_load()
        m.run = lambda command: ran.append(command)
        return m
    monkeypatch.setattr(run, "_load_main", load_and_stub)
    monkeypatch.setattr(sys, "argv", ["run.py", "fetch"])
    monkeypatch.delenv("INVOICES_NO_UPDATE", raising=False)
    run.main()
    assert ran == ["fetch"]
    assert (app / "VERSION").read_text().strip() == "1.0"
    assert "raise RuntimeError" not in (app / "invoices" / "main.py").read_text(encoding="utf-8")


def test_developer_checkout_is_never_overwritten(install, monkeypatch):
    app, mod = install
    (app / ".git").mkdir()
    serve(monkeypatch, mod, "2.0", release_zip(app, "2.0"))
    assert mod.update() is None
