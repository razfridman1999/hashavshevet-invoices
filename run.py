"""Entry point for the desktop shortcuts: update quietly, then run; roll back if the
new version cannot start."""
from __future__ import annotations

import importlib
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
os.chdir(ROOT)
sys.path.insert(0, str(ROOT))

import updater  # noqa: E402


def _load_main():
    for name in [m for m in sys.modules if m == "invoices" or m.startswith("invoices.")]:
        del sys.modules[name]
    return importlib.import_module("invoices.main")


def _install_requirements() -> None:
    import subprocess
    exe = Path(sys.executable)
    python = exe.with_name("python.exe") if exe.name.lower() == "pythonw.exe" else exe
    subprocess.run([str(python), "-m", "pip", "install", "--user", "-q", "-r", str(ROOT / "requirements.txt")],
                   creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0), timeout=600, check=False)


def main() -> None:
    command = sys.argv[1] if len(sys.argv) > 1 else "fetch"
    # setup runs from a .bat that cmd is still reading line by line: never rewrite it mid-run
    skip = os.environ.get("INVOICES_NO_UPDATE") or command == "setup"
    updated = None if skip else updater.update()
    try:
        app = _load_main()
    except ModuleNotFoundError:
        # an update needs a package she does not have yet: install it, then start
        _install_requirements()
        try:
            app = _load_main()
        except Exception:
            if updated and updater.restore():
                app = _load_main()
            else:
                raise
    except Exception:
        if updated and updater.restore():     # the new version is broken: go back and carry on
            app = _load_main()
        else:
            raise
    sys.argv = [sys.argv[0], command]
    app.run(command)


if __name__ == "__main__":
    main()
