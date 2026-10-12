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


def main() -> None:
    command = sys.argv[1] if len(sys.argv) > 1 else "fetch"
    # setup runs from a .bat that cmd is still reading line by line: never rewrite it mid-run
    skip = os.environ.get("INVOICES_NO_UPDATE") or command == "setup"
    updated = None if skip else updater.update()
    try:
        app = _load_main()
    except Exception:
        if updated and updater.restore():     # the new version is broken: go back and carry on
            app = _load_main()
        else:
            raise
    sys.argv = [sys.argv[0], command]
    app.run(command)


if __name__ == "__main__":
    main()
