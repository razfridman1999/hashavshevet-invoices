"""The home page: one icon opens a page in her browser with big buttons and where every
invoice stands. A tiny server on 127.0.0.1 runs the same fetch / export as the old
buttons and hands the result back to the page.

Any website she visits can send requests to 127.0.0.1, so every call needs the random
token that only her page's address carries, and the Host header must be ours."""
from __future__ import annotations

import json
import os
import secrets
import subprocess
import sys
import threading
import time
import traceback
import urllib.request
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from . import clients, insights, journal, main, report, sources, ui

APP = Path(__file__).resolve().parents[1]
PAGE = Path(__file__).with_name("app.html")
IDLE_LIMIT = 30 * 60          # the page pings while open; a forgotten server goes away
LOCK_NAME = "app.json"
CSP = ("default-src 'none'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; "
       "connect-src 'self'; img-src data:; base-uri 'none'; form-action 'none'")


def version() -> str:
    try:
        return (APP / "VERSION").read_text(encoding="utf-8").strip()
    except OSError:
        return ""


def _latest(pattern: str, home: Path | None = None) -> Path | None:
    files = sorted((home or main.HOME).glob(pattern), key=lambda p: p.stat().st_mtime)
    return files[-1] if files else None


def status(home: Path | None = None, root: Path | None = None, client_id: str = "") -> dict:
    home = home or main.HOME
    cfg = {**main.DEFAULT_CONFIG, **main.load(home / "config.json", {})}
    entries = journal.load(home)
    month = datetime.now().strftime("%Y-%m")
    items = sorted(entries.items(), key=lambda kv: kv[1]["updated"], reverse=True)
    items.sort(key=lambda kv: journal.ORDER.get(kv[1]["status"], 9))
    review = _latest("לבדיקה *.xlsx", home)
    last = main.load(home / "state.json", {}).get("last_import")
    imported = Path(last) if last else _latest("לקליטה בחשבשבת *.xlsx", home)
    return {
        "version": version(),
        "client": client_id,
        "clients": clients.list_clients(root or home),
        "settings": clients.settings(home),
        "configured": bool(cfg["own_company_ids"] and cfg["vat_account"] and cfg["ttype"]),
        "counts": {
            "pending": sum(e["status"] == journal.PENDING for _, e in items),
            "ready": sum(e["status"] == journal.READY for _, e in items),
            "exported_month": sum(e["status"] == journal.EXPORTED and e["updated"].startswith(month)
                                  for _, e in items),
        },
        "items": [{**{k: (str(e.get(k)) if e.get(k) is not None else "") for k, _ in journal.COLS},
                   "key": key, "has_file": bool(e.get("file"))}
                  for key, e in items[:300]],
        "review": review.name if review else "",
        "import": imported.name if imported else "",
        "summary": insights.monthly_summary(entries),
        "missing": insights.missing_invoices(entries),
    }


def open_pdf(key: str, home: Path | None = None) -> dict:
    """Only a PDF the tool itself saved, named by its log entry - never a path the
    request supplies, so the page cannot be used to open anything else."""
    home = home or main.HOME
    entry = journal.load(home).get(key)
    inbox = (home / "קבצים").resolve()
    path = Path(entry["file"]).resolve() if entry and entry.get("file") else None
    if not path or path.suffix.lower() != ".pdf" or inbox not in path.parents or not path.exists():
        return {"ok": False, "text": "הקובץ של החשבונית הזו לא נמצא (אולי נמחק מהתיקייה)."}
    _open(path)
    return {"ok": True, "text": ""}


def make_report(note: str, home: Path | None = None) -> dict:
    path = report.write(home or main.HOME, version(), note[:4000])
    _open(path)
    return {"ok": True, "text": f"הדוח נשמר ונפתח לבדיקה:\n{path}\n\n"
                                "אין בו שמות, מספרים או סכומים מהחשבוניות. אפשר לשלוח אותו לרז בוואטסאפ או במייל."}


PICK = """
import sys, tkinter as tk
from pathlib import Path
from tkinter import filedialog
root = tk.Tk(); root.withdraw(); root.attributes("-topmost", True)
if sys.argv[1] == "folder":
    p = filedialog.askdirectory(title="בחירת תיקייה")
else:
    p = filedialog.askopenfilename(title="בחירת קובץ", filetypes=[("Excel", "*.xlsx *.xls"), ("הכול", "*.*")])
Path(sys.argv[2]).write_text(str(Path(p)) if p else "", encoding="utf-8")
"""


def pick(kind: str, home: Path) -> dict:
    """A Windows folder/file dialog, so she never types a path. Its own process: Tk wants
    the main thread, and a request runs in a side thread."""
    if os.environ.get("INVOICES_NO_OPEN"):
        return {"ok": False, "path": ""}
    out = home / "pick.tmp"
    try:
        subprocess.run([sys.executable, "-c", PICK, "folder" if kind == "folder" else "file", str(out)],
                       timeout=600, check=False)
        path = out.read_text(encoding="utf-8").strip() if out.exists() else ""
    except (OSError, subprocess.TimeoutExpired):
        path = ""
    finally:
        out.unlink(missing_ok=True)
    return {"ok": bool(path), "path": path}


def outlook_folders() -> dict:
    com = _com_init()
    try:
        return {"ok": True, "folders": sources.folder_paths()}
    except Exception:
        return {"ok": False, "folders": [], "text": "לא הצלחתי לקרוא את התיקיות מאאוטלוק. לבדוק שאאוטלוק פתוח."}
    finally:
        _com_done(com)


def _open(path: Path) -> None:
    if not os.environ.get("INVOICES_NO_OPEN"):
        os.startfile(path)


def _com_init() -> bool:
    """Every request runs in its own thread, and COM (Outlook, Excel) refuses to work in a
    thread that did not initialise it: "CoInitialize has not been called"."""
    try:
        import pythoncom
    except ImportError:
        return False
    pythoncom.CoInitialize()
    return True


def _com_done(inited: bool) -> None:
    if inited:
        import pythoncom
        pythoncom.CoUninitialize()


def run_action(name: str) -> dict:
    """One button press. Returns what to show on the page: {ok, text}."""
    ui.captured = []
    ok = True
    com = _com_init()
    try:
        if name == "fetch":
            out = main.fetch()
            if out:
                _open(out)
        elif name == "export":
            if main.export():
                _open(main.HOME)
        elif name == "review":
            review = _latest("לבדיקה *.xlsx")
            if not review:
                ui.message('עוד אין טבלת בדיקה. קודם ללחוץ על "משיכת חשבוניות".')
            else:
                _open(review)
        elif name == "journal":
            main.show_journal()
        elif name == "folder":
            _open(main.HOME)
        elif name == "guide":
            _open(APP / "מדריך.html")
        elif name == "setup":   # its own process: the settings form is a Tk window
            if not os.environ.get("INVOICES_NO_OPEN"):
                subprocess.Popen([sys.executable, str(APP / "run.py"), "setup"], cwd=str(APP))
            ui.message("חלון ההגדרות נפתח. אחרי השמירה הדף יתעדכן.")
        else:
            return {"ok": False, "text": "פעולה לא מוכרת"}
    except SystemExit:          # ui.fail: the message is already captured
        ok = False
    except Exception:
        log = main.HOME / "שגיאה אחרונה.txt"
        log.write_text(traceback.format_exc(), encoding="utf-8")
        ui.captured.append((f"משהו השתבש. שום דבר לא נקלט.\n\nלשלוח לרז את הקובץ:\n{log}", True))
        ok = False
    finally:
        messages, ui.captured = ui.captured, None
        _com_done(com)
    return {"ok": ok, "text": "\n\n".join(text for text, _ in messages)}


class App:
    def __init__(self, port: int = 0):
        self.root = main.HOME                 # the main folder; other clients live under it
        self.token = secrets.token_urlsafe(24)
        self.busy = threading.Lock()          # one action at a time: they share ui and files
        self.last_seen = time.monotonic()
        app = self

        class Handler(BaseHTTPRequestHandler):
            def log_message(self, *args):     # pythonw has no console to write to
                pass

            def _allowed(self) -> bool:
                host_ok = self.headers.get("Host") == f"127.0.0.1:{app.port}"
                query = parse_qs(urlparse(self.path).query)
                given = self.headers.get("X-Token") or (query.get("t") or [""])[0]
                return host_ok and secrets.compare_digest(given, app.token)

            def _send(self, code: int, body: bytes, kind: str) -> None:
                self.send_response(code)
                self.send_header("Content-Type", kind)
                self.send_header("Content-Length", str(len(body)))
                self.send_header("Cache-Control", "no-store")
                self.send_header("Content-Security-Policy", CSP)
                self.send_header("X-Content-Type-Options", "nosniff")
                self.send_header("Referrer-Policy", "no-referrer")
                self.end_headers()
                self.wfile.write(body)

            def _json(self, data, code: int = 200) -> None:
                self._send(code, json.dumps(data, ensure_ascii=False).encode("utf-8"),
                           "application/json; charset=utf-8")

            def _client(self) -> tuple[str, Path | None]:
                cid = (parse_qs(urlparse(self.path).query).get("c") or [""])[0]
                return cid, clients.home_of(app.root, cid)

            def _drain(self) -> None:
                """Read and drop what was sent before refusing it: answering over an unread body
                makes Windows reset the connection, and the caller never sees the answer."""
                left = min(int(self.headers.get("Content-Length") or 0), 1_000_000)
                while left > 0:
                    chunk = self.rfile.read(min(left, 65_536))
                    if not chunk:
                        break
                    left -= len(chunk)

            def _body(self) -> dict | None:
                size = int(self.headers.get("Content-Length") or 0)
                if size > 64_000:
                    self._drain()
                    return None
                try:
                    data = json.loads(self.rfile.read(size) or b"{}")
                except ValueError:
                    data = {}
                return data if isinstance(data, dict) else {}

            def do_GET(self):
                if not self._allowed():
                    return self._send(403, b"forbidden", "text/plain")
                app.last_seen = time.monotonic()
                path = urlparse(self.path).path
                if path == "/":
                    return self._send(200, PAGE.read_bytes(), "text/html; charset=utf-8")
                cid, home = self._client()
                if home is None:
                    cid, home = "", app.root          # a client deleted by hand: back to the main one
                if path == "/api/status":
                    return self._json(status(home, app.root, cid))
                self._send(404, b"not found", "text/plain")

            def do_POST(self):
                if not self._allowed() or not self.headers.get("X-Token"):   # POST: header only
                    self._drain()
                    return self._send(403, b"forbidden", "text/plain")
                app.last_seen = time.monotonic()
                url = urlparse(self.path)
                cid, home = self._client()
                if home is None:
                    return self._json({"ok": False, "text": "הלקוח הזה לא נמצא. לבחור לקוח מהרשימה."}, 404)
                if url.path == "/api/pdf":
                    return self._json(open_pdf((parse_qs(url.query).get("key") or [""])[0], home))
                if url.path in ("/api/report", "/api/settings", "/api/clients"):
                    data = self._body()
                    if data is None:
                        return self._send(413, b"too large", "text/plain")
                    if url.path == "/api/report":
                        return self._json(make_report(str(data.get("note", "")), home))
                    if url.path == "/api/settings":
                        ok, text = clients.save_settings(home, data)
                        return self._json({"ok": ok, "text": text})
                    new_id, problem = clients.add(app.root, str(data.get("name", "")))
                    return self._json({"ok": new_id is not None, "id": new_id, "text": problem})
                if url.path == "/api/outlook-folders":
                    return self._json(outlook_folders())
                if url.path == "/api/pick":
                    return self._json(pick((parse_qs(url.query).get("kind") or [""])[0], home))
                if not url.path.startswith("/api/run/"):
                    return self._send(404, b"not found", "text/plain")
                if not app.busy.acquire(blocking=False):
                    return self._json({"ok": False, "text": "פעולה אחרת עוד רצה. לחכות שתסתיים."}, 409)
                try:
                    main.use(home)                    # the action works inside this client's folder
                    result = run_action(url.path.rsplit("/", 1)[1])
                finally:
                    app.busy.release()
                    app.last_seen = time.monotonic()
                self._json(result)

        self.server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
        self.port = self.server.server_address[1]
        self.url = f"http://127.0.0.1:{self.port}/?t={self.token}"

    def watch_idle(self) -> None:
        while True:
            time.sleep(30)
            if time.monotonic() - self.last_seen > IDLE_LIMIT and not self.busy.locked():
                self.server.shutdown()
                return


def _running_url() -> str | None:
    """The page of a server that is already up, so a second click reuses it."""
    lock = main.HOME / LOCK_NAME
    try:
        info = json.loads(lock.read_text(encoding="utf-8"))
        url = f"http://127.0.0.1:{info['port']}/api/status?t={info['token']}"
        with urllib.request.urlopen(url, timeout=2) as r:
            if r.status == 200:
                return f"http://127.0.0.1:{info['port']}/?t={info['token']}"
    except Exception:
        return None
    return None


def serve() -> None:
    import webbrowser
    url = _running_url()
    if url:
        webbrowser.open(url)
        return
    app = App()
    lock = main.HOME / LOCK_NAME
    main.save(lock, {"port": app.port, "token": app.token})
    threading.Thread(target=app.watch_idle, daemon=True).start()
    webbrowser.open(app.url)
    try:
        app.server.serve_forever()
    finally:
        try:
            if json.loads(lock.read_text(encoding="utf-8")).get("token") == app.token:
                lock.unlink()
        except Exception:
            pass
