"""Talking to the bookkeeper in proper windows.

A console window cannot show Hebrew: cmd.exe has no bidi support, so every message
comes out backwards. Native message boxes and Tk windows render RTL correctly.
"""
from __future__ import annotations

import ctypes
import os
import sys

MB_OK, MB_ICONERROR, MB_ICONINFO = 0x0, 0x10, 0x40
MB_RIGHT, MB_RTLREADING, MB_TOPMOST = 0x80000, 0x100000, 0x40000
TITLE = "חשבוניות לחשבשבת"


def message(text: str, error: bool = False) -> None:
    print(text)  # also kept for tests and for anyone running it from a terminal
    if os.environ.get("INVOICES_NO_OPEN") or sys.platform != "win32":
        return
    flags = MB_OK | MB_RIGHT | MB_RTLREADING | MB_TOPMOST | (MB_ICONERROR if error else MB_ICONINFO)
    ctypes.windll.user32.MessageBoxW(None, text, TITLE, flags)


def fail(text: str) -> None:
    message(text, error=True)
    sys.exit(1)


def ask_settings(current: dict) -> dict | None:
    """First-run form: the three things only she knows. Returns None if she cancels."""
    import tkinter as tk

    fields = [("vat_account", 'מספר כרטיס מע"מ תשומות בחשבשבת', True),
              ("own_company_ids", 'ח.פ של החברה שאת עובדת בה (שלא ייחשב כספק)', True),
              ("outlook_folder", "תיקיית החשבוניות באאוטלוק (להשאיר ריק = דואר נכנס)", False)]
    root = tk.Tk()
    root.title(TITLE)
    root.resizable(False, False)
    font = ("Arial", 12)
    entries, result = {}, {}
    tk.Label(root, text="הגדרה ראשונה - פעם אחת בלבד", font=("Arial", 14, "bold")).pack(anchor="e", padx=16, pady=(14, 6))
    for key, label, _ in fields:
        tk.Label(root, text=label, font=font).pack(anchor="e", padx=16, pady=(8, 2))
        e = tk.Entry(root, font=font, width=40, justify="right")
        value = current.get(key)
        e.insert(0, ", ".join(value) if isinstance(value, list) else (value or ""))
        e.pack(anchor="e", padx=16)
        entries[key] = e
    error = tk.Label(root, text="", fg="#b00", font=font)
    error.pack(anchor="e", padx=16)

    def save():
        vals = {k: e.get().strip() for k, e in entries.items()}
        missing = [label for k, label, required in fields if required and not vals[k]]
        if missing:
            error.config(text="חסר: " + missing[0])
            return
        result.update(vals)
        root.destroy()

    tk.Button(root, text="שמירה", font=font, width=12, command=save).pack(pady=14)
    root.bind("<Return>", lambda _: save())
    root.mainloop()
    if not result:
        return None
    result["own_company_ids"] = [x.strip() for x in result["own_company_ids"].replace(",", " ").split() if x.strip()]
    return result
