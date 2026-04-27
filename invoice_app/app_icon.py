from __future__ import annotations

import ctypes
import sys
import tkinter as tk
from pathlib import Path


APP_USER_MODEL_ID = "liuwenhui.invoice.automation"


def resource_path(relative_path: str) -> Path:
    """Return a path that works from source and from a PyInstaller bundle."""
    if getattr(sys, "frozen", False):
        bundle_root = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
    else:
        bundle_root = Path(__file__).resolve().parents[1]
    return bundle_root / relative_path


def configure_windows_app_id() -> None:
    if not hasattr(ctypes, "windll"):
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_USER_MODEL_ID)
    except Exception:
        pass


def apply_window_icon(window: tk.Misc) -> None:
    icon_path = resource_path("assets/invoice_processing.ico")
    if not icon_path.exists():
        return
    try:
        window.iconbitmap(default=str(icon_path))
    except tk.TclError:
        pass
