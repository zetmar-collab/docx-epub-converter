"""Cross-platform UI font selection."""

from __future__ import annotations

import tkinter as tk
import tkinter.font as tkfont


def pick_ui_font(root: tk.Misc, size: int = 10, weight: str = "normal") -> tuple:
    preferred = (
        "Segoe UI",
        "SF Pro Text",
        "Helvetica Neue",
        "Ubuntu",
        "DejaVu Sans",
        "TkDefaultFont",
    )
    available = set(tkfont.families(root))
    for name in preferred:
        if name in available:
            if weight == "bold":
                return (name, size, "bold")
            return (name, size)
    return ("TkDefaultFont", size, "bold") if weight == "bold" else ("TkDefaultFont", size)
