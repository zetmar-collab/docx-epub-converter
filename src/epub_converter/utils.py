"""Shared utilities."""

from __future__ import annotations

import html
import os
import re
import subprocess
import sys
from pathlib import Path


def xml_escape(text: str) -> str:
    return html.escape(text or "", quote=False)


def e(text: str) -> bytes:
    return text.encode("utf-8")


def safe_filename(title: str) -> str:
    name = re.sub(r"[^\w\-]+", "-", title.lower(), flags=re.UNICODE).strip("-")
    return (name or "ebook") + ".epub"


def default_output_dir() -> Path:
    downloads = Path.home() / "Downloads"
    return downloads if downloads.exists() else Path.home()


def html_to_plain_text(fragment: str) -> str:
    text = re.sub(r"</(p|h1|h2|h3|li|blockquote|tr)>", "\n", fragment)
    text = re.sub(r"</(table|ul|ol)>", "\n", text)
    text = re.sub(r"<li>", "- ", text)
    text = re.sub(r"<t[hd][^>]*>", "  ", text)
    text = re.sub(r"<[^>]+>", "", text)
    text = html.unescape(text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def open_path(path: Path) -> None:
    if sys.platform.startswith("win"):
        os.startfile(path)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


def sanitize_dropped_path(raw: str) -> Path | None:
    """Validate drag-and-drop path: must be an existing regular file."""
    cleaned = raw.strip().strip("{}").strip('"')
    try:
        path = Path(cleaned).expanduser().resolve()
    except (OSError, ValueError):
        return None
    if path.is_file():
        return path
    return None
