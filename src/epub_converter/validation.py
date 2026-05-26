"""EPUB validation and Java runtime detection."""

from __future__ import annotations

import os
import shutil
import tempfile
from pathlib import Path

from epubcheck import EpubCheck


def java_available() -> bool:
    return shutil.which("java") is not None


def validate_epub(epub_bytes: bytes) -> tuple[bool, list]:
    with tempfile.NamedTemporaryFile(suffix=".epub", delete=False) as tmp:
        tmp.write(epub_bytes)
        tmp_path = tmp.name
    try:
        result = EpubCheck(tmp_path)
        return result.valid, result.messages
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass
