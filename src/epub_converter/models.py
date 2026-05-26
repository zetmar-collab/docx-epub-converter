"""Shared data structures."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass
class ConversionResult:
    meta: dict
    chapters: list
    images: dict
    epub_bytes: bytes
    cover_bytes: bytes
    cover_ext: str
    saved_path: Path
    valid: bool
    messages: list
    epubcheck_ran: bool = True
