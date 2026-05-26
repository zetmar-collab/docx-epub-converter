#!/usr/bin/env python3
"""Start DOCX EPUB Converter (adds src/ to PYTHONPATH)."""

from __future__ import annotations

import sys
from pathlib import Path

_SRC = Path(__file__).resolve().parent / "src"
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

from epub_converter.app import main  # noqa: E402

if __name__ == "__main__":
    main()
