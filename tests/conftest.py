"""Ensure tests import the package from src/, not the root launcher script."""

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parents[1]
_SRC = _ROOT / "src"

if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

# Drop cached launcher module if pytest picked it up from project root.
sys.modules.pop("epub_converter", None)
