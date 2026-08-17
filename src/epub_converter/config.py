"""User configuration persistence.

Config lives in the per-user application-data folder so the app works when
installed read-only (MSIX / Microsoft Store, Program Files). The legacy
``~/.epub_converter_config.json`` file is migrated on first read.
"""

from __future__ import annotations

import json
import os
from pathlib import Path

_LEGACY_CONFIG_FILE = Path.home() / ".epub_converter_config.json"


def _config_dir() -> Path:
    base = os.environ.get("LOCALAPPDATA") or os.environ.get("XDG_CONFIG_HOME")
    if base:
        return Path(base) / "DocxEpubConverter"
    return Path.home() / ".config" / "DocxEpubConverter"


CONFIG_FILE = _config_dir() / "config.json"


def load_config() -> dict:
    for candidate in (CONFIG_FILE, _LEGACY_CONFIG_FILE):
        try:
            data = json.loads(candidate.read_text(encoding="utf-8"))
        except Exception:
            continue
        if isinstance(data, dict):
            if candidate is _LEGACY_CONFIG_FILE:
                save_config(data)  # migrate forward, keep legacy file untouched
            return data
    return {}


def save_config(data: dict) -> None:
    try:
        CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
        CONFIG_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        pass


def update_config(updates: dict) -> None:
    cfg = load_config()
    cfg.update(updates)
    save_config(cfg)
