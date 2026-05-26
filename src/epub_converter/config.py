"""User configuration persistence."""

import json
from pathlib import Path

CONFIG_FILE = Path.home() / ".epub_converter_config.json"


def load_config() -> dict:
    try:
        return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_config(data: dict) -> None:
    try:
        CONFIG_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        pass


def update_config(updates: dict) -> None:
    cfg = load_config()
    cfg.update(updates)
    save_config(cfg)
