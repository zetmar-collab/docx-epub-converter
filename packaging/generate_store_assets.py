#!/usr/bin/env python3
"""Generate the full Microsoft Store / MSIX asset set from assets/app_icon.png.

Produces every tile, scale factor and target-size variant that the Store
submission and the Windows shell expect, into packaging/Assets/.

Usage:
    python packaging/generate_store_assets.py
"""

from __future__ import annotations

import sys
from pathlib import Path

from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "assets" / "app_icon.png"
OUT = Path(__file__).resolve().parent / "Assets"

# Background used for the non-transparent "plated" Square44x44 variants.
PLATE_BG = (7, 26, 51, 255)  # NAVY_BG

# (base name, width, height) at scale 100
TILES = [
    ("Square44x44Logo", 44, 44),
    ("Square71x71Logo", 71, 71),
    ("Square150x150Logo", 150, 150),
    ("Square310x310Logo", 310, 310),
    ("Wide310x150Logo", 310, 150),
    ("StoreLogo", 50, 50),
    ("SplashScreen", 620, 300),
]

SCALES = (100, 125, 150, 200, 400)

# Square44x44Logo.targetsize-*.png — used in the taskbar, Start search, Alt+Tab.
TARGET_SIZES = (16, 20, 24, 30, 32, 36, 40, 48, 56, 60, 64, 72, 80, 96, 256)


def _fit(src: Image.Image, w: int, h: int, *, pad_ratio: float = 0.0) -> Image.Image:
    """Center `src` on a transparent w×h canvas, contained, with optional padding."""
    canvas = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    inner_w = max(1, int(w * (1 - pad_ratio)))
    inner_h = max(1, int(h * (1 - pad_ratio)))
    art = src.copy()
    art.thumbnail((inner_w, inner_h), Image.LANCZOS)
    canvas.paste(art, ((w - art.width) // 2, (h - art.height) // 2), art)
    return canvas


def main() -> int:
    if not SOURCE.is_file():
        print(f"ERROR: source icon not found: {SOURCE}", file=sys.stderr)
        return 1

    OUT.mkdir(parents=True, exist_ok=True)
    src = Image.open(SOURCE).convert("RGBA")
    written = 0

    for name, base_w, base_h in TILES:
        # Wide/splash tiles keep the art inside a safe area so it is not cropped.
        pad = 0.20 if name in ("Wide310x150Logo", "SplashScreen") else 0.10
        for scale in SCALES:
            w = max(1, round(base_w * scale / 100))
            h = max(1, round(base_h * scale / 100))
            _fit(src, w, h, pad_ratio=pad).save(OUT / f"{name}.scale-{scale}.png")
            written += 1
        # Scale-100 alias without suffix (some tooling expects the bare name).
        _fit(src, base_w, base_h, pad_ratio=pad).save(OUT / f"{name}.png")
        written += 1

    for size in TARGET_SIZES:
        icon = _fit(src, size, size, pad_ratio=0.05)
        icon.save(OUT / f"Square44x44Logo.targetsize-{size}.png")
        icon.save(OUT / f"Square44x44Logo.targetsize-{size}_altform-unplated.png")
        plated = Image.new("RGBA", (size, size), PLATE_BG)
        plated.alpha_composite(icon)
        plated.save(OUT / f"Square44x44Logo.altform-lightunplated_targetsize-{size}.png")
        written += 3

    # 1240x600 Store listing / promotional source image.
    _fit(src, 1240, 600, pad_ratio=0.25).save(OUT / "PromoImage1240x600.png")
    written += 1

    print(f"Wrote {written} asset files to {OUT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
