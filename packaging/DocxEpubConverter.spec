# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for the Microsoft Store (MSIX) build.

Build with:
    pyinstaller packaging/DocxEpubConverter.spec --noconfirm --clean

Produces dist/DocxEpubConverter/ as a one-folder, windowed (no console) app.
One-folder is deliberate: MSIX must be able to sign and inventory the files,
and a one-file build unpacks to temp on every launch, which Store certification
flags for slow start-up.
"""

import os
from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files

SPEC_DIR = Path(os.path.abspath(SPECPATH))
ROOT = SPEC_DIR.parent

datas = []

# Optional drag-and-drop support: ship the tkdnd binaries when installed.
try:
    datas += collect_data_files("tkinterdnd2")
except Exception:
    pass

# Optional EpubCheck jar: only useful on machines that have a Java runtime,
# but bundling it means validation works out of the box where Java exists.
try:
    datas += collect_data_files("epubcheck")
except Exception:
    pass

a = Analysis(
    [str(ROOT / "run_converter.py")],
    pathex=[str(ROOT / "src")],
    binaries=[],
    datas=datas,
    hiddenimports=["PIL._tkinter_finder"],
    hookspath=[],
    runtime_hooks=[],
    excludes=["pytest", "numpy", "matplotlib", "setuptools", "pip"],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="DocxEpubConverter",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(ROOT / "assets" / "app_icon.ico"),
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="DocxEpubConverter",
)
