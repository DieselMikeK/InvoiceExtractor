# -*- mode: python ; coding: utf-8 -*-

import os


updater_asset = os.path.join('dist', 'InvoiceExtractorUpdater.exe')
if not os.path.exists(updater_asset):
    raise SystemExit(
        "Missing dist\\InvoiceExtractorUpdater.exe. "
        "Build the updater helper first with .\\build_release.ps1 or "
        "python -m PyInstaller InvoiceExtractorUpdater.spec."
    )


a = Analysis(
    ['invoice_extractor_gui.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('header.png', '.'),
        ('logo.ico', '.'),
        ('VERSION', '.'),
        (updater_asset, 'update'),
    ],
    hiddenimports=[],
    hookspath=['hooks'],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'pandas',
        'pyarrow',
        'numpy',
        'lxml',
        # Optional stack from pandas hooks that is not used by this app.
        'matplotlib',
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='InvoiceExtractor',
    debug=False,
    icon='logo.ico',
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
