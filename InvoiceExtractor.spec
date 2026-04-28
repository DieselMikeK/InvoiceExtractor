# -*- mode: python ; coding: utf-8 -*-

import os
import selenium


selenium_manager = os.path.join(
    os.path.dirname(selenium.__file__),
    'webdriver',
    'common',
    'windows',
    'selenium-manager.exe',
)
selenium_binaries = []
if os.path.exists(selenium_manager):
    selenium_binaries.append(
        (selenium_manager, os.path.join('selenium', 'webdriver', 'common', 'windows'))
    )


updater_asset = os.path.join('dist', 'InvoiceExtractorUpdater.exe')
if not os.path.exists(updater_asset):
    raise SystemExit(
        "Missing dist\\InvoiceExtractorUpdater.exe. "
        "Build the updater helper first with .\\build_release.ps1 or "
        "python -m PyInstaller InvoiceExtractorUpdater.spec."
    )


a = Analysis(
    ['invoice_extractor_gui.py'],
    pathex=[os.path.abspath('.')],
    binaries=selenium_binaries,
    datas=[
        ('header.png', '.'),
        ('logo.ico', '.'),
        ('VERSION', '.'),
        ('vendors.csv', '.'),
        (updater_asset, 'update'),
    ],
    hiddenimports=[
        'core_detection',
        'gmail_client',
        'invoice_parser',
        'shopify_client',
        'skunexus_client',
        'spreadsheet_writer',
        'update_utils',
    ],
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
