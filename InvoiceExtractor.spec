# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['invoice_extractor_gui.py'],
    pathex=[],
    binaries=[],
    datas=[('header.png', '.'), ('logo.ico', '.'), ('update/updater.py', 'update')],
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
