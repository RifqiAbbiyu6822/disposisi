# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['coba.py'],
    pathex=[],
    binaries=[],
    datas=[('credentials', 'credentials'), ('kop.jpg', '.'), ('assets', 'assets'), ('disposisi_app', 'disposisi_app'), ('email_sender', 'email_sender'), ('logic', 'logic'), ('main_app', 'main_app'), ('constants.py', '.')],
    hiddenimports=['tkinter', 'tkcalendar', 'reportlab', 'PyPDF2', 'gspread'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name='SistemDisposisi',
    debug=False,
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
    uac_admin=True,
    icon=['assets\\JapekELEVATED.ico'],
)
