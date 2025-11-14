# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = ['win32com', 'win32timezone', 'app', 'app.main', 'app.services.excel_copy']
hiddenimports += collect_submodules('win32com')
hiddenimports += collect_submodules('pythoncom')


a = Analysis(
    ['D:\\cobranza_xls_app-13-10-25\\run_app.py'],
    pathex=[],
    binaries=[],
    datas=[('app\\static', 'app\\static'), ('app\\data', 'app\\data')],
    hiddenimports=hiddenimports,
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
    name='CobranzaApp',
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
)
