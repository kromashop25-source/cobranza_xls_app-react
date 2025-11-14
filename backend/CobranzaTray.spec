# -*- mode: python ; coding: utf-8 -*-
import sys
from pathlib import Path
from PyInstaller.utils.hooks import collect_submodules

# ── Rutas robustas aunque __file__ no exista ──────────────────────────────────
try:
    SPEC_DIR = Path(__file__).resolve().parent
except NameError:
    CWD = Path.cwd()
    SPEC_DIR = CWD / "backend" if (CWD / "backend").is_dir() else CWD
PROJ_ROOT = SPEC_DIR.parent

block_cipher = None

# ── Datos que deben ir dentro del bundle (si existen) ─────────────────────────
datas = []
for src, dst in [
    (SPEC_DIR / "app" / "static",           "app/static"),
    (SPEC_DIR / "app" / "data",             "app/data"),
    (PROJ_ROOT / "frontend" / "dist",       "frontend/dist"),
]:
    if src.exists():
        datas.append((str(src), dst))

# Pydantic v2 (y core) bajo PyInstaller + win32timezone para pywin32
hidden = (
    collect_submodules("pydantic")
    + collect_submodules("pydantic_core")
    + ["win32timezone"]
)

a = Analysis(
    [str(SPEC_DIR / "run_tray.py")],   # <-- el script está junto al .spec
    pathex=[str(PROJ_ROOT)],           # para resolver imports "backend.*"
    binaries=[],
    datas=datas,
    hiddenimports=hidden,
    hookspath=[],
    plugins=[],
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# ── EXE (ONEDIR) ───────────────────────────────────────────────────────────────
# Importante: primero argumentos posicionales (pyz, a.scripts), luego keywords.
exe = EXE(
    pyz,
    a.scripts,
    exclude_binaries=True,
    name="CobranzaTray",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # bandeja del sistema (sin consola)
    icon=str(SPEC_DIR / "app" / "static" / "cobranza.ico"),
)

# ── COLLECT: empaqueta todo en dist/CobranzaTray ──────────────────────────────
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    name="CobranzaTray",
)
