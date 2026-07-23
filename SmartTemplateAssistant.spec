# -*- mode: python ; coding: utf-8 -*-
# Portable PyInstaller configuration for Smart Template Assistant.
from pathlib import Path

from PyInstaller.utils.hooks import collect_submodules
from PyInstaller.utils.hooks import collect_all

project_root = Path(SPECPATH)

datas = [
    (str(project_root / 'assets' / 'app_icon.png'), 'assets'),
]
binaries = []
hiddenimports = []
hiddenimports += collect_submodules('openpyxl')
tmp_ret = collect_all('psd_tools')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    [str(project_root / 'app.py')],
    pathex=[str(project_root)],
    binaries=binaries,
    datas=datas,
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
    name='智能套版助手',
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
    icon=str(project_root / 'assets' / 'app_icon.ico'),
)
