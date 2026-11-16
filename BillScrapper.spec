# -*- mode: python ; coding: utf-8 -*-
import os
import glob
import sys

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

compiled_binaries = []
for pyd_path in glob.glob('*.pyd'):
    compiled_binaries.append((pyd_path, '.'))

entry_script = 'launcher.py' if os.path.exists('launcher.py') else 'main.py'

ddddocr_datas = []
try:
    ddddocr_datas = collect_data_files('ddddocr')
except Exception:
    pass

a = Analysis(
    [entry_script],
    pathex=[],
    binaries=compiled_binaries,
    datas=[('assets', 'assets')] + ddddocr_datas,
    hiddenimports=['openpyxl', 'bs4', 'requests', 'lxml', 'main', 'ddddocr', 'onnxruntime'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe_icon = ['assets/billscrapper.ico'] if sys.platform == 'win32' else None

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='BillScrapper',
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
    icon=exe_icon,
)
