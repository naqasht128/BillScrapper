# -*- mode: python ; coding: utf-8 -*-
import os
import glob
import sys
import importlib.util
from PyInstaller.utils.hooks import collect_submodules

compiled_binaries = []
for pyd_path in glob.glob('*.pyd'):
    compiled_binaries.append((pyd_path, '.'))

entry_script = 'launcher.py' if os.path.exists('launcher.py') else 'main.py'

# Bundle only the lightweight 13MB default OCR model (common_old.onnx), omitting the 52MB beta model (common.onnx) and 19MB detection model (common_det.onnx)
ddddocr_datas = []
try:
    dd_spec = importlib.util.find_spec('ddddocr')
    if dd_spec and dd_spec.submodule_search_locations:
        dd_path = list(dd_spec.submodule_search_locations)[0]
        for fname in ['common_old.onnx', 'charsets.py', 'logo.png']:
            fpath = os.path.join(dd_path, fname)
            if os.path.exists(fpath):
                ddddocr_datas.append((fpath, 'ddddocr'))
except Exception:
    pass

hidden_imports = ['openpyxl', 'bs4', 'requests', 'lxml', 'main', 'ddddocr', 'onnxruntime']
hidden_imports += collect_submodules('ddddocr')

a = Analysis(
    [entry_script],
    pathex=[],
    binaries=compiled_binaries,
    datas=[('assets', 'assets')] + ddddocr_datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=['rthook_cv2.py'],
    excludes=['cv2', 'opencv'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe_icon = ['assets/billscrapper.ico'] if sys.platform == 'win32' else None

if sys.platform == 'win32':
    exe = EXE(
        pyz,
        a.scripts,
        [],
        exclude_binaries=True,
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
    coll = COLLECT(
        exe,
        a.binaries,
        a.datas,
        strip=False,
        upx=True,
        upx_exclude=[],
        name='BillScrapper',
    )
else:
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
        console=True,
        disable_windowed_traceback=False,
        argv_emulation=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
        icon=exe_icon,
    )
