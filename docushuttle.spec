# -*- mode: python ; coding: utf-8 -*-


import os
import sys
import glob

# Find VC runtime DLLs dynamically
vcruntime_dlls = []
search_paths = [
    os.path.join(os.environ.get('SYSTEMROOT', r'C:\Windows'), 'System32'),
    os.path.dirname(sys.executable),
    os.path.join(os.path.dirname(sys.executable), 'DLLs'),
]
for search_dir in search_paths:
    for dll_name in ['vcruntime140.dll', 'vcruntime140_1.dll']:
        dll_path = os.path.join(search_dir, dll_name)
        if os.path.exists(dll_path) and dll_path not in [d[0] for d in vcruntime_dlls]:
            vcruntime_dlls.append((dll_path, '.'))

a = Analysis(
    ['docushuttle.py'],
    pathex=[],
    binaries=vcruntime_dlls,
    datas=[('myicon.ico', '.'), ('myicon.png', '.')],
    hiddenimports=['win32timezone'],
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
    name='DocuShuttle',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['myicon.ico'],
)
