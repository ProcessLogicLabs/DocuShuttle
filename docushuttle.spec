# -*- mode: python ; coding: utf-8 -*-

import os
import sys

# Collect VC runtime and Python DLLs dynamically
extra_binaries = []
search_paths = [
    os.path.join(os.environ.get('SYSTEMROOT', r'C:\Windows'), 'System32'),
    os.path.dirname(sys.executable),
    os.path.join(os.path.dirname(sys.executable), 'DLLs'),
]
dll_names = [
    'vcruntime140.dll',
    'vcruntime140_1.dll',
    'python3.dll',
    'msvcp140.dll',
]
collected = set()
for search_dir in search_paths:
    for dll_name in dll_names:
        if dll_name in collected:
            continue
        dll_path = os.path.join(search_dir, dll_name)
        if os.path.exists(dll_path):
            extra_binaries.append((dll_path, '.'))
            collected.add(dll_name)

a = Analysis(
    ['docushuttle.py'],
    pathex=[],
    binaries=extra_binaries,
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
