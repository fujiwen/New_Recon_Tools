# -*- mode: python ; coding: utf-8 -*-
import sys
import os

# 从MC_Recon_UI.py中导入VERSION
sys.path.append(os.path.dirname(os.path.abspath('MC_Recon_UI.py')))
from MC_Recon_UI import VERSION

a = Analysis(
    ['MC_Recon_UI.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
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
    name=f'MC对账明细工具_v{VERSION}',
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
    icon=os.path.join(os.path.dirname(os.path.abspath(SPEC)), 'favicon.ico'),
    version='file_version_info.txt',
)
