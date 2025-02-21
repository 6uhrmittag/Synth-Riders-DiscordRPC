# -*- mode: python ; coding: utf-8 -*-


rpc = Analysis(
    ['src/bin/rpc.py'],
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

rpc_pyz = PYZ(rpc.pure)

rpc_exe = EXE(
    rpc_pyz,
    rpc.scripts,
    rpc.binaries,
    rpc.datas,
    [],
    uac_admin=True,
    name='Synth Riders RPC',
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
    icon=['assets\\logo.ico'],
)

