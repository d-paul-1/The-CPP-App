# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['src\\base_app.py'],
    pathex=[],
    binaries=[],
    datas=[('assets/cpp_heart_logo.png', 'assets'), ('assets/cpp_heart_logo.ico', 'assets'), ('assets/templates', 'assets/templates')],
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
    name='CPP APP',
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
    icon=['assets\\cpp_heart_logo.ico'],
)
