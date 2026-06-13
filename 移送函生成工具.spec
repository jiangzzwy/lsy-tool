# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['/Users/jiangzhongzhou/Desktop/lsy/app.py'],
    pathex=[],
    binaries=[],
    datas=[('/Users/jiangzhongzhou/Desktop/lsy/demands', 'demands')],
    hiddenimports=['dearpygui', 'openpyxl', 'docx', 'requests'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'matplotlib', 'numpy', 'pandas'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='移送函生成工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='移送函生成工具',
)
app = BUNDLE(
    coll,
    name='移送函生成工具.app',
    icon=None,
    bundle_identifier=None,
)
