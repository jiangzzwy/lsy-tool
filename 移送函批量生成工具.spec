# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['/Users/jiangzhongzhou/Desktop/李双义需求/web.py'],
    pathex=[],
    binaries=[],
    datas=[('/Users/jiangzhongzhou/Desktop/李双义需求/demands', 'demands'), ('/Users/jiangzhongzhou/Desktop/李双义需求/web/templates', 'web/templates')],
    hiddenimports=['flask', 'webview', 'openpyxl', 'docx', 'requests', 'api_client', 'config', 'excel_parser', 'word_generator', 'ledger_generator'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'dearpygui', 'matplotlib', 'numpy', 'pandas'],
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
    name='移送函批量生成工具',
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
)
app = BUNDLE(
    exe,
    name='移送函批量生成工具.app',
    icon=None,
    bundle_identifier=None,
)
