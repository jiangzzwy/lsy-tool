# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec file for 案件线索移送函批量生成工具
# Build on Windows:  pyinstaller build.spec
# Build on macOS:    pyinstaller build.spec

import sys
from pathlib import Path

block_cipher = None

base_dir = Path(SPECPATH)

# Web templates
web_templates = str(base_dir / 'web' / 'templates')
web_static = str(base_dir / 'web' / 'static')

# Demands directory (templates + sample Excel)
demands_dir = str(base_dir / 'demands')

a = Analysis(
    [str(base_dir / 'web.py')],
    pathex=[str(base_dir)],
    binaries=[],
    datas=[
        (demands_dir, 'demands'),
        (web_templates, 'web/templates'),
        (web_static, 'web/static'),
    ],
    hiddenimports=[
        'flask',
        'webview',
        'openpyxl',
        'docx',
        'requests',
        'api_client',
        'config',
        'excel_parser',
        'word_generator',
        'ledger_generator',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'dearpygui',
        'matplotlib',
        'numpy',
        'pandas',
        'scipy',
        'PIL',
        'pytest',
        'selenium',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
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
    icon=None,
)
