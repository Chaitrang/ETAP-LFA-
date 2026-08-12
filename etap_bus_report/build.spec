# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller specification.

    pip install pyinstaller
    pyinstaller build.spec          # -> dist/ETAP Report Filler(.exe)

Both Word templates are bundled as data and located at runtime through
utils.resource_path(), so the user never has to supply one.
"""

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('assets/LoadFlow_Template.docx', 'assets'),
        ('assets/ShortCircuit_Template.docx', 'assets'),
        ('assets/Bus_Template.docx', 'assets'),      # kept for older builds
    ],
    hiddenimports=[
        'pdfminer.pdfinterp', 'pdfminer.converter', 'pdfminer.layout',
        'openpyxl.cell._writer',      # pulled in dynamically by openpyxl
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=['tkinter', 'matplotlib', 'PySide6.QtWebEngineCore'],
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
    name='ETAP Report Filler',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    runtime_tmpdir=None,
    console=False,          # windowed application
    icon=None,              # put an .ico here if you have one
)
