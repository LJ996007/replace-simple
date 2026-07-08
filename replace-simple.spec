# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files

datas = [('icon-256.png', '.')]
datas += collect_data_files('tksheet')


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=['openpyxl', 'xlrd', 'docx', 'pptx', 'tksheet'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['numpy', 'pandas', 'scipy', 'matplotlib', 'win32com', 'pythoncom', 'pywintypes', 'pywin32', 'win32evtlog', 'win32evtlogutil', 'win32api', 'bs4', 'charset_normalizer', 'soupsieve', 'pyreadline3', 'lxml.isoschematron', 'lxml.html', 'lxml.objectify', 'lxml.sax', 'pythonnet', 'clr_loader', 'clr', 'jinja2', 'yaml', 'PyYAML'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='replace-simple',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['icon.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='replace-simple',
)
