# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_submodules


python_root = r'C:\Users\Admin\AppData\Local\Programs\Python\Python310'
tkinter_runtime_hook = python_root + r'\Lib\site-packages\PyInstaller\hooks\rthooks\pyi_rth__tkinter.py'

a = Analysis(
    ['xoadon_fix.py'],
    pathex=[],
    binaries=[
        (python_root + r'\DLLs\_tkinter.pyd', '.'),
        (python_root + r'\DLLs\tcl86t.dll', '.'),
        (python_root + r'\DLLs\tk86t.dll', '.'),
    ],
    datas=[
        (python_root + r'\Lib\tkinter', 'tkinter'),
        (python_root + r'\tcl\tcl8.6', '_tcl_data'),
        (python_root + r'\tcl\tk8.6', '_tk_data'),
    ],
    hiddenimports=['_tkinter'] + collect_submodules('selenium.webdriver'),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[tkinter_runtime_hook],
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
    name='xoadon_2804',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
