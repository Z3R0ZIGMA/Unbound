# -*- mode: python ; coding: utf-8 -*-
# How Unbound.exe is built. Run from this folder:
#
#     .venv\Scripts\pyinstaller --noconfirm --clean Unbound.spec
#
# The result is dist\Unbound.exe. GitHub Actions builds releases from this
# same file (.github/workflows/release.yml).

import os
import re

from PyInstaller.utils.win32.versioninfo import (
    FixedFileInfo, StringFileInfo, StringStruct, StringTable, VarFileInfo, VarStruct,
    VSVersionInfo,
)

# The version is written once, in unbound/__init__.py
with open(os.path.join(SPECPATH, "unbound", "__init__.py"), encoding="utf-8") as fh:
    VERSION = re.search(r'^__version__ = "([^"]+)"', fh.read(), re.M).group(1)
numbers = tuple(int(n) for n in VERSION.split("."))
numbers += (0,) * (4 - len(numbers))

# Shown under Properties > Details, and in Task Manager
version_info = VSVersionInfo(
    ffi=FixedFileInfo(filevers=numbers, prodvers=numbers, mask=0x3F, flags=0x0,
                      OS=0x40004, fileType=0x1, subtype=0x0, date=(0, 0)),
    kids=[
        StringFileInfo([StringTable("040904B0", [
            StringStruct("FileDescription", "Unbound"),
            StringStruct("FileVersion", VERSION),
            StringStruct("InternalName", "Unbound"),
            StringStruct("OriginalFilename", "Unbound.exe"),
            StringStruct("ProductName", "Unbound"),
            StringStruct("ProductVersion", VERSION),
        ])]),
        VarFileInfo([VarStruct("Translation", [0x0409, 1200])]),
    ],
)

a = Analysis(
    ['unbound_app.py'],
    pathex=[],
    binaries=[],
    datas=[('assets/icon.ico', 'assets')],
    # PyInstaller only bundles modules it can see being imported. pywin32
    # loads win32timezone on demand, the first time a date marked with a
    # time zone goes through COM; without it calendar items fail and emails
    # get the wrong date, although the code works from source.
    hiddenimports=['pythoncom', 'pywintypes', 'win32com.client', 'win32timezone'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter'],
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
    name='Unbound',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    # No UPX: packed executables are flagged far more often by virus
    # scanners, and packing can damage Qt's libraries
    upx=False,
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['assets/icon.ico'],
    version=version_info,
)
