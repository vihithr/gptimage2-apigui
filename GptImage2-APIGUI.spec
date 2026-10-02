# -*- mode: python ; coding: utf-8 -*-

import glob
import os
import sys

# A venv created on top of an Anaconda install keeps `_tkinter.pyd` in the base
# install's `DLLs`, while the Tcl/Tk DLLs it links against live in `<base>/Library/bin`.
# PyInstaller adds that folder to its search path only when `conda-meta` sits next to
# `sys.prefix`, so in such a venv both DLLs resolve to None and get silently dropped --
# producing an exe that dies at `import _tkinter`. Collect them here; a non-conda Python
# has no `Library/bin`, so this contributes nothing there.
_tcl_tk_binaries = [
    (path, '.')
    for pattern in ('tcl*.dll', 'tk*.dll')
    for path in glob.glob(os.path.join(sys.base_prefix, 'Library', 'bin', pattern))
]

a = Analysis(
    ['GptImage2-APIGUI.py'],
    pathex=[],
    binaries=_tcl_tk_binaries,
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
    name='GptImage2-APIGUI',
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
