# -*- mode: python ; coding: utf-8 -*-
import os
import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules, copy_metadata

streamlit_datas = collect_data_files("streamlit", include_py_files=True)

meta_datas = (
    copy_metadata("streamlit")
    + copy_metadata("pandas")
    + copy_metadata("openpyxl")
    + copy_metadata("altair")
    + copy_metadata("packaging")
    + copy_metadata("pystray")
    + copy_metadata("pillow")
)

app_datas = [
    ("automationtoolstreamlit19.py", "."),
]

all_datas = streamlit_datas + app_datas + meta_datas

hidden_imports = (
    collect_submodules("streamlit")
    + collect_submodules("altair")
    + collect_submodules("pystray")
    + collect_submodules("PIL")
    + collect_submodules("multiprocessing")
    + [
        "streamlit.web.cli",
        "streamlit.runtime.scriptrunner.magic_funcs",
        "streamlit.runtime.scriptrunner",
        "streamlit.web.server",
        "pandas",
        "openpyxl",
        "openpyxl.cell._cell",
        "difflib",
        "pkg_resources.py2_warn",
        "packaging",
        "packaging.version",
        "packaging.specifiers",
        "packaging.requirements",
        "pystray",
        "pystray._win32",
        "PIL",
        "PIL.Image",
        "PIL.ImageDraw",
        "multiprocessing",
        "multiprocessing.freeze_support",
        "socket",
        "urllib.request",
    ]
)

# ── Runtime hook: freeze_support must fire before launcher.py runs ────────────
# SPECPATH is PyInstaller's built-in variable for the .spec file's directory
_hook_dir = os.path.join(SPECPATH, "runtime_hooks")
os.makedirs(_hook_dir, exist_ok=True)
_hook_path = os.path.join(_hook_dir, "hook_freeze_support.py")
with open(_hook_path, "w") as f:
    f.write("import multiprocessing\nmultiprocessing.freeze_support()\n")
# ─────────────────────────────────────────────────────────────────────────────

block_cipher = None

a = Analysis(
    ["launcher.py"],
    pathex=[SPECPATH],            # SPECPATH instead of "."
    binaries=[],
    datas=all_datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[_hook_path],
    excludes=[
        "tkinter",
        "matplotlib",
        "scipy",
        "notebook",
        "IPython",
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
    [],
    exclude_binaries=True,
    name="GEP_PackageTool",
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
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[
        "python*.dll",
        "vcruntime*.dll",
        "_uuid*.pyd",
        "charset_normalizer*.pyd",
    ],
    name="GEP_PackageTool",
)
