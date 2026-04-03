# -*- mode: python ; coding: utf-8 -*-
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
    + collect_submodules("multiprocessing")   # ← needed for freeze_support
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
        "pystray._win32",                     # ← Windows tray backend
        "PIL",
        "PIL.Image",
        "PIL.ImageDraw",
        "multiprocessing",
        "multiprocessing.freeze_support",     # ← prevents multiple instances
        "socket",
        "urllib.request",
    ]
)

# ── Runtime hook: ensures freeze_support() is called before anything else ────
# Write the hook inline so no extra file is needed
import os, textwrap, tempfile

_hook_content = textwrap.dedent("""\
    import multiprocessing
    multiprocessing.freeze_support()
""")
_hook_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "runtime_hooks")
os.makedirs(_hook_dir, exist_ok=True)
_hook_path = os.path.join(_hook_dir, "hook_freeze_support.py")
with open(_hook_path, "w") as f:
    f.write(_hook_content)
# ─────────────────────────────────────────────────────────────────────────────

block_cipher = None

a = Analysis(
    ["launcher.py"],
    pathex=["."],
    binaries=[],
    datas=all_datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[_hook_path],   # ← freeze_support fires before launcher.py
    excludes=[
        "tkinter",                # not needed, saves space
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
        "python*.dll",            # don't UPX Python DLLs — causes crashes
        "vcruntime*.dll",
        "_uuid*.pyd",
        "charset_normalizer*.pyd",  # was causing your WinError 5 earlier
    ],
    name="GEP_PackageTool",
)
