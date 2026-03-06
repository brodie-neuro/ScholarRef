# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

from PyInstaller.utils.hooks import copy_metadata, collect_data_files, collect_submodules


project_root = Path(SPECPATH)

datas = collect_data_files("customtkinter")
datas += copy_metadata("scholarref")
datas += [
    (str(project_root / "pyproject.toml"), "."),
    (str(project_root / "logo" / "logo.png"), "logo"),
    (str(project_root / "logo" / "logo_removebg.png"), "logo"),
    (str(project_root / "logo" / "scholarref-mark.ico"), "logo"),
]

hiddenimports = collect_submodules("customtkinter")
hiddenimports += collect_submodules("docx")
hiddenimports += [
    "darkdetect",
    "scholarref",
    "scholarref_runtime",
    "verify_reference_integrity",
]


a = Analysis(
    ["scholarref_gui.py"],
    pathex=[str(project_root)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="ScholarRef",
    icon=str(project_root / "logo" / "scholarref-mark.ico"),
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="ScholarRef",
)
