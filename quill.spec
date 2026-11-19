# -*- mode: python ; coding: utf-8 -*-
# Build with:  pyinstaller quill.spec --noconfirm --clean     (or just run build_exe.bat)
from PyInstaller.utils.hooks import collect_data_files

datas = [("assets", "assets"), ("themes", "themes")] + collect_data_files("docx")

a = Analysis(
    ["quill.py"],
    pathex=["."],
    datas=datas,
    hiddenimports=["quill_paste", "quill_paste.parser", "quill_paste.mathconv", "quill_paste.apply",
                   "quill_themes", "quill_layout", "quill_layout_ui"],
    excludes=["tkinter", "scipy", "numpy", "PIL", "pytest", "unittest"],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, [],
    exclude_binaries=True,
    name="Quill",
    icon="assets/quill.ico",
    console=False,          # no black terminal window
    upx=False,
)
coll = COLLECT(exe, a.binaries, a.datas, name="Quill", upx=False)
