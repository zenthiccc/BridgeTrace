# PyInstaller spec for the BridgeTrace GUI.
#
#     pip install -r requirements.txt -r requirements-build.txt
#     pyinstaller build_exe.spec --noconfirm --clean
#
# Produces a single windowed executable: dist/BridgeTrace.exe
# customtkinter's themes and fonts are bundled by pyinstaller-hooks-contrib.

a = Analysis(
    ["src/gui.py"],
    pathex=["src"],
    hiddenimports=["xlrd", "openpyxl"],  # lazy-imported by pandas
    excludes=["pytest", "_pytest"],
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="BridgeTrace",
    upx=True,
    console=False,
)
