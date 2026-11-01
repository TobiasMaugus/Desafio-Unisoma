# Build with: ..\.venv\Scripts\python.exe -m PyInstaller --noconfirm ABC_Aprendiz.spec
# EXE receives binaries/datas directly: ONEFILE, with no COLLECT/ONEDIR stage.
import os
from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files, collect_submodules, copy_metadata
import pulp

root = Path(SPECPATH)
datas = []
for name in ("streamlit", "st_aggrid"):
    datas += collect_data_files(name)
for name in ("streamlit", "streamlit-aggrid", "pulp", "openpyxl", "pandas"):
    datas += copy_metadata(name, recursive=True)
for name in (".streamlit", "assets", "static"):
    datas.append((str(root / name), name))
for name in ("streamlit_app.py", "app.py", "pdf_report.py", "abc-logo-1024x798.webp",
             "abc1024x279.webp", "MANUAL_USUARIO.md", "MANUAL_USUARIO.pdf"):
    datas.append((str(root / name), "."))
cbc = Path(pulp.PULP_CBC_CMD().path)
if not cbc.is_file():
    raise RuntimeError(f"CBC não encontrado: {cbc}")
datas += collect_data_files("pulp", includes=["pulp.cfg.*", "solverdir/cbc/win/64/*"], excludes=["**/cbc.exe"])

a = Analysis(
    [str(root / "launcher.py")],
    pathex=[str(root)],
    binaries=[(str(cbc), "pulp/solverdir/cbc/win/64")],
    datas=datas,
    hiddenimports=["app", "pdf_report", "streamlit_app"]
        + collect_submodules("streamlit", filter=lambda name: not name.startswith("streamlit.testing"))
        + collect_submodules("st_aggrid"),
    hookspath=[],
    runtime_hooks=[],
    excludes=["pytest", "playwright", "IPython", "notebook"],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, a.binaries, a.datas, [],
    name="ABC_Aprendiz",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=os.environ.get("ABC_BUILD_CONSOLE") == "1",
    disable_windowed_traceback=False,
)
