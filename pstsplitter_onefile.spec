# One-file PyInstaller spec for PST Splitter
# Build with:  pyinstaller pstsplitter_onefile.spec
# Result: a single self-extracting EXE (slower first start, extracts to temp)

import sys
from pathlib import Path
from PyInstaller.utils.hooks import collect_submodules

project_root = Path.cwd()
src_dir = project_root / "src"
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))

# Broad hidden imports (can be trimmed later for smaller size)
hidden = collect_submodules("win32com") + [
    "pythoncom",
    "pywintypes",
]
try:
    hidden += collect_submodules("win32timezone")
except Exception:
    pass

block_cipher = None

a = Analysis(
    ["src/launch_pstsplitter.py"],
    pathex=[str(src_dir)],
    binaries=[],
    datas=[],
    hiddenimports=hidden,
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=1,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="PSTSplitterOneFile",
    debug=False,
    strip=False,
    upx=True,
    console=False,
    version="version_info.py" if Path("version_info.py").is_file() else None,
    icon=None,  # Add icon path here if you have one
)
