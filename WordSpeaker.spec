# -*- mode: python ; coding: utf-8 -*-
import sys
from pathlib import Path

from PyInstaller.utils.hooks import collect_all


ROOT = Path(globals().get("SPECPATH", ".")).resolve()
VENDOR = ROOT / "vendor" / "site-packages"

for path_entry in (str(VENDOR), str(ROOT)):
    if path_entry in sys.path:
        sys.path.remove(path_entry)
    sys.path.insert(0, path_entry)

pathex = [
    str(VENDOR),
    str(ROOT),
]

datas = [
    (str(ROOT / "data" / "models"), "data/models"),
    (str(ROOT / "data" / "nltk_data"), "data/nltk_data"),
]
binaries = []
hiddenimports = []

for package_name in (
    "argostranslate",
    "kokoro_onnx",
    "nltk",
    "onnxruntime",
    "numpy",
    "phonemizer",
    "espeakng_loader",
    "piper",
    "spacy",
    "spacy_wordnet",
):
    try:
        pkg_datas, pkg_bins, pkg_hidden = collect_all(package_name)
    except Exception:
        continue
    datas += pkg_datas
    binaries += pkg_bins
    hiddenimports += pkg_hidden

hiddenimports += [
    "docx",
    "fitz",
    "pkg_resources._vendor.appdirs",
    "pkg_resources._vendor.packaging",
    "pkg_resources._vendor.pyparsing",
]

a = Analysis(
    ["app.py"],
    pathex=pathex,
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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
    [],
    exclude_binaries=True,
    name="WordSpeaker",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
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
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="WordSpeaker",
)
