# -*- mode: python ; coding: utf-8 -*-
import sys
from pathlib import Path

from PyInstaller.utils.hooks import (
    collect_all,
    collect_data_files,
    collect_dynamic_libs,
    collect_submodules,
)


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
    (str(ROOT / "version.json"), "."),
]
binaries = []
hiddenimports = [
    "docx",
    "fitz",
    "argostranslate.package",
    "argostranslate.translate",
    "spacy_wordnet.wordnet_annotator",
]

for package_name in ("piper", "phonemizer", "espeakng_loader"):
    try:
        pkg_datas, pkg_bins, pkg_hidden = collect_all(package_name)
    except Exception:
        continue
    datas += pkg_datas
    binaries += pkg_bins
    hiddenimports += pkg_hidden

for package_name in ("argostranslate", "spacy", "spacy_wordnet"):
    try:
        datas += collect_data_files(package_name)
    except Exception:
        continue

try:
    binaries += collect_dynamic_libs("onnxruntime")
except Exception:
    pass

try:
    hiddenimports += collect_submodules("spacy_wordnet")
except Exception:
    pass

a = Analysis(
    ["app.py"],
    pathex=pathex,
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "torch",
        "functorch",
        "tensorflow",
        "tensorboard",
        "mxnet",
        "cupy",
        "cupyx",
        "matplotlib",
        "mpl_toolkits",
        "IPython",
        "ipykernel",
        "jupyter_client",
        "jupyter_core",
        "zmq",
        "pyzmq",
        "nbconvert",
        "nbformat",
        "jsonschema",
        "jsonschema_specifications",
        "referencing",
        "rpds",
        "numba",
        "llvmlite",
        "scipy",
        "setuptools",
        "pkg_resources",
        "spacy.tests",
        "spacy.cli.package",
        "nltk.test",
        "nltk.book",
        "nltk.chat",
        "onnxruntime.tools",
        "argostranslate.cli",
    ],
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
