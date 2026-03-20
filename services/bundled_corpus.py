# -*- coding: utf-8 -*-
import json
import os
import threading
from pathlib import Path

from services.corpus_search import corpus_stats, import_corpus_files, remove_document


_LOCK = threading.Lock()
_PREPARE_STARTED = False
_STATE_CACHE = None

_BASE_DIR = Path(__file__).resolve().parent.parent
_BUNDLED_CORPUS_DIR = _BASE_DIR / "data" / "bundled_corpus"
_STATE_PATH = _BASE_DIR / "data" / "bundled_corpus_state.json"
_SUPPORTED_SUFFIXES = {".txt", ".docx", ".pdf"}


def _iter_bundled_files():
    if not _BUNDLED_CORPUS_DIR.exists():
        return []
    files = []
    for path in _BUNDLED_CORPUS_DIR.rglob("*"):
        if not path.is_file():
            continue
        if path.suffix.lower() not in _SUPPORTED_SUFFIXES:
            continue
        files.append(path.resolve())
    return sorted(files)


def _file_stamp(path):
    stat = path.stat()
    return {
        "size": int(stat.st_size or 0),
        "mtime_ns": int(getattr(stat, "st_mtime_ns", int(stat.st_mtime * 1_000_000_000))),
    }


def _load_state_locked():
    global _STATE_CACHE
    if _STATE_CACHE is not None:
        return _STATE_CACHE
    if not _STATE_PATH.exists():
        _STATE_CACHE = {"files": {}}
        return _STATE_CACHE
    try:
        with open(_STATE_PATH, "r", encoding="utf-8") as fp:
            payload = json.load(fp)
    except Exception:
        payload = {}
    if not isinstance(payload, dict):
        payload = {}
    files = payload.get("files")
    if not isinstance(files, dict):
        files = {}
    _STATE_CACHE = {"files": files}
    return _STATE_CACHE


def _save_state_locked():
    _STATE_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(_STATE_PATH, "w", encoding="utf-8") as fp:
        json.dump(_STATE_CACHE or {"files": {}}, fp, ensure_ascii=False, indent=2)


def ensure_bundled_corpus_imported():
    bundled_files = _iter_bundled_files()
    if not bundled_files:
        return {"imported": 0, "available": 0}

    try:
        stats = corpus_stats()
        force_import = int(stats.get("documents", 0) or 0) <= 0
    except Exception:
        force_import = True

    with _LOCK:
        state = _load_state_locked()
        known_files = dict(state.get("files") or {})
        current_files = {}
        to_import = []
        removed_paths = [path for path in known_files.keys() if path not in {str(item) for item in bundled_files}]
        for path in bundled_files:
            stamp = _file_stamp(path)
            path_key = str(path)
            current_files[path_key] = stamp
            if force_import or known_files.get(path_key) != stamp:
                to_import.append(path_key)

        if not to_import and not removed_paths:
            return {"imported": 0, "available": len(bundled_files)}

    for path in removed_paths:
        try:
            remove_document(path)
        except Exception:
            pass

    result = import_corpus_files(to_import)

    with _LOCK:
        state = _load_state_locked()
        state["files"] = current_files
        _save_state_locked()

    return {
        "imported": int(result.get("files", 0) or 0),
        "available": len(bundled_files),
        "errors": list(result.get("errors") or []),
    }


def import_bundled_corpus_package(package_path):
    source_path = Path(str(package_path or "").strip())
    if not source_path.exists() or not source_path.is_file():
        raise FileNotFoundError("Bundled corpus package not found.")

    import zipfile

    extracted = []
    with zipfile.ZipFile(source_path, "r") as zf:
        members = [item for item in zf.infolist() if not item.is_dir()]
        valid_members = []
        for member in members:
            relative = Path(str(member.filename or "").replace("\\", "/").strip("/"))
            if not str(relative):
                continue
            if relative.name.startswith("."):
                continue
            if relative.suffix.lower() not in _SUPPORTED_SUFFIXES:
                continue
            valid_members.append((member, relative))
        if not valid_members:
            raise RuntimeError("Bundled corpus package does not contain any supported files.")

        _BUNDLED_CORPUS_DIR.mkdir(parents=True, exist_ok=True)
        for path in list(_BUNDLED_CORPUS_DIR.rglob("*")):
            if not path.is_file():
                continue
            if path.suffix.lower() not in _SUPPORTED_SUFFIXES:
                continue
            try:
                path.unlink()
            except Exception:
                pass

        for member, relative in valid_members:
            target_path = (_BUNDLED_CORPUS_DIR / relative).resolve()
            if _BUNDLED_CORPUS_DIR.resolve() not in target_path.parents and target_path != _BUNDLED_CORPUS_DIR.resolve():
                continue
            target_path.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(member, "r") as src, open(target_path, "wb") as dst:
                dst.write(src.read())
            extracted.append(str(target_path))

    import_result = ensure_bundled_corpus_imported()
    import_result["files"] = len(extracted)
    return import_result


def prepare_async():
    global _PREPARE_STARTED
    with _LOCK:
        if _PREPARE_STARTED:
            return
        _PREPARE_STARTED = True

    def _run():
        global _PREPARE_STARTED
        try:
            ensure_bundled_corpus_imported()
        finally:
            with _LOCK:
                _PREPARE_STARTED = False

    threading.Thread(target=_run, daemon=True).start()
