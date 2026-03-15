# -*- coding: utf-8 -*-
import base64
import hashlib
import json
import os
import re
import shutil
import tempfile
import threading
import time
import urllib.error
import urllib.request
import wave
import winsound

import numpy as np
from tkinter import messagebox

from services.app_config import get_gemini_api_key
from services.voice_catalog import (
    get_kokoro_paths,
    get_piper_voice_profile,
    get_voice_profile,
    kokoro_ready,
    piper_ready,
)
from services.voice_manager import SOURCE_GEMINI, SOURCE_KOKORO, SOURCE_PIPER, get_voice_id, get_voice_source


BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
AUDIO_CACHE_ROOT_DIR = os.path.join(BASE_DIR, "data", "audio_cache")
LEGACY_WORD_CACHE_WRAPPER_DIR = os.path.join(AUDIO_CACHE_ROOT_DIR, "words")
GLOBAL_WORD_CACHE_DIR = AUDIO_CACHE_ROOT_DIR
SHARED_WORD_CACHE_DIR = os.path.join(AUDIO_CACHE_ROOT_DIR, "global")
SOURCE_WORD_CACHE_ROOT_DIR = os.path.join(AUDIO_CACHE_ROOT_DIR, "sources")
PENDING_GEMINI_QUEUE_PATH = os.path.join(BASE_DIR, "data", "audio_cache", "pending_gemini_replacements.json")
RECENT_WRONG_SOURCE_KEY = "__recent_wrong_words__"
MANUAL_SESSION_SOURCE_KEY = "__manual_session__"
KOKORO_SAMPLE_RATE = 24000

_lock = threading.Lock()
_token = 0
_shown_errors = set()
_current_wav = None
_kokoro = None
_kokoro_lock = threading.Lock()
_piper_voices = {}
_piper_lock = threading.Lock()
_backend_lock = threading.Lock()
_backend_status = {}
_pending_gemini_replacements = {}
_pending_gemini_lock = threading.Lock()
_pending_gemini_worker_running = False
_manual_session_cache_paths = set()
_manual_session_cache_lock = threading.Lock()
_cache_metadata_memory = {}
_cache_metadata_lock = threading.Lock()

TTS_MODEL = "gemini-2.5-flash-preview-tts"
TTS_URL = f"https://generativelanguage.googleapis.com/v1beta/models/{TTS_MODEL}:generateContent"
TTS_SAMPLE_RATE = 24000
TTS_CHANNELS = 1
TTS_SAMPLE_WIDTH = 2
TTS_VOICE_NAME = "Kore"
TTS_STYLE_SHORT = "Read the word clearly in a neutral British English accent for IELTS vocabulary practice."
TTS_STYLE_LONG = "Read this passage clearly in a natural British English accent for IELTS listening practice."
GEMINI_RATE_LIMIT_COOLDOWN_SECONDS = 25


def _clamp(value, low, high):
    return max(low, min(high, float(value)))


def _normalize_text(text, ensure_sentence_end=False):
    raw = re.sub(r"\s+", " ", str(text or "").strip())
    if ensure_sentence_end and raw and raw[-1] not in ".!?;:":
        raw = f"{raw}."
    return raw


def _safe_name(text, limit=40):
    raw = re.sub(r"[^A-Za-z0-9._-]+", "_", str(text or "").strip()).strip("._-")
    return (raw or "audio")[:limit]


def _normalize_source_path(source_path):
    raw = str(source_path or "").strip()
    if not raw:
        return None
    if raw in {RECENT_WRONG_SOURCE_KEY, MANUAL_SESSION_SOURCE_KEY}:
        return raw
    return os.path.abspath(raw)


def _source_bucket_name(source_path=None):
    normalized = _normalize_source_path(source_path)
    if not normalized:
        return MANUAL_SESSION_SOURCE_KEY
    if normalized == RECENT_WRONG_SOURCE_KEY:
        return "recent_wrong"
    if normalized == MANUAL_SESSION_SOURCE_KEY:
        return MANUAL_SESSION_SOURCE_KEY
    base = os.path.splitext(os.path.basename(normalized))[0]
    label = _safe_name(base, limit=24) or "source"
    suffix = hashlib.sha1(normalized.casefold().encode("utf-8")).hexdigest()[:10]
    return f"{label}_{suffix}"


def _source_word_cache_dir(text=None, source_path=None):
    base_dir = os.path.join(SOURCE_WORD_CACHE_ROOT_DIR, _source_bucket_name(source_path))
    if text is None:
        return base_dir
    return os.path.join(base_dir, _shared_letter_bucket(text))


def _shared_letter_bucket(text):
    normalized = _normalize_text(text, ensure_sentence_end=False).casefold()
    for ch in normalized:
        if "a" <= ch <= "z":
            return ch
    return "other"


def _shared_word_cache_dir(text):
    return os.path.join(SHARED_WORD_CACHE_DIR, _shared_letter_bucket(text))


def _filename_letter_bucket(name):
    raw = str(name or "").strip().casefold()
    for ch in raw:
        if "a" <= ch <= "z":
            return ch
    return "other"


def _flat_compat_word_cache_path(text, source_path=None):
    source_key = str(_normalize_source_path(source_path) or "").casefold()
    normalized = _normalize_text(text, ensure_sentence_end=False).casefold()
    key = hashlib.sha1(
        f"gemini|{source_key}|{TTS_MODEL}|{TTS_VOICE_NAME}|{normalized}".encode("utf-8")
    ).hexdigest()[:16]
    name = _safe_name(normalized)
    return os.path.join(GLOBAL_WORD_CACHE_DIR, f"{name}_{key}.wav")


def _source_flat_compat_word_cache_path(text, source_path=None):
    source_key = str(_normalize_source_path(source_path) or "").casefold()
    normalized = _normalize_text(text, ensure_sentence_end=False).casefold()
    key = hashlib.sha1(
        f"gemini|{source_key}|{TTS_MODEL}|{TTS_VOICE_NAME}|{normalized}".encode("utf-8")
    ).hexdigest()[:16]
    name = _safe_name(normalized)
    return os.path.join(os.path.join(SOURCE_WORD_CACHE_ROOT_DIR, _source_bucket_name(source_path)), f"{name}_{key}.wav")


def _legacy_word_cache_path(text, source_path=None):
    source = str(source_path or "").strip()
    if not (source and os.path.isfile(source)):
        return None
    source_name = os.path.basename(source)
    legacy_dir = os.path.join(os.path.dirname(source), f".{source_name}.wordspeaker_audio")
    source_key = get_voice_source()
    voice_id = get_voice_id()
    normalized = _normalize_text(text, ensure_sentence_end=False).casefold()
    key = hashlib.sha1(
        f"{source_key}|{voice_id}|{TTS_MODEL}|{TTS_VOICE_NAME}|{normalized}".encode("utf-8")
    ).hexdigest()[:16]
    name = _safe_name(normalized)
    return os.path.join(legacy_dir, f"{name}_{key}.wav")


def _word_cache_path(text, source_path=None):
    source_key = str(_normalize_source_path(source_path) or "").casefold()
    normalized = _normalize_text(text, ensure_sentence_end=False).casefold()
    key = hashlib.sha1(
        f"gemini|{source_key}|{TTS_MODEL}|{TTS_VOICE_NAME}|{normalized}".encode("utf-8")
    ).hexdigest()[:16]
    name = _safe_name(normalized)
    return os.path.join(_source_word_cache_dir(normalized, source_path=source_path), f"{name}_{key}.wav")


def _shared_word_cache_path(text):
    normalized = _normalize_text(text, ensure_sentence_end=False).casefold()
    key = hashlib.sha1(
        f"gemini|shared|{TTS_MODEL}|{TTS_VOICE_NAME}|{normalized}".encode("utf-8")
    ).hexdigest()[:16]
    name = _safe_name(normalized)
    return os.path.join(_shared_word_cache_dir(normalized), f"{name}_{key}.wav")


def _is_recent_wrong_source(source_path=None):
    return _normalize_source_path(source_path) == RECENT_WRONG_SOURCE_KEY


def _is_lightweight_file_source(source_path=None):
    normalized = _normalize_source_path(source_path)
    return bool(normalized and normalized not in {RECENT_WRONG_SOURCE_KEY, MANUAL_SESSION_SOURCE_KEY})


def _cache_meta_path(cache_path):
    return f"{cache_path}.json"


def _canonicalize_cache_path(cache_path, *, text=None, source_path=None, metadata=None):
    raw = str(cache_path or "").strip()
    if not raw:
        return raw
    normalized = os.path.abspath(raw)
    legacy_wrapper = os.path.abspath(LEGACY_WORD_CACHE_WRAPPER_DIR)
    if normalized.startswith(legacy_wrapper + os.sep):
        rel_path = os.path.relpath(normalized, legacy_wrapper)
        parts = rel_path.split(os.sep)
        if parts and parts[0] == "sources":
            return os.path.join(SOURCE_WORD_CACHE_ROOT_DIR, *parts[1:])
        if parts and parts[0] == "global":
            return os.path.join(SHARED_WORD_CACHE_DIR, *parts[1:])
        text_value = _normalize_text(text or (metadata or {}).get("text"), ensure_sentence_end=False)
        source_value = _normalize_source_path(source_path or (metadata or {}).get("source_path"))
        if text_value:
            if source_value == "shared":
                return _shared_word_cache_path(text_value)
            return _word_cache_path(text_value, source_path=source_value)
    return normalized


def _backend_key(source=None, fallback_backend=None):
    source_name = str(source or "").strip().lower()
    fallback_name = str(fallback_backend or "").strip().lower()
    if fallback_name in {"gemini", "kokoro", "piper"}:
        return fallback_name
    if source_name == SOURCE_KOKORO:
        return "kokoro"
    if source_name == SOURCE_PIPER:
        return "piper"
    return "gemini"


def _backend_label_from_key(backend_key):
    backend = str(backend_key or "").strip().lower()
    if backend == "kokoro":
        return "Kokoro (Offline)"
    if backend == "piper":
        return "Piper (Local)"
    return "Gemini TTS"


def _selected_source_backend_key():
    return _backend_key(source=get_voice_source())


def _actual_backend_key_for_result(*, fallback=False):
    selected = _selected_source_backend_key()
    if selected == "gemini" and fallback:
        return "kokoro"
    return selected


def _load_json_file(path, default):
    try:
        with open(path, "r", encoding="utf-8") as fp:
            data = json.load(fp)
        if isinstance(default, dict) and isinstance(data, dict):
            return data
        if isinstance(default, list) and isinstance(data, list):
            return data
    except Exception:
        pass
    return default


def _migrate_flat_root_cache_layout():
    renamed_paths = {}
    scan_roots = []
    for root in (LEGACY_WORD_CACHE_WRAPPER_DIR, GLOBAL_WORD_CACHE_DIR):
        if os.path.isdir(root) and root not in scan_roots:
            scan_roots.append(root)

    for scan_root in scan_roots:
        entries = list(os.listdir(scan_root))
        for name in entries:
            full_path = os.path.join(scan_root, name)
            if os.path.isdir(full_path):
                continue
            if name.lower().endswith(".wav"):
                cache_path = full_path
            elif name.lower().endswith(".wav.json"):
                cache_path = full_path[:-5]
            else:
                continue

            cache_name = os.path.basename(cache_path)
            metadata = _load_json_file(_cache_meta_path(cache_path), {})
            source_hint = _normalize_source_path((metadata or {}).get("source_path"))
            if source_hint == "shared":
                target_dir = os.path.join(SHARED_WORD_CACHE_DIR, _filename_letter_bucket(cache_name))
            else:
                target_dir = os.path.join(
                    SOURCE_WORD_CACHE_ROOT_DIR,
                    _source_bucket_name(source_hint),
                    _filename_letter_bucket(cache_name),
                )
            target_cache_path = os.path.join(target_dir, cache_name)
            if os.path.abspath(target_cache_path) == os.path.abspath(cache_path):
                continue

            os.makedirs(target_dir, exist_ok=True)
            wav_src = cache_path
            wav_dst = target_cache_path
            meta_src = _cache_meta_path(cache_path)
            meta_dst = _cache_meta_path(target_cache_path)

            try:
                if os.path.exists(wav_src):
                    if os.path.exists(wav_dst):
                        os.remove(wav_src)
                    else:
                        shutil.move(wav_src, wav_dst)
                if os.path.exists(meta_src):
                    meta_payload = _load_json_file(meta_src, {})
                    if isinstance(meta_payload, dict):
                        meta_payload["cache_path"] = target_cache_path
                    if os.path.exists(meta_dst):
                        os.remove(meta_src)
                        if isinstance(meta_payload, dict) and meta_payload:
                            _write_json_file(meta_dst, meta_payload)
                    else:
                        _write_json_file(meta_dst, meta_payload if isinstance(meta_payload, dict) else {})
                        os.remove(meta_src)
                renamed_paths[cache_path] = target_cache_path
            except Exception:
                continue

    if not renamed_paths:
        return

    pending_items = _load_json_file(PENDING_GEMINI_QUEUE_PATH, [])
    if isinstance(pending_items, list):
        changed = False
        for item in pending_items:
            if not isinstance(item, dict):
                continue
            old_cache_path = str(item.get("cache_path") or "").strip()
            if old_cache_path in renamed_paths:
                item["cache_path"] = renamed_paths[old_cache_path]
                changed = True
        if changed:
            _write_json_file(PENDING_GEMINI_QUEUE_PATH, pending_items)


def _migrate_legacy_word_wrapper_layout():
    if not os.path.isdir(LEGACY_WORD_CACHE_WRAPPER_DIR):
        return

    renamed_paths = {}
    move_plan = [
        (os.path.join(LEGACY_WORD_CACHE_WRAPPER_DIR, "sources"), SOURCE_WORD_CACHE_ROOT_DIR),
        (os.path.join(LEGACY_WORD_CACHE_WRAPPER_DIR, "global"), SHARED_WORD_CACHE_DIR),
    ]
    for legacy_dir, target_dir in move_plan:
        if not os.path.isdir(legacy_dir):
            continue
        os.makedirs(target_dir, exist_ok=True)
        for root, dirs, files in os.walk(legacy_dir, topdown=False):
            rel_root = os.path.relpath(root, legacy_dir)
            current_target_root = target_dir if rel_root == "." else os.path.join(target_dir, rel_root)
            os.makedirs(current_target_root, exist_ok=True)
            for name in files:
                src = os.path.join(root, name)
                dst = os.path.join(current_target_root, name)
                try:
                    if os.path.exists(dst):
                        os.remove(src)
                    else:
                        shutil.move(src, dst)
                    if name.lower().endswith(".wav"):
                        renamed_paths[src] = dst
                except Exception:
                    continue
            for name in dirs:
                old_dir = os.path.join(root, name)
                try:
                    if os.path.isdir(old_dir) and not os.listdir(old_dir):
                        os.rmdir(old_dir)
                except Exception:
                    pass
        try:
            if os.path.isdir(legacy_dir) and not os.listdir(legacy_dir):
                os.rmdir(legacy_dir)
        except Exception:
            pass

    pending_items = _load_json_file(PENDING_GEMINI_QUEUE_PATH, [])
    if isinstance(pending_items, list) and renamed_paths:
        changed = False
        for item in pending_items:
            if not isinstance(item, dict):
                continue
            old_cache_path = str(item.get("cache_path") or "").strip()
            if old_cache_path in renamed_paths:
                item["cache_path"] = renamed_paths[old_cache_path]
                changed = True
        if changed:
            _write_json_file(PENDING_GEMINI_QUEUE_PATH, pending_items)

    try:
        if os.path.isdir(LEGACY_WORD_CACHE_WRAPPER_DIR) and not os.listdir(LEGACY_WORD_CACHE_WRAPPER_DIR):
            os.rmdir(LEGACY_WORD_CACHE_WRAPPER_DIR)
    except Exception:
        pass


def _collapse_existing_lightweight_source_caches():
    if not os.path.isdir(SOURCE_WORD_CACHE_ROOT_DIR):
        return
    for root, _, names in os.walk(SOURCE_WORD_CACHE_ROOT_DIR):
        for name in names:
            if not name.lower().endswith(".wav.json"):
                continue
            meta_path = os.path.join(root, name)
            cache_path = meta_path[:-5]
            metadata = _load_json_file(meta_path, {})
            if not isinstance(metadata, dict):
                continue
            source_path = _normalize_source_path(metadata.get("source_path"))
            if not _is_lightweight_file_source(source_path):
                continue
            backend = str(metadata.get("backend") or "").strip().lower()
            desired_backend = str(metadata.get("desired_backend") or "").strip().lower()
            if backend != "gemini" or desired_backend != "gemini":
                continue
            if not os.path.exists(cache_path):
                continue
            text_value = _normalize_text(metadata.get("text"), ensure_sentence_end=False)
            if not text_value:
                continue
            _collapse_source_cache_to_alias(text_value, source_path=source_path)


def _cache_group_key(cache_path, metadata):
    text_value = _normalize_text((metadata or {}).get("text"), ensure_sentence_end=False)
    if text_value:
        return text_value.casefold()
    stem = os.path.splitext(os.path.basename(str(cache_path or "").strip()))[0]
    base = stem.rsplit("_", 1)[0] if "_" in stem else stem
    return str(base or "").casefold()


def _cache_group_source_bucket(cache_path):
    try:
        rel_path = os.path.relpath(str(cache_path or "").strip(), SOURCE_WORD_CACHE_ROOT_DIR)
    except Exception:
        return ""
    parts = rel_path.split(os.sep)
    return parts[0] if parts else ""


def _cache_sort_key(cache_path, metadata):
    metadata = dict(metadata or {})
    desired_backend = str(metadata.get("desired_backend") or "").strip().lower()
    backend = str(metadata.get("backend") or "").strip().lower()
    updated_at = int(metadata.get("updated_at") or 0)
    pending = 1 if _is_pending_gemini(cache_path) else 0
    playable = 1 if _resolve_cache_audio_path(cache_path) else 0
    actual_wav = 1 if os.path.exists(cache_path) else 0
    try:
        latest_mtime = max(
            os.path.getmtime(path)
            for path in (cache_path, _cache_meta_path(cache_path))
            if os.path.exists(path)
        )
    except Exception:
        latest_mtime = 0
    return (
        1 if desired_backend == "gemini" else 0,
        1 if backend == "gemini" else 0,
        pending,
        playable,
        actual_wav,
        updated_at,
        latest_mtime,
    )


def _cleanup_duplicate_source_cache_entries():
    if not os.path.isdir(SOURCE_WORD_CACHE_ROOT_DIR):
        return 0

    grouped = {}
    for root, _, names in os.walk(SOURCE_WORD_CACHE_ROOT_DIR):
        for name in names:
            lower_name = name.lower()
            if lower_name.endswith(".wav"):
                cache_path = os.path.join(root, name)
            elif lower_name.endswith(".wav.json"):
                cache_path = os.path.join(root, name[:-5])
            else:
                continue
            metadata = _load_cache_metadata(cache_path)
            bucket = _cache_group_source_bucket(cache_path)
            key = _cache_group_key(cache_path, metadata)
            if not (bucket and key):
                continue
            grouped.setdefault((bucket, key), {})[cache_path] = metadata

    removed = 0
    for (_bucket, _key), item_map in grouped.items():
        cache_paths = list(item_map.keys())
        if len(cache_paths) <= 1:
            continue
        keep_path = max(cache_paths, key=lambda path: _cache_sort_key(path, item_map.get(path)))
        keep_metadata = item_map.get(keep_path) or {}
        keep_source = _normalize_source_path(keep_metadata.get("source_path"))
        keep_text = _normalize_text(keep_metadata.get("text"), ensure_sentence_end=False)
        pending_seen = False
        for cache_path in cache_paths:
            if cache_path == keep_path:
                continue
            pending_seen = pending_seen or _is_pending_gemini(cache_path)
            try:
                if os.path.exists(cache_path):
                    os.remove(cache_path)
                    removed += 1
            except Exception:
                pass
            _remove_cache_metadata(cache_path)
            _remove_pending_gemini(cache_path)
        if pending_seen and keep_text and keep_source and not _is_pending_gemini(keep_path):
            keep_backend = str(keep_metadata.get("backend") or "").strip().lower()
            keep_desired = str(keep_metadata.get("desired_backend") or "").strip().lower()
            if keep_backend != "gemini" and keep_desired == "gemini":
                _enqueue_existing_cache_for_gemini_replacement(keep_text, keep_path, source_path=keep_source)
    return removed


def _write_json_file(path, payload):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as fp:
        json.dump(payload, fp, ensure_ascii=False, indent=2)


def _load_cache_metadata(cache_path):
    canonical_path = _canonicalize_cache_path(cache_path)
    key = str(canonical_path or "").strip()
    if not key:
        return {}
    with _cache_metadata_lock:
        cached = _cache_metadata_memory.get(key)
        if isinstance(cached, dict):
            return dict(cached)
    data = _load_json_file(_cache_meta_path(canonical_path), {})
    payload = data if isinstance(data, dict) else {}
    if payload:
        payload["cache_path"] = canonical_path
        source_value = _normalize_source_path(payload.get("source_path"))
        linked_shared = str(payload.get("linked_shared_path") or "").strip()
        if linked_shared:
            payload["linked_shared_path"] = _canonicalize_cache_path(linked_shared, metadata=payload)
        if payload.get("cache_path") != str(cache_path or "").strip():
            _write_json_file(_cache_meta_path(canonical_path), payload)
    with _cache_metadata_lock:
        _cache_metadata_memory[key] = dict(payload)
    return payload


def _save_cache_metadata(cache_path, metadata):
    cache_path = _canonicalize_cache_path(cache_path, metadata=metadata)
    payload = dict(metadata or {})
    payload["cache_path"] = cache_path
    _write_json_file(_cache_meta_path(cache_path), payload)
    with _cache_metadata_lock:
        _cache_metadata_memory[str(cache_path or "").strip()] = dict(payload)


def _remove_cache_metadata(cache_path):
    cache_path = _canonicalize_cache_path(cache_path)
    key = str(cache_path or "").strip()
    try:
        meta_path = _cache_meta_path(cache_path)
        if os.path.exists(meta_path):
            os.remove(meta_path)
    except Exception:
        pass
    with _cache_metadata_lock:
        _cache_metadata_memory.pop(key, None)


def _copy_cache_file(src_path, dst_path, metadata=None):
    os.makedirs(os.path.dirname(dst_path), exist_ok=True)
    shutil.copyfile(src_path, dst_path)
    payload = dict(metadata or {})
    if not payload:
        payload = _load_cache_metadata(src_path)
    if payload:
        payload["updated_at"] = int(os.path.getmtime(dst_path)) if os.path.exists(dst_path) else None
        _save_cache_metadata(dst_path, payload)


def _resolve_cache_audio_path(cache_path):
    cache_path = _canonicalize_cache_path(cache_path)
    if os.path.exists(cache_path):
        return cache_path
    metadata = _load_cache_metadata(cache_path)
    linked = str(metadata.get("linked_shared_path") or "").strip()
    if linked and os.path.exists(linked):
        return linked
    return ""


def _write_source_alias_metadata(text, source_path, shared_path, *, backend="gemini", desired_backend="gemini"):
    source_cache_path = _word_cache_path(text, source_path=source_path)
    metadata = {
        "text": _normalize_text(text, ensure_sentence_end=False),
        "backend": backend,
        "desired_backend": desired_backend,
        "source_path": str(_normalize_source_path(source_path) or "").strip() or None,
        "linked_shared_path": shared_path,
        "updated_at": int(os.path.getmtime(shared_path)) if shared_path and os.path.exists(shared_path) else None,
    }
    if os.path.exists(source_cache_path):
        try:
            os.remove(source_cache_path)
        except Exception:
            pass
    _save_cache_metadata(source_cache_path, metadata)
    return source_cache_path


def _collapse_source_cache_to_alias(text, source_path=None):
    if not _is_lightweight_file_source(source_path):
        return False
    source_cache_path = _word_cache_path(text, source_path=source_path)
    shared_cache_path = _shared_word_cache_path(text)
    if not _has_valid_gemini_cache(shared_cache_path):
        return False
    try:
        _write_source_alias_metadata(
            text,
            source_path,
            shared_cache_path,
            backend="gemini",
            desired_backend="gemini",
        )
        return True
    except Exception:
        return False


def _migrate_legacy_cache_if_needed(text, source_path=None):
    cache_path = _word_cache_path(text, source_path=source_path)
    if os.path.exists(cache_path):
        return cache_path

    legacy_candidates = [
        _source_flat_compat_word_cache_path(text, source_path=source_path),
        _flat_compat_word_cache_path(text, source_path=source_path),
        _legacy_word_cache_path(text, source_path=source_path),
    ]
    for legacy_path in legacy_candidates:
        if not legacy_path or not os.path.exists(legacy_path):
            continue
        try:
            os.makedirs(os.path.dirname(cache_path), exist_ok=True)
            shutil.move(legacy_path, cache_path)
        except Exception:
            try:
                _copy_cache_file(legacy_path, cache_path)
            except Exception:
                continue
        legacy_meta = _load_cache_metadata(legacy_path)
        if legacy_meta:
            legacy_meta["source_path"] = str(_normalize_source_path(source_path) or "").strip() or None
            _save_cache_metadata(cache_path, legacy_meta)
            _remove_cache_metadata(legacy_path)
        return cache_path
    return cache_path


def _hydrate_source_cache_from_shared(text, source_path=None):
    shared_cache_path = _shared_word_cache_path(text)
    source_cache_path = _word_cache_path(text, source_path=source_path)
    if not _has_valid_gemini_cache(shared_cache_path):
        return False
    try:
        metadata = _load_cache_metadata(shared_cache_path)
        if _is_lightweight_file_source(source_path):
            _write_source_alias_metadata(
                text,
                source_path,
                shared_cache_path,
                backend=str(metadata.get("backend") or "gemini"),
                desired_backend=str(metadata.get("desired_backend") or "gemini"),
            )
        else:
            metadata["source_path"] = str(_normalize_source_path(source_path) or "").strip() or None
            _copy_cache_file(shared_cache_path, source_cache_path, metadata=metadata)
        if not str(source_path or "").strip():
            with _manual_session_cache_lock:
                _manual_session_cache_paths.add(source_cache_path)
        return True
    except Exception:
        return False


def _save_pending_gemini_queue_locked():
    payload = []
    for cache_path, item in _pending_gemini_replacements.items():
        if not isinstance(item, dict):
            continue
        text = _normalize_text(item.get("text"), ensure_sentence_end=False)
        if not text:
            continue
        payload.append(
            {
                "cache_path": cache_path,
                "text": text,
                "source_path": str(item.get("source_path") or "").strip() or None,
                "created_at": item.get("created_at"),
                "desired_backend": "gemini",
            }
        )
    _write_json_file(PENDING_GEMINI_QUEUE_PATH, payload)


def _load_pending_gemini_queue():
    items = _load_json_file(PENDING_GEMINI_QUEUE_PATH, [])
    if not isinstance(items, list):
        return
    changed = False
    with _pending_gemini_lock:
        _pending_gemini_replacements.clear()
        for item in items:
            if not isinstance(item, dict):
                continue
            cache_path = _canonicalize_cache_path(
                item.get("cache_path"),
                text=item.get("text"),
                source_path=item.get("source_path"),
                metadata=item,
            )
            text = _normalize_text(item.get("text"), ensure_sentence_end=False)
            if not cache_path or not text:
                continue
            if cache_path != str(item.get("cache_path") or "").strip():
                item["cache_path"] = cache_path
                changed = True
            _pending_gemini_replacements[cache_path] = {
                "text": text,
                "source_path": str(item.get("source_path") or "").strip() or None,
                "created_at": item.get("created_at"),
            }
        if changed:
            _write_json_file(PENDING_GEMINI_QUEUE_PATH, items)


def _remove_pending_gemini(cache_path):
    with _pending_gemini_lock:
        if cache_path in _pending_gemini_replacements:
            _pending_gemini_replacements.pop(cache_path, None)
            _save_pending_gemini_queue_locked()


def _is_pending_gemini(cache_path):
    with _pending_gemini_lock:
        return cache_path in _pending_gemini_replacements


def _enqueue_existing_cache_for_gemini_replacement(text, cache_path, source_path=None):
    normalized = _normalize_text(text, ensure_sentence_end=False)
    if not normalized or not cache_path:
        return
    if not str(source_path or "").strip():
        with _manual_session_cache_lock:
            _manual_session_cache_paths.add(cache_path)
    with _pending_gemini_lock:
        _pending_gemini_replacements[cache_path] = {
            "text": normalized,
            "source_path": str(source_path or "").strip() or None,
            "created_at": int(time.time()),
        }
        _save_pending_gemini_queue_locked()
    _start_pending_gemini_worker()


def _cache_requires_gemini_replacement(cache_path):
    metadata = _load_cache_metadata(cache_path)
    backend = str(metadata.get("backend") or "").strip().lower()
    return backend != "gemini"


def _has_valid_gemini_cache(cache_path):
    if _cache_requires_gemini_replacement(cache_path):
        return False
    return bool(_resolve_cache_audio_path(cache_path))


def _save_word_cache_file(cache_path, wav_path, *, text=None, source_path=None, backend=None, desired_backend=None):
    actual_backend = _backend_key(fallback_backend=backend)
    wanted_backend = _backend_key(fallback_backend=desired_backend or actual_backend)
    normalized_source = _normalize_source_path(source_path)
    payload = {
        "text": _normalize_text(text, ensure_sentence_end=False),
        "backend": actual_backend,
        "desired_backend": wanted_backend,
        "source_path": str(normalized_source or "").strip() or None,
    }

    if actual_backend == "gemini":
        shared_cache_path = _shared_word_cache_path(text or cache_path)
        try:
            shared_payload = dict(payload)
            shared_payload["source_path"] = "shared"
            os.makedirs(os.path.dirname(shared_cache_path), exist_ok=True)
            shutil.copyfile(wav_path, shared_cache_path)
            shared_payload["updated_at"] = int(os.path.getmtime(shared_cache_path)) if os.path.exists(shared_cache_path) else None
            _save_cache_metadata(shared_cache_path, shared_payload)
        except Exception:
            shared_cache_path = ""
        if _is_lightweight_file_source(source_path) and shared_cache_path:
            _write_source_alias_metadata(
                text or cache_path,
                source_path,
                shared_cache_path,
                backend=actual_backend,
                desired_backend=wanted_backend,
            )
            return

    os.makedirs(os.path.dirname(cache_path), exist_ok=True)
    shutil.copyfile(wav_path, cache_path)
    payload["updated_at"] = int(os.path.getmtime(cache_path)) if os.path.exists(cache_path) else None
    _save_cache_metadata(cache_path, payload)
    if not str(source_path or "").strip():
        with _manual_session_cache_lock:
            _manual_session_cache_paths.add(cache_path)


def cleanup_manual_session_cache():
    with _manual_session_cache_lock:
        paths = list(_manual_session_cache_paths)
        _manual_session_cache_paths.clear()
    for path in paths:
        try:
            if path and os.path.exists(path):
                os.remove(path)
        except Exception:
            pass
        _remove_cache_metadata(path)
        _remove_pending_gemini(path)


def cleanup_cache_for_source_path(source_path):
    target = _normalize_source_path(source_path)
    if not target:
        return 0
    removed = 0
    source_dir = _source_word_cache_dir(source_path=target)
    if not os.path.isdir(source_dir):
        with _pending_gemini_lock:
            pending_paths = [
                cache_path
                for cache_path, item in _pending_gemini_replacements.items()
                if _normalize_source_path((item or {}).get("source_path")) == target
            ]
        for cache_path in pending_paths:
            _remove_pending_gemini(cache_path)
        return 0
    for root, _, names in os.walk(source_dir, topdown=False):
        for name in names:
            full_path = os.path.join(root, name)
            if name.lower().endswith(".wav"):
                cache_path = full_path
            elif name.lower().endswith(".wav.json"):
                cache_path = full_path[:-5]
            else:
                continue
            try:
                if os.path.exists(full_path):
                    os.remove(full_path)
                    removed += 1
            except Exception:
                pass
            if name.lower().endswith(".wav"):
                _remove_cache_metadata(cache_path)
            elif name.lower().endswith(".wav.json"):
                _remove_pending_gemini(cache_path)
        try:
            if os.path.isdir(root) and not os.listdir(root):
                os.rmdir(root)
        except Exception:
            pass
    with _pending_gemini_lock:
        pending_paths = [
            cache_path
            for cache_path, item in _pending_gemini_replacements.items()
            if _normalize_source_path((item or {}).get("source_path")) == target
        ]
    for cache_path in pending_paths:
        _remove_pending_gemini(cache_path)
    try:
        if os.path.isdir(source_dir) and not os.listdir(source_dir):
            os.rmdir(source_dir)
    except Exception:
        pass
    return removed


def _is_gemini_rate_limited_error(error):
    message = str(error or "").lower()
    return (
        "rate limit" in message
        or "resource_exhausted" in message
        or "429" in message
        or "quota exceeded" in message
    )


def _clone_to_temp(path):
    fd, temp_path = tempfile.mkstemp(prefix="wordspeaker_", suffix=".wav")
    os.close(fd)
    shutil.copyfile(path, temp_path)
    return temp_path


def _ensure_source_gemini_cache(text, source_path=None):
    source_cache_path = _migrate_legacy_cache_if_needed(text, source_path=source_path)
    if _has_valid_gemini_cache(source_cache_path):
        shared_path = _shared_word_cache_path(text)
        if not _has_valid_gemini_cache(shared_path):
            try:
                metadata = _load_cache_metadata(source_cache_path)
                shared_payload = dict(metadata) if isinstance(metadata, dict) else {}
                shared_payload["source_path"] = "shared"
                actual_source_path = _resolve_cache_audio_path(source_cache_path)
                if actual_source_path:
                    _copy_cache_file(actual_source_path, shared_path, metadata=shared_payload)
            except Exception:
                pass
        if _is_lightweight_file_source(source_path) and _has_valid_gemini_cache(shared_path):
            try:
                _collapse_source_cache_to_alias(text, source_path=source_path)
            except Exception:
                pass
        return source_cache_path
    if _hydrate_source_cache_from_shared(text, source_path=source_path):
        return source_cache_path
    return source_cache_path


def has_cached_word_audio(text, source_path=None):
    return _has_valid_gemini_cache(_ensure_source_gemini_cache(text, source_path=source_path))


def get_word_audio_cache_info(text, source_path=None):
    cache_path = _ensure_source_gemini_cache(text, source_path=source_path)
    playable_cache = _resolve_cache_audio_path(cache_path)
    exists = bool(playable_cache)
    metadata = _load_cache_metadata(cache_path)
    backend = str(metadata.get("backend") or "").strip().lower()
    desired_backend = str(metadata.get("desired_backend") or "").strip().lower()
    with _pending_gemini_lock:
        pending = cache_path in _pending_gemini_replacements
    shared_path = _shared_word_cache_path(text)
    return {
        "exists": exists,
        "cache_path": cache_path,
        "playable_cache_path": playable_cache or cache_path,
        "meta_path": _cache_meta_path(cache_path),
        "shared_cache_path": shared_path,
        "shared_exists": os.path.exists(shared_path),
        "uses_shared_cache": bool(playable_cache and os.path.abspath(playable_cache) == os.path.abspath(shared_path) and os.path.abspath(cache_path) != os.path.abspath(shared_path)),
        "backend": backend or None,
        "backend_label": _backend_label_from_key(backend) if backend else "",
        "desired_backend": (desired_backend or ("gemini" if pending else None)),
        "desired_backend_label": _backend_label_from_key(desired_backend or "gemini") if (desired_backend or pending) else "",
        "pending_gemini_replacement": bool(pending),
        "metadata": dict(metadata) if isinstance(metadata, dict) else {},
    }


def get_recent_wrong_cache_source():
    return RECENT_WRONG_SOURCE_KEY


def queue_word_audio_generation(text, source_path=None):
    normalized = _normalize_text(text, ensure_sentence_end=False)
    if not normalized:
        return False
    cache_path = _ensure_source_gemini_cache(normalized, source_path=source_path)
    if _has_valid_gemini_cache(cache_path) or _is_pending_gemini(cache_path):
        return False
    _enqueue_existing_cache_for_gemini_replacement(normalized, cache_path, source_path=source_path)
    return True


def _copy_cache_between_sources(text, *, from_source_path=None, to_source_path=None):
    source_cache = _ensure_source_gemini_cache(text, source_path=from_source_path)
    target_cache = _word_cache_path(text, source_path=to_source_path)
    playable_source_cache = _resolve_cache_audio_path(source_cache)
    if not playable_source_cache:
        return False
    try:
        os.makedirs(os.path.dirname(target_cache), exist_ok=True)
        shutil.copyfile(playable_source_cache, target_cache)
        metadata = _load_cache_metadata(source_cache)
        payload = dict(metadata) if isinstance(metadata, dict) else {}
        payload["source_path"] = str(to_source_path or "").strip() or None
        if "backend" not in payload:
            payload["backend"] = "unknown"
        if "desired_backend" not in payload:
            payload["desired_backend"] = "gemini"
        _save_cache_metadata(target_cache, payload)
        return True
    except Exception:
        return False


def promote_word_audio_to_recent_wrong(text, source_path=None):
    normalized = _normalize_text(text, ensure_sentence_end=False)
    if not normalized:
        return False
    wrong_source = get_recent_wrong_cache_source()
    wrong_cache = _word_cache_path(normalized, source_path=wrong_source)
    source_cache = _word_cache_path(normalized, source_path=source_path)
    copied = False
    source_exists = bool(source_path) and bool(_resolve_cache_audio_path(source_cache))
    wrong_exists = bool(_resolve_cache_audio_path(wrong_cache))
    if source_exists and (not wrong_exists or (_has_valid_gemini_cache(source_cache) and not _has_valid_gemini_cache(wrong_cache))):
        copied = _copy_cache_between_sources(
            normalized,
            from_source_path=source_path,
            to_source_path=wrong_source,
        )
    if copied:
        source_pending = _is_pending_gemini(source_cache)
        if source_pending or _cache_requires_gemini_replacement(wrong_cache):
            _enqueue_existing_cache_for_gemini_replacement(normalized, wrong_cache, source_path=wrong_source)
        return True
    return queue_word_audio_generation(normalized, source_path=wrong_source)


def _set_backend_status(token, label, *, from_cache=False, fallback=False):
    with _backend_lock:
        _backend_status[token] = {
            "label": str(label or ""),
            "from_cache": bool(from_cache),
            "fallback": bool(fallback),
        }


def get_backend_status(token):
    with _backend_lock:
        data = _backend_status.get(int(token or 0))
        return dict(data) if isinstance(data, dict) else None


def _stop_locked():
    global _current_wav
    try:
        winsound.PlaySound(None, winsound.SND_PURGE)
    except Exception:
        pass
    if _current_wav:
        try:
            os.remove(_current_wav)
        except Exception:
            pass
        _current_wav = None


def _show_error_once(message):
    if not message:
        return
    with _lock:
        if message in _shown_errors:
            return
        _shown_errors.add(message)
    try:
        messagebox.showerror("Speech Error", f"Error: {message}")
    except Exception:
        pass


def _extract_error_message(http_error):
    try:
        raw = http_error.read().decode("utf-8", errors="ignore")
    except Exception:
        return str(http_error)
    try:
        data = json.loads(raw)
    except Exception:
        return raw or str(http_error)
    error = data.get("error") or {}
    return str(error.get("message") or data.get("message") or raw or http_error)


def _request_gemini_tts(text, *, short_text):
    api_key = get_gemini_api_key()
    if not api_key:
        raise RuntimeError("Gemini API key is empty.")

    prompt = TTS_STYLE_SHORT if short_text else TTS_STYLE_LONG
    payload = {
        "contents": [{"parts": [{"text": f"{prompt}\n\nText:\n{text}"}]}],
        "generationConfig": {
            "responseModalities": ["AUDIO"],
            "speechConfig": {
                "voiceConfig": {
                    "prebuiltVoiceConfig": {
                        "voiceName": TTS_VOICE_NAME,
                    }
                }
            },
        },
    }
    req = urllib.request.Request(
        TTS_URL,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "x-goog-api-key": api_key,
            "Content-Type": "application/json",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=180) as response:
            return json.loads(response.read().decode("utf-8", errors="ignore"))
    except urllib.error.HTTPError as e:
        raise RuntimeError(_extract_error_message(e)) from e
    except urllib.error.URLError as e:
        raise RuntimeError(f"Gemini TTS request failed: {e.reason}") from e
    except Exception as e:
        raise RuntimeError(f"Gemini TTS request failed: {e}") from e


def _extract_pcm_bytes(data):
    candidates = data.get("candidates") or []
    for candidate in candidates:
        content = candidate.get("content") or {}
        for part in content.get("parts") or []:
            inline = part.get("inlineData") or {}
            encoded = inline.get("data")
            mime_type = str(inline.get("mimeType") or "")
            if encoded and "pcm" in mime_type:
                return base64.b64decode(encoded)
    raise RuntimeError("Gemini TTS returned no audio.")


def _write_pcm_to_wav_path(pcm_bytes, sample_rate, volume):
    gain = _clamp(volume, 0.0, 1.0)
    if gain <= 0.0:
        pcm_bytes = b""
    elif abs(gain - 1.0) > 1e-6:
        import array

        samples = array.array("h")
        samples.frombytes(pcm_bytes)
        for idx, sample in enumerate(samples):
            scaled = int(sample * gain)
            if scaled > 32767:
                scaled = 32767
            elif scaled < -32768:
                scaled = -32768
            samples[idx] = scaled
        pcm_bytes = samples.tobytes()

    fd, path = tempfile.mkstemp(prefix="wordspeaker_", suffix=".wav")
    os.close(fd)
    with wave.open(path, "wb") as wav_fp:
        wav_fp.setnchannels(TTS_CHANNELS)
        wav_fp.setsampwidth(TTS_SAMPLE_WIDTH)
        wav_fp.setframerate(sample_rate)
        wav_fp.writeframes(pcm_bytes)
    return path


def _write_float_audio_to_wav_path(audio, sample_rate, volume):
    gain = _clamp(volume, 0.0, 1.0)
    pcm = np.asarray(audio, dtype=np.float32) * gain
    pcm = np.clip(pcm, -1.0, 1.0)
    pcm16 = (pcm * 32767.0).astype(np.int16)
    return _write_pcm_to_wav_path(pcm16.tobytes(), sample_rate=sample_rate, volume=1.0)


def _ensure_kokoro():
    global _kokoro
    if _kokoro is not None:
        return _kokoro
    if not kokoro_ready():
        model_path, voices_path = get_kokoro_paths()
        raise RuntimeError(
            "Kokoro model files are missing.\n"
            f"Expected:\n{model_path}\n{voices_path}"
        )
    with _kokoro_lock:
        if _kokoro is not None:
            return _kokoro
        from kokoro_onnx import Kokoro

        model_path, voices_path = get_kokoro_paths()
        _kokoro = Kokoro(model_path=model_path, voices_path=voices_path)
        return _kokoro


def _ensure_piper_voice(voice_id=None):
    voice_key = str(voice_id or get_voice_id() or "").strip()
    if not voice_key:
        raise RuntimeError("Piper voice id is empty.")
    with _piper_lock:
        if voice_key in _piper_voices:
            return _piper_voices[voice_key]
    if not piper_ready():
        raise RuntimeError("Piper is not ready. Add a Piper model under data/models/piper.")
    profile = get_piper_voice_profile(voice_key)
    model_path = str(profile.get("model_path") or "").strip()
    config_path = str(profile.get("config_path") or "").strip()
    if not model_path or not config_path:
        raise RuntimeError("Piper model files are missing or incomplete.")
    from piper import PiperVoice

    voice = PiperVoice.load(model_path=model_path, config_path=config_path)
    with _piper_lock:
        _piper_voices[voice_key] = voice
    return voice


def _synthesize_with_gemini(text, volume, *, short_text):
    data = _request_gemini_tts(text, short_text=short_text)
    pcm_bytes = _extract_pcm_bytes(data)
    return _write_pcm_to_wav_path(pcm_bytes, sample_rate=TTS_SAMPLE_RATE, volume=volume), "Gemini TTS", True


def _synthesize_with_kokoro(text, volume, rate_ratio):
    kokoro = _ensure_kokoro()
    voice_id = get_voice_id()
    profile = get_voice_profile(get_voice_source(), voice_id)
    lang = str((profile.get("languages") or ["en-GB"])[0]).strip().lower().replace("_", "-")
    speed = _clamp(rate_ratio, 0.7, 1.4)
    audio, sample_rate = kokoro.create(text, voice=voice_id, speed=speed, lang=lang)
    return (
        _write_float_audio_to_wav_path(audio, sample_rate=sample_rate or KOKORO_SAMPLE_RATE, volume=volume),
        "Kokoro (Offline)",
        False,
    )


def _synthesize_with_kokoro_voice(text, volume, rate_ratio, voice_id, lang="en-gb"):
    kokoro = _ensure_kokoro()
    speed = _clamp(rate_ratio, 0.7, 1.4)
    audio, sample_rate = kokoro.create(text, voice=voice_id, speed=speed, lang=lang)
    return (
        _write_float_audio_to_wav_path(audio, sample_rate=sample_rate or KOKORO_SAMPLE_RATE, volume=volume),
        "Kokoro (Offline)",
        False,
    )


def _synthesize_with_piper(text, volume, rate_ratio):
    if not piper_ready():
        raise RuntimeError("Piper is not ready. Add a Piper model under data/models/piper.")
    profile = get_piper_voice_profile(get_voice_id())
    voice = _ensure_piper_voice(profile.get("id"))
    from piper import SynthesisConfig

    speed = _clamp(rate_ratio, 0.7, 1.4)
    syn_config = SynthesisConfig(length_scale=max(0.1, 1.0 / speed), volume=_clamp(volume, 0.0, 1.0))
    audio_chunks = list(voice.synthesize(str(text or ""), syn_config=syn_config))
    if not audio_chunks:
        raise RuntimeError("Piper returned no audio.")
    audio = np.concatenate([chunk.audio_float_array for chunk in audio_chunks])
    sample_rate = int(audio_chunks[0].sample_rate or TTS_SAMPLE_RATE)
    return _write_float_audio_to_wav_path(audio, sample_rate=sample_rate, volume=1.0), "Piper (Local)", False


def _synthesize_with_local_placeholder(text, volume, rate_ratio):
    if piper_ready():
        return _synthesize_with_piper(text, volume=volume, rate_ratio=rate_ratio), "piper"
    if kokoro_ready():
        return (
            _synthesize_with_kokoro_voice(
                text,
                volume=volume,
                rate_ratio=rate_ratio,
                voice_id="bf_emma",
                lang="en-gb",
            ),
            "kokoro",
        )
    raise RuntimeError("No local TTS backend is ready for Gemini fallback.")


def _synthesize_with_selected_source(text, volume, rate_ratio, *, short_text):
    source = get_voice_source()
    if source == SOURCE_KOKORO:
        return _synthesize_with_kokoro(text, volume=volume, rate_ratio=rate_ratio), False
    if source == SOURCE_PIPER:
        return _synthesize_with_piper(text, volume=volume, rate_ratio=rate_ratio), False

    try:
        return _synthesize_with_gemini(text, volume=volume, short_text=short_text), False
    except Exception:
        if kokoro_ready():
            return (
                _synthesize_with_kokoro_voice(
                    text,
                    volume=volume,
                    rate_ratio=rate_ratio,
                    voice_id="bf_emma",
                    lang="en-gb",
                ),
                True,
            )
        raise


def _synthesize_to_wav(text, volume, rate_ratio, *, short_text=False, source_path=None, request_token=None):
    normalized = _normalize_text(text, ensure_sentence_end=short_text)
    if not normalized:
        raise RuntimeError("Text is empty.")

    if short_text:
        cache_path = _ensure_source_gemini_cache(text, source_path=source_path)
        if _selected_source_backend_key() == "gemini" and _has_valid_gemini_cache(cache_path):
            playable_cache = _resolve_cache_audio_path(cache_path) or cache_path
            if request_token is not None:
                _set_backend_status(request_token, "Gemini TTS", from_cache=True, fallback=False)
            return _clone_to_temp(playable_cache)
        if _selected_source_backend_key() == "gemini" and os.path.exists(cache_path):
            metadata = _load_cache_metadata(cache_path)
            backend_key = str(metadata.get("backend") or "").strip().lower()
            if backend_key in {"piper", "kokoro"}:
                _enqueue_existing_cache_for_gemini_replacement(text, cache_path, source_path=source_path)
                if request_token is not None:
                    _set_backend_status(
                        request_token,
                        _backend_label_from_key(backend_key),
                        from_cache=True,
                        fallback=True,
                    )
                playable_cache = _resolve_cache_audio_path(cache_path) or cache_path
                return _clone_to_temp(playable_cache)
        legacy_cache_path = _legacy_word_cache_path(text, source_path=source_path)
        if legacy_cache_path and os.path.exists(legacy_cache_path) and not os.path.exists(cache_path):
            try:
                os.makedirs(os.path.dirname(cache_path), exist_ok=True)
                shutil.move(legacy_cache_path, cache_path)
            except Exception:
                try:
                    shutil.copyfile(legacy_cache_path, cache_path)
                except Exception:
                    cache_path = legacy_cache_path
            if not _load_cache_metadata(cache_path):
                _save_cache_metadata(
                    cache_path,
                    {
                        "backend": "unknown",
                        "desired_backend": "gemini",
                        "source_path": str(source_path or "").strip() or None,
                    },
                )
        if short_text and not _has_valid_gemini_cache(cache_path):
            _enqueue_existing_cache_for_gemini_replacement(text, cache_path, source_path=source_path)

    selected_backend = _selected_source_backend_key()
    if selected_backend == "kokoro":
        wav_path, backend_label, _can_cache = _synthesize_with_kokoro(
            normalized,
            volume=volume,
            rate_ratio=rate_ratio,
        )
        if request_token is not None:
            _set_backend_status(request_token, backend_label, from_cache=False, fallback=False)
        return wav_path
    if selected_backend == "piper":
        wav_path, backend_label, _can_cache = _synthesize_with_piper(
            normalized,
            volume=volume,
            rate_ratio=rate_ratio,
        )
        if request_token is not None:
            _set_backend_status(request_token, backend_label, from_cache=False, fallback=False)
        return wav_path

    try:
        wav_path, backend_label, _can_cache = _synthesize_with_gemini(
            normalized,
            volume=volume,
            short_text=short_text,
        )
        if request_token is not None:
            _set_backend_status(request_token, backend_label, from_cache=False, fallback=False)
        if short_text:
            try:
                _save_word_cache_file(
                    cache_path,
                    wav_path,
                    text=normalized,
                    source_path=source_path,
                    backend="gemini",
                    desired_backend="gemini",
                )
            except Exception:
                pass
        return wav_path
    except Exception:
        if short_text:
            try:
                (wav_path, backend_label, _can_cache), placeholder_backend = _synthesize_with_local_placeholder(
                    normalized,
                    volume=volume,
                    rate_ratio=rate_ratio,
                )
                try:
                    _save_word_cache_file(
                        cache_path,
                        wav_path,
                        text=normalized,
                        source_path=source_path,
                        backend=placeholder_backend,
                        desired_backend="gemini",
                    )
                except Exception:
                    pass
                _enqueue_existing_cache_for_gemini_replacement(text, cache_path, source_path=source_path)
                if request_token is not None:
                    _set_backend_status(request_token, backend_label, from_cache=False, fallback=True)
                return wav_path
            except Exception:
                _enqueue_existing_cache_for_gemini_replacement(text, cache_path, source_path=source_path)
        if kokoro_ready():
            wav_path, backend_label, _can_cache = _synthesize_with_kokoro_voice(
                normalized,
                volume=volume,
                rate_ratio=rate_ratio,
                voice_id="bf_emma",
                lang="en-gb",
            )
            if request_token is not None:
                _set_backend_status(request_token, backend_label, from_cache=False, fallback=True)
            return wav_path
        raise


def _start_pending_gemini_worker():
    global _pending_gemini_worker_running
    with _pending_gemini_lock:
        if _pending_gemini_worker_running or not _pending_gemini_replacements:
            return
        _pending_gemini_worker_running = True

    def _worker():
        global _pending_gemini_worker_running
        try:
            while True:
                with _pending_gemini_lock:
                    pending_items = list(_pending_gemini_replacements.items())
                if not pending_items:
                    return
                cache_path_local, item = pending_items[0]
                if not isinstance(item, dict):
                    _remove_pending_gemini(cache_path_local)
                    continue
                normalized_text = _normalize_text(item.get("text"), ensure_sentence_end=False)
                source_path = str(item.get("source_path") or "").strip() or None
                if not normalized_text:
                    _remove_pending_gemini(cache_path_local)
                    continue
                try:
                    wav_path, _label, _can_cache = _synthesize_with_gemini(
                        normalized_text,
                        volume=1.0,
                        short_text=True,
                    )
                    try:
                        _save_word_cache_file(
                            cache_path_local,
                            wav_path,
                            text=normalized_text,
                            source_path=source_path,
                            backend="gemini",
                            desired_backend="gemini",
                        )
                    finally:
                        try:
                            os.remove(wav_path)
                        except Exception:
                            pass
                    _remove_pending_gemini(cache_path_local)
                except Exception as exc:
                    if _is_gemini_rate_limited_error(exc):
                        if not os.path.exists(cache_path_local):
                            try:
                                (wav_path, _label, _can_cache), placeholder_backend = _synthesize_with_local_placeholder(
                                    normalized_text,
                                    volume=1.0,
                                    rate_ratio=1.0,
                                )
                                try:
                                    _save_word_cache_file(
                                        cache_path_local,
                                        wav_path,
                                        text=normalized_text,
                                        source_path=source_path,
                                        backend=placeholder_backend,
                                        desired_backend="gemini",
                                    )
                                finally:
                                    try:
                                        os.remove(wav_path)
                                    except Exception:
                                        pass
                            except Exception:
                                pass
                        threading.Event().wait(GEMINI_RATE_LIMIT_COOLDOWN_SECONDS)
                        continue
                    _remove_pending_gemini(cache_path_local)
        finally:
            with _pending_gemini_lock:
                _pending_gemini_worker_running = False
        _start_pending_gemini_worker()

    threading.Thread(target=_worker, daemon=True).start()


def _enqueue_gemini_replacement(text, cache_path, source_path=None):
    _enqueue_existing_cache_for_gemini_replacement(text, cache_path, source_path=source_path)


def stop():
    with _lock:
        _stop_locked()


def speak_async(text, volume=1.0, rate_ratio=1.0, cancel_before=False, source_path=None):
    global _token
    if cancel_before:
        stop()
    with _lock:
        _token += 1
        my_token = _token

    def _run():
        global _current_wav
        try:
            with _lock:
                if my_token != _token:
                    return
            wav_path = _synthesize_to_wav(
                text=text,
                volume=volume,
                rate_ratio=rate_ratio,
                short_text=True,
                source_path=source_path,
                request_token=my_token,
            )
            with _lock:
                if my_token != _token:
                    try:
                        os.remove(wav_path)
                    except Exception:
                        pass
                    return
                _stop_locked()
                _current_wav = wav_path
            winsound.PlaySound(
                wav_path,
                winsound.SND_FILENAME | winsound.SND_ASYNC | winsound.SND_NODEFAULT,
            )
        except Exception as e:
            _show_error_once(str(e))

    threading.Thread(target=_run, daemon=True).start()
    return my_token


def _split_long_text(text, chunk_chars=1200):
    raw = _normalize_text(text)
    if not raw:
        return []

    sentences = re.split(r"(?<=[.!?;:])\s+", raw)
    chunks = []
    buf = ""
    limit = max(400, int(chunk_chars))
    for sentence in sentences:
        part = sentence.strip()
        if not part:
            continue
        candidate = f"{buf} {part}".strip() if buf else part
        if len(candidate) <= limit:
            buf = candidate
            continue
        if buf:
            chunks.append(buf)
        buf = part
    if buf:
        chunks.append(buf)
    return chunks


def speak_stream_async(text, volume=1.0, rate_ratio=1.0, cancel_before=False, chunk_chars=220):
    global _token
    if cancel_before:
        stop()
    with _lock:
        _token += 1
        my_token = _token

    def _run():
        global _current_wav
        try:
            with _lock:
                if my_token != _token:
                    return
            chunks = _split_long_text(text, chunk_chars=max(400, int(chunk_chars)))
            if not chunks:
                return
            wav_paths = []
            fallback = False
            if get_voice_source() == SOURCE_KOKORO:
                backend_label = "Kokoro (Offline)"
                synth_mode = "kokoro"
            else:
                first_result, fallback = _synthesize_with_selected_source(
                    chunks[0],
                    volume=volume,
                    rate_ratio=rate_ratio,
                    short_text=False,
                )
                first_path, backend_label, _can_cache = first_result
                wav_paths.append(first_path)
                synth_mode = "kokoro" if fallback else "gemini"
                if my_token is not None:
                    _set_backend_status(my_token, backend_label, from_cache=False, fallback=fallback)

            start_index = 0 if get_voice_source() == SOURCE_KOKORO else 1
            for chunk in chunks[start_index:]:
                with _lock:
                    if my_token != _token:
                        return
                if synth_mode == "kokoro":
                    wav_path, _label, _can_cache = _synthesize_with_kokoro_voice(
                        chunk,
                        volume=volume,
                        rate_ratio=rate_ratio,
                        voice_id="bf_emma",
                        lang="en-gb",
                    )
                else:
                    wav_path, _label, _can_cache = _synthesize_with_gemini(
                        chunk,
                        volume=volume,
                        short_text=False,
                    )
                wav_paths.append(wav_path)

            if get_voice_source() == SOURCE_KOKORO:
                _set_backend_status(my_token, backend_label, from_cache=False, fallback=False)

            fd, merged_path = tempfile.mkstemp(prefix="wordspeaker_", suffix=".wav")
            os.close(fd)
            with wave.open(merged_path, "wb") as out_fp:
                out_fp.setnchannels(TTS_CHANNELS)
                out_fp.setsampwidth(TTS_SAMPLE_WIDTH)
                out_fp.setframerate(TTS_SAMPLE_RATE)
                for wav_path in wav_paths:
                    with wave.open(wav_path, "rb") as in_fp:
                        out_fp.writeframes(in_fp.readframes(in_fp.getnframes()))
            for wav_path in wav_paths:
                try:
                    os.remove(wav_path)
                except Exception:
                    pass

            with _lock:
                if my_token != _token:
                    try:
                        os.remove(merged_path)
                    except Exception:
                        pass
                    return
                _stop_locked()
                _current_wav = merged_path
            winsound.PlaySound(
                merged_path,
                winsound.SND_FILENAME | winsound.SND_ASYNC | winsound.SND_NODEFAULT,
            )
        except Exception as e:
            _show_error_once(str(e))

    threading.Thread(target=_run, daemon=True).start()
    return my_token


def cancel_all():
    global _token
    with _lock:
        _token += 1
        _stop_locked()


def precache_word_audio_async(words, source_path=None, rate_ratio=1.0, on_progress=None, on_done=None):
    items = []
    seen = set()
    for word in words or []:
        text = _normalize_text(word, ensure_sentence_end=False)
        if not text:
            continue
        key = text.casefold()
        if key in seen:
            continue
        seen.add(key)
        items.append(text)

    def _emit_progress(done_count, total_count, current_text):
        if callable(on_progress):
            try:
                on_progress(done_count, total_count, current_text)
            except Exception:
                pass

    def _emit_done(success_count, skipped_count, pending_count, error_count):
        if callable(on_done):
            try:
                on_done(success_count, skipped_count, pending_count, error_count)
            except Exception:
                pass

    def _run():
        success_count = 0
        skipped_count = 0
        pending_count = 0
        error_count = 0
        total_count = len(items)
        for index, text in enumerate(items, start=1):
            cache_path = _ensure_source_gemini_cache(text, source_path=source_path)
            if _has_valid_gemini_cache(cache_path):
                skipped_count += 1
                _emit_progress(index, total_count, text)
                continue
            try:
                if not _is_pending_gemini(cache_path):
                    _enqueue_gemini_replacement(text, cache_path, source_path=source_path)
                pending_count += 1
            except Exception:
                error_count += 1
            _emit_progress(index, total_count, text)
        _emit_done(success_count, skipped_count, pending_count, error_count)

    threading.Thread(target=_run, daemon=True).start()


def prepare_async():
    return None


def get_runtime_label():
    source = get_voice_source()
    if source == SOURCE_KOKORO:
        return "Kokoro (Offline)"
    if source == SOURCE_PIPER:
        return "Piper (Local)"
    return "Gemini API"


_migrate_legacy_word_wrapper_layout()
_migrate_flat_root_cache_layout()
_load_pending_gemini_queue()
_cleanup_duplicate_source_cache_entries()
_collapse_existing_lightweight_source_caches()
_start_pending_gemini_worker()
