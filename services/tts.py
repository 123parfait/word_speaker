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

from services.app_config import get_llm_api_key, get_tts_api_key, get_tts_api_provider
from services.text_normalization import normalize_ielts_tts_text
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
PENDING_ONLINE_TTS_QUEUE_PATH = os.path.join(BASE_DIR, "data", "audio_cache", "pending_online_tts_replacements.json")
LEGACY_PENDING_GEMINI_QUEUE_PATH = os.path.join(BASE_DIR, "data", "audio_cache", "pending_gemini_replacements.json")
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
_preferred_pending_source = None
_manual_session_cache_paths = set()
_manual_session_cache_lock = threading.Lock()
_cache_metadata_memory = {}
_cache_metadata_lock = threading.Lock()
_gemini_queue_status_lock = threading.Lock()
_gemini_queue_status = {
    "state": "idle",
    "next_retry_at": 0.0,
    "last_success_at": 0.0,
    "last_error": "",
    "worker_running": False,
    "queue_count": 0,
}
_gemini_queue_attempt_lock = threading.Lock()
_gemini_queue_last_attempt_at = 0.0

TTS_MODEL = "gemini-2.5-flash-preview-tts"
TTS_URL = f"https://generativelanguage.googleapis.com/v1beta/models/{TTS_MODEL}:generateContent"
TTS_SAMPLE_RATE = 24000
TTS_CHANNELS = 1
TTS_SAMPLE_WIDTH = 2
TTS_VOICE_NAME = "Kore"
TTS_STYLE_SHORT = "Read the word clearly in a neutral British English accent for IELTS vocabulary practice."
TTS_STYLE_LONG = "Read this passage clearly in a natural British English accent for IELTS listening practice."
ELEVENLABS_VOICE_ID = "JBFqnCBsd6RMkjVDRZzb"
ELEVENLABS_MODEL_ID = "eleven_multilingual_v2"
ELEVENLABS_URL = f"https://api.elevenlabs.io/v1/text-to-speech/{ELEVENLABS_VOICE_ID}?output_format=pcm_24000"
GEMINI_QUEUE_REQUEST_INTERVAL_SECONDS = 40
GEMINI_RATE_LIMIT_COOLDOWN_SECONDS = 120
GEMINI_MANUAL_REQUEST_COOLDOWN_SECONDS = 60
ELEVENLABS_QUEUE_REQUEST_INTERVAL_SECONDS = 1.5
ELEVENLABS_RATE_LIMIT_COOLDOWN_SECONDS = 45
ELEVENLABS_MANUAL_REQUEST_COOLDOWN_SECONDS = 3
_QUEUE_THROTTLE_CONFIG = {
    "gemini": {
        "base_interval": GEMINI_QUEUE_REQUEST_INTERVAL_SECONDS,
        "min_interval": 35.0,
        "max_interval": 90.0,
        "success_step": 1.5,
        "success_streak": 2,
        "soft_fail_step": 8.0,
        "rate_limit_step": 12.0,
    },
    "elevenlabs": {
        "base_interval": ELEVENLABS_QUEUE_REQUEST_INTERVAL_SECONDS,
        "min_interval": 1.2,
        "max_interval": 8.0,
        "success_step": 0.15,
        "success_streak": 3,
        "soft_fail_step": 0.75,
        "rate_limit_step": 1.25,
    },
}
_queue_throttle_state_lock = threading.Lock()
_queue_throttle_state = {
    provider: {
        "current_interval": float(config["base_interval"]),
        "success_streak": 0,
        "last_event": "idle",
    }
    for provider, config in _QUEUE_THROTTLE_CONFIG.items()
}


def _clamp(value, low, high):
    return max(low, min(high, float(value)))


def _set_gemini_queue_status(**updates):
    with _gemini_queue_status_lock:
        _gemini_queue_status.update(updates)


def _provider_key(provider):
    return "elevenlabs" if str(provider or "").strip().lower() == "elevenlabs" else "gemini"


def _queue_throttle_config(provider):
    return _QUEUE_THROTTLE_CONFIG[_provider_key(provider)]


def _get_queue_throttle_snapshot(provider):
    key = _provider_key(provider)
    config = _queue_throttle_config(key)
    with _queue_throttle_state_lock:
        state = dict(_queue_throttle_state.get(key) or {})
    if not state:
        state = {
            "current_interval": float(config["base_interval"]),
            "success_streak": 0,
            "last_event": "idle",
        }
    return state


def _queue_interval_for_provider(provider):
    state = _get_queue_throttle_snapshot(provider)
    return float(state.get("current_interval") or _queue_throttle_config(provider)["base_interval"])


def _record_queue_success(provider):
    key = _provider_key(provider)
    config = _queue_throttle_config(key)
    with _queue_throttle_state_lock:
        state = _queue_throttle_state.setdefault(
            key,
            {
                "current_interval": float(config["base_interval"]),
                "success_streak": 0,
                "last_event": "idle",
            },
        )
        state["success_streak"] = int(state.get("success_streak") or 0) + 1
        if state["success_streak"] >= int(config["success_streak"]):
            state["current_interval"] = max(
                float(config["min_interval"]),
                float(state.get("current_interval") or config["base_interval"]) - float(config["success_step"]),
            )
            state["success_streak"] = 0
        state["last_event"] = "success"


def _record_queue_soft_failure(provider):
    key = _provider_key(provider)
    config = _queue_throttle_config(key)
    with _queue_throttle_state_lock:
        state = _queue_throttle_state.setdefault(
            key,
            {
                "current_interval": float(config["base_interval"]),
                "success_streak": 0,
                "last_event": "idle",
            },
        )
        state["success_streak"] = 0
        state["current_interval"] = min(
            float(config["max_interval"]),
            max(
                float(config["base_interval"]),
                float(state.get("current_interval") or config["base_interval"]) + float(config["soft_fail_step"]),
            ),
        )
        state["last_event"] = "soft_failure"


def _record_queue_rate_limit(provider):
    key = _provider_key(provider)
    config = _queue_throttle_config(key)
    with _queue_throttle_state_lock:
        state = _queue_throttle_state.setdefault(
            key,
            {
                "current_interval": float(config["base_interval"]),
                "success_streak": 0,
                "last_event": "idle",
            },
        )
        state["success_streak"] = 0
        state["current_interval"] = min(
            float(config["max_interval"]),
            max(
                float(config["base_interval"]),
                float(state.get("current_interval") or config["base_interval"]) + float(config["rate_limit_step"]),
            ),
        )
        state["last_event"] = "rate_limited"


def _refresh_gemini_queue_status_counts():
    with _pending_gemini_lock:
        queue_count = len(_pending_gemini_replacements)
        worker_running = bool(_pending_gemini_worker_running)
    try:
        disk_payload = _load_pending_queue_disk_payload()
        if isinstance(disk_payload, list):
            queue_count = len(disk_payload)
    except Exception:
        pass
    _set_gemini_queue_status(queue_count=queue_count, worker_running=worker_running)


def get_gemini_queue_status():
    _refresh_gemini_queue_status_counts()
    with _gemini_queue_status_lock:
        return dict(_gemini_queue_status)


def get_online_tts_queue_status():
    return get_gemini_queue_status()


def _defer_gemini_queue(wait_seconds, *, state=None, provider=None):
    global _gemini_queue_last_attempt_at
    wait_seconds = max(0.0, float(wait_seconds or 0.0))
    interval_seconds = _queue_interval_for_provider(provider)
    with _gemini_queue_attempt_lock:
        now = time.time()
        current_next = _gemini_queue_last_attempt_at + interval_seconds
        target_next = max(current_next, now + wait_seconds)
        _gemini_queue_last_attempt_at = target_next - interval_seconds
    updates = {"next_retry_at": target_next}
    if state:
        updates["state"] = state
    _set_gemini_queue_status(**updates)


def _wait_for_gemini_queue_slot(provider=None):
    global _gemini_queue_last_attempt_at
    interval_seconds = _queue_interval_for_provider(provider)
    while True:
        with _gemini_queue_attempt_lock:
            now = time.time()
            wait_seconds = interval_seconds - (now - _gemini_queue_last_attempt_at)
            if wait_seconds <= 0:
                _gemini_queue_last_attempt_at = now
                return
            next_retry_at = now + wait_seconds
        _set_gemini_queue_status(state="ok", next_retry_at=next_retry_at)
        threading.Event().wait(wait_seconds)


def _synthesize_with_user_online(text, volume, *, short_text):
    provider = _primary_online_provider()
    _defer_gemini_queue(_manual_request_cooldown_for_provider(provider), state="ok", provider=provider)
    try:
        result, used_provider = _synthesize_with_online(text, volume=volume, short_text=short_text)
        if _provider_key(used_provider) != _provider_key(provider):
            _record_queue_soft_failure(provider)
        _record_queue_success(used_provider)
        return result
    except Exception as exc:
        if _is_gemini_rate_limited_error(exc):
            _record_queue_rate_limit(provider)
            _set_gemini_queue_status(
                state="rate_limited",
                next_retry_at=time.time() + _rate_limit_cooldown_for_provider(provider),
                last_error=str(exc),
            )
        else:
            _record_queue_soft_failure(provider)
        raise


def _normalize_text(text, ensure_sentence_end=False):
    raw = re.sub(r"\s+", " ", str(text or "").strip())
    if ensure_sentence_end and raw and raw[-1] not in ".!?;:":
        raw = f"{raw}."
    return raw


def _normalize_tts_spoken_text(text):
    return normalize_ielts_tts_text(text)


def _normalize_cache_key_text(text):
    return _normalize_text(text, ensure_sentence_end=False).rstrip(".!?;:").casefold()


def _current_online_provider():
    value = str(get_tts_api_provider() or "gemini").strip().lower()
    return "elevenlabs" if value == "elevenlabs" else "gemini"


def _is_online_backend(backend_key):
    return str(backend_key or "").strip().lower() in {"gemini", "elevenlabs"}


def _online_provider_label(provider=None):
    return "ElevenLabs TTS" if str(provider or _current_online_provider()).strip().lower() == "elevenlabs" else "Gemini TTS"


def _primary_online_provider():
    return _current_online_provider()


def _secondary_online_provider(primary=None):
    first = str(primary or _primary_online_provider()).strip().lower()
    if first == "elevenlabs" and get_llm_api_key():
        return "gemini"
    return None


def _online_provider_candidates(primary=None):
    first = _provider_key(primary or _primary_online_provider())
    providers = [first]
    secondary = _secondary_online_provider(first)
    if secondary and _provider_key(secondary) not in providers:
        providers.append(_provider_key(secondary))
    return providers


def _rate_limit_cooldown_for_provider(provider):
    return ELEVENLABS_RATE_LIMIT_COOLDOWN_SECONDS if _provider_key(provider) == "elevenlabs" else GEMINI_RATE_LIMIT_COOLDOWN_SECONDS


def _manual_request_cooldown_for_provider(provider):
    return ELEVENLABS_MANUAL_REQUEST_COOLDOWN_SECONDS if _provider_key(provider) == "elevenlabs" else GEMINI_MANUAL_REQUEST_COOLDOWN_SECONDS


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
    normalized = _normalize_cache_key_text(text)
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
    normalized = _normalize_cache_key_text(text)
    provider = _current_online_provider()
    key = hashlib.sha1(
        f"{provider}|{source_key}|{TTS_MODEL}|{TTS_VOICE_NAME}|{normalized}".encode("utf-8")
    ).hexdigest()[:16]
    name = _safe_name(normalized)
    return os.path.join(GLOBAL_WORD_CACHE_DIR, f"{name}_{key}.wav")


def _source_flat_compat_word_cache_path(text, source_path=None):
    source_key = str(_normalize_source_path(source_path) or "").casefold()
    normalized = _normalize_cache_key_text(text)
    provider = _current_online_provider()
    key = hashlib.sha1(
        f"{provider}|{source_key}|{TTS_MODEL}|{TTS_VOICE_NAME}|{normalized}".encode("utf-8")
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
    normalized = _normalize_cache_key_text(text)
    key = hashlib.sha1(
        f"{source_key}|{voice_id}|{TTS_MODEL}|{TTS_VOICE_NAME}|{normalized}".encode("utf-8")
    ).hexdigest()[:16]
    name = _safe_name(normalized)
    return os.path.join(legacy_dir, f"{name}_{key}.wav")


def _word_cache_path(text, source_path=None):
    source_key = str(_normalize_source_path(source_path) or "").casefold()
    normalized = _normalize_cache_key_text(text)
    provider = _current_online_provider()
    key = hashlib.sha1(
        f"{provider}|{source_key}|{TTS_MODEL}|{TTS_VOICE_NAME}|{normalized}".encode("utf-8")
    ).hexdigest()[:16]
    name = _safe_name(normalized)
    return os.path.join(_source_word_cache_dir(normalized, source_path=source_path), f"{name}_{key}.wav")


def _shared_word_cache_path(text):
    normalized = _normalize_cache_key_text(text)
    provider = _current_online_provider()
    key = hashlib.sha1(
        f"{provider}|shared|{TTS_MODEL}|{TTS_VOICE_NAME}|{normalized}".encode("utf-8")
    ).hexdigest()[:16]
    name = _safe_name(normalized)
    return os.path.join(_shared_word_cache_dir(normalized), f"{name}_{key}.wav")


def _provider_source_word_cache_path(text, source_path=None, provider=None):
    normalized = _normalize_cache_key_text(text)
    backend = str(provider or _current_online_provider()).strip().lower()
    key = hashlib.sha1(
        f"{backend}|{str(_normalize_source_path(source_path) or '').casefold()}|{TTS_MODEL}|{TTS_VOICE_NAME}|{normalized}".encode("utf-8")
    ).hexdigest()[:16]
    return os.path.join(_source_word_cache_dir(normalized, source_path=source_path), f"{_safe_name(normalized)}_{key}.wav")


def _provider_shared_word_cache_path(text, provider=None):
    normalized = _normalize_cache_key_text(text)
    backend = str(provider or _current_online_provider()).strip().lower()
    key = hashlib.sha1(
        f"{backend}|shared|{TTS_MODEL}|{TTS_VOICE_NAME}|{normalized}".encode("utf-8")
    ).hexdigest()[:16]
    return os.path.join(_shared_word_cache_dir(normalized), f"{_safe_name(normalized)}_{key}.wav")


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
        provider_value = str(
            (metadata or {}).get("desired_backend")
            or (metadata or {}).get("backend")
            or _current_online_provider()
        ).strip().lower()
        if not _is_online_backend(provider_value):
            provider_value = _current_online_provider()
        if text_value:
            if source_value == "shared":
                return _provider_shared_word_cache_path(text_value, provider=provider_value)
            return _provider_source_word_cache_path(text_value, source_path=source_value, provider=provider_value)
    return normalized


def _backend_key(source=None, fallback_backend=None):
    source_name = str(source or "").strip().lower()
    fallback_name = str(fallback_backend or "").strip().lower()
    if fallback_name in {"gemini", "elevenlabs", "kokoro", "piper"}:
        return fallback_name
    if source_name == SOURCE_KOKORO:
        return "kokoro"
    if source_name == SOURCE_PIPER:
        return "piper"
    return _current_online_provider()


def _backend_label_from_key(backend_key):
    backend = str(backend_key or "").strip().lower()
    if backend == "kokoro":
        return "Kokoro (Offline)"
    if backend == "piper":
        return "Piper (Local)"
    if backend == "elevenlabs":
        return "ElevenLabs TTS"
    return "Gemini TTS"


def _backend_key_from_label(label):
    value = str(label or "").strip().lower()
    if "elevenlabs" in value:
        return "elevenlabs"
    if "kokoro" in value:
        return "kokoro"
    if "piper" in value:
        return "piper"
    return "gemini"


def _selected_source_backend_key():
    return _backend_key(source=get_voice_source())


def _actual_backend_key_for_result(*, fallback=False):
    selected = _selected_source_backend_key()
    if _is_online_backend(selected) and fallback:
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
            if source_path == "shared":
                continue
            text_value = _infer_text_from_cache_filename(cache_path, metadata)
            if not text_value:
                continue
            linked_shared = str(metadata.get("linked_shared_path") or "").strip()
            backend = str(metadata.get("backend") or "").strip().lower()
            desired_backend = str(metadata.get("desired_backend") or "").strip().lower()
            if linked_shared:
                _alias_source_cache_to_shared(
                    text_value,
                    source_path=source_path,
                    shared_path=linked_shared,
                    backend=backend or _current_online_provider(),
                    desired_backend=desired_backend or backend or _current_online_provider(),
                    metadata=metadata,
                    cache_path=cache_path,
                )
                continue
            playable_path = _resolve_cache_audio_path(cache_path)
            if not playable_path:
                continue
            try:
                _alias_source_cache_to_shared(
                    text_value,
                    source_path=source_path,
                    backend=backend or _current_online_provider(),
                    desired_backend=desired_backend or backend or _current_online_provider(),
                    metadata=metadata,
                    cache_path=cache_path,
                )
            except Exception:
                continue


def _collapse_all_source_cache_entities_to_aliases():
    if not os.path.isdir(SOURCE_WORD_CACHE_ROOT_DIR):
        return 0
    collapsed = 0
    for root, _, names in os.walk(SOURCE_WORD_CACHE_ROOT_DIR):
        for name in names:
            if not name.lower().endswith(".wav"):
                continue
            cache_path = os.path.join(root, name)
            metadata = _load_cache_metadata(cache_path)
            if not isinstance(metadata, dict):
                continue
            source_path = _normalize_source_path(metadata.get("source_path"))
            if source_path == "shared":
                continue
            text_value = _infer_text_from_cache_filename(cache_path, metadata)
            if not text_value:
                continue
            backend = str(metadata.get("backend") or "").strip().lower()
            desired_backend = str(metadata.get("desired_backend") or backend or _current_online_provider()).strip().lower()
            try:
                alias_path = _alias_source_cache_to_shared(
                    text_value,
                    source_path=source_path,
                    backend=backend or _current_online_provider(),
                    desired_backend=desired_backend or backend or _current_online_provider(),
                    metadata=metadata,
                    cache_path=cache_path,
                )
                if alias_path:
                    collapsed += 1
            except Exception:
                continue
    return collapsed


def _cache_group_key(cache_path, metadata):
    stem = os.path.splitext(os.path.basename(str(cache_path or "").strip()))[0]
    base = stem.rsplit("_", 1)[0] if "_" in stem else stem
    filename_guess = _normalize_text(str(base or "").replace("_", " "), ensure_sentence_end=False)
    if filename_guess:
        return filename_guess.rstrip(".!?;:").casefold()
    text_value = _normalize_text((metadata or {}).get("text"), ensure_sentence_end=False)
    if text_value:
        return text_value.rstrip(".!?;:").casefold()
    return str(base or "").rstrip(".!?;:").casefold()


def _infer_text_from_cache_filename(cache_path, metadata=None):
    text_value = _normalize_text((metadata or {}).get("text"), ensure_sentence_end=False)
    if text_value:
        return text_value
    stem = os.path.splitext(os.path.basename(str(cache_path or "").strip()))[0]
    base = stem.rsplit("_", 1)[0] if "_" in stem else stem
    guess = str(base or "").replace("_", " ").strip()
    return _normalize_text(guess, ensure_sentence_end=False)


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
        1 if _is_online_backend(desired_backend) else 0,
        1 if backend == desired_backend and _is_online_backend(backend) else 0,
        pending,
        playable,
        actual_wav,
        updated_at,
        latest_mtime,
    )


def _cleanup_duplicate_source_cache_entries():
    def _collect_grouped(root_dir, root_kind):
        grouped = {}
        if not os.path.isdir(root_dir):
            return grouped
        for root, _, names in os.walk(root_dir):
            for name in names:
                lower_name = name.lower()
                if lower_name.endswith(".wav"):
                    cache_path = os.path.join(root, name)
                elif lower_name.endswith(".wav.json"):
                    cache_path = os.path.join(root, name[:-5])
                else:
                    continue
                metadata = _load_cache_metadata(cache_path)
                bucket = _cache_group_source_bucket(cache_path) if root_kind == "source" else "shared"
                key = _cache_group_key(cache_path, metadata)
                if not (bucket and key):
                    continue
                grouped.setdefault((bucket, key), {})[cache_path] = metadata
        return grouped

    def _remove_cache_entry(cache_path):
        removed_local = 0
        try:
            if os.path.exists(cache_path):
                os.remove(cache_path)
                removed_local += 1
        except Exception:
            pass
        meta_path = _cache_meta_path(cache_path)
        try:
            if os.path.exists(meta_path):
                os.remove(meta_path)
                removed_local += 1
        except Exception:
            pass
        with _cache_metadata_lock:
            _cache_metadata_memory.pop(str(cache_path or "").strip(), None)
        _remove_pending_gemini(cache_path)
        return removed_local

    if not os.path.isdir(SOURCE_WORD_CACHE_ROOT_DIR) and not os.path.isdir(SHARED_WORD_CACHE_DIR):
        return 0

    removed = 0
    shared_path_rewrites = {}
    shared_grouped = _collect_grouped(SHARED_WORD_CACHE_DIR, "shared")
    for (_bucket, _key), item_map in shared_grouped.items():
        cache_paths = list(item_map.keys())
        if len(cache_paths) <= 1:
            continue
        keep_path = max(cache_paths, key=lambda path: _cache_sort_key(path, item_map.get(path)))
        for cache_path in cache_paths:
            if cache_path == keep_path:
                continue
            shared_path_rewrites[cache_path] = keep_path
            removed += _remove_cache_entry(cache_path)

    if shared_path_rewrites and os.path.isdir(SOURCE_WORD_CACHE_ROOT_DIR):
        for root, _, names in os.walk(SOURCE_WORD_CACHE_ROOT_DIR):
            for name in names:
                if not name.lower().endswith(".wav.json"):
                    continue
                cache_path = os.path.join(root, name[:-5])
                metadata = _load_cache_metadata(cache_path)
                if not isinstance(metadata, dict):
                    continue
                linked_shared = str(metadata.get("linked_shared_path") or "").strip()
                rewritten = shared_path_rewrites.get(linked_shared)
                if not rewritten:
                    continue
                metadata["linked_shared_path"] = rewritten
                _save_cache_metadata(cache_path, metadata)

    source_grouped = _collect_grouped(SOURCE_WORD_CACHE_ROOT_DIR, "source")
    for (_bucket, _key), item_map in source_grouped.items():
        cache_paths = list(item_map.keys())
        if len(cache_paths) <= 1:
            continue
        keep_path = max(cache_paths, key=lambda path: _cache_sort_key(path, item_map.get(path)))
        keep_metadata = item_map.get(keep_path) or {}
        keep_source = _normalize_source_path(keep_metadata.get("source_path"))
        keep_text = _infer_text_from_cache_filename(keep_path, keep_metadata)
        pending_seen = False
        for cache_path in cache_paths:
            if cache_path == keep_path:
                continue
            pending_seen = pending_seen or _is_pending_gemini(cache_path)
            removed += _remove_cache_entry(cache_path)
        if pending_seen and keep_text and keep_source and not _is_pending_gemini(keep_path):
            keep_backend = str(keep_metadata.get("backend") or "").strip().lower()
            keep_desired = str(keep_metadata.get("desired_backend") or "").strip().lower()
            if keep_backend != keep_desired and _is_online_backend(keep_desired):
                _enqueue_existing_cache_for_gemini_replacement(keep_text, keep_path, source_path=keep_source)
    return removed


def _normalize_cache_metadata_texts():
    roots = []
    if os.path.isdir(SHARED_WORD_CACHE_DIR):
        roots.append(("shared", SHARED_WORD_CACHE_DIR))
    if os.path.isdir(SOURCE_WORD_CACHE_ROOT_DIR):
        roots.append(("source", SOURCE_WORD_CACHE_ROOT_DIR))
    if not roots:
        return 0

    updated = 0
    for root_kind, root_dir in roots:
        for root, _, names in os.walk(root_dir):
            for name in names:
                if not name.lower().endswith(".wav.json"):
                    continue
                cache_path = os.path.join(root, name[:-5])
                metadata = _load_cache_metadata(cache_path)
                if not isinstance(metadata, dict):
                    continue
                normalized_guess = _infer_text_from_cache_filename(cache_path, {})
                if not normalized_guess:
                    continue
                payload = dict(metadata)
                current_text = _normalize_text(payload.get("text"), ensure_sentence_end=False)
                if current_text == normalized_guess and payload.get("text") == normalized_guess:
                    continue
                payload["text"] = normalized_guess
                if root_kind == "shared":
                    payload["source_path"] = "shared"
                _save_cache_metadata(cache_path, payload)
                updated += 1
    return updated


def _write_json_file(path, payload):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as fp:
        json.dump(payload, fp, ensure_ascii=False, indent=2)


def _load_pending_queue_disk_payload():
    data = _load_json_file(PENDING_ONLINE_TTS_QUEUE_PATH, [])
    if isinstance(data, list) and data:
        return data
    legacy = _load_json_file(LEGACY_PENDING_GEMINI_QUEUE_PATH, [])
    if isinstance(legacy, list) and legacy:
        return legacy
    return data if isinstance(data, list) else []


def _migrate_pending_queue_path():
    legacy_exists = os.path.exists(LEGACY_PENDING_GEMINI_QUEUE_PATH)
    new_exists = os.path.exists(PENDING_ONLINE_TTS_QUEUE_PATH)
    if legacy_exists and not new_exists:
        try:
            os.makedirs(os.path.dirname(PENDING_ONLINE_TTS_QUEUE_PATH), exist_ok=True)
            shutil.move(LEGACY_PENDING_GEMINI_QUEUE_PATH, PENDING_ONLINE_TTS_QUEUE_PATH)
            return
        except Exception:
            pass
    if legacy_exists:
        try:
            os.remove(LEGACY_PENDING_GEMINI_QUEUE_PATH)
        except Exception:
            pass


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
        original_cache_path = str(payload.get("cache_path") or "").strip()
        source_value = _normalize_source_path(payload.get("source_path"))
        if source_value != payload.get("source_path"):
            payload["source_path"] = source_value
        linked_shared = str(payload.get("linked_shared_path") or "").strip()
        if linked_shared:
            payload["linked_shared_path"] = _canonicalize_cache_path(linked_shared, metadata=payload)
        payload["cache_path"] = canonical_path
        if (
            original_cache_path != canonical_path
            or payload.get("source_path") != data.get("source_path")
            or payload.get("linked_shared_path") != data.get("linked_shared_path")
        ):
            _write_json_file(_cache_meta_path(canonical_path), payload)
    with _cache_metadata_lock:
        _cache_metadata_memory[key] = dict(payload)
    return payload


def _save_cache_metadata(cache_path, metadata):
    cache_path = _canonicalize_cache_path(cache_path)
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


def _best_shared_online_cache(text, *, preferred_provider=None):
    normalized_text = _normalize_text(text, ensure_sentence_end=False)
    if not normalized_text:
        return ("", {}, "")
    for provider in _online_provider_candidates(preferred_provider):
        shared_cache_path = _provider_shared_word_cache_path(normalized_text, provider=provider)
        metadata = _load_cache_metadata(shared_cache_path)
        if _resolve_cache_audio_path(shared_cache_path):
            return (shared_cache_path, metadata if isinstance(metadata, dict) else {}, provider)
    return ("", {}, "")


def _write_source_alias_metadata(text, source_path, shared_path, *, backend=None, desired_backend=None, cache_path=None):
    source_cache_path = _canonicalize_cache_path(
        cache_path,
        text=text,
        source_path=source_path,
        metadata={"text": text, "source_path": source_path, "desired_backend": desired_backend, "backend": backend},
    ) if cache_path else _word_cache_path(text, source_path=source_path)
    actual_backend = _backend_key(fallback_backend=backend or _current_online_provider())
    wanted_backend = _backend_key(fallback_backend=desired_backend or actual_backend)
    metadata = {
        "text": _normalize_text(text, ensure_sentence_end=False),
        "backend": actual_backend,
        "desired_backend": wanted_backend,
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


def _shared_cache_target_provider(*, backend=None, desired_backend=None):
    actual_backend = _backend_key(fallback_backend=backend)
    wanted_backend = _backend_key(fallback_backend=desired_backend or actual_backend)
    if _is_online_backend(actual_backend):
        return actual_backend
    if _is_online_backend(wanted_backend):
        return wanted_backend
    if actual_backend in {"piper", "kokoro"}:
        return actual_backend
    if wanted_backend in {"piper", "kokoro"}:
        return wanted_backend
    return ""


def _ensure_shared_cache_from_playable(
    text,
    playable_path,
    *,
    backend=None,
    desired_backend=None,
    metadata=None,
):
    normalized_text = _normalize_text(text, ensure_sentence_end=False)
    if not normalized_text or not playable_path or not os.path.exists(playable_path):
        return ""
    shared_provider = _shared_cache_target_provider(
        backend=backend,
        desired_backend=desired_backend,
    )
    if not shared_provider:
        return ""
    shared_cache_path = _provider_shared_word_cache_path(normalized_text, provider=shared_provider)
    shared_payload = dict(metadata or {})
    shared_payload["text"] = normalized_text
    shared_payload["backend"] = _backend_key(fallback_backend=backend or shared_payload.get("backend"))
    shared_payload["desired_backend"] = _backend_key(
        fallback_backend=desired_backend or shared_payload.get("desired_backend") or shared_payload["backend"]
    )
    shared_payload["source_path"] = "shared"
    os.makedirs(os.path.dirname(shared_cache_path), exist_ok=True)
    if os.path.abspath(playable_path) != os.path.abspath(shared_cache_path):
        shutil.copyfile(playable_path, shared_cache_path)
    shared_payload["updated_at"] = int(os.path.getmtime(shared_cache_path)) if os.path.exists(shared_cache_path) else None
    _save_cache_metadata(shared_cache_path, shared_payload)
    return shared_cache_path


def _alias_source_cache_to_shared(
    text,
    *,
    source_path=None,
    shared_path="",
    backend=None,
    desired_backend=None,
    metadata=None,
    cache_path=None,
):
    normalized_text = _normalize_text(text, ensure_sentence_end=False)
    if not normalized_text:
        return ""
    resolved_shared = str(shared_path or "").strip()
    if not resolved_shared:
        source_cache_path = _canonicalize_cache_path(
            cache_path,
            text=normalized_text,
            source_path=source_path,
            metadata=metadata,
        ) if cache_path else _word_cache_path(normalized_text, source_path=source_path)
        source_metadata = dict(metadata or _load_cache_metadata(source_cache_path) or {})
        playable_path = _resolve_cache_audio_path(source_cache_path)
        resolved_shared = _ensure_shared_cache_from_playable(
            normalized_text,
            playable_path,
            backend=backend or source_metadata.get("backend"),
            desired_backend=desired_backend or source_metadata.get("desired_backend"),
            metadata=source_metadata,
        )
    if not resolved_shared:
        return ""
    return _write_source_alias_metadata(
        normalized_text,
        source_path,
        resolved_shared,
        backend=backend,
        desired_backend=desired_backend,
        cache_path=cache_path,
    )


def _update_source_cache_desired_backend(text, source_path, *, backend=None, desired_backend=None, linked_shared_path=None):
    source_cache_path = _word_cache_path(text, source_path=source_path)
    metadata = _load_cache_metadata(source_cache_path)
    payload = dict(metadata) if isinstance(metadata, dict) else {}
    payload["text"] = _normalize_text(text, ensure_sentence_end=False)
    payload["source_path"] = str(_normalize_source_path(source_path) or "").strip() or None
    if backend:
        payload["backend"] = _backend_key(fallback_backend=backend)
    if desired_backend:
        payload["desired_backend"] = _backend_key(fallback_backend=desired_backend)
    if linked_shared_path:
        payload["linked_shared_path"] = linked_shared_path
    _save_cache_metadata(source_cache_path, payload)
    return source_cache_path


def _collapse_source_cache_to_alias(text, source_path=None):
    if _normalize_source_path(source_path) == "shared":
        return False
    source_cache_path = _word_cache_path(text, source_path=source_path)
    metadata = _load_cache_metadata(source_cache_path)
    shared_cache_path = str(metadata.get("linked_shared_path") or "").strip()
    if not shared_cache_path:
        shared_cache_path, _shared_meta, _shared_provider = _best_shared_online_cache(
            text,
            preferred_provider=str(metadata.get("desired_backend") or _current_online_provider()),
        )
    if not shared_cache_path:
        return False
    try:
        _alias_source_cache_to_shared(
            text,
            source_path=source_path,
            shared_path=shared_cache_path,
            backend=str(metadata.get("backend") or _current_online_provider()),
            desired_backend=str(metadata.get("desired_backend") or _current_online_provider()),
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
    if _normalize_source_path(source_path) == "shared":
        return False
    source_cache_path = _word_cache_path(text, source_path=source_path)
    shared_cache_path, metadata, shared_provider = _best_shared_online_cache(
        text,
        preferred_provider=_current_online_provider(),
    )
    if not shared_cache_path:
        return False
    try:
        _alias_source_cache_to_shared(
            text,
            source_path=source_path,
            shared_path=shared_cache_path,
            backend=str(metadata.get("backend") or shared_provider or _current_online_provider()),
            desired_backend=str(metadata.get("desired_backend") or _current_online_provider()),
            metadata=metadata,
        )
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
                "desired_backend": str(item.get("desired_backend") or _current_online_provider()).strip().lower(),
            }
        )
    _write_json_file(PENDING_ONLINE_TTS_QUEUE_PATH, payload)
    try:
        if os.path.exists(LEGACY_PENDING_GEMINI_QUEUE_PATH):
            os.remove(LEGACY_PENDING_GEMINI_QUEUE_PATH)
    except Exception:
        pass


def _load_pending_gemini_queue():
    items = _load_pending_queue_disk_payload()
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
                "desired_backend": str(item.get("desired_backend") or _current_online_provider()).strip().lower(),
            }
        if _dedupe_pending_gemini_locked(preferred_provider=_current_online_provider()):
            changed = True
        if changed:
            _save_pending_gemini_queue_locked()
    _refresh_gemini_queue_status_counts()


def _dedupe_pending_gemini_locked(preferred_provider=None):
    preferred_provider = str(preferred_provider or _current_online_provider() or "").strip().lower()
    preferred_provider = preferred_provider if preferred_provider in {"gemini", "elevenlabs"} else ""
    changed = False
    deduped = {}

    def _entry_key(item):
        source_key = _normalize_source_path((item or {}).get("source_path")) or str((item or {}).get("source_path") or "").strip()
        text_key = _normalize_text((item or {}).get("text"), ensure_sentence_end=False)
        return (source_key, text_key)

    def _score(item):
        item = item if isinstance(item, dict) else {}
        desired = str(item.get("desired_backend") or "").strip().lower()
        created = int(item.get("created_at") or 0)
        return (
            1 if preferred_provider and desired == preferred_provider else 0,
            1 if _is_online_backend(desired) else 0,
            created,
        )

    for cache_path, item in list(_pending_gemini_replacements.items()):
        if not isinstance(item, dict):
            continue
        key = _entry_key(item)
        if not key[1]:
            continue
        existing = deduped.get(key)
        if existing is None:
            deduped[key] = (cache_path, item)
            continue
        existing_path, existing_item = existing
        if _score(item) >= _score(existing_item):
            deduped[key] = (cache_path, item)
            if existing_path != cache_path:
                changed = True
        else:
            changed = True

    if changed:
        _pending_gemini_replacements.clear()
        for cache_path, item in deduped.values():
            _pending_gemini_replacements[cache_path] = item
    return changed


def dedupe_pending_online_queue(preferred_provider=None):
    with _pending_gemini_lock:
        changed = _dedupe_pending_gemini_locked(preferred_provider=preferred_provider)
        if changed:
            _save_pending_gemini_queue_locked()
    _refresh_gemini_queue_status_counts()
    return changed


def _remove_pending_gemini(cache_path):
    with _pending_gemini_lock:
        if cache_path in _pending_gemini_replacements:
            _pending_gemini_replacements.pop(cache_path, None)
            _save_pending_gemini_queue_locked()
    _refresh_gemini_queue_status_counts()


def _is_pending_gemini(cache_path):
    with _pending_gemini_lock:
        return cache_path in _pending_gemini_replacements


def set_preferred_pending_source(source_path=None):
    global _preferred_pending_source
    _preferred_pending_source = _normalize_source_path(source_path)


def _next_pending_gemini_item():
    preferred_source = _preferred_pending_source
    with _pending_gemini_lock:
        pending_items = list(_pending_gemini_replacements.items())
    if not pending_items:
        return None

    def _sort_key(entry):
        cache_path, item = entry
        item = item if isinstance(item, dict) else {}
        source_value = _normalize_source_path(item.get("source_path"))
        created_at = int(item.get("created_at") or 0)
        preferred_rank = 0 if preferred_source and source_value == preferred_source else 1
        return (preferred_rank, created_at, str(cache_path or ""))

    return min(pending_items, key=_sort_key)


def _move_pending_gemini_entry(old_cache_path, new_cache_path, *, text=None, source_path=None):
    old_cache_path = _canonicalize_cache_path(old_cache_path, text=text, source_path=None)
    new_cache_path = _canonicalize_cache_path(new_cache_path, text=text, source_path=source_path)
    if not old_cache_path or not new_cache_path:
        return False
    with _pending_gemini_lock:
        existing = _pending_gemini_replacements.pop(old_cache_path, None)
        if not isinstance(existing, dict):
            return False
        existing["text"] = _normalize_text(text or existing.get("text"), ensure_sentence_end=False)
        existing["source_path"] = str(source_path or "").strip() or None
        _pending_gemini_replacements[new_cache_path] = existing
        _save_pending_gemini_queue_locked()
    _refresh_gemini_queue_status_counts()
    return True


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
            "desired_backend": _current_online_provider(),
        }
        _dedupe_pending_gemini_locked(preferred_provider=_current_online_provider())
        _save_pending_gemini_queue_locked()
    _refresh_gemini_queue_status_counts()
    _start_pending_gemini_worker()


def _cache_requires_gemini_replacement(cache_path):
    metadata = _load_cache_metadata(cache_path)
    backend = str(metadata.get("backend") or "").strip().lower()
    desired_backend = str(metadata.get("desired_backend") or _current_online_provider()).strip().lower()
    if _is_online_backend(desired_backend):
        return backend != desired_backend
    return backend != _current_online_provider()


def _cache_requires_online_replacement(cache_path):
    return _cache_requires_gemini_replacement(cache_path)


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

    shared_cache_path = _ensure_shared_cache_from_playable(
        text or cache_path,
        wav_path,
        backend=actual_backend,
        desired_backend=wanted_backend,
        metadata=payload,
    )
    if normalized_source != "shared" and shared_cache_path:
        _alias_source_cache_to_shared(
            text or cache_path,
            source_path=source_path,
            shared_path=shared_cache_path,
            backend=actual_backend,
            desired_backend=wanted_backend,
            metadata=payload,
            cache_path=cache_path,
        )
        if not str(source_path or "").strip():
            with _manual_session_cache_lock:
                _manual_session_cache_paths.add(cache_path)
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


def rebind_manual_session_cache_to_source(texts, source_path):
    target_source = _normalize_source_path(source_path)
    if not target_source or target_source in {RECENT_WRONG_SOURCE_KEY, MANUAL_SESSION_SOURCE_KEY}:
        return {"migrated": 0, "queued": 0, "removed": 0}

    moved_cache = 0
    moved_queue = 0
    removed_manual = 0
    seen = set()

    for raw_text in texts or []:
        normalized = _normalize_text(raw_text, ensure_sentence_end=False)
        if not normalized:
            continue
        folded = normalized.casefold()
        if folded in seen:
            continue
        seen.add(folded)

        target_cache = _word_cache_path(normalized, source_path=target_source)
        shared_cache = _shared_word_cache_path(normalized)
        manual_candidates = []
        for manual_source in (None, MANUAL_SESSION_SOURCE_KEY):
            candidate = _word_cache_path(normalized, source_path=manual_source)
            if candidate not in manual_candidates:
                manual_candidates.append(candidate)

        manual_cache = ""
        manual_metadata = {}
        manual_audio = ""
        manual_pending = False
        for candidate in manual_candidates:
            candidate_metadata = _load_cache_metadata(candidate)
            candidate_audio = _resolve_cache_audio_path(candidate)
            candidate_pending = _is_pending_gemini(candidate)
            if candidate_audio or candidate_metadata or candidate_pending:
                manual_cache = candidate
                manual_metadata = candidate_metadata
                manual_audio = candidate_audio
                manual_pending = candidate_pending
                break
        if not manual_cache:
            manual_cache = manual_candidates[0]

        target_has_valid_gemini = _has_valid_gemini_cache(target_cache)
        target_has_any_audio = bool(_resolve_cache_audio_path(target_cache))

        if manual_pending:
            if not _is_pending_gemini(target_cache):
                if _move_pending_gemini_entry(
                    manual_cache,
                    target_cache,
                    text=normalized,
                    source_path=target_source,
                ):
                    moved_queue += 1
            else:
                _remove_pending_gemini(manual_cache)

        should_copy = bool(manual_audio) and (
            not target_has_any_audio
            or (
                not target_has_valid_gemini
                and _is_online_backend(str(manual_metadata.get("desired_backend") or "").strip().lower())
            )
        )
        if should_copy:
            payload = dict(manual_metadata or {})
            payload["text"] = normalized
            payload["source_path"] = target_source
            if "backend" not in payload:
                payload["backend"] = "unknown"
            if "desired_backend" not in payload:
                payload["desired_backend"] = _current_online_provider()
            backend = str(payload.get("backend") or "").strip().lower()
            desired_backend = str(payload.get("desired_backend") or "").strip().lower() or _current_online_provider()
            try:
                if _is_online_backend(backend) and _has_valid_gemini_cache(shared_cache):
                    _write_source_alias_metadata(
                        normalized,
                        target_source,
                        shared_cache,
                        backend=backend,
                        desired_backend=desired_backend,
                    )
                else:
                    _copy_cache_file(manual_audio, target_cache, metadata=payload)
                moved_cache += 1
            except Exception:
                pass
        elif (
            not target_has_any_audio
            and _is_online_backend(str(manual_metadata.get("backend") or "").strip().lower())
            and _has_valid_gemini_cache(shared_cache)
        ):
            try:
                _write_source_alias_metadata(
                    normalized,
                    target_source,
                    shared_cache,
                    backend=str(manual_metadata.get("backend") or _current_online_provider()),
                    desired_backend=str(manual_metadata.get("desired_backend") or _current_online_provider()),
                )
                moved_cache += 1
            except Exception:
                pass

        # Clear the manual-session entry once the saved file has taken ownership.
        try:
            if os.path.exists(manual_cache):
                os.remove(manual_cache)
                removed_manual += 1
        except Exception:
            pass
        if manual_metadata:
            _remove_cache_metadata(manual_cache)
            removed_manual += 1
        with _manual_session_cache_lock:
            for candidate in manual_candidates:
                _manual_session_cache_paths.discard(candidate)

    _cleanup_duplicate_source_cache_entries()
    _collapse_existing_lightweight_source_caches()
    return {"migrated": moved_cache, "queued": moved_queue, "removed": removed_manual}


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


def cleanup_word_audio_cache(text, source_path=None):
    normalized = _normalize_text(text, ensure_sentence_end=False)
    if not normalized:
        return False
    cache_path = _word_cache_path(normalized, source_path=source_path)
    removed = False
    try:
        if os.path.exists(cache_path):
            os.remove(cache_path)
            removed = True
    except Exception:
        pass
    metadata = _load_cache_metadata(cache_path)
    if metadata:
        _remove_cache_metadata(cache_path)
        removed = True
    if _is_pending_gemini(cache_path):
        _remove_pending_gemini(cache_path)
        removed = True
    if not str(source_path or "").strip():
        with _manual_session_cache_lock:
            _manual_session_cache_paths.discard(cache_path)
    return removed


def rename_cache_source_path(old_source_path, new_source_path):
    old_source = _normalize_source_path(old_source_path)
    new_source = _normalize_source_path(new_source_path)
    if not old_source or not new_source or old_source == new_source:
        return {"migrated": 0, "queued": 0}

    old_dir = _source_word_cache_dir(source_path=old_source)
    moved_count = 0
    queued_count = 0
    migrated_pending = {}

    if os.path.isdir(old_dir):
        cache_paths = set()
        for root, _, names in os.walk(old_dir):
            for name in names:
                lower_name = name.lower()
                if lower_name.endswith(".wav"):
                    cache_paths.add(os.path.join(root, name))
                elif lower_name.endswith(".wav.json"):
                    cache_paths.add(os.path.join(root, name[:-5]))

        for old_cache_path in sorted(cache_paths):
            old_metadata = _load_cache_metadata(old_cache_path)
            text_value = _normalize_text((old_metadata or {}).get("text"), ensure_sentence_end=False)
            if text_value:
                new_cache_path = _word_cache_path(text_value, source_path=new_source)
            else:
                rel_path = os.path.relpath(old_cache_path, old_dir)
                new_cache_path = os.path.join(_source_word_cache_dir(source_path=new_source), rel_path)
            new_metadata = dict(old_metadata or {})
            new_metadata["source_path"] = new_source
            linked_shared = str(new_metadata.get("linked_shared_path") or "").strip()
            old_audio_exists = os.path.exists(old_cache_path)
            try:
                if linked_shared and not old_audio_exists:
                    if text_value:
                        _write_source_alias_metadata(
                            text_value,
                            new_source,
                            linked_shared,
                            backend=str(new_metadata.get("backend") or _current_online_provider()),
                            desired_backend=str(new_metadata.get("desired_backend") or _current_online_provider()),
                        )
                        moved_count += 1
                elif old_audio_exists:
                    _copy_cache_file(old_cache_path, new_cache_path, metadata=new_metadata)
                    moved_count += 1
                elif new_metadata:
                    _save_cache_metadata(new_cache_path, new_metadata)
                    moved_count += 1
            except Exception:
                continue

            if _is_pending_gemini(old_cache_path):
                pending_item = dict(_pending_gemini_replacements.get(old_cache_path) or {})
                if pending_item:
                    pending_item["source_path"] = new_source
                    migrated_pending[new_cache_path] = pending_item
                    queued_count += 1
                _remove_pending_gemini(old_cache_path)

            try:
                if os.path.exists(old_cache_path):
                    os.remove(old_cache_path)
            except Exception:
                pass
            _remove_cache_metadata(old_cache_path)

        for root, _, _ in os.walk(old_dir, topdown=False):
            try:
                if os.path.isdir(root) and not os.listdir(root):
                    os.rmdir(root)
            except Exception:
                pass
        try:
            if os.path.isdir(old_dir) and not os.listdir(old_dir):
                os.rmdir(old_dir)
        except Exception:
            pass

    with _pending_gemini_lock:
        updated = False
        for cache_path, item in list(_pending_gemini_replacements.items()):
            if not isinstance(item, dict):
                continue
            if _normalize_source_path(item.get("source_path")) != old_source:
                continue
            text_value = _normalize_text(item.get("text"), ensure_sentence_end=False)
            if text_value:
                new_cache_path = _word_cache_path(text_value, source_path=new_source)
            else:
                old_dir = _source_word_cache_dir(source_path=old_source)
                rel_path = os.path.relpath(cache_path, old_dir) if old_dir and os.path.commonpath([old_dir, cache_path]) == old_dir else os.path.basename(cache_path)
                new_cache_path = os.path.join(_source_word_cache_dir(source_path=new_source), rel_path)
            migrated_pending[new_cache_path] = {
                "text": text_value or item.get("text"),
                "source_path": new_source,
                "created_at": int(item.get("created_at") or time.time()),
            }
            _pending_gemini_replacements.pop(cache_path, None)
            queued_count += 1
            updated = True
        for cache_path, item in migrated_pending.items():
            _pending_gemini_replacements[cache_path] = item
            updated = True
        if updated:
            _save_pending_gemini_queue_locked()

    _cleanup_duplicate_source_cache_entries()
    _collapse_existing_lightweight_source_caches()
    return {"migrated": moved_count, "queued": queued_count}


def _is_gemini_rate_limited_error(error):
    message = str(error or "").lower()
    return (
        "rate limit" in message
        or "resource_exhausted" in message
        or "429" in message
        or "quota exceeded" in message
        or "too many requests" in message
    )


def _clone_to_temp(path):
    fd, temp_path = tempfile.mkstemp(prefix="wordspeaker_", suffix=".wav")
    os.close(fd)
    shutil.copyfile(path, temp_path)
    return temp_path


def _ensure_source_gemini_cache(text, source_path=None):
    source_cache_path = _migrate_legacy_cache_if_needed(text, source_path=source_path)
    if _has_valid_gemini_cache(source_cache_path):
        metadata = _load_cache_metadata(source_cache_path)
        backend = str(metadata.get("backend") or "").strip().lower()
        desired_backend = str(metadata.get("desired_backend") or backend or _current_online_provider()).strip().lower()
        if not _is_online_backend(backend):
            backend = _current_online_provider()
        if not _is_online_backend(desired_backend):
            desired_backend = _current_online_provider()
        shared_path = _provider_shared_word_cache_path(text, provider=backend)
        if not _has_valid_gemini_cache(shared_path):
            try:
                shared_payload = dict(metadata) if isinstance(metadata, dict) else {}
                shared_payload["source_path"] = "shared"
                actual_source_path = _resolve_cache_audio_path(source_cache_path)
                if actual_source_path:
                    _copy_cache_file(actual_source_path, shared_path, metadata=shared_payload)
            except Exception:
                pass
        current_provider = _current_online_provider()
        if backend != current_provider:
            try:
                if source_path != "shared" and _has_valid_gemini_cache(shared_path):
                    _alias_source_cache_to_shared(
                        text,
                        source_path=source_path,
                        shared_path=shared_path,
                        backend=backend,
                        desired_backend=current_provider,
                        metadata=metadata,
                        cache_path=source_cache_path,
                    )
                else:
                    _update_source_cache_desired_backend(
                        text,
                        source_path,
                        backend=backend,
                        desired_backend=current_provider,
                        linked_shared_path=shared_path if _has_valid_gemini_cache(shared_path) else None,
                    )
                _enqueue_existing_cache_for_gemini_replacement(text, source_cache_path, source_path=source_path)
            except Exception:
                pass
            return source_cache_path
        if source_path != "shared" and _has_valid_gemini_cache(shared_path):
            try:
                _collapse_source_cache_to_alias(text, source_path=source_path)
            except Exception:
                pass
        return source_cache_path
    if _hydrate_source_cache_from_shared(text, source_path=source_path):
        return source_cache_path
    shared_cache_path, shared_metadata, shared_provider = _best_shared_online_cache(
        text,
        preferred_provider=_current_online_provider(),
    )
    if shared_cache_path:
        try:
            _alias_source_cache_to_shared(
                text,
                source_path=source_path,
                shared_path=shared_cache_path,
                backend=str(shared_metadata.get("backend") or shared_provider),
                desired_backend=_current_online_provider(),
                metadata=shared_metadata,
                cache_path=source_cache_path,
            )
            _enqueue_existing_cache_for_gemini_replacement(text, source_cache_path, source_path=source_path)
            return source_cache_path
        except Exception:
            pass
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
    shared_path, _shared_meta, _shared_provider = _best_shared_online_cache(
        text,
        preferred_provider=desired_backend or _current_online_provider(),
    )
    return {
        "exists": exists,
        "cache_path": cache_path,
        "playable_cache_path": playable_cache or cache_path,
        "meta_path": _cache_meta_path(cache_path),
        "shared_cache_path": shared_path,
        "shared_exists": bool(shared_path and os.path.exists(shared_path)),
        "uses_shared_cache": bool(shared_path and playable_cache and os.path.abspath(playable_cache) == os.path.abspath(shared_path) and os.path.abspath(cache_path) != os.path.abspath(shared_path)),
        "backend": backend or None,
        "backend_label": _backend_label_from_key(backend) if backend else "",
        "desired_backend": (desired_backend or (_current_online_provider() if pending else None)),
        "desired_backend_label": _backend_label_from_key(desired_backend or _current_online_provider()) if (desired_backend or pending) else "",
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
    source_metadata = _load_cache_metadata(source_cache)
    linked_shared = str(source_metadata.get("linked_shared_path") or "").strip()
    shared_cache_path = linked_shared
    if not shared_cache_path:
        playable_source_cache = _resolve_cache_audio_path(source_cache)
        if not playable_source_cache:
            return False
        shared_cache_path = _ensure_shared_cache_from_playable(
            text,
            playable_source_cache,
            backend=source_metadata.get("backend"),
            desired_backend=source_metadata.get("desired_backend"),
            metadata=source_metadata,
        )
    if not shared_cache_path:
        return False
    try:
        _write_source_alias_metadata(
            text,
            to_source_path,
            shared_cache_path,
            backend=source_metadata.get("backend") or "unknown",
            desired_backend=source_metadata.get("desired_backend") or _current_online_provider(),
        )
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


def _play_wav_async(path):
    last_error = None
    for attempt in range(2):
        try:
            try:
                winsound.PlaySound(None, winsound.SND_PURGE)
            except Exception:
                pass
            if attempt:
                time.sleep(0.08)
            else:
                time.sleep(0.03)
            winsound.PlaySound(
                path,
                winsound.SND_FILENAME | winsound.SND_ASYNC | winsound.SND_NODEFAULT,
            )
            return
        except Exception as e:
            last_error = e
            time.sleep(0.05)
    if last_error:
        raise last_error


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


def _request_gemini_tts(text, *, short_text, api_key=None):
    api_key = str(api_key or get_tts_api_key()).strip()
    if not api_key:
        raise RuntimeError("TTS API key is empty.")
    spoken_text = _normalize_tts_spoken_text(text)

    prompt = TTS_STYLE_SHORT if short_text else TTS_STYLE_LONG
    payload = {
        "contents": [{"parts": [{"text": f"{prompt}\n\nText:\n{spoken_text}"}]}],
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


def _request_elevenlabs_tts(text, *, short_text, api_key=None):
    api_key = str(api_key or get_tts_api_key()).strip()
    if not api_key:
        raise RuntimeError("TTS API key is empty.")
    spoken_text = _normalize_tts_spoken_text(text)

    payload = {
        "text": spoken_text,
        "model_id": ELEVENLABS_MODEL_ID,
        "voice_settings": {
            "stability": 0.45,
            "similarity_boost": 0.75,
            "style": 0.1 if short_text else 0.25,
            "use_speaker_boost": True,
        },
    }
    req = urllib.request.Request(
        ELEVENLABS_URL,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "xi-api-key": api_key,
            "Content-Type": "application/json",
            "Accept": "audio/pcm",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=180) as response:
            return response.read()
    except urllib.error.HTTPError as e:
        raise RuntimeError(_extract_error_message(e)) from e
    except urllib.error.URLError as e:
        raise RuntimeError(f"ElevenLabs TTS request failed: {e.reason}") from e
    except Exception as e:
        raise RuntimeError(f"ElevenLabs TTS request failed: {e}") from e


def _request_online_tts(text, *, short_text, provider=None, api_key=None):
    backend = str(provider or _primary_online_provider()).strip().lower()
    if backend == "elevenlabs":
        return _request_elevenlabs_tts(text, short_text=short_text, api_key=api_key), "elevenlabs"
    return _request_gemini_tts(text, short_text=short_text, api_key=api_key), "gemini"


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


def validate_gemini_tts_api_key(api_key, timeout=25):
    if not str(api_key or "").strip():
        raise RuntimeError("TTS API key is empty.")
    prompt = TTS_STYLE_SHORT
    payload = {
        "contents": [{"parts": [{"text": f"{prompt}\n\nText:\nTest"}]}],
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
            "x-goog-api-key": str(api_key).strip(),
            "Content-Type": "application/json",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as response:
            data = json.loads(response.read().decode("utf-8", errors="ignore"))
    except urllib.error.HTTPError as e:
        raise RuntimeError(_extract_error_message(e)) from e
    except urllib.error.URLError as e:
        raise RuntimeError(f"Gemini TTS request failed: {e.reason}") from e
    except Exception as e:
        raise RuntimeError(f"Gemini TTS request failed: {e}") from e
    _extract_pcm_bytes(data)
    return True


def validate_elevenlabs_tts_api_key(api_key, timeout=25):
    if not str(api_key or "").strip():
        raise RuntimeError("TTS API key is empty.")
    payload = {
        "text": "Test",
        "model_id": ELEVENLABS_MODEL_ID,
        "voice_settings": {
            "stability": 0.45,
            "similarity_boost": 0.75,
            "style": 0.1,
            "use_speaker_boost": True,
        },
    }
    req = urllib.request.Request(
        ELEVENLABS_URL,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "xi-api-key": str(api_key).strip(),
            "Content-Type": "application/json",
            "Accept": "audio/pcm",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as response:
            data = response.read()
    except urllib.error.HTTPError as e:
        raise RuntimeError(_extract_error_message(e)) from e
    except urllib.error.URLError as e:
        raise RuntimeError(f"ElevenLabs TTS request failed: {e.reason}") from e
    except Exception as e:
        raise RuntimeError(f"ElevenLabs TTS request failed: {e}") from e
    if not data:
        raise RuntimeError("ElevenLabs TTS returned no audio.")
    return True


def validate_tts_api_key(api_key, provider, timeout=25):
    backend = str(provider or "").strip().lower()
    if backend == "elevenlabs":
        return validate_elevenlabs_tts_api_key(api_key, timeout=timeout)
    return validate_gemini_tts_api_key(api_key, timeout=timeout)


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


def _synthesize_with_elevenlabs(text, volume, *, short_text):
    pcm_bytes = _request_elevenlabs_tts(text, short_text=short_text)
    return _write_pcm_to_wav_path(pcm_bytes, sample_rate=TTS_SAMPLE_RATE, volume=volume), "ElevenLabs TTS", True


def _synthesize_with_gemini_fallback(text, volume, *, short_text, api_key=None):
    data = _request_gemini_tts(text, short_text=short_text, api_key=api_key)
    pcm_bytes = _extract_pcm_bytes(data)
    return _write_pcm_to_wav_path(pcm_bytes, sample_rate=TTS_SAMPLE_RATE, volume=volume), "Gemini TTS", True


def _synthesize_with_elevenlabs_fallback(text, volume, *, short_text, api_key=None):
    pcm_bytes = _request_elevenlabs_tts(text, short_text=short_text, api_key=api_key)
    return _write_pcm_to_wav_path(pcm_bytes, sample_rate=TTS_SAMPLE_RATE, volume=volume), "ElevenLabs TTS", True


def _synthesize_with_online_provider(text, volume, *, short_text, provider, api_key=None):
    backend = str(provider or _primary_online_provider()).strip().lower()
    if backend == "elevenlabs":
        return _synthesize_with_elevenlabs_fallback(text, volume=volume, short_text=short_text, api_key=api_key)
    return _synthesize_with_gemini_fallback(text, volume=volume, short_text=short_text, api_key=api_key)


def _synthesize_with_online(text, volume, *, short_text):
    primary = _primary_online_provider()
    try:
        return _synthesize_with_online_provider(text, volume=volume, short_text=short_text, provider=primary), primary
    except Exception:
        secondary = _secondary_online_provider(primary)
        if secondary:
            fallback_key = get_llm_api_key() if secondary == "gemini" else get_tts_api_key()
            return (
                _synthesize_with_online_provider(
                    text,
                    volume=volume,
                    short_text=short_text,
                    provider=secondary,
                    api_key=fallback_key,
                ),
                secondary,
            )
        raise


def _synthesize_with_kokoro(text, volume, rate_ratio):
    kokoro = _ensure_kokoro()
    voice_id = get_voice_id()
    profile = get_voice_profile(get_voice_source(), voice_id)
    lang = str((profile.get("languages") or ["en-GB"])[0]).strip().lower().replace("_", "-")
    speed = _clamp(rate_ratio, 0.7, 1.4)
    audio, sample_rate = kokoro.create(_normalize_tts_spoken_text(text), voice=voice_id, speed=speed, lang=lang)
    return (
        _write_float_audio_to_wav_path(audio, sample_rate=sample_rate or KOKORO_SAMPLE_RATE, volume=volume),
        "Kokoro (Offline)",
        False,
    )


def _synthesize_with_kokoro_voice(text, volume, rate_ratio, voice_id, lang="en-gb"):
    kokoro = _ensure_kokoro()
    speed = _clamp(rate_ratio, 0.7, 1.4)
    audio, sample_rate = kokoro.create(_normalize_tts_spoken_text(text), voice=voice_id, speed=speed, lang=lang)
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
    audio_chunks = list(voice.synthesize(_normalize_tts_spoken_text(text), syn_config=syn_config))
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
        return _synthesize_with_user_online(text, volume=volume, short_text=short_text), False
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
        if _is_online_backend(_selected_source_backend_key()) and _has_valid_gemini_cache(cache_path):
            playable_cache = _resolve_cache_audio_path(cache_path) or cache_path
            if request_token is not None:
                _set_backend_status(request_token, _online_provider_label(), from_cache=True, fallback=False)
            return _clone_to_temp(playable_cache)
        if _is_online_backend(_selected_source_backend_key()) and os.path.exists(cache_path):
            metadata = _load_cache_metadata(cache_path)
            backend_key = str(metadata.get("backend") or "").strip().lower()
            desired_backend = str(metadata.get("desired_backend") or _current_online_provider()).strip().lower()
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
            if _is_online_backend(backend_key) and desired_backend != backend_key:
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
                        "desired_backend": _current_online_provider(),
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
        wav_path, backend_label, _can_cache = _synthesize_with_user_online(
            normalized,
            volume=volume,
            short_text=short_text,
        )
        actual_online_backend = _backend_key_from_label(backend_label)
        desired_online_backend = _primary_online_provider()
        if request_token is not None:
            _set_backend_status(request_token, backend_label, from_cache=False, fallback=False)
        if short_text:
            try:
                _save_word_cache_file(
                    cache_path,
                    wav_path,
                    text=normalized,
                    source_path=source_path,
                    backend=actual_online_backend,
                    desired_backend=desired_online_backend,
                )
                if actual_online_backend != desired_online_backend:
                    _enqueue_existing_cache_for_gemini_replacement(text, cache_path, source_path=source_path)
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
                        desired_backend=_current_online_provider(),
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
    _set_gemini_queue_status(worker_running=True)
    _refresh_gemini_queue_status_counts()
    with _gemini_queue_status_lock:
        if _gemini_queue_status.get("state") in {"idle", ""}:
            _gemini_queue_status["state"] = "ok"

    def _worker():
        global _pending_gemini_worker_running
        try:
            while True:
                next_item = _next_pending_gemini_item()
                if not next_item:
                    return
                cache_path_local, item = next_item
                if not isinstance(item, dict):
                    _remove_pending_gemini(cache_path_local)
                    continue
                normalized_text = _normalize_text(item.get("text"), ensure_sentence_end=False)
                source_path = str(item.get("source_path") or "").strip() or None
                if not normalized_text:
                    _remove_pending_gemini(cache_path_local)
                    continue
                desired_backend = str(item.get("desired_backend") or _current_online_provider()).strip().lower()
                if desired_backend not in {"gemini", "elevenlabs"}:
                    desired_backend = _current_online_provider()
                try:
                    _wait_for_gemini_queue_slot(provider=desired_backend)
                    wav_path, _label, _can_cache = (
                        _synthesize_with_elevenlabs(normalized_text, volume=1.0, short_text=True)
                        if desired_backend == "elevenlabs"
                        else _synthesize_with_gemini(normalized_text, volume=1.0, short_text=True)
                    )
                    _record_queue_success(desired_backend)
                    _set_gemini_queue_status(
                        state="ok",
                        next_retry_at=0.0,
                        last_success_at=time.time(),
                        last_error="",
                    )
                    try:
                        _save_word_cache_file(
                            cache_path_local,
                            wav_path,
                            text=normalized_text,
                            source_path=source_path,
                            backend=desired_backend,
                            desired_backend=desired_backend,
                        )
                    finally:
                        try:
                            os.remove(wav_path)
                        except Exception:
                            pass
                    _remove_pending_gemini(cache_path_local)
                except Exception as exc:
                    secondary_backend = _secondary_online_provider(desired_backend)
                    primary_rate_limited = _is_gemini_rate_limited_error(exc)
                    if primary_rate_limited:
                        _record_queue_rate_limit(desired_backend)
                    else:
                        _record_queue_soft_failure(desired_backend)
                    if secondary_backend:
                        try:
                            fallback_key = get_llm_api_key() if secondary_backend == "gemini" else get_tts_api_key()
                            wav_path, _label, _can_cache = _synthesize_with_online_provider(
                                normalized_text,
                                volume=1.0,
                                short_text=True,
                                provider=secondary_backend,
                                api_key=fallback_key,
                            )
                            try:
                                _save_word_cache_file(
                                    cache_path_local,
                                    wav_path,
                                    text=normalized_text,
                                    source_path=source_path,
                                    backend=secondary_backend,
                                    desired_backend=desired_backend,
                                )
                            finally:
                                try:
                                    os.remove(wav_path)
                                except Exception:
                                    pass
                            _record_queue_success(secondary_backend)
                            if primary_rate_limited:
                                cooldown_seconds = _rate_limit_cooldown_for_provider(desired_backend)
                                _defer_gemini_queue(
                                    cooldown_seconds,
                                    state="rate_limited",
                                    provider=desired_backend,
                                )
                                _set_gemini_queue_status(
                                    last_error=str(exc),
                                    next_retry_at=time.time() + cooldown_seconds,
                                )
                            else:
                                _set_gemini_queue_status(
                                    state="ok",
                                    next_retry_at=time.time() + _queue_interval_for_provider(desired_backend),
                                    last_success_at=time.time(),
                                    last_error="",
                                )
                            continue
                        except Exception:
                            _record_queue_soft_failure(secondary_backend)
                            pass
                    if primary_rate_limited:
                        cooldown_seconds = _rate_limit_cooldown_for_provider(desired_backend)
                        _set_gemini_queue_status(
                            state="rate_limited",
                            next_retry_at=time.time() + cooldown_seconds,
                            last_error=str(exc),
                        )
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
                                        desired_backend=desired_backend,
                                    )
                                finally:
                                    try:
                                        os.remove(wav_path)
                                    except Exception:
                                        pass
                            except Exception:
                                pass
                        threading.Event().wait(cooldown_seconds)
                        continue
                    _set_gemini_queue_status(
                        state="error",
                        next_retry_at=0.0,
                        last_error=str(exc),
                    )
                    _remove_pending_gemini(cache_path_local)
        finally:
            with _pending_gemini_lock:
                _pending_gemini_worker_running = False
            _refresh_gemini_queue_status_counts()
            with _pending_gemini_lock:
                still_pending = bool(_pending_gemini_replacements)
            if not still_pending:
                with _gemini_queue_status_lock:
                    if _gemini_queue_status.get("state") != "rate_limited":
                        _gemini_queue_status["state"] = "idle"
                    _gemini_queue_status["next_retry_at"] = 0.0
                    _gemini_queue_status["worker_running"] = False
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
            _play_wav_async(wav_path)
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
                    wav_path, _label, _can_cache = _synthesize_with_user_online(
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
            _play_wav_async(merged_path)
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
    return _online_provider_label()


_migrate_legacy_word_wrapper_layout()
_migrate_flat_root_cache_layout()
_migrate_pending_queue_path()
_load_pending_gemini_queue()
_cleanup_duplicate_source_cache_entries()
_normalize_cache_metadata_texts()
_collapse_existing_lightweight_source_caches()
_collapse_all_source_cache_entities_to_aliases()
_start_pending_gemini_worker()
