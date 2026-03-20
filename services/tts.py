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
from services.runtime_log import log_error as _log_error, log_info as _log_info, log_warning as _log_warning
from services.shared_metadata import export_shared_metadata_payload, import_shared_metadata_payload
from services.tts_backend_strategy import (
    backend_key as _strategy_backend_key,
    backend_key_from_label as _strategy_backend_key_from_label,
    backend_label_from_key as _strategy_backend_label_from_key,
    current_online_provider as _strategy_current_online_provider,
    interactive_online_timeout as _strategy_interactive_online_timeout,
    local_fallback_ready as _strategy_local_fallback_ready,
    manual_request_cooldown_for_provider as _strategy_manual_request_cooldown_for_provider,
    online_provider_candidates as _strategy_online_provider_candidates,
    online_provider_label as _strategy_online_provider_label,
    primary_online_provider as _strategy_primary_online_provider,
    rate_limit_cooldown_for_provider as _strategy_rate_limit_cooldown_for_provider,
    secondary_online_provider as _strategy_secondary_online_provider,
    synthesize_with_online as _strategy_synthesize_with_online,
    synthesize_with_selected_source as _strategy_synthesize_with_selected_source,
)
from services.text_normalization import normalize_ielts_tts_text
from services.tts_audio import (
    cleanup_temp_wavs as _cleanup_temp_wavs,
    prepend_silence_to_wav as _prepend_silence_to_wav,
    wav_duration_seconds as _wav_duration_seconds,
)
from services.tts_persistence import (
    CacheMetadataStore,
    cache_meta_path as _cache_meta_path_for_file,
    load_json_file as _load_json_file,
    load_pending_queue_disk_payload as _load_pending_queue_disk_payload_from_disk,
    load_word_audio_overrides as _load_word_audio_overrides_from_disk,
    migrate_pending_queue_path as _migrate_pending_queue_path_on_disk,
    save_word_audio_overrides as _save_word_audio_overrides_to_disk,
    write_json_file as _write_json_file_to_disk,
)
from services.tts_queue import OnlineTtsQueueManager
from services.tts_shared_cache import (
    export_shared_audio_cache_package as _export_shared_audio_cache_package_impl,
    import_shared_audio_cache_package as _import_shared_audio_cache_package_impl,
)
from services.tts_synth_cache import resolve_short_text_cache as _resolve_short_text_cache
from services.tts_synth_execute import (
    execute_local_backend as _execute_local_backend,
    execute_online_with_fallback as _execute_online_with_fallback,
)
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
WORD_AUDIO_OVERRIDE_PATH = os.path.join(BASE_DIR, "data", "word_audio_overrides.json")
RECENT_WRONG_SOURCE_KEY = "__recent_wrong_words__"
MANUAL_SESSION_SOURCE_KEY = "__manual_session__"
KOKORO_SAMPLE_RATE = 24000
SHARED_CACHE_PACKAGE_VERSION = 2
SHARED_CACHE_PACKAGE_MANIFEST = "manifest.json"
SHARED_CACHE_METADATA_FILE = "global/metadata.json"

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
_word_audio_override_lock = threading.Lock()
_word_audio_override_memory = None
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
ONLINE_TTS_REQUEST_TIMEOUT_SECONDS = 180
INTERACTIVE_FAST_TIMEOUT_SHORT_SECONDS = 12
INTERACTIVE_FAST_TIMEOUT_LONG_SECONDS = 20
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
_online_queue_manager = OnlineTtsQueueManager(throttle_config=_QUEUE_THROTTLE_CONFIG)
_cache_metadata_store = None


def _clamp(value, low, high):
    return max(low, min(high, float(value)))


def _set_gemini_queue_status(**updates):
    _online_queue_manager.set_status(**updates)


def _provider_key(provider):
    return _online_queue_manager.provider_key(provider)


def _queue_throttle_config(provider):
    return _online_queue_manager.throttle_config(provider)


def _get_queue_throttle_snapshot(provider):
    return _online_queue_manager.get_queue_throttle_snapshot(provider)


def _queue_interval_for_provider(provider):
    return _online_queue_manager.queue_interval_for_provider(provider)


def _record_queue_success(provider):
    _online_queue_manager.record_queue_success(provider)


def _record_queue_soft_failure(provider):
    _online_queue_manager.record_queue_soft_failure(provider)


def _record_queue_rate_limit(provider):
    _online_queue_manager.record_queue_rate_limit(provider)


def _refresh_gemini_queue_status_counts():
    with _pending_gemini_lock:
        queue_count = len(_pending_gemini_replacements)
        worker_running = bool(_pending_gemini_worker_running)
    try:
        disk_payload = _load_pending_queue_disk_payload()
        if isinstance(disk_payload, list):
            queue_count = len(disk_payload)
    except Exception as exc:
        _log_warning("tts_queue_status_disk_payload_failed", error=exc)
    _online_queue_manager.refresh_counts(queue_count=queue_count, worker_running=worker_running)


def get_gemini_queue_status():
    _refresh_gemini_queue_status_counts()
    return _online_queue_manager.get_status()


def get_online_tts_queue_status():
    return get_gemini_queue_status()


def _defer_gemini_queue(wait_seconds, *, state=None, provider=None):
    _online_queue_manager.defer(wait_seconds, state=state, provider=provider)


def _wait_for_gemini_queue_slot(provider=None):
    _online_queue_manager.wait_for_slot(provider=provider)


def _synthesize_with_user_online(text, volume, *, short_text, timeout_seconds=ONLINE_TTS_REQUEST_TIMEOUT_SECONDS):
    provider = _primary_online_provider()
    _defer_gemini_queue(_manual_request_cooldown_for_provider(provider), state="ok", provider=provider)
    try:
        result, used_provider = _synthesize_with_online(
            text,
            volume=volume,
            short_text=short_text,
            timeout_seconds=timeout_seconds,
        )
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
    return _strategy_current_online_provider(get_tts_api_provider())


def _is_online_backend(backend_key):
    return str(backend_key or "").strip().lower() in {"gemini", "elevenlabs"}


def _online_provider_label(provider=None):
    return _strategy_online_provider_label(provider, current_provider=_current_online_provider())


def _primary_online_provider():
    return _strategy_primary_online_provider(get_tts_api_provider())


def _secondary_online_provider(primary=None):
    return _strategy_secondary_online_provider(primary or _primary_online_provider(), has_llm_api_key=bool(get_llm_api_key()))


def _online_provider_candidates(primary=None):
    return _strategy_online_provider_candidates(primary or _primary_online_provider(), has_llm_api_key=bool(get_llm_api_key()))


def _rate_limit_cooldown_for_provider(provider):
    return _strategy_rate_limit_cooldown_for_provider(
        provider,
        gemini_seconds=GEMINI_RATE_LIMIT_COOLDOWN_SECONDS,
        elevenlabs_seconds=ELEVENLABS_RATE_LIMIT_COOLDOWN_SECONDS,
    )


def _manual_request_cooldown_for_provider(provider):
    return _strategy_manual_request_cooldown_for_provider(
        provider,
        gemini_seconds=GEMINI_MANUAL_REQUEST_COOLDOWN_SECONDS,
        elevenlabs_seconds=ELEVENLABS_MANUAL_REQUEST_COOLDOWN_SECONDS,
    )


def _local_fallback_ready():
    return _strategy_local_fallback_ready(piper_ready=piper_ready(), kokoro_ready=kokoro_ready())


def _interactive_online_timeout(short_text):
    return _strategy_interactive_online_timeout(
        short_text,
        local_fallback_ready=_local_fallback_ready(),
        online_timeout_seconds=ONLINE_TTS_REQUEST_TIMEOUT_SECONDS,
        fast_short_timeout_seconds=INTERACTIVE_FAST_TIMEOUT_SHORT_SECONDS,
        fast_long_timeout_seconds=INTERACTIVE_FAST_TIMEOUT_LONG_SECONDS,
    )


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
    return _cache_meta_path_for_file(cache_path)


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
    return _strategy_backend_key(
        source=source,
        fallback_backend=fallback_backend,
        source_kokoro=SOURCE_KOKORO,
        source_piper=SOURCE_PIPER,
        current_online_provider=_current_online_provider(),
    )


def _backend_label_from_key(backend_key):
    return _strategy_backend_label_from_key(backend_key)


def _backend_key_from_label(label):
    return _strategy_backend_key_from_label(label)


def _selected_source_backend_key():
    return _backend_key(source=get_voice_source())


def get_word_backend_override(text, source_path=None):
    key = _word_audio_override_key(text, source_path=source_path)
    if not key:
        return ""
    return str(_load_word_audio_overrides().get(key) or "").strip().lower()


def set_word_backend_override(text, source_path=None, backend=None):
    key = _word_audio_override_key(text, source_path=source_path)
    backend_key = str(backend or "").strip().lower()
    if not key or backend_key not in {"gemini", "elevenlabs", "kokoro", "piper"}:
        return False
    payload = _load_word_audio_overrides()
    payload[key] = backend_key
    _save_word_audio_overrides(payload)
    return True


def clear_word_backend_override(text, source_path=None):
    key = _word_audio_override_key(text, source_path=source_path)
    if not key:
        return False
    payload = _load_word_audio_overrides()
    existed = key in payload
    if existed:
        payload.pop(key, None)
        _save_word_audio_overrides(payload)
    cache_path = _word_cache_path(text, source_path=source_path)
    metadata = _load_cache_metadata(cache_path)
    backend = str(metadata.get("backend") or "").strip().lower()
    if backend in {"piper", "kokoro"}:
        playable_path = _resolve_cache_audio_path(cache_path)
        for path in {str(cache_path or "").strip(), str(playable_path or "").strip(), _cache_meta_path(cache_path)}:
            if not path:
                continue
            try:
                if os.path.exists(path):
                    os.remove(path)
            except Exception as exc:
                _log_warning("tts_clear_backend_override_remove_failed", path=path, error=exc)
        _remove_cache_metadata(cache_path)
        return True
    return existed


def _selected_word_backend_key(text, *, source_path=None, short_text=False):
    if short_text:
        override_backend = get_word_backend_override(text, source_path=source_path)
        if override_backend in {"gemini", "elevenlabs", "kokoro", "piper"}:
            return override_backend
    return _selected_source_backend_key()


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


def _word_audio_override_key(text, source_path=None):
    normalized_text = _normalize_text(text, ensure_sentence_end=False)
    normalized_source = _normalize_source_path(source_path)
    if not normalized_text or not normalized_source:
        return ""
    return f"{str(normalized_source).casefold()}|{normalized_text.casefold()}"


def _load_word_audio_overrides():
    global _word_audio_override_memory
    with _word_audio_override_lock:
        if isinstance(_word_audio_override_memory, dict):
            return dict(_word_audio_override_memory)
        payload = _load_word_audio_overrides_from_disk(
            WORD_AUDIO_OVERRIDE_PATH,
            allowed_backends={"gemini", "elevenlabs", "kokoro", "piper"},
        )
        _word_audio_override_memory = dict(payload)
        return dict(payload)


def _save_word_audio_overrides(data):
    payload = dict(data or {})
    _save_word_audio_overrides_to_disk(WORD_AUDIO_OVERRIDE_PATH, payload)
    with _word_audio_override_lock:
        global _word_audio_override_memory
        _word_audio_override_memory = dict(payload)


def _safe_rel_path(path):
    value = str(path or "").replace("\\", "/").strip().lstrip("/")
    normalized = os.path.normpath(value).replace("\\", "/")
    if not normalized or normalized in {".", ""}:
        return ""
    if normalized.startswith("../") or normalized == "..":
        return ""
    return normalized


def _zip_entry_path(*parts):
    cleaned = [str(part or "").strip("/\\") for part in parts if str(part or "").strip("/\\")]
    return "/".join(cleaned)


def _sha1_file(path):
    sha1 = hashlib.sha1()
    with open(path, "rb") as fp:
        while True:
            chunk = fp.read(1024 * 1024)
            if not chunk:
                break
            sha1.update(chunk)
    return sha1.hexdigest()


def _shared_cache_export_entries():
    entries = []
    seen_targets = set()
    if not os.path.isdir(SHARED_WORD_CACHE_DIR):
        return entries
    for root, _, names in os.walk(SHARED_WORD_CACHE_DIR):
        for name in sorted(names):
            if not name.lower().endswith(".wav"):
                continue
            cache_path = os.path.join(root, name)
            rel_path = _safe_rel_path(os.path.relpath(cache_path, SHARED_WORD_CACHE_DIR))
            if not rel_path:
                continue
            metadata = _load_cache_metadata(cache_path)
            playable_path = _resolve_cache_audio_path(cache_path)
            if not playable_path or not os.path.exists(playable_path):
                continue
            payload = dict(metadata or {})
            payload["source_path"] = "shared"
            target_path = _shared_cache_target_path(relative_path=rel_path, metadata=payload) or cache_path
            target_key = os.path.abspath(target_path)
            if target_key in seen_targets:
                continue
            seen_targets.add(target_key)
            relative_path = _safe_rel_path(os.path.relpath(target_path, SHARED_WORD_CACHE_DIR)) or rel_path
            entries.append(
                {
                    "cache_path": target_path,
                    "audio_path": playable_path,
                    "relative_path": relative_path,
                    "meta_relative_path": f"{relative_path}.json",
                    "metadata": payload,
                }
            )
    return entries


def _shared_cache_target_path(relative_path="", metadata=None):
    payload = dict(metadata or {})
    payload["source_path"] = "shared"
    text_value = _normalize_text(payload.get("text"), ensure_sentence_end=False)
    backend = _backend_key(
        fallback_backend=payload.get("backend") or payload.get("desired_backend") or _current_online_provider()
    )
    if text_value:
        return _provider_shared_word_cache_path(text_value, provider=backend)
    safe_rel = _safe_rel_path(relative_path)
    if not safe_rel:
        return ""
    return os.path.join(SHARED_WORD_CACHE_DIR, safe_rel.replace("/", os.sep))


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
                _log_warning(
                    "tts_rename_cache_source_migrate_entry_failed",
                    old_cache_path=old_cache_path,
                    new_cache_path=new_cache_path,
                )
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
        except Exception as exc:
            _log_warning("tts_remove_cache_entry_audio_failed", cache_path=cache_path, error=exc)
        meta_path = _cache_meta_path(cache_path)
        try:
            if os.path.exists(meta_path):
                os.remove(meta_path)
                removed_local += 1
        except Exception as exc:
            _log_warning("tts_remove_cache_entry_meta_failed", meta_path=meta_path, error=exc)
        _remove_cache_metadata(cache_path)
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
    _write_json_file_to_disk(path, payload)


def _load_pending_queue_disk_payload():
    return _load_pending_queue_disk_payload_from_disk(
        PENDING_ONLINE_TTS_QUEUE_PATH,
        LEGACY_PENDING_GEMINI_QUEUE_PATH,
    )


def _migrate_pending_queue_path():
    _migrate_pending_queue_path_on_disk(
        PENDING_ONLINE_TTS_QUEUE_PATH,
        LEGACY_PENDING_GEMINI_QUEUE_PATH,
    )


def _get_cache_metadata_store():
    global _cache_metadata_store
    if _cache_metadata_store is None:
        _cache_metadata_store = CacheMetadataStore(
            canonicalize=_canonicalize_cache_path,
            normalize_source_path=_normalize_source_path,
        )
    return _cache_metadata_store


def _load_cache_metadata(cache_path):
    return _get_cache_metadata_store().load(cache_path)


def _save_cache_metadata(cache_path, metadata):
    _get_cache_metadata_store().save(cache_path, metadata)


def _remove_cache_metadata(cache_path):
    _get_cache_metadata_store().remove(cache_path)


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
        except Exception as exc:
            _log_warning("tts_cleanup_manual_session_remove_failed", path=path, error=exc)
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
            except Exception as exc:
                _log_warning(
                    "tts_rebind_manual_cache_copy_failed",
                    source_path=target_source,
                    cache_path=target_cache,
                    text=normalized,
                    error=exc,
                )
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
            except Exception as exc:
                _log_warning(
                    "tts_rebind_manual_cache_alias_failed",
                    source_path=target_source,
                    cache_path=target_cache,
                    text=normalized,
                    error=exc,
                )

        # Clear the manual-session entry once the saved file has taken ownership.
        try:
            if os.path.exists(manual_cache):
                os.remove(manual_cache)
                removed_manual += 1
        except Exception as exc:
            _log_warning("tts_rebind_manual_cache_cleanup_failed", cache_path=manual_cache, error=exc)
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
            except Exception as exc:
                _log_warning("tts_cleanup_source_cache_remove_failed", path=full_path, error=exc)
            if name.lower().endswith(".wav"):
                _remove_cache_metadata(cache_path)
            elif name.lower().endswith(".wav.json"):
                _remove_pending_gemini(cache_path)
        try:
            if os.path.isdir(root) and not os.listdir(root):
                os.rmdir(root)
        except Exception as exc:
            _log_warning("tts_cleanup_source_cache_rmdir_failed", path=root, error=exc)
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
    except Exception as exc:
        _log_warning("tts_cleanup_source_cache_root_rmdir_failed", path=source_dir, error=exc)
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
    except Exception as exc:
        _log_warning("tts_cleanup_word_cache_remove_failed", cache_path=cache_path, error=exc)
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
            except Exception as exc:
                _log_warning("tts_rename_cache_source_cleanup_failed", cache_path=old_cache_path, error=exc)
            _remove_cache_metadata(old_cache_path)

        for root, _, _ in os.walk(old_dir, topdown=False):
            try:
                if os.path.isdir(root) and not os.listdir(root):
                    os.rmdir(root)
            except Exception as exc:
                _log_warning("tts_rename_cache_source_rmdir_failed", path=root, error=exc)
        try:
            if os.path.isdir(old_dir) and not os.listdir(old_dir):
                os.rmdir(old_dir)
        except Exception as exc:
            _log_warning("tts_rename_cache_source_root_rmdir_failed", path=old_dir, error=exc)

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


def _clone_to_temp(path, *, volume=1.0):
    gain = _clamp(volume, 0.0, 6.0)
    if abs(gain - 1.0) <= 1e-6:
        fd, temp_path = tempfile.mkstemp(prefix="wordspeaker_", suffix=".wav")
        os.close(fd)
        shutil.copyfile(path, temp_path)
        return temp_path
    try:
        with wave.open(path, "rb") as wav_fp:
            sample_rate = int(wav_fp.getframerate() or TTS_SAMPLE_RATE)
            pcm_bytes = wav_fp.readframes(wav_fp.getnframes())
        return _write_pcm_to_wav_path(pcm_bytes, sample_rate=sample_rate, volume=gain)
    except Exception as exc:
        _log_warning("tts_clone_temp_gain_failed", path=path, volume=gain, error=exc)
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
            except Exception as exc:
                _log_warning(
                    "tts_source_cache_retarget_failed",
                    cache_path=source_cache_path,
                    source_path=source_path or "",
                    text=text[:120],
                    error=exc,
                )
            return source_cache_path
        if source_path != "shared" and _has_valid_gemini_cache(shared_path):
            try:
                _collapse_source_cache_to_alias(text, source_path=source_path)
            except Exception as exc:
                _log_warning(
                    "tts_source_cache_alias_collapse_failed",
                    cache_path=source_cache_path,
                    source_path=source_path or "",
                    text=text[:120],
                    error=exc,
                )
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
        except Exception as exc:
            _log_warning(
                "tts_source_cache_alias_from_shared_failed",
                cache_path=source_cache_path,
                shared_cache_path=shared_cache_path,
                source_path=source_path or "",
                text=text[:120],
                error=exc,
            )
    return source_cache_path


def has_cached_word_audio(text, source_path=None):
    override_backend = get_word_backend_override(text, source_path=source_path)
    if override_backend in {"kokoro", "piper"}:
        cache_path = _word_cache_path(text, source_path=source_path)
        metadata = _load_cache_metadata(cache_path)
        backend = str(metadata.get("backend") or "").strip().lower()
        return bool(_resolve_cache_audio_path(cache_path) and backend == override_backend)
    return _has_valid_gemini_cache(_ensure_source_gemini_cache(text, source_path=source_path))


def get_word_audio_cache_info(text, source_path=None):
    override_backend = get_word_backend_override(text, source_path=source_path)
    if override_backend in {"kokoro", "piper"}:
        cache_path = _word_cache_path(text, source_path=source_path)
    else:
        cache_path = _ensure_source_gemini_cache(text, source_path=source_path)
    playable_cache = _resolve_cache_audio_path(cache_path)
    exists = bool(playable_cache)
    metadata = _load_cache_metadata(cache_path)
    backend = str(metadata.get("backend") or "").strip().lower()
    desired_backend = str(override_backend or metadata.get("desired_backend") or "").strip().lower()
    with _pending_gemini_lock:
        pending = bool(not override_backend and cache_path in _pending_gemini_replacements)
    shared_path = ""
    if not override_backend:
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


def export_shared_audio_cache_package(package_path):
    return _export_shared_audio_cache_package_impl(
        package_path,
        get_export_entries=_shared_cache_export_entries,
        zip_entry_path=_zip_entry_path,
        shared_cache_metadata_file=SHARED_CACHE_METADATA_FILE,
        shared_cache_package_manifest=SHARED_CACHE_PACKAGE_MANIFEST,
        shared_cache_package_version=SHARED_CACHE_PACKAGE_VERSION,
        export_shared_metadata_payload=export_shared_metadata_payload,
        sha1_file=_sha1_file,
    )


def import_shared_audio_cache_package(package_path):
    return _import_shared_audio_cache_package_impl(
        package_path,
        shared_cache_package_manifest=SHARED_CACHE_PACKAGE_MANIFEST,
        shared_cache_metadata_file=SHARED_CACHE_METADATA_FILE,
        safe_rel_path=_safe_rel_path,
        zip_entry_path=_zip_entry_path,
        shared_cache_target_path=_shared_cache_target_path,
        sha1_file=_sha1_file,
        load_cache_metadata=_load_cache_metadata,
        save_cache_metadata=_save_cache_metadata,
        import_shared_metadata_payload=import_shared_metadata_payload,
        cleanup_duplicate_source_cache_entries=_cleanup_duplicate_source_cache_entries,
        collapse_existing_lightweight_source_caches=_collapse_existing_lightweight_source_caches,
        collapse_all_source_cache_entities_to_aliases=_collapse_all_source_cache_entities_to_aliases,
    )


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
    except Exception as exc:
        _log_warning(
            "tts_copy_cache_between_sources_failed",
            text=text[:120],
            from_source_path=from_source_path or "",
            to_source_path=to_source_path or "",
            error=exc,
        )
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


def _set_backend_status(token, label, *, from_cache=False, fallback=False, duration_seconds=0.0):
    with _backend_lock:
        _backend_status[token] = {
            "label": str(label or ""),
            "from_cache": bool(from_cache),
            "fallback": bool(fallback),
            "duration_seconds": max(0.0, float(duration_seconds or 0.0)),
        }


def get_backend_status(token):
    with _backend_lock:
        data = _backend_status.get(int(token or 0))
        return dict(data) if isinstance(data, dict) else None


def _stop_locked():
    global _current_wav
    try:
        winsound.PlaySound(None, winsound.SND_PURGE)
    except Exception as exc:
        _log_warning("tts_stop_purge_failed", error=exc)
    if _current_wav:
        try:
            os.remove(_current_wav)
        except Exception as exc:
            _log_warning("tts_stop_temp_remove_failed", path=_current_wav, error=exc)
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
            _log_warning("tts_play_retry", path=path, attempt=attempt + 1, error=e)
            time.sleep(0.05)
    if last_error:
        _log_error("tts_play_failed", path=path, error=last_error)
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
    _log_error("tts_user_visible_error", message=message)


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


def _request_gemini_tts(text, *, short_text, api_key=None, timeout=ONLINE_TTS_REQUEST_TIMEOUT_SECONDS):
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
        _log_info("tts_request_start", provider="gemini", short_text=short_text, timeout=timeout, text=spoken_text[:120])
        with urllib.request.urlopen(req, timeout=timeout) as response:
            return json.loads(response.read().decode("utf-8", errors="ignore"))
    except urllib.error.HTTPError as e:
        _log_error("tts_request_http_error", provider="gemini", short_text=short_text, error=_extract_error_message(e))
        raise RuntimeError(_extract_error_message(e)) from e
    except urllib.error.URLError as e:
        _log_error("tts_request_url_error", provider="gemini", short_text=short_text, error=e.reason)
        raise RuntimeError(f"Gemini TTS request failed: {e.reason}") from e
    except Exception as e:
        _log_error("tts_request_error", provider="gemini", short_text=short_text, error=e)
        raise RuntimeError(f"Gemini TTS request failed: {e}") from e


def _request_elevenlabs_tts(text, *, short_text, api_key=None, timeout=ONLINE_TTS_REQUEST_TIMEOUT_SECONDS):
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
        _log_info("tts_request_start", provider="elevenlabs", short_text=short_text, timeout=timeout, text=spoken_text[:120])
        with urllib.request.urlopen(req, timeout=timeout) as response:
            return response.read()
    except urllib.error.HTTPError as e:
        _log_error("tts_request_http_error", provider="elevenlabs", short_text=short_text, error=_extract_error_message(e))
        raise RuntimeError(_extract_error_message(e)) from e
    except urllib.error.URLError as e:
        _log_error("tts_request_url_error", provider="elevenlabs", short_text=short_text, error=e.reason)
        raise RuntimeError(f"ElevenLabs TTS request failed: {e.reason}") from e
    except Exception as e:
        _log_error("tts_request_error", provider="elevenlabs", short_text=short_text, error=e)
        raise RuntimeError(f"ElevenLabs TTS request failed: {e}") from e


def _request_online_tts(text, *, short_text, provider=None, api_key=None, timeout=ONLINE_TTS_REQUEST_TIMEOUT_SECONDS):
    backend = str(provider or _primary_online_provider()).strip().lower()
    if backend == "elevenlabs":
        return _request_elevenlabs_tts(text, short_text=short_text, api_key=api_key, timeout=timeout), "elevenlabs"
    return _request_gemini_tts(text, short_text=short_text, api_key=api_key, timeout=timeout), "gemini"


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
    gain = _clamp(volume, 0.0, 6.0)
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
    gain = _clamp(volume, 0.0, 6.0)
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


def _synthesize_with_gemini(text, volume, *, short_text, timeout_seconds=ONLINE_TTS_REQUEST_TIMEOUT_SECONDS):
    data = _request_gemini_tts(text, short_text=short_text, timeout=timeout_seconds)
    pcm_bytes = _extract_pcm_bytes(data)
    return _write_pcm_to_wav_path(pcm_bytes, sample_rate=TTS_SAMPLE_RATE, volume=volume), "Gemini TTS", True


def _synthesize_with_elevenlabs(text, volume, *, short_text, timeout_seconds=ONLINE_TTS_REQUEST_TIMEOUT_SECONDS):
    pcm_bytes = _request_elevenlabs_tts(text, short_text=short_text, timeout=timeout_seconds)
    return _write_pcm_to_wav_path(pcm_bytes, sample_rate=TTS_SAMPLE_RATE, volume=volume), "ElevenLabs TTS", True


def _synthesize_with_gemini_fallback(text, volume, *, short_text, api_key=None, timeout_seconds=ONLINE_TTS_REQUEST_TIMEOUT_SECONDS):
    data = _request_gemini_tts(text, short_text=short_text, api_key=api_key, timeout=timeout_seconds)
    pcm_bytes = _extract_pcm_bytes(data)
    return _write_pcm_to_wav_path(pcm_bytes, sample_rate=TTS_SAMPLE_RATE, volume=volume), "Gemini TTS", True


def _synthesize_with_elevenlabs_fallback(text, volume, *, short_text, api_key=None, timeout_seconds=ONLINE_TTS_REQUEST_TIMEOUT_SECONDS):
    pcm_bytes = _request_elevenlabs_tts(text, short_text=short_text, api_key=api_key, timeout=timeout_seconds)
    return _write_pcm_to_wav_path(pcm_bytes, sample_rate=TTS_SAMPLE_RATE, volume=volume), "ElevenLabs TTS", True


def _synthesize_with_online_provider(text, volume, *, short_text, provider, api_key=None, timeout_seconds=ONLINE_TTS_REQUEST_TIMEOUT_SECONDS):
    backend = str(provider or _primary_online_provider()).strip().lower()
    if backend == "elevenlabs":
        return _synthesize_with_elevenlabs_fallback(
            text,
            volume=volume,
            short_text=short_text,
            api_key=api_key,
            timeout_seconds=timeout_seconds,
        )
    return _synthesize_with_gemini_fallback(
        text,
        volume=volume,
        short_text=short_text,
        api_key=api_key,
        timeout_seconds=timeout_seconds,
    )


def _synthesize_with_online(text, volume, *, short_text, timeout_seconds=ONLINE_TTS_REQUEST_TIMEOUT_SECONDS):
    primary = _primary_online_provider()
    secondary = _secondary_online_provider(primary)
    return _strategy_synthesize_with_online(
        text,
        volume,
        short_text=short_text,
        timeout_seconds=timeout_seconds,
        primary_provider=primary,
        secondary_provider=secondary,
        get_fallback_key=lambda provider: get_llm_api_key() if provider == "gemini" else get_tts_api_key(),
        synthesize_with_online_provider=_synthesize_with_online_provider,
    )


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
    syn_config = SynthesisConfig(length_scale=max(0.1, 1.0 / speed), volume=1.0)
    audio_chunks = list(voice.synthesize(_normalize_tts_spoken_text(text), syn_config=syn_config))
    if not audio_chunks:
        raise RuntimeError("Piper returned no audio.")
    audio = np.concatenate([chunk.audio_float_array for chunk in audio_chunks])
    sample_rate = int(audio_chunks[0].sample_rate or TTS_SAMPLE_RATE)
    return _write_float_audio_to_wav_path(audio, sample_rate=sample_rate, volume=volume), "Piper (Local)", False


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


def _synthesize_with_selected_source(text, volume, rate_ratio, *, short_text, timeout_seconds=ONLINE_TTS_REQUEST_TIMEOUT_SECONDS):
    return _strategy_synthesize_with_selected_source(
        text,
        volume,
        rate_ratio,
        short_text=short_text,
        timeout_seconds=timeout_seconds,
        source=get_voice_source(),
        source_kokoro=SOURCE_KOKORO,
        source_piper=SOURCE_PIPER,
        synthesize_with_kokoro=_synthesize_with_kokoro,
        synthesize_with_piper=_synthesize_with_piper,
        synthesize_with_user_online=_synthesize_with_user_online,
        local_fallback_ready=_local_fallback_ready(),
        synthesize_with_local_placeholder=_synthesize_with_local_placeholder,
    )


def _synthesize_to_wav(text, volume, rate_ratio, *, short_text=False, source_path=None, request_token=None):
    normalized = _normalize_text(text, ensure_sentence_end=short_text)
    if not normalized:
        raise RuntimeError("Text is empty.")
    online_timeout_seconds = _interactive_online_timeout(short_text)
    selected_backend = _selected_word_backend_key(text, source_path=source_path, short_text=short_text)
    _log_info(
        "tts_synthesize_start",
        request_token=request_token,
        short_text=short_text,
        backend=selected_backend,
        source_path=source_path or "",
        text=normalized[:120],
    )
    cache_path = None

    if short_text:
        cache_state = _resolve_short_text_cache(
            text=text,
            source_path=source_path,
            selected_backend=selected_backend,
            request_token=request_token,
            volume=volume,
            current_online_provider=_current_online_provider,
            word_cache_path=_word_cache_path,
            resolve_cache_audio_path=_resolve_cache_audio_path,
            load_cache_metadata=_load_cache_metadata,
            set_backend_status=_set_backend_status,
            backend_label_from_key=_backend_label_from_key,
            wav_duration_seconds=_wav_duration_seconds,
            clone_to_temp=_clone_to_temp,
            ensure_source_online_cache=_ensure_source_gemini_cache,
            is_online_backend=_is_online_backend,
            has_valid_online_cache=_has_valid_gemini_cache,
            online_provider_label=_online_provider_label,
            enqueue_existing_cache_for_online_replacement=_enqueue_existing_cache_for_gemini_replacement,
            legacy_word_cache_path=_legacy_word_cache_path,
            save_cache_metadata=_save_cache_metadata,
        )
        cache_path = cache_state.get("cache_path") or None
        if cache_state.get("hit"):
            _log_info(
                "tts_cache_hit",
                request_token=request_token,
                backend=selected_backend,
                cache_path=cache_path or "",
                text=normalized[:120],
            )
            return cache_state.get("wav_path")

    local_result = _execute_local_backend(
        selected_backend=selected_backend,
        normalized=normalized,
        volume=volume,
        rate_ratio=rate_ratio,
        short_text=short_text,
        cache_path=cache_path,
        source_path=source_path,
        request_token=request_token,
        synthesize_with_kokoro=_synthesize_with_kokoro,
        synthesize_with_piper=_synthesize_with_piper,
        save_word_cache_file=_save_word_cache_file,
        set_backend_status=_set_backend_status,
        wav_duration_seconds=_wav_duration_seconds,
    )
    if local_result.get("handled"):
        _log_info(
            "tts_local_backend_handled",
            request_token=request_token,
            backend=selected_backend,
            wav_path=local_result.get("wav_path") or "",
        )
        return local_result.get("wav_path")

    return _execute_online_with_fallback(
        text=text,
        normalized=normalized,
        volume=volume,
        rate_ratio=rate_ratio,
        short_text=short_text,
        cache_path=cache_path,
        source_path=source_path,
        request_token=request_token,
        online_timeout_seconds=online_timeout_seconds,
        current_online_provider=_current_online_provider,
        primary_online_provider=_primary_online_provider,
        backend_key_from_label=_backend_key_from_label,
        set_backend_status=_set_backend_status,
        wav_duration_seconds=_wav_duration_seconds,
        synthesize_with_user_online=_synthesize_with_user_online,
        save_word_cache_file=_save_word_cache_file,
        enqueue_existing_cache_for_online_replacement=_enqueue_existing_cache_for_gemini_replacement,
        synthesize_with_local_placeholder=_synthesize_with_local_placeholder,
        kokoro_ready=kokoro_ready,
        synthesize_with_kokoro_voice=_synthesize_with_kokoro_voice,
    )


def _start_pending_gemini_worker():
    global _pending_gemini_worker_running
    with _pending_gemini_lock:
        if _pending_gemini_worker_running or not _pending_gemini_replacements:
            return
        _pending_gemini_worker_running = True
    _set_gemini_queue_status(worker_running=True)
    _refresh_gemini_queue_status_counts()
    current_status = _online_queue_manager.get_status()
    if current_status.get("state") in {"idle", ""}:
        _online_queue_manager.set_status(state="ok")

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
                    _log_info(
                        "tts_queue_item_start",
                        provider=desired_backend,
                        cache_path=cache_path_local,
                        source_path=source_path or "",
                        text=normalized_text[:120],
                    )
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
                    _log_info("tts_queue_item_done", provider=desired_backend, cache_path=cache_path_local)
                except Exception as exc:
                    _log_warning(
                        "tts_queue_item_failed",
                        provider=desired_backend,
                        cache_path=cache_path_local,
                        error=exc,
                    )
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
                            _log_info(
                                "tts_queue_item_fallback_success",
                                provider=desired_backend,
                                fallback_provider=secondary_backend,
                                cache_path=cache_path_local,
                            )
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
                        except Exception as fallback_exc:
                            _record_queue_soft_failure(secondary_backend)
                            _log_error(
                                "tts_queue_item_fallback_failed",
                                provider=desired_backend,
                                fallback_provider=secondary_backend,
                                cache_path=cache_path_local,
                                error=fallback_exc,
                            )
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
                                    except Exception as cleanup_exc:
                                        _log_warning("tts_queue_placeholder_cleanup_failed", path=wav_path, error=cleanup_exc)
                            except Exception as placeholder_exc:
                                _log_warning(
                                    "tts_queue_placeholder_fallback_failed",
                                    provider=desired_backend,
                                    cache_path=cache_path_local,
                                    error=placeholder_exc,
                                )
                        _log_warning(
                            "tts_queue_item_rate_limited",
                            provider=desired_backend,
                            cache_path=cache_path_local,
                            cooldown_seconds=cooldown_seconds,
                        )
                        threading.Event().wait(cooldown_seconds)
                        continue
                    _set_gemini_queue_status(
                        state="error",
                        next_retry_at=0.0,
                        last_error=str(exc),
                    )
                    _log_error(
                        "tts_queue_item_abandoned",
                        provider=desired_backend,
                        cache_path=cache_path_local,
                        error=exc,
                    )
                    _remove_pending_gemini(cache_path_local)
        finally:
            with _pending_gemini_lock:
                _pending_gemini_worker_running = False
            _refresh_gemini_queue_status_counts()
            with _pending_gemini_lock:
                still_pending = bool(_pending_gemini_replacements)
            if not still_pending:
                current_status = _online_queue_manager.get_status()
                updates = {
                    "next_retry_at": 0.0,
                    "worker_running": False,
                }
                if current_status.get("state") != "rate_limited":
                    updates["state"] = "idle"
                _online_queue_manager.set_status(**updates)
        _start_pending_gemini_worker()

    threading.Thread(target=_worker, daemon=True).start()


def _enqueue_gemini_replacement(text, cache_path, source_path=None):
    _enqueue_existing_cache_for_gemini_replacement(text, cache_path, source_path=source_path)


def stop():
    with _lock:
        _stop_locked()


def speak_async(text, volume=1.0, rate_ratio=1.0, cancel_before=False, source_path=None, pre_silence_ms=0):
    global _token
    if cancel_before:
        stop()
    with _lock:
        _token += 1
        my_token = _token
    _log_info(
        "tts_speak_async_start",
        token=my_token,
        source_path=source_path or "",
        volume=volume,
        rate_ratio=rate_ratio,
        pre_silence_ms=pre_silence_ms,
        text=str(text or "")[:120],
    )

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
            wav_path = _prepend_silence_to_wav(
                wav_path,
                silence_ms=pre_silence_ms,
                default_channels=TTS_CHANNELS,
                default_sample_width=TTS_SAMPLE_WIDTH,
                default_sample_rate=TTS_SAMPLE_RATE,
            )
            with _lock:
                if my_token != _token:
                    try:
                        os.remove(wav_path)
                    except Exception as exc:
                        _log_warning("tts_speak_async_cancel_cleanup_failed", path=wav_path, error=exc)
                    _log_warning("tts_speak_async_cancelled", token=my_token)
                    return
                _stop_locked()
                _current_wav = wav_path
            _play_wav_async(wav_path)
            _log_info("tts_speak_async_playing", token=my_token, wav_path=wav_path)
        except Exception as e:
            _log_error("tts_speak_async_failed", token=my_token, error=e, text=str(text or "")[:120])
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
    _log_info(
        "tts_stream_start",
        token=my_token,
        volume=volume,
        rate_ratio=rate_ratio,
        chunk_chars=chunk_chars,
        text=str(text or "")[:160],
    )

    def _run():
        global _current_wav
        try:
            with _lock:
                if my_token != _token:
                    return
            chunks = _split_long_text(text, chunk_chars=max(400, int(chunk_chars)))
            if not chunks:
                return
            _log_info("tts_stream_chunks_ready", token=my_token, chunk_count=len(chunks))
            wav_paths = []
            fallback = False
            online_timeout_seconds = _interactive_online_timeout(False)

            def _synthesize_locally(chunk_text):
                (local_path, local_label, _can_cache), local_backend = _synthesize_with_local_placeholder(
                    chunk_text,
                    volume=volume,
                    rate_ratio=rate_ratio,
                )
                return local_path, local_label, local_backend

            selected_source = get_voice_source()
            if selected_source == SOURCE_KOKORO:
                synth_mode = "kokoro"
                backend_label = "Kokoro (Offline)"
            elif selected_source == SOURCE_PIPER:
                synth_mode = "piper"
                backend_label = "Piper (Local)"
            else:
                first_result, fallback = _synthesize_with_selected_source(
                    chunks[0],
                    volume=volume,
                    rate_ratio=rate_ratio,
                    short_text=False,
                    timeout_seconds=online_timeout_seconds,
                )
                first_path, backend_label, _can_cache = first_result
                wav_paths.append(first_path)
                backend_key = _backend_key_from_label(backend_label)
                synth_mode = backend_key if backend_key in {"kokoro", "piper"} else "online"

            try:
                start_index = 0 if selected_source in {SOURCE_KOKORO, SOURCE_PIPER} else 1
                for chunk in chunks[start_index:]:
                    with _lock:
                        if my_token != _token:
                            _cleanup_temp_wavs(wav_paths)
                            _log_warning("tts_stream_cancelled", token=my_token)
                            return
                    if synth_mode == "kokoro":
                        wav_path, _label, _can_cache = _synthesize_with_kokoro_voice(
                            chunk,
                            volume=volume,
                            rate_ratio=rate_ratio,
                            voice_id="bf_emma",
                            lang="en-gb",
                        )
                    elif synth_mode == "piper":
                        wav_path, _label, _can_cache = _synthesize_with_piper(
                            chunk,
                            volume=volume,
                            rate_ratio=rate_ratio,
                        )
                    else:
                        wav_path, _label, _can_cache = _synthesize_with_user_online(
                            chunk,
                            volume=volume,
                            short_text=False,
                            timeout_seconds=online_timeout_seconds,
                        )
                    wav_paths.append(wav_path)
            except Exception:
                if synth_mode == "online" and _local_fallback_ready():
                    _cleanup_temp_wavs(wav_paths)
                    wav_paths = []
                    fallback = True
                    _log_warning("tts_stream_local_fallback", token=my_token)
                    local_backend = ""
                    backend_label = ""
                    try:
                        for chunk in chunks:
                            with _lock:
                                if my_token != _token:
                                    _cleanup_temp_wavs(wav_paths)
                                    _log_warning("tts_stream_cancelled", token=my_token)
                                    return
                            wav_path, backend_label, local_backend = _synthesize_locally(chunk)
                            wav_paths.append(wav_path)
                    except Exception:
                        _cleanup_temp_wavs(wav_paths)
                        raise
                    synth_mode = local_backend or "kokoro"
                else:
                    _cleanup_temp_wavs(wav_paths)
                    raise

            fd, merged_path = tempfile.mkstemp(prefix="wordspeaker_", suffix=".wav")
            os.close(fd)
            with wave.open(merged_path, "wb") as out_fp:
                out_fp.setnchannels(TTS_CHANNELS)
                out_fp.setsampwidth(TTS_SAMPLE_WIDTH)
                out_fp.setframerate(TTS_SAMPLE_RATE)
                for wav_path in wav_paths:
                    with wave.open(wav_path, "rb") as in_fp:
                        out_fp.writeframes(in_fp.readframes(in_fp.getnframes()))
            _cleanup_temp_wavs(wav_paths)

            if my_token is not None:
                _set_backend_status(
                    my_token,
                    backend_label,
                    from_cache=False,
                    fallback=fallback,
                    duration_seconds=_wav_duration_seconds(merged_path),
                )

            with _lock:
                if my_token != _token:
                    try:
                        os.remove(merged_path)
                    except Exception:
                        pass
                    _log_warning("tts_stream_cancelled", token=my_token)
                    return
                _stop_locked()
                _current_wav = merged_path
            _play_wav_async(merged_path)
            _log_info("tts_stream_playing", token=my_token, wav_path=merged_path, fallback=fallback)
        except Exception as e:
            _log_error("tts_stream_failed", token=my_token, error=e, text=str(text or "")[:160])
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
            except Exception as exc:
                _log_warning("tts_precache_progress_callback_failed", current_text=current_text, error=exc)

    def _emit_done(success_count, skipped_count, pending_count, error_count):
        if callable(on_done):
            try:
                on_done(success_count, skipped_count, pending_count, error_count)
            except Exception as exc:
                _log_warning(
                    "tts_precache_done_callback_failed",
                    success_count=success_count,
                    skipped_count=skipped_count,
                    pending_count=pending_count,
                    error_count=error_count,
                    error=exc,
                )

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
            except Exception as exc:
                error_count += 1
                _log_warning(
                    "tts_precache_enqueue_failed",
                    cache_path=cache_path,
                    source_path=source_path or "",
                    text=text,
                    error=exc,
                )
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
