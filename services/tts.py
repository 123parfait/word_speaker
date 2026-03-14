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
GLOBAL_WORD_CACHE_DIR = os.path.join(BASE_DIR, "data", "audio_cache", "words")
PENDING_GEMINI_QUEUE_PATH = os.path.join(BASE_DIR, "data", "audio_cache", "pending_gemini_replacements.json")
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


def _word_cache_dir(source_path=None):
    return GLOBAL_WORD_CACHE_DIR


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
    source = get_voice_source()
    voice_id = get_voice_id()
    normalized = _normalize_text(text, ensure_sentence_end=False).casefold()
    key = hashlib.sha1(
        f"{source}|{voice_id}|{TTS_MODEL}|{TTS_VOICE_NAME}|{normalized}".encode("utf-8")
    ).hexdigest()[:16]
    name = _safe_name(normalized)
    return os.path.join(_word_cache_dir(source_path), f"{name}_{key}.wav")


def _cache_meta_path(cache_path):
    return f"{cache_path}.json"


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


def _write_json_file(path, payload):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as fp:
        json.dump(payload, fp, ensure_ascii=False, indent=2)


def _load_cache_metadata(cache_path):
    data = _load_json_file(_cache_meta_path(cache_path), {})
    return data if isinstance(data, dict) else {}


def _save_cache_metadata(cache_path, metadata):
    payload = dict(metadata or {})
    payload["cache_path"] = cache_path
    _write_json_file(_cache_meta_path(cache_path), payload)


def _remove_cache_metadata(cache_path):
    try:
        meta_path = _cache_meta_path(cache_path)
        if os.path.exists(meta_path):
            os.remove(meta_path)
    except Exception:
        pass


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
    with _pending_gemini_lock:
        _pending_gemini_replacements.clear()
        for item in items:
            if not isinstance(item, dict):
                continue
            cache_path = str(item.get("cache_path") or "").strip()
            text = _normalize_text(item.get("text"), ensure_sentence_end=False)
            if not cache_path or not text:
                continue
            _pending_gemini_replacements[cache_path] = {
                "text": text,
                "source_path": str(item.get("source_path") or "").strip() or None,
                "created_at": item.get("created_at"),
            }


def _remove_pending_gemini(cache_path):
    with _pending_gemini_lock:
        if cache_path in _pending_gemini_replacements:
            _pending_gemini_replacements.pop(cache_path, None)
            _save_pending_gemini_queue_locked()


def _enqueue_existing_cache_for_gemini_replacement(text, cache_path, source_path=None):
    normalized = _normalize_text(text, ensure_sentence_end=False)
    if not normalized or not cache_path:
        return
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
    desired_backend = str(metadata.get("desired_backend") or "").strip().lower()
    if backend in {"kokoro", "piper"} and desired_backend == "gemini":
        return True
    return not metadata


def _save_word_cache_file(cache_path, wav_path, *, source_path=None, backend=None, desired_backend=None):
    os.makedirs(os.path.dirname(cache_path), exist_ok=True)
    shutil.copyfile(wav_path, cache_path)
    actual_backend = _backend_key(fallback_backend=backend)
    wanted_backend = _backend_key(fallback_backend=desired_backend or actual_backend)
    _save_cache_metadata(
        cache_path,
        {
            "backend": actual_backend,
            "desired_backend": wanted_backend,
            "source_path": str(source_path or "").strip() or None,
            "updated_at": int(os.path.getmtime(cache_path)) if os.path.exists(cache_path) else None,
        },
    )
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
    target = str(source_path or "").strip()
    if not target:
        return 0
    removed = 0
    if not os.path.isdir(GLOBAL_WORD_CACHE_DIR):
        return 0
    for name in os.listdir(GLOBAL_WORD_CACHE_DIR):
        if not name.lower().endswith(".wav"):
            continue
        cache_path = os.path.join(GLOBAL_WORD_CACHE_DIR, name)
        metadata = _load_cache_metadata(cache_path)
        meta_source = str(metadata.get("source_path") or "").strip()
        if meta_source != target:
            continue
        try:
            if os.path.exists(cache_path):
                os.remove(cache_path)
                removed += 1
        except Exception:
            pass
        _remove_cache_metadata(cache_path)
        _remove_pending_gemini(cache_path)
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


def has_cached_word_audio(text, source_path=None):
    return os.path.exists(_word_cache_path(text, source_path=source_path))


def get_word_audio_cache_info(text, source_path=None):
    cache_path = _word_cache_path(text, source_path=source_path)
    exists = os.path.exists(cache_path)
    metadata = _load_cache_metadata(cache_path) if exists else {}
    backend = str(metadata.get("backend") or "").strip().lower()
    desired_backend = str(metadata.get("desired_backend") or "").strip().lower()
    pending = False
    if exists:
        with _pending_gemini_lock:
            pending = cache_path in _pending_gemini_replacements
    return {
        "exists": exists,
        "cache_path": cache_path,
        "meta_path": _cache_meta_path(cache_path),
        "backend": backend or None,
        "backend_label": _backend_label_from_key(backend) if backend else "",
        "desired_backend": desired_backend or None,
        "desired_backend_label": _backend_label_from_key(desired_backend) if desired_backend else "",
        "pending_gemini_replacement": bool(pending),
        "metadata": dict(metadata) if isinstance(metadata, dict) else {},
    }


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
        True,
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
    voice = _ensure_piper_voice(get_voice_id())
    from piper import SynthesisConfig

    speed = _clamp(rate_ratio, 0.7, 1.4)
    syn_config = SynthesisConfig(length_scale=max(0.1, 1.0 / speed), volume=_clamp(volume, 0.0, 1.0))
    audio_chunks = list(voice.synthesize(str(text or ""), syn_config=syn_config))
    if not audio_chunks:
        raise RuntimeError("Piper returned no audio.")
    audio = np.concatenate([chunk.audio_float_array for chunk in audio_chunks])
    sample_rate = int(audio_chunks[0].sample_rate or TTS_SAMPLE_RATE)
    return _write_float_audio_to_wav_path(audio, sample_rate=sample_rate, volume=1.0), "Piper (Local)", True


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
        cache_path = _word_cache_path(text, source_path=source_path)
        if os.path.exists(cache_path):
            metadata = _load_cache_metadata(cache_path)
            backend_key = str(metadata.get("backend") or "").strip().lower()
            label = _backend_label_from_key(backend_key or _selected_source_backend_key())
            if _selected_source_backend_key() == "gemini" and _cache_requires_gemini_replacement(cache_path):
                _enqueue_existing_cache_for_gemini_replacement(text, cache_path, source_path=source_path)
            if request_token is not None:
                _set_backend_status(request_token, label, from_cache=True, fallback=False)
            return _clone_to_temp(cache_path)
        legacy_cache_path = _legacy_word_cache_path(text, source_path=source_path)
        if legacy_cache_path and os.path.exists(legacy_cache_path):
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
                        "desired_backend": _selected_source_backend_key(),
                        "source_path": str(source_path or "").strip() or None,
                    },
                )
            label = _backend_label_from_key(_selected_source_backend_key())
            if _selected_source_backend_key() == "gemini":
                _enqueue_existing_cache_for_gemini_replacement(text, cache_path, source_path=source_path)
            if request_token is not None:
                _set_backend_status(request_token, label, from_cache=True, fallback=False)
            return _clone_to_temp(cache_path)

    (wav_path, backend_label, can_cache), fallback = _synthesize_with_selected_source(
        normalized,
        volume=volume,
        rate_ratio=rate_ratio,
        short_text=short_text,
    )

    if request_token is not None:
        _set_backend_status(request_token, backend_label, from_cache=False, fallback=fallback)

    if short_text and can_cache:
        try:
            os.makedirs(os.path.dirname(cache_path), exist_ok=True)
            actual_backend = _actual_backend_key_for_result(fallback=fallback)
            desired_backend = _selected_source_backend_key()
            _save_word_cache_file(
                cache_path,
                wav_path,
                source_path=source_path,
                backend=actual_backend,
                desired_backend=desired_backend,
            )
        except Exception:
            pass

    return wav_path


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

    def _emit_done(success_count, skipped_count, error_count):
        if callable(on_done):
            try:
                on_done(success_count, skipped_count, error_count)
            except Exception:
                pass

    def _run():
        success_count = 0
        skipped_count = 0
        error_count = 0
        total_count = len(items)
        use_kokoro_until_gemini_recovers = False
        for index, text in enumerate(items, start=1):
            cache_path = _word_cache_path(text, source_path=source_path)
            if os.path.exists(cache_path):
                if get_voice_source() == SOURCE_GEMINI and _cache_requires_gemini_replacement(cache_path):
                    _enqueue_gemini_replacement(text, cache_path, source_path=source_path)
                skipped_count += 1
                _emit_progress(index, total_count, text)
                continue
            try:
                if get_voice_source() == SOURCE_GEMINI and kokoro_ready():
                    if use_kokoro_until_gemini_recovers:
                        wav_path, _label, _can_cache = _synthesize_with_kokoro_voice(
                            text,
                            volume=1.0,
                            rate_ratio=rate_ratio,
                            voice_id="bf_emma",
                            lang="en-gb",
                        )
                        _save_word_cache_file(
                            cache_path,
                            wav_path,
                            source_path=source_path,
                            backend="kokoro",
                            desired_backend="gemini",
                        )
                        _enqueue_gemini_replacement(text, cache_path, source_path=source_path)
                    else:
                        try:
                            wav_path, _label, _can_cache = _synthesize_with_gemini(
                                text,
                                volume=1.0,
                                short_text=True,
                            )
                            _save_word_cache_file(
                                cache_path,
                                wav_path,
                                source_path=source_path,
                                backend="gemini",
                                desired_backend="gemini",
                            )
                        except Exception as exc:
                            if not _is_gemini_rate_limited_error(exc):
                                raise
                            use_kokoro_until_gemini_recovers = True
                            wav_path, _label, _can_cache = _synthesize_with_kokoro_voice(
                                text,
                                volume=1.0,
                                rate_ratio=rate_ratio,
                                voice_id="bf_emma",
                                lang="en-gb",
                            )
                            _save_word_cache_file(
                                cache_path,
                                wav_path,
                                source_path=source_path,
                                backend="kokoro",
                                desired_backend="gemini",
                            )
                            _enqueue_gemini_replacement(text, cache_path, source_path=source_path)
                else:
                    wav_path = _synthesize_to_wav(
                        text=text,
                        volume=1.0,
                        rate_ratio=rate_ratio,
                        short_text=True,
                        source_path=source_path,
                        request_token=None,
                    )
                success_count += 1
                try:
                    os.remove(wav_path)
                except Exception:
                    pass
            except Exception:
                error_count += 1
            _emit_progress(index, total_count, text)
        _emit_done(success_count, skipped_count, error_count)

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


_load_pending_gemini_queue()
_start_pending_gemini_worker()
