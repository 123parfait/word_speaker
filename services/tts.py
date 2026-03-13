# -*- coding: utf-8 -*-
import os
import queue
import re
import threading
import tempfile
import urllib.error
import urllib.request
import wave
import winsound
from pathlib import Path

from tkinter import messagebox

from services.voice_catalog import get_voice_profile
from services.voice_manager import get_voice_id

_lock = threading.Lock()
_token = 0
_kokoro = None
_download_started = False
_download_failed_reason = None
_shown_errors = set()
_current_wav = None

_MODEL_URL = "https://github.com/thewh1teagle/kokoro-onnx/releases/download/model-files-v1.0/kokoro-v1.0.onnx"
_VOICES_URL = "https://github.com/thewh1teagle/kokoro-onnx/releases/download/model-files-v1.0/voices-v1.0.bin"
_MODELS_DIR = Path(__file__).resolve().parent.parent / "data" / "models" / "kokoro"
_MODEL_PATH = _MODELS_DIR / "kokoro-v1.0.onnx"
_VOICES_PATH = _MODELS_DIR / "voices-v1.0.bin"


def _download_file(url, target_path):
    target_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = Path(str(target_path) + ".part")
    start_offset = temp_path.stat().st_size if temp_path.exists() else 0
    headers = {}
    mode = "wb"
    if start_offset > 0:
        headers["Range"] = f"bytes={start_offset}-"
        mode = "ab"

    request = urllib.request.Request(url, headers=headers)
    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            status = getattr(response, "status", 200)
            if start_offset > 0 and status != 206:
                # Server did not accept range request; restart from zero.
                start_offset = 0
                mode = "wb"
            with open(temp_path, mode) as fp:
                while True:
                    chunk = response.read(1024 * 1024)
                    if not chunk:
                        break
                    fp.write(chunk)
    except urllib.error.HTTPError as e:
        # 416 can happen when the partial file already reached file end.
        if e.code == 416 and temp_path.exists() and temp_path.stat().st_size > 0:
            os.replace(temp_path, target_path)
            return
        raise
    os.replace(temp_path, target_path)


def _models_ready():
    return _MODEL_PATH.exists() and _VOICES_PATH.exists()


def _download_models_worker():
    global _download_started, _download_failed_reason
    try:
        if not _MODEL_PATH.exists():
            _download_file(_MODEL_URL, _MODEL_PATH)
        if not _VOICES_PATH.exists():
            _download_file(_VOICES_URL, _VOICES_PATH)
        with _lock:
            _download_failed_reason = None
    except Exception as e:
        with _lock:
            _download_failed_reason = str(e)
    finally:
        with _lock:
            _download_started = False


def _start_model_download_if_needed():
    global _download_started
    with _lock:
        if _models_ready() or _download_started:
            return
        _download_started = True
    threading.Thread(target=_download_models_worker, daemon=True).start()


def _get_kokoro_not_ready_reason():
    with _lock:
        if _download_failed_reason:
            return f"Kokoro model download failed: {_download_failed_reason}"
        if _download_started:
            return "Kokoro model is downloading in background. Please wait and try again."
    return "Kokoro model files are not ready yet."


def _ensure_kokoro():
    global _kokoro
    if _kokoro is not None:
        return _kokoro
    try:
        from kokoro_onnx import Kokoro  # imported lazily to keep app startup fast
    except Exception as e:
        raise RuntimeError(
            "Kokoro is not available. Install a 64-bit Python environment and run: "
            "pip install kokoro-onnx onnxruntime"
        ) from e
    if not _models_ready():
        raise RuntimeError(_get_kokoro_not_ready_reason())
    _kokoro = Kokoro(str(_MODEL_PATH), str(_VOICES_PATH))
    return _kokoro


def _write_temp_wav(audio, sample_rate):
    import numpy as np

    pcm = np.asarray(audio, dtype=np.float32)
    pcm = np.clip(pcm, -1.0, 1.0)
    pcm16 = (pcm * 32767.0).astype(np.int16)
    fd, path = tempfile.mkstemp(prefix="wordspeaker_", suffix=".wav")
    os.close(fd)
    with wave.open(path, "wb") as wav_fp:
        wav_fp.setnchannels(1)
        wav_fp.setsampwidth(2)
        wav_fp.setframerate(int(sample_rate))
        wav_fp.writeframes(pcm16.tobytes())
    return path


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


def _synthesize_to_wav(text, volume, rate_ratio):
    voice = get_voice_profile(get_voice_id())
    voice_id = voice.get("id", "af_heart")
    lang = (voice.get("languages") or ["en-us"])[0]
    ratio = max(0.5, min(2.0, float(rate_ratio)))
    gain = max(0.0, min(1.0, float(volume)))

    kokoro = _ensure_kokoro()
    audio, sample_rate = kokoro.create(
        text=str(text),
        voice=voice_id,
        speed=ratio,
        lang=lang,
        trim=True,
    )
    import numpy as np

    audio = np.asarray(audio, dtype=np.float32) * gain
    return _write_temp_wav(audio, sample_rate)


def stop():
    with _lock:
        _stop_locked()


def speak_async(text, volume=1.0, rate_ratio=1.0, cancel_before=False):
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
            if not _models_ready():
                _start_model_download_if_needed()
                raise RuntimeError(_get_kokoro_not_ready_reason())
            wav_path = _synthesize_to_wav(text=str(text), volume=volume, rate_ratio=rate_ratio)
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


def _split_long_text(text, chunk_chars=220):
    raw = re.sub(r"\s+", " ", str(text or "").strip())
    if not raw:
        return []

    sentences = re.split(r"(?<=[.!?;:])\s+", raw)
    chunks = []
    buf = ""
    limit = max(80, int(chunk_chars))

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
            buf = ""
        if len(part) <= limit:
            buf = part
            continue

        words = part.split(" ")
        inner = ""
        for w in words:
            c2 = f"{inner} {w}".strip() if inner else w
            if len(c2) <= limit:
                inner = c2
            else:
                if inner:
                    chunks.append(inner)
                inner = w
        if inner:
            buf = inner

    if buf:
        chunks.append(buf)
    return chunks


def speak_stream_async(text, volume=1.0, rate_ratio=1.0, cancel_before=False, chunk_chars=220):
    """
    Speak long text in chunks so playback can start faster than one-shot synthesis.
    """
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
            if not _models_ready():
                _start_model_download_if_needed()
                raise RuntimeError(_get_kokoro_not_ready_reason())

            chunks = _split_long_text(text, chunk_chars=chunk_chars)
            if not chunks:
                return

            _stop_locked()
            wav_queue = queue.Queue(maxsize=2)
            sentinel = object()

            def _queue_put(item):
                while True:
                    with _lock:
                        if my_token != _token:
                            return False
                    try:
                        wav_queue.put(item, timeout=0.2)
                        return True
                    except queue.Full:
                        continue

            def _producer():
                try:
                    for chunk in chunks:
                        with _lock:
                            if my_token != _token:
                                return
                        wav_path = _synthesize_to_wav(text=chunk, volume=volume, rate_ratio=rate_ratio)
                        if not _queue_put(wav_path):
                            try:
                                os.remove(wav_path)
                            except Exception:
                                pass
                            return
                except Exception as e:
                    _queue_put(e)
                finally:
                    _queue_put(sentinel)

            threading.Thread(target=_producer, daemon=True).start()

            while True:
                with _lock:
                    if my_token != _token:
                        return
                try:
                    item = wav_queue.get(timeout=0.2)
                except queue.Empty:
                    continue

                if item is sentinel:
                    return
                if isinstance(item, Exception):
                    raise item

                wav_path = item
                with _lock:
                    if my_token != _token:
                        try:
                            os.remove(wav_path)
                        except Exception:
                            pass
                        return
                    _current_wav = wav_path
                winsound.PlaySound(
                    wav_path,
                    winsound.SND_FILENAME | winsound.SND_NODEFAULT,
                )
                with _lock:
                    if _current_wav == wav_path:
                        _current_wav = None
                try:
                    os.remove(wav_path)
                except Exception:
                    pass
        except Exception as e:
            _show_error_once(str(e))

    threading.Thread(target=_run, daemon=True).start()


def cancel_all():
    global _token
    with _lock:
        _token += 1
        _stop_locked()


def prepare_async():
    """Preload models/engine in background to reduce first-play latency."""

    def _run():
        try:
            if not _models_ready():
                _start_model_download_if_needed()
                return
            kokoro = _ensure_kokoro()
            # Warm-up one short inference so first real play feels instant.
            kokoro.create("hello", voice="af_heart", speed=1.0, lang="en-us", trim=True)
        except Exception:
            # Startup warm-up should stay silent; actual play will surface errors.
            return

    threading.Thread(target=_run, daemon=True).start()
