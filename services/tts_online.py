import base64
import json
import urllib.error
import urllib.request


def extract_error_message(http_error):
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


def request_gemini_tts(
    text,
    *,
    short_text,
    api_key,
    timeout,
    url,
    voice_name,
    style_short,
    style_long,
    normalize_text,
    log_info,
    log_error,
):
    api_key = str(api_key or "").strip()
    if not api_key:
        raise RuntimeError("TTS API key is empty.")
    spoken_text = normalize_text(text)
    prompt = style_short if short_text else style_long
    payload = {
        "contents": [{"parts": [{"text": f"{prompt}\n\nText:\n{spoken_text}"}]}],
        "generationConfig": {
            "responseModalities": ["AUDIO"],
            "speechConfig": {
                "voiceConfig": {
                    "prebuiltVoiceConfig": {
                        "voiceName": voice_name,
                    }
                }
            },
        },
    }
    req = urllib.request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "x-goog-api-key": api_key,
            "Content-Type": "application/json",
        },
        method="POST",
    )
    try:
        log_info("tts_request_start", provider="gemini", short_text=short_text, timeout=timeout, text=spoken_text[:120])
        with urllib.request.urlopen(req, timeout=timeout) as response:
            return json.loads(response.read().decode("utf-8", errors="ignore"))
    except urllib.error.HTTPError as exc:
        message = extract_error_message(exc)
        log_error("tts_request_http_error", provider="gemini", short_text=short_text, error=message)
        raise RuntimeError(message) from exc
    except urllib.error.URLError as exc:
        log_error("tts_request_url_error", provider="gemini", short_text=short_text, error=exc.reason)
        raise RuntimeError(f"Gemini TTS request failed: {exc.reason}") from exc
    except Exception as exc:
        log_error("tts_request_error", provider="gemini", short_text=short_text, error=exc)
        raise RuntimeError(f"Gemini TTS request failed: {exc}") from exc


def request_elevenlabs_tts(
    text,
    *,
    short_text,
    api_key,
    timeout,
    url,
    model_id,
    normalize_text,
    log_info,
    log_error,
):
    api_key = str(api_key or "").strip()
    if not api_key:
        raise RuntimeError("TTS API key is empty.")
    spoken_text = normalize_text(text)
    payload = {
        "text": spoken_text,
        "model_id": model_id,
        "voice_settings": {
            "stability": 0.45,
            "similarity_boost": 0.75,
            "style": 0.1 if short_text else 0.25,
            "use_speaker_boost": True,
        },
    }
    req = urllib.request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "xi-api-key": api_key,
            "Content-Type": "application/json",
            "Accept": "audio/pcm",
        },
        method="POST",
    )
    try:
        log_info("tts_request_start", provider="elevenlabs", short_text=short_text, timeout=timeout, text=spoken_text[:120])
        with urllib.request.urlopen(req, timeout=timeout) as response:
            return response.read()
    except urllib.error.HTTPError as exc:
        message = extract_error_message(exc)
        log_error("tts_request_http_error", provider="elevenlabs", short_text=short_text, error=message)
        raise RuntimeError(message) from exc
    except urllib.error.URLError as exc:
        log_error("tts_request_url_error", provider="elevenlabs", short_text=short_text, error=exc.reason)
        raise RuntimeError(f"ElevenLabs TTS request failed: {exc.reason}") from exc
    except Exception as exc:
        log_error("tts_request_error", provider="elevenlabs", short_text=short_text, error=exc)
        raise RuntimeError(f"ElevenLabs TTS request failed: {exc}") from exc


def request_online_tts(
    text,
    *,
    short_text,
    provider,
    api_key,
    timeout,
    primary_online_provider,
    request_gemini,
    request_elevenlabs,
):
    backend = str(provider or primary_online_provider()).strip().lower()
    if backend == "elevenlabs":
        return request_elevenlabs(text, short_text=short_text, api_key=api_key, timeout=timeout), "elevenlabs"
    return request_gemini(text, short_text=short_text, api_key=api_key, timeout=timeout), "gemini"


def extract_pcm_bytes(data):
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


def validate_gemini_tts_api_key(
    api_key,
    *,
    timeout,
    url,
    voice_name,
    style_short,
    extract_pcm,
):
    if not str(api_key or "").strip():
        raise RuntimeError("TTS API key is empty.")
    payload = {
        "contents": [{"parts": [{"text": f"{style_short}\n\nText:\nTest"}]}],
        "generationConfig": {
            "responseModalities": ["AUDIO"],
            "speechConfig": {
                "voiceConfig": {
                    "prebuiltVoiceConfig": {
                        "voiceName": voice_name,
                    }
                }
            },
        },
    }
    req = urllib.request.Request(
        url,
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
    except urllib.error.HTTPError as exc:
        raise RuntimeError(extract_error_message(exc)) from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Gemini TTS request failed: {exc.reason}") from exc
    except Exception as exc:
        raise RuntimeError(f"Gemini TTS request failed: {exc}") from exc
    extract_pcm(data)
    return True


def validate_elevenlabs_tts_api_key(api_key, *, timeout, url, model_id):
    if not str(api_key or "").strip():
        raise RuntimeError("TTS API key is empty.")
    payload = {
        "text": "Test",
        "model_id": model_id,
        "voice_settings": {
            "stability": 0.45,
            "similarity_boost": 0.75,
            "style": 0.1,
            "use_speaker_boost": True,
        },
    }
    req = urllib.request.Request(
        url,
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
    except urllib.error.HTTPError as exc:
        raise RuntimeError(extract_error_message(exc)) from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"ElevenLabs TTS request failed: {exc.reason}") from exc
    except Exception as exc:
        raise RuntimeError(f"ElevenLabs TTS request failed: {exc}") from exc
    if not data:
        raise RuntimeError("ElevenLabs TTS returned no audio.")
    return True


def validate_tts_api_key(api_key, provider, *, timeout, validate_gemini, validate_elevenlabs):
    backend = str(provider or "").strip().lower()
    if backend == "elevenlabs":
        return validate_elevenlabs(api_key, timeout=timeout)
    return validate_gemini(api_key, timeout=timeout)
