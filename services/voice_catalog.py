# -*- coding: utf-8 -*-
from pathlib import Path

from services.app_config import get_tts_api_provider


BASE_DIR = Path(__file__).resolve().parent.parent
KOKORO_DIR = BASE_DIR / "data" / "models" / "kokoro"
KOKORO_MODEL = KOKORO_DIR / "kokoro-v1.0.onnx"
KOKORO_VOICES = KOKORO_DIR / "voices-v1.0.bin"
PIPER_DIR = BASE_DIR / "data" / "models" / "piper"


def kokoro_ready():
    return KOKORO_MODEL.exists() and KOKORO_VOICES.exists()


def get_kokoro_paths():
    return str(KOKORO_MODEL), str(KOKORO_VOICES)


def get_kokoro_placeholder_voice():
    return {
        "source": "kokoro",
        "id": "kokoro:not-ready",
        "name": "Kokoro (Not Ready)",
        "languages": ["local"],
    }


def _piper_model_candidates():
    if not PIPER_DIR.exists():
        return []
    return sorted(PIPER_DIR.glob("*.onnx"))


def _guess_language_from_name(name):
    value = str(name or "").lower()
    if "en-gb" in value or "en_gb" in value or "british" in value or "uk" in value:
        return "en-GB"
    if "en-us" in value or "en_us" in value or "american" in value or "lessac" in value:
        return "en-US"
    return "en"


def get_piper_voices():
    voices = []
    for model_path in _piper_model_candidates():
        stem = model_path.stem
        config_path = model_path.with_suffix(".onnx.json")
        if not config_path.exists():
            continue
        voice_id = f"piper:{stem}"
        lang = _guess_language_from_name(stem)
        display = stem.replace("_", " ").replace("-", " ").strip() or "Piper Voice"
        voices.append(
            {
                "source": "piper",
                "id": voice_id,
                "name": f"Piper {display.title()}",
                "languages": [lang],
                "model_path": str(model_path),
                "config_path": str(config_path),
            }
        )
    return voices


def piper_ready():
    return bool(get_piper_voices())


def get_piper_voice_profile(voice_id=None):
    voices = get_piper_voices()
    if not voices:
        return {}
    target_id = str(voice_id or "").strip()
    for profile in voices:
        if profile.get("id") == target_id:
            return dict(profile)
    return dict(voices[0])


def get_piper_placeholder_voice():
    return {
        "source": "piper",
        "id": "piper:not-ready",
        "name": "Piper (Not Ready)",
        "languages": ["local"],
    }


def list_system_voices():
    online_name = "ElevenLabs TTS" if get_tts_api_provider() == "elevenlabs" else "Gemini TTS (UK)"
    voices = [
        {
            "source": "gemini",
            "id": "gemini-kore",
            "name": online_name,
            "languages": ["en-GB"],
        },
    ]
    if kokoro_ready():
        voices.append(
            {
                "source": "kokoro",
                "id": "bf_emma",
                "name": "Kokoro English (UK)",
                "languages": ["en-GB"],
            }
        )
    else:
        voices.append(get_kokoro_placeholder_voice())
    if piper_ready():
        voices.extend(get_piper_voices())
    else:
        voices.append(get_piper_placeholder_voice())
    return voices


def get_voice_profile(source, voice_id):
    for profile in list_system_voices():
        if profile.get("source") == source and profile.get("id") == voice_id:
            return dict(profile)
    return dict(list_system_voices()[0])
