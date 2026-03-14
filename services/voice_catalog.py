# -*- coding: utf-8 -*-
import os
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent.parent
KOKORO_DIR = BASE_DIR / "data" / "models" / "kokoro"
KOKORO_MODEL = KOKORO_DIR / "kokoro-v1.0.onnx"
KOKORO_VOICES = KOKORO_DIR / "voices-v1.0.bin"


def kokoro_ready():
    return KOKORO_MODEL.exists() and KOKORO_VOICES.exists()


def get_kokoro_paths():
    return str(KOKORO_MODEL), str(KOKORO_VOICES)


def list_system_voices():
    voices = [
        {
            "source": "gemini",
            "id": "gemini-kore",
            "name": "Gemini TTS (UK)",
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
    return voices


def get_voice_profile(source, voice_id):
    for profile in list_system_voices():
        if profile.get("source") == source and profile.get("id") == voice_id:
            return dict(profile)
    return dict(list_system_voices()[0])
