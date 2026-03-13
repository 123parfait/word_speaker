# -*- coding: utf-8 -*-

# Kokoro voices focused on English accents.
_VOICE_PROFILES = [
    {
        "id": "af_heart",
        "name": "English (US)",
        "languages": ["en-us"],
    },
    {
        "id": "bf_emma",
        "name": "English (UK)",
        "languages": ["en-gb"],
    },
]


def list_system_voices():
    # Keep function name for backward compatibility with the current UI call site.
    return [dict(v) for v in _VOICE_PROFILES]


def get_voice_profile(voice_id):
    for profile in _VOICE_PROFILES:
        if profile["id"] == voice_id:
            return dict(profile)
    for profile in _VOICE_PROFILES:
        if profile.get("id") == "bf_emma":
            return dict(profile)
    return dict(_VOICE_PROFILES[0])
