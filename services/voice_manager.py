# -*- coding: utf-8 -*-
SOURCE_GEMINI = "gemini"
SOURCE_KOKORO = "kokoro"
SOURCE_PIPER = "piper"

_voice_source = SOURCE_GEMINI
_voice_id = "gemini-kore"
_voice_label = "Gemini TTS (UK)"


def get_voice_source():
    return _voice_source


def get_voice_id():
    return _voice_id


def get_voice_label():
    return _voice_label


def set_voice_source(source, voice_id=None, voice_label=None):
    global _voice_source, _voice_id, _voice_label
    if source not in {SOURCE_GEMINI, SOURCE_KOKORO, SOURCE_PIPER}:
        return False
    _voice_source = source
    if _voice_source == SOURCE_KOKORO:
        _voice_id = voice_id if voice_id else "bf_emma"
        _voice_label = voice_label if voice_label else "Kokoro English (UK)"
    elif _voice_source == SOURCE_PIPER:
        _voice_id = voice_id if voice_id else "piper-en-gb"
        _voice_label = voice_label if voice_label else "Piper English"
    else:
        _voice_id = voice_id if voice_id else "gemini-kore"
        _voice_label = voice_label if voice_label else "Gemini TTS (UK)"
    return True
