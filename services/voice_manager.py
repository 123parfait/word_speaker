# -*- coding: utf-8 -*-
SOURCE_KOKORO = "kokoro"

_voice_source = SOURCE_KOKORO
_voice_id = "bf_emma"
_voice_label = "English (UK)"


def get_voice_source():
    return _voice_source


def get_voice_id():
    return _voice_id


def get_voice_label():
    return _voice_label


def set_voice_source(source, voice_id=None, voice_label=None):
    global _voice_source, _voice_id, _voice_label
    if source != SOURCE_KOKORO:
        return False
    _voice_source = source
    _voice_id = voice_id if voice_id else "bf_emma"
    _voice_label = voice_label if voice_label else "English (UK)"
    return True
