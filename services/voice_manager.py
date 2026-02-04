# -*- coding: utf-8 -*-
SOURCE_PYTTSX3 = "pyttsx3"
SOURCE_SYSTEM = "system"

_voice_source = SOURCE_PYTTSX3
_voice_id = None
_voice_label = None


def get_voice_source():
    return _voice_source


def get_voice_id():
    return _voice_id


def get_voice_label():
    return _voice_label


def set_voice_source(source, voice_id=None, voice_label=None):
    global _voice_source, _voice_id, _voice_label
    if source not in (SOURCE_PYTTSX3, SOURCE_SYSTEM):
        return False
    _voice_source = source
    _voice_id = voice_id if source == SOURCE_SYSTEM else None
    _voice_label = voice_label if source == SOURCE_SYSTEM else None
    return True


def select_voice(engine):
    if _voice_source != SOURCE_SYSTEM or not _voice_id:
        return None
    engine.setProperty("voice", _voice_id)
    return _voice_label
