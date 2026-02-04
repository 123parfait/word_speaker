# -*- coding: utf-8 -*-
import pyttsx3


def _safe_attr(obj, name):
    value = getattr(obj, name, None)
    return value if value is not None else ""


def list_system_voices():
    engine = None
    try:
        engine = pyttsx3.init(driverName="sapi5")
        voices = engine.getProperty("voices") or []
        result = []
        for v in voices:
            voice_id = _safe_attr(v, "id")
            name = _safe_attr(v, "name")
            langs = _safe_attr(v, "languages")
            gender = _safe_attr(v, "gender")
            result.append(
                {
                    "id": voice_id,
                    "name": name,
                    "languages": langs,
                    "gender": gender,
                }
            )
        return result
    except Exception:
        return []
    finally:
        if engine:
            try:
                engine.stop()
            except Exception:
                pass
