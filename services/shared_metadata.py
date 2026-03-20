# -*- coding: utf-8 -*-
import json
from pathlib import Path

from services.phonetics import apply_cached_phonetics
from services.translation import apply_cached_translations
from services.word_analysis import apply_cached_pos


_DATA_DIR = Path(__file__).resolve().parent.parent / "data"
_TRANSLATION_CACHE = _DATA_DIR / "translation_cache.json"
_POS_CACHE = _DATA_DIR / "pos_cache.json"
_PHONETIC_CACHE = _DATA_DIR / "phonetics_cache.json"


def _read_json_dict(path):
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}
    return data if isinstance(data, dict) else {}


def export_shared_metadata_payload():
    return {
        "translations": _read_json_dict(_TRANSLATION_CACHE),
        "pos": _read_json_dict(_POS_CACHE),
        "phonetics": _read_json_dict(_PHONETIC_CACHE),
    }


def import_shared_metadata_payload(payload):
    data = payload if isinstance(payload, dict) else {}
    translations = dict(data.get("translations") or {})
    pos = dict(data.get("pos") or {})
    phonetics = dict(data.get("phonetics") or {})
    return {
        "translations": int(apply_cached_translations(translations) or 0),
        "pos": int(apply_cached_pos(pos) or 0),
        "phonetics": int(apply_cached_phonetics(phonetics) or 0),
    }
