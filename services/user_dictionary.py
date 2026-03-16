# -*- coding: utf-8 -*-
import json
import threading
from pathlib import Path


_DICT_PATH = Path(__file__).resolve().parent.parent / "data" / "user_dictionary.json"
_lock = threading.Lock()
_dict_cache = None


def _normalize_key(text):
    return str(text or "").strip().casefold()


def _clean_text(value):
    return str(value or "").strip()


def _load_dict_locked():
    global _dict_cache
    if _dict_cache is not None:
        return _dict_cache
    if not _DICT_PATH.exists():
        _dict_cache = {}
        return _dict_cache
    try:
        with open(_DICT_PATH, "r", encoding="utf-8") as fp:
            data = json.load(fp)
        _dict_cache = data if isinstance(data, dict) else {}
    except Exception:
        _dict_cache = {}
    cleaned = {}
    changed = False
    for raw_key, raw_value in list((_dict_cache or {}).items()):
        key = _normalize_key(raw_key)
        if not key:
            changed = True
            continue
        payload = raw_value if isinstance(raw_value, dict) else {}
        translation = _clean_text(payload.get("translation"))
        pos = _clean_text(payload.get("pos"))
        if not translation and not pos:
            changed = True
            continue
        cleaned[key] = {}
        if translation:
            cleaned[key]["translation"] = translation
        if pos:
            cleaned[key]["pos"] = pos
        if raw_key != key or payload != cleaned[key]:
            changed = True
    _dict_cache = cleaned
    if changed:
        _save_dict_locked()
    return _dict_cache


def _save_dict_locked():
    _DICT_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(_DICT_PATH, "w", encoding="utf-8") as fp:
        json.dump(_dict_cache or {}, fp, ensure_ascii=False, indent=2)


def get_entry(word):
    key = _normalize_key(word)
    if not key:
        return {}
    with _lock:
        data = _load_dict_locked().get(key)
        return dict(data) if isinstance(data, dict) else {}


def get_entries(words):
    result = {}
    with _lock:
        cache = _load_dict_locked()
        for word in words or []:
            key = _normalize_key(word)
            payload = cache.get(key)
            if key and isinstance(payload, dict):
                result[word] = dict(payload)
    return result


def set_entry(word, *, translation=None, pos=None):
    key = _normalize_key(word)
    if not key:
        return False
    translation_value = _clean_text(translation)
    pos_value = _clean_text(pos)
    with _lock:
        cache = _load_dict_locked()
        existing = dict(cache.get(key) or {})
        changed = False
        if translation is not None:
            if translation_value:
                if existing.get("translation") != translation_value:
                    existing["translation"] = translation_value
                    changed = True
            elif "translation" in existing:
                existing.pop("translation", None)
                changed = True
        if pos is not None:
            if pos_value:
                if existing.get("pos") != pos_value:
                    existing["pos"] = pos_value
                    changed = True
            elif "pos" in existing:
                existing.pop("pos", None)
                changed = True
        if existing:
            cache[key] = existing
        elif key in cache:
            cache.pop(key, None)
            changed = True
        if changed:
            _save_dict_locked()
    return True


def apply_entries(entries):
    changed = False
    with _lock:
        cache = _load_dict_locked()
        for item in entries or []:
            if not isinstance(item, dict):
                continue
            key = _normalize_key(item.get("word"))
            if not key:
                continue
            translation = _clean_text(item.get("translation"))
            pos = _clean_text(item.get("pos"))
            payload = dict(cache.get(key) or {})
            if translation:
                payload["translation"] = translation
            if pos:
                payload["pos"] = pos
            if not payload:
                continue
            if cache.get(key) == payload:
                continue
            cache[key] = payload
            changed = True
        if changed:
            _save_dict_locked()
    return changed
