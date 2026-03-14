# -*- coding: utf-8 -*-
import json
import threading
from pathlib import Path

from services.corpus_search import get_nlp


_CACHE_PATH = Path(__file__).resolve().parent.parent / "data" / "pos_cache.json"
_lock = threading.Lock()
_cache_data = None

_POS_LABELS = {
    "NOUN": "n.",
    "PROPN": "propn.",
    "VERB": "v.",
    "AUX": "aux.",
    "ADJ": "adj.",
    "ADV": "adv.",
    "PRON": "pron.",
    "DET": "det.",
    "ADP": "prep.",
    "CCONJ": "conj.",
    "SCONJ": "conj.",
    "NUM": "num.",
    "PART": "part.",
    "INTJ": "int.",
}


def _normalize_key(text):
    return str(text or "").strip().casefold()


def _load_cache_locked():
    global _cache_data
    if _cache_data is not None:
        return _cache_data
    if not _CACHE_PATH.exists():
        _cache_data = {}
        return _cache_data
    try:
        with open(_CACHE_PATH, "r", encoding="utf-8") as fp:
            data = json.load(fp)
        _cache_data = data if isinstance(data, dict) else {}
    except Exception:
        _cache_data = {}
    return _cache_data


def _save_cache_locked():
    _CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(_CACHE_PATH, "w", encoding="utf-8") as fp:
        json.dump(_cache_data or {}, fp, ensure_ascii=False, indent=2)


def get_cached_pos(words):
    result = {}
    with _lock:
        cache = _load_cache_locked()
        for word in words:
            key = _normalize_key(word)
            if key and key in cache:
                result[word] = str(cache.get(key) or "")
    return result


def _guess_pos_label(text):
    value = str(text or "").strip()
    if not value:
        return ""
    nlp, _mode = get_nlp()
    doc = nlp(value)
    tokens = [token for token in doc if not token.is_space and not token.is_punct]
    if not tokens:
        return ""
    if any(str(token.tag_ or "").upper() == "MD" for token in tokens):
        return "modal."
    root = None
    for token in tokens:
        if token.head == token:
            root = token
            break
    if root is None:
        root = tokens[0]
    pos = str(root.pos_ or "").upper()
    return _POS_LABELS.get(pos, pos.lower() + "." if pos else "")


def analyze_words(words):
    result = get_cached_pos(words)
    missing = []
    for word in words:
        if word in result:
            continue
        key = _normalize_key(word)
        if key:
            missing.append(word)

    updates = {}
    for word in missing:
        try:
            label = _guess_pos_label(word)
        except Exception:
            label = ""
        result[word] = label
        if label:
            updates[word] = label

    if updates:
        with _lock:
            cache = _load_cache_locked()
            changed = False
            for word, label in updates.items():
                key = _normalize_key(word)
                if not key:
                    continue
                if cache.get(key) == label:
                    continue
                cache[key] = label
                changed = True
            if changed:
                _save_cache_locked()
    return result


def set_cached_pos(word, pos_label):
    key = _normalize_key(word)
    if not key:
        return False
    value = str(pos_label or "").strip()
    with _lock:
        cache = _load_cache_locked()
        changed = False
        if value:
            if cache.get(key) != value:
                cache[key] = value
                changed = True
        elif key in cache:
            cache.pop(key, None)
            changed = True
        if changed:
            _save_cache_locked()
    return True
