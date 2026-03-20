# -*- coding: utf-8 -*-
import json
import threading
from pathlib import Path

from services.corpus_search import get_nlp
from services.metadata_repository import JsonMetadataRepository
from services.user_dictionary import get_entries as get_user_dictionary_entries, set_entry as set_user_dictionary_entry


_CACHE_PATH = Path(__file__).resolve().parent.parent / "data" / "pos_cache.json"
_lock = threading.Lock()
_cache_data = None
_repo = None

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
    _cache_data = _get_repo().cleanup()
    return _cache_data


def _save_cache_locked():
    global _cache_data
    _cache_data = _get_repo().export_payload()


def _normalize_value(text):
    return str(text or "").strip()


def _get_repo():
    global _repo
    if _repo is None:
        _repo = JsonMetadataRepository(
            _CACHE_PATH,
            key_normalizer=_normalize_key,
            value_normalizer=_normalize_value,
        )
    return _repo


def get_cached_pos(words):
    result = {}
    overrides = get_user_dictionary_entries(words)
    for word, payload in overrides.items():
        value = str((payload or {}).get("pos") or "").strip()
        if value:
            result[word] = value
    cache_result = _get_repo().get_many(words)
    for word in words:
        if word in result:
            continue
        if word in cache_result:
            result[word] = cache_result[word]
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
        _get_repo().apply_many(updates)
    return result


def set_cached_pos(word, pos_label):
    key = _normalize_key(word)
    if not key:
        return False
    value = _normalize_value(pos_label)
    set_user_dictionary_entry(word, translation=None, pos=value if value else "")
    _get_repo().set_one(word, value)
    return True


def apply_cached_pos(pairs):
    return _get_repo().apply_many(pairs)
