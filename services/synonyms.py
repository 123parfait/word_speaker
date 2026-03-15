# -*- coding: utf-8 -*-
import json
import threading
from pathlib import Path

import nltk

from services.corpus_search import get_nlp


_CACHE_PATH = Path(__file__).resolve().parent.parent / "data" / "synonyms_cache.json"
_NLTK_DATA_DIR = Path(__file__).resolve().parent.parent / "data" / "nltk_data"
_lock = threading.Lock()
_cache_data = None
_wordnet_ready = False


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


def _ensure_wordnet_ready():
    global _wordnet_ready
    if _wordnet_ready:
        return
    data_path = str(_NLTK_DATA_DIR)
    if data_path not in nltk.data.path:
        nltk.data.path.insert(0, data_path)
    try:
        import spacy_wordnet.wordnet_annotator  # noqa: F401
        from nltk.corpus import wordnet as wn

        wn.ensure_loaded()
    except LookupError as exc:
        raise RuntimeError("WordNet data is missing. Download 'wordnet' and 'omw-1.4' into data/nltk_data.") from exc
    except Exception as exc:
        raise RuntimeError(f"spacy-wordnet is not ready: {exc}") from exc
    _wordnet_ready = True


def _ensure_wordnet_pipe(nlp):
    if "spacy_wordnet" in nlp.pipe_names:
        return
    if "wordnet" in nlp.pipe_names:
        return
    if "tagger" in nlp.pipe_names:
        nlp.add_pipe("spacy_wordnet", after="tagger")
    else:
        nlp.add_pipe("spacy_wordnet")


def get_cached_synonyms(word):
    key = _normalize_key(word)
    if not key:
        return []
    with _lock:
        cache = _load_cache_locked()
        values = cache.get(key)
        return list(values) if isinstance(values, list) else []


def _set_cached_synonyms(word, synonyms):
    key = _normalize_key(word)
    if not key:
        return
    values = [str(item or "").strip() for item in (synonyms or []) if str(item or "").strip()]
    with _lock:
        cache = _load_cache_locked()
        cache[key] = values
        _save_cache_locked()


def _pick_focus_token(doc):
    tokens = [token for token in doc if not token.is_space and not token.is_punct]
    if not tokens:
        return None
    for token in tokens:
        if token.head == token:
            return token
    return tokens[0]


def get_synonyms(word, limit=12):
    target = str(word or "").strip()
    if not target:
        return {"word": "", "focus": "", "synonyms": []}

    cached = get_cached_synonyms(target)
    if cached:
        return {"word": target, "focus": target, "synonyms": cached[: max(1, int(limit or 12))]}

    _ensure_wordnet_ready()
    nlp, _mode = get_nlp()
    _ensure_wordnet_pipe(nlp)

    doc = nlp(target)
    focus = _pick_focus_token(doc)
    if focus is None:
        return {"word": target, "focus": "", "synonyms": []}

    seen = set()
    results = []
    original_forms = {
        _normalize_key(target),
        _normalize_key(focus.text),
        _normalize_key(focus.lemma_),
    }

    for lemma in focus._.wordnet.lemmas():
        try:
            name = lemma.name()
        except Exception:
            name = str(lemma or "")
        value = str(name or "").replace("_", " ").strip()
        key = _normalize_key(value)
        if not value or key in original_forms or key in seen:
            continue
        seen.add(key)
        results.append(value)
        if len(results) >= max(1, int(limit or 12)):
            break

    _set_cached_synonyms(target, results)
    return {"word": target, "focus": focus.text, "synonyms": results}
