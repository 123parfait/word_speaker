# -*- coding: utf-8 -*-
import json
import threading
from pathlib import Path

import nltk

from services.app_config import get_generation_model, get_llm_api_key
from services.corpus_search import get_nlp
from services.gemini_writer import DEFAULT_GEMINI_MODEL, _request_gemini


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


def get_cached_synonym_result(word):
    key = _normalize_key(word)
    if not key:
        return None
    with _lock:
        cache = _load_cache_locked()
        values = cache.get(key)
        if isinstance(values, dict):
            payload = dict(values)
            payload["synonyms"] = list(payload.get("synonyms") or [])
            payload["source"] = str(payload.get("source") or "local").strip().lower()
            return payload
        if isinstance(values, list):
            target = str(word or "").strip()
            return {
                "word": target,
                "focus": target,
                "source": "local",
                "synonyms": list(values),
            }
    return None


def get_cached_synonyms(word):
    payload = get_cached_synonym_result(word)
    if not isinstance(payload, dict):
        return []
    return list(payload.get("synonyms") or [])


def _set_cached_synonym_result(word, result):
    key = _normalize_key(word)
    if not key:
        return
    payload = dict(result or {})
    payload["word"] = str(payload.get("word") or word or "").strip()
    payload["focus"] = str(payload.get("focus") or payload["word"] or "").strip()
    payload["source"] = str(payload.get("source") or "local").strip().lower()
    payload["synonyms"] = [
        str(item or "").strip()
        for item in (payload.get("synonyms") or [])
        if str(item or "").strip()
    ]
    with _lock:
        cache = _load_cache_locked()
        cache[key] = payload
        _save_cache_locked()


def _pick_focus_token(doc):
    tokens = [token for token in doc if not token.is_space and not token.is_punct]
    if not tokens:
        return None
    for token in tokens:
        if token.head == token:
            return token
    return tokens[0]


def _build_synonym_prompt(word, limit):
    target = str(word or "").strip()
    capped_limit = max(3, min(int(limit or 12), 15))
    return (
        "You are an IELTS English coach.\n"
        "Return near-synonyms or closely related interchangeable words for the target word or phrase.\n"
        "Rules:\n"
        "- English only\n"
        "- Return plain JSON only\n"
        "- Use this exact JSON schema: {\"focus\":\"...\",\"synonyms\":[\"...\", \"...\"]}\n"
        f"- Return at most {capped_limit} synonyms\n"
        "- Prefer IELTS-useful, semantically close, practical words or phrases\n"
        "- Avoid vague associations and topic words\n"
        "- Keep the results highly relevant to the target meaning\n"
        "- If the target is a phrase, prefer phrase-level alternatives when possible\n"
        "- Do not include the original target in the synonyms array\n"
        "- Do not include explanations\n\n"
        f"Target: {target}\n"
    )


def _parse_gemini_synonym_response(word, text, limit):
    raw = str(text or "").strip()
    if not raw:
        raise RuntimeError("Gemini returned empty synonym content.")
    data = json.loads(raw)
    if not isinstance(data, dict):
        raise RuntimeError("Gemini synonym response is not an object.")
    focus = str(data.get("focus") or word or "").strip()
    values = data.get("synonyms") or []
    if not isinstance(values, list):
        raise RuntimeError("Gemini synonym response has no synonym list.")
    seen = set()
    original_forms = {_normalize_key(word), _normalize_key(focus)}
    results = []
    for item in values:
        value = str(item or "").replace("_", " ").strip()
        key = _normalize_key(value)
        if not value or key in seen or key in original_forms:
            continue
        seen.add(key)
        results.append(value)
        if len(results) >= max(1, int(limit or 12)):
            break
    return {"word": str(word or "").strip(), "focus": focus, "source": "gemini", "synonyms": results}


def _get_synonyms_with_gemini(word, limit=12):
    api_key = get_llm_api_key()
    if not str(api_key or "").strip():
        raise RuntimeError("LLM API key is empty.")
    model = get_generation_model() or DEFAULT_GEMINI_MODEL
    text = _request_gemini(
        prompt=_build_synonym_prompt(word, limit),
        api_key=api_key,
        model=model,
        timeout=35,
        temperature=0.2,
        max_tokens=180,
    )
    return _parse_gemini_synonym_response(word, text, limit)


def _get_synonyms_local(word, limit=12):
    target = str(word or "").strip()
    if not target:
        return {"word": "", "focus": "", "source": "local", "synonyms": []}

    _ensure_wordnet_ready()
    nlp, _mode = get_nlp()
    _ensure_wordnet_pipe(nlp)

    doc = nlp(target)
    focus = _pick_focus_token(doc)
    if focus is None:
        return {"word": target, "focus": "", "source": "local", "synonyms": []}

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

    return {"word": target, "focus": focus.text, "source": "local", "synonyms": results}


def get_synonyms(word, limit=12):
    target = str(word or "").strip()
    if not target:
        return {"word": "", "focus": "", "source": "local", "synonyms": []}

    cached = get_cached_synonym_result(target)
    if cached:
        cached["synonyms"] = list(cached.get("synonyms") or [])[: max(1, int(limit or 12))]
        return cached

    try:
        result = _get_synonyms_with_gemini(target, limit=limit)
    except Exception:
        result = _get_synonyms_local(target, limit=limit)

    _set_cached_synonym_result(target, result)
    return result
