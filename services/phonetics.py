# -*- coding: utf-8 -*-
import json
import re
import threading
import urllib.error
import urllib.request
from pathlib import Path

from services.app_config import get_generation_model, get_llm_api_key
from services.metadata_repository import JsonMetadataRepository


_CACHE_PATH = Path(__file__).resolve().parent.parent / "data" / "phonetics_cache.json"
_GEMINI_URL = "https://generativelanguage.googleapis.com/v1beta/openai/chat/completions"
_lock = threading.Lock()
_cache_data = None
_repo = None


def _normalize_key(text):
    return str(text or "").strip().casefold()


def _normalize_phonetic(text):
    value = str(text or "").strip()
    if not value:
        return ""
    value = re.sub(r"\s+", " ", value)
    value = value.replace("／", "/")
    if value.startswith("[") and value.endswith("]"):
        value = f"/{value[1:-1].strip()}/"
    if value and not value.startswith("/"):
        value = "/" + value.lstrip("/")
    if value and not value.endswith("/"):
        value = value.rstrip("/") + "/"
    value = re.sub(r"/{2,}", "/", value)
    return value.strip()


def _strip_code_fence(text):
    raw = str(text or "").strip()
    if raw.startswith("```"):
        raw = re.sub(r"^```[a-zA-Z0-9_-]*\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)
    return raw.strip()


def _extract_json_object(text):
    raw = _strip_code_fence(text)
    if not raw:
        return {}
    try:
        data = json.loads(raw)
        return data if isinstance(data, dict) else {}
    except Exception:
        pass
    match = re.search(r"\{[\s\S]*\}", raw)
    if not match:
        return {}
    try:
        data = json.loads(match.group(0))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _load_cache_locked():
    global _cache_data
    _cache_data = _get_repo().cleanup()
    return _cache_data


def _save_cache_locked():
    global _cache_data
    _cache_data = _get_repo().export_payload()


def _get_repo():
    global _repo
    if _repo is None:
        _repo = JsonMetadataRepository(
            _CACHE_PATH,
            key_normalizer=_normalize_key,
            value_normalizer=_normalize_phonetic,
        )
    return _repo


def get_cached_phonetics(words):
    return _get_repo().get_many(words)


def _update_cache(pairs):
    return bool(_get_repo().apply_many(pairs))


def set_cached_phonetic(word, phonetic):
    key = _normalize_key(word)
    if not key:
        return False
    value = _normalize_phonetic(phonetic)
    _get_repo().set_one(word, value)
    return True


def apply_cached_phonetics(pairs):
    return _get_repo().apply_many(pairs)


def _request_gemini_phonetic_map(words, timeout=35):
    api_key = str(get_llm_api_key() or "").strip()
    if not api_key:
        raise RuntimeError("Gemini API key is empty.")
    requested_words = [str(word or "").strip() for word in words if str(word or "").strip()]
    if not requested_words:
        return {}

    prompt = (
        "Return the British English IPA pronunciation for each input word or short phrase.\n"
        "Rules:\n"
        "- Return JSON only.\n"
        "- Keep every input string exactly as the JSON key.\n"
        "- Each value must be a single UK IPA string wrapped in forward slashes.\n"
        "- No explanations, no labels, no American IPA, no example sentences.\n"
        "- If uncertain, use an empty string.\n\n"
        f"Input terms:\n{json.dumps(requested_words, ensure_ascii=False)}"
    )
    payload = {
        "model": str(get_generation_model() or "gemini-2.5-flash").strip(),
        "messages": [
            {
                "role": "system",
                "content": "You are an English pronunciation assistant. Output valid JSON only.",
            },
            {"role": "user", "content": prompt},
        ],
        "temperature": 0.1,
        "max_tokens": max(256, 40 * len(requested_words)),
    }
    req = urllib.request.Request(
        _GEMINI_URL,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as response:
            data = json.loads(response.read().decode("utf-8", errors="ignore"))
    except urllib.error.HTTPError as e:
        try:
            raw = e.read().decode("utf-8", errors="ignore")
        except Exception:
            raw = str(e)
        raise RuntimeError(raw or str(e)) from e
    except urllib.error.URLError as e:
        raise RuntimeError(f"Gemini request failed: {e.reason}") from e
    except Exception as e:
        raise RuntimeError(f"Gemini request failed: {e}") from e

    choices = data.get("choices") or []
    if not choices:
        raise RuntimeError("Gemini returned no choices.")
    message = (choices[0] or {}).get("message") or {}
    content = message.get("content")
    if isinstance(content, list):
        content = "".join(str((item or {}).get("text") or "") for item in content if isinstance(item, dict))
    mapping = _extract_json_object(content)
    result = {}
    for word in requested_words:
        phonetic = _normalize_phonetic(mapping.get(word) or "")
        if phonetic:
            result[word] = phonetic
    return result


def get_phonetics(words):
    result = get_cached_phonetics(words)
    missing = []
    seen_missing = set()
    for word in words or []:
        if word in result:
            continue
        key = _normalize_key(word)
        if not key or key in seen_missing:
            continue
        seen_missing.add(key)
        missing.append(word)
    if not missing:
        return result
    try:
        updates = _request_gemini_phonetic_map(missing)
    except Exception:
        updates = {}
    for word in missing:
        result[word] = _normalize_phonetic(updates.get(word) or "")
    if updates:
        _update_cache(updates)
    return result
