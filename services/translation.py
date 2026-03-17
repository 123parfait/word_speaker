# -*- coding: utf-8 -*-
import json
import os
import re
import threading
import unicodedata
from pathlib import Path

from services.user_dictionary import get_entries as get_user_dictionary_entries, set_entry as set_user_dictionary_entry

_lock = threading.Lock()
_translation = None
_init_error = None
_prepare_started = False
_CACHE_PATH = Path(__file__).resolve().parent.parent / "data" / "translation_cache.json"
_cache_data = None
_RUNTIME_PATHS_READY = False

_DUPLICATE_BRACKET_PATTERNS = [
    ("(", ")"),
    ("（", "）"),
    ("[", "]"),
    ("【", "】"),
]


def _find_lang(langs, prefix):
    prefix = str(prefix).lower()
    for lang in langs:
        code = str(getattr(lang, "code", "")).lower()
        if code == prefix or code.startswith(prefix + "_") or code.startswith(prefix + "-"):
            return lang
    return None


def _runtime_base_dir():
    import sys

    if getattr(sys, "frozen", False):
        meipass = getattr(sys, "_MEIPASS", "")
        if meipass:
            return Path(meipass)
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def _configure_runtime_paths():
    global _RUNTIME_PATHS_READY
    if _RUNTIME_PATHS_READY:
        return

    base_dir = _runtime_base_dir()
    candidate_dirs = [
        base_dir / "data" / "argos_packages",
        base_dir / "vendor" / "argos_packages",
    ]
    for candidate in candidate_dirs:
        if not candidate.exists() or not candidate.is_dir():
            continue
        try:
            entries = [item for item in candidate.iterdir() if item.is_dir()]
        except Exception:
            entries = []
        if not entries:
            continue
        os.environ["ARGOS_PACKAGES_DIR"] = str(candidate)
        break

    _RUNTIME_PATHS_READY = True


def _ensure_translation():
    global _translation, _init_error
    with _lock:
        if _translation is not None:
            return _translation

    try:
        _configure_runtime_paths()
        import argostranslate.package
        import argostranslate.translate
    except Exception as e:
        err = (
            "Argos Translate is not available. Install it with: "
            "pip install argostranslate"
        )
        with _lock:
            _init_error = err
        raise RuntimeError(err) from e

    try:
        installed_languages = argostranslate.translate.get_installed_languages()
        from_lang = _find_lang(installed_languages, "en")
        to_lang = _find_lang(installed_languages, "zh")
        if not (from_lang and to_lang):
            argostranslate.package.update_package_index()
            available_packages = argostranslate.package.get_available_packages()
            package = None
            for pkg in available_packages:
                if pkg.from_code == "en" and pkg.to_code.startswith("zh"):
                    package = pkg
                    break
            if package is None:
                raise RuntimeError("No Argos package found for English -> Chinese.")
            download_path = package.download()
            argostranslate.package.install_from_path(download_path)
            installed_languages = argostranslate.translate.get_installed_languages()
            from_lang = _find_lang(installed_languages, "en")
            to_lang = _find_lang(installed_languages, "zh")
            if not (from_lang and to_lang):
                raise RuntimeError("Failed to initialize English -> Chinese translation package.")

        translation = from_lang.get_translation(to_lang)
        with _lock:
            _translation = translation
            _init_error = None
        return translation
    except Exception as e:
        with _lock:
            _init_error = str(e)
        raise


def _normalize_cache_key(text):
    return str(text or "").strip().casefold()


def _normalize_translation_text(text):
    value = unicodedata.normalize("NFKC", str(text or "")).strip()
    if not value:
        return ""

    value = re.sub(r"\s+", " ", value)
    value = re.sub(r"\s*([,;:，；：])\s*", r"\1 ", value)
    value = re.sub(r"\s+\)", ")", value)
    value = re.sub(r"\(\s+", "(", value)
    value = re.sub(r"\s+）", "）", value)
    value = re.sub(r"（\s+", "（", value)
    value = value.strip(" ,;:，；：")

    for left, right in _DUPLICATE_BRACKET_PATTERNS:
        escaped_left = re.escape(left)
        escaped_right = re.escape(right)
        pattern = rf"^(?P<outer>.+?)\s*{escaped_left}\s*(?P<inner>.+?)\s*{escaped_right}$"
        match = re.match(pattern, value)
        if not match:
            continue
        outer = str(match.group("outer") or "").strip()
        inner = str(match.group("inner") or "").strip()
        if outer and inner and outer.casefold() == inner.casefold():
            value = outer
            break

    separators = [" | ", "; ", "；", ", ", "，", "/"]
    for separator in separators:
        if separator not in value:
            continue
        pieces = [piece.strip() for piece in value.split(separator)]
        deduped = []
        seen = set()
        for piece in pieces:
            if not piece:
                continue
            key = piece.casefold()
            if key in seen:
                continue
            seen.add(key)
            deduped.append(piece)
        if deduped:
            value = separator.join(deduped)

    return value.strip()


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
    if not isinstance(_cache_data, dict):
        _cache_data = {}
    cleaned = {}
    changed = False
    for key, value in list(_cache_data.items()):
        clean_key = _normalize_cache_key(key)
        clean_value = _normalize_translation_text(value)
        if not clean_key or not clean_value:
            changed = True
            continue
        if clean_key in cleaned and cleaned[clean_key] == clean_value:
            changed = True
            continue
        cleaned[clean_key] = clean_value
        if key != clean_key or str(value or "") != clean_value:
            changed = True
    _cache_data = cleaned
    if changed:
        _save_cache_locked()
    return _cache_data


def _save_cache_locked():
    _CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(_CACHE_PATH, "w", encoding="utf-8") as fp:
        json.dump(_cache_data or {}, fp, ensure_ascii=False, indent=2)


def get_cached_translations(words):
    result = {}
    overrides = get_user_dictionary_entries(words)
    for word, payload in overrides.items():
        value = _normalize_translation_text((payload or {}).get("translation"))
        if value:
            result[word] = value
    with _lock:
        cache = _load_cache_locked()
        for word in words:
            if word in result:
                continue
            key = _normalize_cache_key(word)
            if key in cache:
                result[word] = _normalize_translation_text(cache.get(key) or "")
    return result


def _update_translation_cache(pairs):
    changed = False
    with _lock:
        cache = _load_cache_locked()
        for word, zh_text in pairs.items():
            key = _normalize_cache_key(word)
            value = _normalize_translation_text(zh_text)
            if not key or not value:
                continue
            if cache.get(key) == value:
                continue
            cache[key] = value
            changed = True
        if changed:
            _save_cache_locked()
    return changed


def set_cached_translation(word, zh_text):
    key = _normalize_cache_key(word)
    if not key:
        return False
    value = _normalize_translation_text(zh_text)
    set_user_dictionary_entry(word, translation=value if value else "", pos=None)
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


def prepare_async():
    global _prepare_started
    with _lock:
        if _prepare_started:
            return
        _prepare_started = True

    def _run():
        global _prepare_started
        try:
            _ensure_translation()
        finally:
            with _lock:
                _prepare_started = False

    threading.Thread(target=_run, daemon=True).start()


def translate_text(text):
    translation = _ensure_translation()
    return _normalize_translation_text(translation.translate(str(text)))


def translate_words(words):
    result = get_cached_translations(words)
    missing = []
    seen_missing = set()
    for w in words:
        if w in result:
            continue
        key = _normalize_cache_key(w)
        if not key or key in seen_missing:
            continue
        seen_missing.add(key)
        missing.append(w)

    new_pairs = {}
    for w in missing:
        try:
            translated = translate_text(w)
        except Exception:
            translated = ""
        result[w] = translated
        if translated:
            new_pairs[w] = translated
    if new_pairs:
        _update_translation_cache(new_pairs)
    return result
