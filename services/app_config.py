# -*- coding: utf-8 -*-
import json
from pathlib import Path


_CONFIG_PATH = Path(__file__).resolve().parent.parent / "data" / "app_config.json"
_DEFAULT_CONFIG = {
    "gemini_api_key": "",
    "llm_api_provider": "gemini",
    "llm_api_key": "",
    "tts_api_provider": "gemini",
    "tts_api_key": "",
    "gemini_model": "gemini-2.5-flash",
    "ui_language": "zh",
    "update_manifest_url": "",
    "shared_cache_manifest_url": "",
}


def _normalize_tts_provider(provider):
    value = str(provider or "").strip().lower()
    if value == "elevenlabs":
        return "elevenlabs"
    return "gemini"


def load_config():
    if not _CONFIG_PATH.exists():
        return dict(_DEFAULT_CONFIG)
    try:
        with open(_CONFIG_PATH, "r", encoding="utf-8") as fp:
            data = json.load(fp)
    except Exception:
        return dict(_DEFAULT_CONFIG)
    if not isinstance(data, dict):
        return dict(_DEFAULT_CONFIG)
    merged = dict(_DEFAULT_CONFIG)
    merged.update(data)
    legacy_key = str(merged.get("gemini_api_key") or "").strip()
    if not str(merged.get("llm_api_key") or "").strip() and legacy_key:
        merged["llm_api_key"] = legacy_key
    if not str(merged.get("tts_api_key") or "").strip() and legacy_key:
        merged["tts_api_key"] = legacy_key
    merged["llm_api_provider"] = "gemini" if str(merged.get("llm_api_provider") or "").strip().lower() == "gemini" else "gemini"
    merged["tts_api_provider"] = _normalize_tts_provider(merged.get("tts_api_provider"))
    return merged


def save_config(config):
    merged = dict(_DEFAULT_CONFIG)
    if isinstance(config, dict):
        merged.update(config)
    _CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(_CONFIG_PATH, "w", encoding="utf-8") as fp:
        json.dump(merged, fp, ensure_ascii=False, indent=2)


def get_gemini_api_key():
    return get_llm_api_key()


def set_gemini_api_key(api_key):
    set_llm_api_key(api_key)


def get_llm_api_provider():
    return str(load_config().get("llm_api_provider") or "gemini").strip().lower() or "gemini"


def set_llm_api_provider(provider):
    config = load_config()
    config["llm_api_provider"] = "gemini" if str(provider or "").strip().lower() == "gemini" else "gemini"
    save_config(config)


def get_llm_api_key():
    config = load_config()
    return str(config.get("llm_api_key") or config.get("gemini_api_key") or "").strip()


def set_llm_api_key(api_key):
    config = load_config()
    value = str(api_key or "").strip()
    config["llm_api_key"] = value
    if not str(config.get("gemini_api_key") or "").strip():
        config["gemini_api_key"] = value
    save_config(config)


def get_tts_api_provider():
    return _normalize_tts_provider(load_config().get("tts_api_provider"))


def set_tts_api_provider(provider):
    config = load_config()
    config["tts_api_provider"] = _normalize_tts_provider(provider)
    save_config(config)


def get_tts_api_key():
    config = load_config()
    return str(config.get("tts_api_key") or config.get("gemini_api_key") or "").strip()


def set_tts_api_key(api_key):
    config = load_config()
    value = str(api_key or "").strip()
    config["tts_api_key"] = value
    if not str(config.get("gemini_api_key") or "").strip():
        config["gemini_api_key"] = value
    save_config(config)


def get_generation_model():
    return str(load_config().get("gemini_model") or _DEFAULT_CONFIG["gemini_model"]).strip()


def set_generation_model(model_name):
    config = load_config()
    config["gemini_model"] = str(model_name or _DEFAULT_CONFIG["gemini_model"]).strip()
    save_config(config)


def get_ui_language():
    language = str(load_config().get("ui_language") or _DEFAULT_CONFIG["ui_language"]).strip().lower()
    return "en" if language == "en" else "zh"


def set_ui_language(language):
    config = load_config()
    config["ui_language"] = "en" if str(language or "").strip().lower() == "en" else "zh"
    save_config(config)


def get_update_manifest_url():
    return str(load_config().get("update_manifest_url") or "").strip()


def set_update_manifest_url(url):
    config = load_config()
    config["update_manifest_url"] = str(url or "").strip()
    save_config(config)


def get_shared_cache_manifest_url():
    return str(load_config().get("shared_cache_manifest_url") or "").strip()


def set_shared_cache_manifest_url(url):
    config = load_config()
    config["shared_cache_manifest_url"] = str(url or "").strip()
    save_config(config)
