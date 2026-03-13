# -*- coding: utf-8 -*-
import json
from pathlib import Path


_CONFIG_PATH = Path(__file__).resolve().parent.parent / "data" / "app_config.json"
_DEFAULT_CONFIG = {
    "gemini_api_key": "",
    "gemini_model": "gemini-2.5-flash",
}


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
    return merged


def save_config(config):
    merged = dict(_DEFAULT_CONFIG)
    if isinstance(config, dict):
        merged.update(config)
    _CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(_CONFIG_PATH, "w", encoding="utf-8") as fp:
        json.dump(merged, fp, ensure_ascii=False, indent=2)


def get_gemini_api_key():
    return str(load_config().get("gemini_api_key") or "").strip()


def set_gemini_api_key(api_key):
    config = load_config()
    config["gemini_api_key"] = str(api_key or "").strip()
    save_config(config)


def get_generation_model():
    return str(load_config().get("gemini_model") or _DEFAULT_CONFIG["gemini_model"]).strip()


def set_generation_model(model_name):
    config = load_config()
    config["gemini_model"] = str(model_name or _DEFAULT_CONFIG["gemini_model"]).strip()
    save_config(config)
