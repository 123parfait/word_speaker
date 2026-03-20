# -*- coding: utf-8 -*-
import json
import os
import shutil
import threading


def load_json_file(path, default):
    try:
        with open(path, "r", encoding="utf-8") as fp:
            data = json.load(fp)
        if isinstance(default, dict) and isinstance(data, dict):
            return data
        if isinstance(default, list) and isinstance(data, list):
            return data
    except Exception:
        pass
    return default


def write_json_file(path, payload):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as fp:
        json.dump(payload, fp, ensure_ascii=False, indent=2)


def cache_meta_path(cache_path):
    return f"{cache_path}.json"


def load_word_audio_overrides(path, *, allowed_backends):
    data = load_json_file(path, {})
    payload = {}
    if isinstance(data, dict):
        for key, value in data.items():
            backend = str(value or "").strip().lower()
            if backend in allowed_backends:
                payload[str(key or "").strip()] = backend
    return payload


def save_word_audio_overrides(path, data):
    payload = dict(data or {})
    target_dir = os.path.dirname(path)
    if target_dir:
        os.makedirs(target_dir, exist_ok=True)
    with open(path, "w", encoding="utf-8") as fp:
        json.dump(payload, fp, ensure_ascii=False, indent=2)


def load_pending_queue_disk_payload(new_path, legacy_path):
    data = load_json_file(new_path, [])
    if isinstance(data, list) and data:
        return data
    legacy = load_json_file(legacy_path, [])
    if isinstance(legacy, list) and legacy:
        return legacy
    return data if isinstance(data, list) else []


def migrate_pending_queue_path(new_path, legacy_path):
    legacy_exists = os.path.exists(legacy_path)
    new_exists = os.path.exists(new_path)
    if legacy_exists and not new_exists:
        try:
            os.makedirs(os.path.dirname(new_path), exist_ok=True)
            shutil.move(legacy_path, new_path)
            return
        except Exception:
            pass
    if legacy_exists:
        try:
            os.remove(legacy_path)
        except Exception:
            pass


class CacheMetadataStore:
    def __init__(self, *, canonicalize, normalize_source_path):
        self._canonicalize = canonicalize
        self._normalize_source_path = normalize_source_path
        self._memory = {}
        self._lock = threading.Lock()

    def load(self, cache_path):
        canonical_path = self._canonicalize(cache_path)
        key = str(canonical_path or "").strip()
        if not key:
            return {}
        with self._lock:
            cached = self._memory.get(key)
            if isinstance(cached, dict):
                return dict(cached)
        data = load_json_file(cache_meta_path(canonical_path), {})
        payload = data if isinstance(data, dict) else {}
        if payload:
            original_cache_path = str(payload.get("cache_path") or "").strip()
            source_value = self._normalize_source_path(payload.get("source_path"))
            if source_value != payload.get("source_path"):
                payload["source_path"] = source_value
            linked_shared = str(payload.get("linked_shared_path") or "").strip()
            if linked_shared:
                payload["linked_shared_path"] = self._canonicalize(linked_shared, metadata=payload)
            payload["cache_path"] = canonical_path
            if (
                original_cache_path != canonical_path
                or payload.get("source_path") != data.get("source_path")
                or payload.get("linked_shared_path") != data.get("linked_shared_path")
            ):
                write_json_file(cache_meta_path(canonical_path), payload)
        with self._lock:
            self._memory[key] = dict(payload)
        return payload

    def save(self, cache_path, metadata):
        cache_path = self._canonicalize(cache_path)
        payload = dict(metadata or {})
        payload["cache_path"] = cache_path
        write_json_file(cache_meta_path(cache_path), payload)
        with self._lock:
            self._memory[str(cache_path or "").strip()] = dict(payload)

    def remove(self, cache_path):
        cache_path = self._canonicalize(cache_path)
        key = str(cache_path or "").strip()
        try:
            meta_path = cache_meta_path(cache_path)
            if os.path.exists(meta_path):
                os.remove(meta_path)
        except Exception:
            pass
        with self._lock:
            self._memory.pop(key, None)
