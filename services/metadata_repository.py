# -*- coding: utf-8 -*-
import json
import threading
from pathlib import Path


class JsonMetadataRepository:
    def __init__(self, path, *, key_normalizer, value_normalizer):
        self.path = Path(path)
        self._key_normalizer = key_normalizer
        self._value_normalizer = value_normalizer
        self._lock = threading.RLock()
        self._cache_data = None

    def _load_locked(self):
        if self._cache_data is not None:
            return self._cache_data
        if not self.path.exists():
            self._cache_data = {}
            return self._cache_data
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
            self._cache_data = data if isinstance(data, dict) else {}
        except Exception:
            self._cache_data = {}
        return self._cache_data

    def _save_locked(self):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(json.dumps(self._cache_data or {}, ensure_ascii=False, indent=2), encoding="utf-8")

    def normalize_pairs(self, pairs):
        normalized = {}
        for word, value in dict(pairs or {}).items():
            key = self._key_normalizer(word)
            text = self._value_normalizer(value)
            if key and text:
                normalized[key] = text
        return normalized

    def cleanup(self):
        changed = False
        with self._lock:
            cache = self._load_locked()
            cleaned = {}
            for key, value in list(cache.items()):
                clean_key = self._key_normalizer(key)
                clean_value = self._value_normalizer(value)
                if not clean_key or not clean_value:
                    changed = True
                    continue
                if cleaned.get(clean_key) == clean_value:
                    changed = True
                    continue
                cleaned[clean_key] = clean_value
                if key != clean_key or str(value or "") != clean_value:
                    changed = True
            self._cache_data = cleaned
            if changed:
                self._save_locked()
        return dict(self._cache_data or {})

    def get_many(self, words):
        result = {}
        with self._lock:
            cache = self.cleanup()
            for word in words or []:
                key = self._key_normalizer(word)
                if key and key in cache:
                    result[word] = self._value_normalizer(cache.get(key) or "")
        return result

    def apply_many(self, pairs):
        normalized = self.normalize_pairs(pairs)
        if not normalized:
            return 0
        changed = False
        with self._lock:
            cache = self._load_locked()
            for key, value in normalized.items():
                if cache.get(key) == value:
                    continue
                cache[key] = value
                changed = True
            if changed:
                self._save_locked()
        return len(normalized)

    def set_one(self, word, value):
        key = self._key_normalizer(word)
        if not key:
            return False
        text = self._value_normalizer(value)
        changed = False
        with self._lock:
            cache = self._load_locked()
            if text:
                if cache.get(key) != text:
                    cache[key] = text
                    changed = True
            elif key in cache:
                cache.pop(key, None)
                changed = True
            if changed:
                self._save_locked()
        return True

    def export_payload(self):
        with self._lock:
            return dict(self.cleanup())
