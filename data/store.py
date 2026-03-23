# -*- coding: utf-8 -*-
import csv
import json
import os
from datetime import datetime
from difflib import SequenceMatcher
from typing import Any


DICTATION_META_KEY = "__meta__"

WRONG_TYPE_SPACE = "空格错误"
WRONG_TYPE_MISSING = "缺字母"
WRONG_TYPE_EXTRA = "多字母"
WRONG_TYPE_SPELLING = "拼写错误"
WRONG_TYPE_ORDER = "顺序错误"
WRONG_TYPE_PHONETIC = "发音型错误"
WRONG_TYPE_SIMILAR = "近似拼写"
WRONG_TYPE_MANUAL = "手动加入"


def _now_minute_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def _now_second_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _abs_path(path: str) -> str:
    return os.path.abspath(str(path or "").strip())


def _clean_token(value: Any) -> str:
    return str(value or "").strip()


def _default_dictation_entry() -> dict[str, Any]:
    return {
        "wrong_count": 0,
        "correct_count": 0,
        "last_wrong_input": "",
        "last_wrong_type": "",
        "last_result": "",
        "last_seen": "",
        "note": "",
        "phonetic": "",
    }


class JsonFileRepository:
    def __init__(self, path: str, default_factory, backup_on_error: bool = False):
        self.path = path
        self.default_factory = default_factory
        self.backup_on_error = backup_on_error

    def load(self):
        default_value = self.default_factory()
        if not os.path.exists(self.path):
            return default_value
        try:
            with open(self.path, "r", encoding="utf-8") as fp:
                data = json.load(fp)
            if isinstance(default_value, dict) and isinstance(data, dict):
                return data
            if isinstance(default_value, list) and isinstance(data, list):
                return data
        except Exception:
            if self.backup_on_error:
                self._backup_corrupted_file()
            return default_value
        return default_value

    def save(self, payload) -> None:
        try:
            with open(self.path, "w", encoding="utf-8") as fp:
                json.dump(payload, fp, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _backup_corrupted_file(self) -> None:
        try:
            bad_path = self.path + ".bad"
            if os.path.exists(bad_path):
                os.remove(bad_path)
            os.rename(self.path, bad_path)
        except Exception:
            pass


class WordListFileStore:
    @staticmethod
    def load(path: str) -> tuple[list[str], list[str]]:
        words: list[str] = []
        notes: list[str] = []
        lower_path = str(path or "").lower()
        if lower_path.endswith(".txt"):
            with open(path, "r", encoding="utf-8") as fp:
                for line in fp:
                    raw = str(line or "").rstrip("\r\n")
                    if "\t" in raw:
                        word, note = raw.split("\t", 1)
                    else:
                        word, note = raw, ""
                    token = _clean_token(word)
                    if token:
                        words.append(token)
                        notes.append(_clean_token(note))
        elif lower_path.endswith(".csv"):
            with open(path, "r", encoding="utf-8") as fp:
                reader = csv.reader(fp)
                for row in reader:
                    if not row:
                        continue
                    token = _clean_token(row[0])
                    if token:
                        words.append(token)
                        notes.append(_clean_token(row[1]) if len(row) > 1 else "")
        return words, notes

    @staticmethod
    def save(path: str, words: list[str], notes: list[str]) -> bool:
        target = _abs_path(path)
        if not target:
            return False
        if target.lower().endswith(".txt"):
            with open(target, "w", encoding="utf-8", newline="") as fp:
                for idx, word in enumerate(words):
                    token = _clean_token(word)
                    if not token:
                        continue
                    note = _clean_token(notes[idx]) if idx < len(notes) else ""
                    fp.write(f"{token}\t{note}\n" if note else f"{token}\n")
            return True
        if target.lower().endswith(".csv"):
            with open(target, "w", encoding="utf-8", newline="") as fp:
                writer = csv.writer(fp)
                for idx, word in enumerate(words):
                    token = _clean_token(word)
                    if not token:
                        continue
                    note = _clean_token(notes[idx]) if idx < len(notes) else ""
                    writer.writerow([token, note] if note else [token])
            return True
        raise ValueError("Unsupported file type for save.")


class HistoryRepository:
    def __init__(self, path: str):
        self._repo = JsonFileRepository(path, default_factory=list, backup_on_error=True)

    def load(self) -> list[dict[str, Any]]:
        return self._repo.load()

    def add(self, path: str) -> list[dict[str, Any]]:
        target = _abs_path(path)
        if not target:
            return self.load()
        history = [item for item in self.load() if _clean_token(item.get("path")) != target]
        history.insert(
            0,
            {
                "path": target,
                "name": os.path.basename(target),
                "time": _now_minute_text(),
            },
        )
        self._repo.save(history)
        return history

    def remove(self, path: str) -> list[dict[str, Any]]:
        target = _abs_path(path)
        if not target:
            return self.load()
        history = [item for item in self.load() if _clean_token(item.get("path")) != target]
        self._repo.save(history)
        return history

    def rename_path(self, old_path: str, new_path: str) -> list[dict[str, Any]]:
        old_target = _abs_path(old_path)
        new_target = _abs_path(new_path)
        if not old_target or not new_target:
            return self.load()
        history = self.load()
        for item in history:
            if _abs_path(item.get("path")) != old_target:
                continue
            item["path"] = new_target
            item["name"] = os.path.basename(new_target)
            item["time"] = _now_minute_text()
            break
        self._repo.save(history)
        return history


class WordStatsRepository:
    def __init__(self, path: str):
        self._repo = JsonFileRepository(path, default_factory=dict)

    def load(self) -> dict[str, int]:
        data = self._repo.load()
        return {str(key): int(value or 0) for key, value in data.items()}

    def save(self, stats: dict[str, int]) -> None:
        payload = {str(key): int(value or 0) for key, value in (stats or {}).items()}
        self._repo.save(payload)

    def record_loaded_words(self, words: list[str]) -> None:
        stats = self.load()
        for word in words:
            token = _clean_token(word)
            if not token:
                continue
            stats[token] = int(stats.get(token, 0) or 0) + 1
        self.save(stats)


class DictationStatsRepository:
    def __init__(self, path: str):
        self._repo = JsonFileRepository(path, default_factory=dict)

    def load(self) -> dict[str, Any]:
        data = self._repo.load()
        return data if isinstance(data, dict) else {}

    def save(self, stats: dict[str, Any]) -> None:
        self._repo.save(stats or {})

    def get_entry(self, word: str) -> dict[str, Any]:
        token = _clean_token(word)
        if not token:
            return _default_dictation_entry()
        entry = self.load().get(token)
        if not isinstance(entry, dict):
            return _default_dictation_entry()
        payload = _default_dictation_entry()
        payload.update(
            {
                "wrong_count": int(entry.get("wrong_count", 0) or 0),
                "correct_count": int(entry.get("correct_count", 0) or 0),
                "last_wrong_input": _clean_token(entry.get("last_wrong_input")),
                "last_wrong_type": _clean_token(entry.get("last_wrong_type")),
                "last_result": _clean_token(entry.get("last_result")),
                "last_seen": _clean_token(entry.get("last_seen")),
                "note": _clean_token(entry.get("note")),
                "phonetic": _clean_token(entry.get("phonetic")),
            }
        )
        return payload

    def snapshot_entry(self, word: str) -> dict[str, Any] | None:
        token = _clean_token(word)
        if not token:
            return None
        entry = self.load().get(token)
        return dict(entry) if isinstance(entry, dict) else None

    def restore_entry(self, word: str, snapshot: dict[str, Any] | None) -> bool:
        token = _clean_token(word)
        if not token:
            return False
        stats = self.load()
        if isinstance(snapshot, dict):
            stats[token] = dict(snapshot)
        else:
            stats.pop(token, None)
        self.save(stats)
        return True

    def save_result(
        self,
        word: str,
        user_input: str,
        correct: bool,
        wrong_type: str,
        note: str = "",
        phonetic: str = "",
    ) -> None:
        token = _clean_token(word)
        if not token:
            return
        stats = self.load()
        entry = self.get_entry(token)
        if correct:
            entry["correct_count"] = int(entry.get("correct_count", 0) or 0) + 1
            entry["last_result"] = "correct"
        else:
            entry["wrong_count"] = int(entry.get("wrong_count", 0) or 0) + 1
            entry["last_wrong_input"] = _clean_token(user_input)
            entry["last_wrong_type"] = wrong_type
            entry["last_result"] = "wrong"
        if _clean_token(note):
            entry["note"] = _clean_token(note)
        if _clean_token(phonetic):
            entry["phonetic"] = _clean_token(phonetic)
        entry["last_seen"] = _now_second_text()
        stats[token] = entry
        self.save(stats)

    def add_wrong_word(
        self,
        word: str,
        user_input: str,
        wrong_type: str,
        note: str = "",
        phonetic: str = "",
    ) -> bool:
        token = _clean_token(word)
        if not token:
            return False
        stats = self.load()
        entry = self.get_entry(token)
        entry["wrong_count"] = int(entry.get("wrong_count", 0) or 0) + 1
        entry["last_wrong_input"] = _clean_token(user_input)
        entry["last_wrong_type"] = wrong_type
        entry["last_result"] = "wrong"
        if _clean_token(note):
            entry["note"] = _clean_token(note)
        if _clean_token(phonetic):
            entry["phonetic"] = _clean_token(phonetic)
        entry["last_seen"] = _now_second_text()
        stats[token] = entry
        self.save(stats)
        return True

    def clear_wrong_word(self, word: str) -> bool:
        token = _clean_token(word)
        if not token:
            return False
        stats = self.load()
        entry = stats.get(token)
        if not isinstance(entry, dict):
            return False
        payload = self.get_entry(token)
        payload["wrong_count"] = 0
        payload["last_wrong_input"] = ""
        payload["last_wrong_type"] = ""
        payload["last_result"] = "correct"
        payload["last_seen"] = _now_second_text()
        stats[token] = payload
        self.save(stats)
        return True

    def set_note(self, word: str, note: str) -> bool:
        token = _clean_token(word)
        if not token:
            return False
        stats = self.load()
        entry = self.get_entry(token)
        entry["note"] = _clean_token(note)
        entry["last_seen"] = _now_second_text()
        stats[token] = entry
        self.save(stats)
        return True

    def rename_word(self, old_word: str, new_word: str) -> bool:
        old_token = _clean_token(old_word)
        new_token = _clean_token(new_word)
        if not old_token or not new_token:
            return False
        if old_token == new_token:
            return True
        stats = self.load()
        source_entry = stats.get(old_token)
        if not isinstance(source_entry, dict):
            return False
        target_entry = self.get_entry(new_token)
        source_payload = self.get_entry(old_token)
        merged = dict(target_entry)
        merged["wrong_count"] = int(target_entry.get("wrong_count", 0) or 0) + int(source_payload.get("wrong_count", 0) or 0)
        merged["correct_count"] = int(target_entry.get("correct_count", 0) or 0) + int(source_payload.get("correct_count", 0) or 0)
        for field in ("last_wrong_input", "last_wrong_type", "last_result", "note", "phonetic"):
            merged[field] = _clean_token(source_payload.get(field)) or _clean_token(target_entry.get(field))
        merged["last_seen"] = _now_second_text()
        stats[new_token] = merged
        stats.pop(old_token, None)
        self.save(stats)
        return True

    def recent_wrong_words(self, words=None, limit: int = 50) -> list[dict[str, Any]]:
        stats = self.load()
        allowed = {_clean_token(word) for word in (words or []) if _clean_token(word)}
        items = []
        for word, entry in stats.items():
            if str(word).startswith("__") or not isinstance(entry, dict):
                continue
            if allowed and word not in allowed:
                continue
            payload = self.get_entry(word)
            if int(payload.get("wrong_count", 0) or 0) <= 0:
                continue
            items.append({"word": word, **payload})
        items.sort(key=lambda item: (item.get("wrong_count", 0), item.get("last_seen", "")), reverse=True)
        return items[: max(1, int(limit or 50))]

    def get_last_accuracy(self) -> float | None:
        meta = self.load().get(DICTATION_META_KEY)
        if not isinstance(meta, dict):
            return None
        try:
            return float(meta.get("last_session_accuracy"))
        except Exception:
            return None

    def save_last_accuracy(self, accuracy: Any) -> None:
        stats = self.load()
        meta = stats.get(DICTATION_META_KEY)
        if not isinstance(meta, dict):
            meta = {}
        try:
            meta["last_session_accuracy"] = float(accuracy)
        except Exception:
            meta["last_session_accuracy"] = 0.0
        meta["updated_at"] = _now_second_text()
        stats[DICTATION_META_KEY] = meta
        self.save(stats)


class WordStore:
    def __init__(self):
        self.words: list[str] = []
        self.notes: list[str] = []
        self.current_source_path: str | None = None
        self.temp_source_active = False

        base_dir = os.path.dirname(__file__)
        self.history_path = os.path.join(base_dir, "history.json")
        self.stats_path = os.path.join(base_dir, "word_stats.json")
        self.dictation_stats_path = os.path.join(base_dir, "dictation_stats.json")
        self.temp_source_dir = os.path.join(base_dir, "temp_word_lists")
        self.temp_source_path = os.path.join(self.temp_source_dir, "current_manual_list.csv")

        self.history_repo = HistoryRepository(self.history_path)
        self.word_stats_repo = WordStatsRepository(self.stats_path)
        self.dictation_repo = DictationStatsRepository(self.dictation_stats_path)

    def clear(self):
        self.words = []
        self.notes = []
        self.current_source_path = None
        self.temp_source_active = False

    def set_words(self, words, notes=None, preserve_source=False):
        self.words = list(words or [])
        base_notes = list(notes or [])
        if len(base_notes) < len(self.words):
            base_notes.extend([""] * (len(self.words) - len(base_notes)))
        self.notes = base_notes[: len(self.words)]
        if not preserve_source:
            self.current_source_path = None
            self.temp_source_active = False

    def load_from_file(self, path):
        words, notes = WordListFileStore.load(path)
        self.words = words
        self.notes = notes
        self.current_source_path = _abs_path(path)
        self.temp_source_active = False
        self.history_repo.add(path)
        self.word_stats_repo.record_loaded_words(words)
        return words

    def get_current_source_path(self):
        return self.current_source_path

    def get_display_source_path(self):
        if self.temp_source_active:
            return None
        return self.current_source_path

    def has_current_source_file(self):
        return bool(
            self.current_source_path
            and not self.temp_source_active
            and os.path.exists(self.current_source_path)
        )

    def has_bound_source_path(self):
        return bool(self.current_source_path and os.path.exists(self.current_source_path))

    def is_temp_source_active(self):
        return bool(self.temp_source_active)

    def detach_current_source(self):
        self.current_source_path = None
        self.temp_source_active = False

    def ensure_temp_source_binding(self):
        os.makedirs(self.temp_source_dir, exist_ok=True)
        WordListFileStore.save(self.temp_source_path, self.words, self.notes)
        self.current_source_path = _abs_path(self.temp_source_path)
        self.temp_source_active = True
        return self.current_source_path

    def clear_temp_source_file(self):
        temp_path = _abs_path(self.temp_source_path)
        try:
            if os.path.exists(temp_path):
                os.remove(temp_path)
        except Exception:
            pass
        try:
            if os.path.isdir(self.temp_source_dir) and not os.listdir(self.temp_source_dir):
                os.rmdir(self.temp_source_dir)
        except Exception:
            pass
        if self.current_source_path and _abs_path(self.current_source_path) == temp_path:
            self.current_source_path = None
        self.temp_source_active = False

    def save_to_current_file(self):
        if not self.current_source_path:
            return False
        return self.save_to_file(self.current_source_path)

    def save_to_file(self, path):
        saved = WordListFileStore.save(path, self.words, self.notes)
        if saved:
            normalized_path = _abs_path(path)
            self.current_source_path = normalized_path
            self.temp_source_active = normalized_path == _abs_path(self.temp_source_path)
        return saved

    def load_history(self):
        return self.history_repo.load()

    def load_stats(self) -> dict[str, int]:
        return self.word_stats_repo.load()

    def save_stats(self, stats: dict[str, int]):
        self.word_stats_repo.save(stats)

    def load_dictation_stats(self):
        return self.dictation_repo.load()

    def save_dictation_stats(self, stats):
        self.dictation_repo.save(stats)

    def _analyze_spelling_error(self, word, user_input):
        target = _clean_token(word).casefold()
        typed = _clean_token(user_input).casefold()
        if not target or typed == target:
            return ""
        compact_target = target.replace(" ", "")
        compact_typed = typed.replace(" ", "")
        if compact_target == compact_typed:
            return WRONG_TYPE_SPACE
        if len(compact_typed) < len(compact_target):
            base = WRONG_TYPE_MISSING
        elif len(compact_typed) > len(compact_target):
            base = WRONG_TYPE_EXTRA
        else:
            base = WRONG_TYPE_SPELLING
        if sorted(compact_typed) == sorted(compact_target):
            return WRONG_TYPE_ORDER
        phonetic_target = (
            compact_target.replace("ph", "f")
            .replace("tion", "shun")
            .replace("sion", "zhun")
            .replace("ck", "k")
            .replace("c", "k")
            .replace("q", "k")
        )
        phonetic_typed = (
            compact_typed.replace("ph", "f")
            .replace("tion", "shun")
            .replace("sion", "zhun")
            .replace("ck", "k")
            .replace("c", "k")
            .replace("q", "k")
        )
        if phonetic_target == phonetic_typed:
            return WRONG_TYPE_PHONETIC
        similarity = SequenceMatcher(None, compact_typed, compact_target).ratio()
        if similarity >= 0.75 and base == WRONG_TYPE_SPELLING:
            return WRONG_TYPE_SIMILAR
        return base

    def _get_note_for_word(self, word):
        token = _clean_token(word)
        if not token:
            return ""
        try:
            index = self.words.index(token)
        except ValueError:
            return ""
        if 0 <= index < len(self.notes):
            return _clean_token(self.notes[index])
        return ""

    def _get_phonetic_for_word(self, word):
        token = _clean_token(word)
        if not token:
            return ""
        try:
            from services.phonetics import get_cached_phonetics

            return _clean_token(get_cached_phonetics([token]).get(token))
        except Exception:
            return ""

    def record_dictation_result(self, word, user_input, correct):
        wrong_type = "" if correct else self._analyze_spelling_error(word, user_input)
        self.dictation_repo.save_result(
            word,
            user_input,
            bool(correct),
            wrong_type,
            note=self._get_note_for_word(word),
            phonetic=self._get_phonetic_for_word(word),
        )

    def snapshot_dictation_word_stats(self, word):
        return self.dictation_repo.snapshot_entry(word)

    def restore_dictation_word_stats(self, word, snapshot):
        return self.dictation_repo.restore_entry(word, snapshot)

    def get_dictation_word_stats(self, word):
        return self.dictation_repo.get_entry(word)

    def get_last_dictation_accuracy(self):
        return self.dictation_repo.get_last_accuracy()

    def save_last_dictation_accuracy(self, accuracy):
        self.dictation_repo.save_last_accuracy(accuracy)

    def recent_wrong_words(self, words=None, limit=50):
        return self.dictation_repo.recent_wrong_words(words=words, limit=limit)

    def clear_wrong_word(self, word):
        return self.dictation_repo.clear_wrong_word(word)

    def set_recent_wrong_note(self, word, note):
        return self.dictation_repo.set_note(word, note)

    def rename_recent_wrong_word(self, old_word, new_word):
        return self.dictation_repo.rename_word(old_word, new_word)

    def add_history(self, path):
        return self.history_repo.add(path)

    def remove_history(self, path):
        return self.history_repo.remove(path)

    def rename_history_path(self, old_path, new_path):
        history = self.history_repo.rename_path(old_path, new_path)
        if self.current_source_path and _abs_path(self.current_source_path) == _abs_path(old_path):
            self.current_source_path = _abs_path(new_path)
        return history

    def add_wrong_word(self, word, user_input=""):
        wrong_type = self._analyze_spelling_error(word, user_input) if user_input else WRONG_TYPE_MANUAL
        return self.dictation_repo.add_wrong_word(
            word,
            user_input,
            wrong_type,
            note=self._get_note_for_word(word),
            phonetic=self._get_phonetic_for_word(word),
        )
