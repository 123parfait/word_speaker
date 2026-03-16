# -*- coding: utf-8 -*-
import csv
import json
import os
from difflib import SequenceMatcher
from datetime import datetime


DICTATION_META_KEY = "__meta__"


class WordStore:
    def __init__(self):
        self.words = []
        self.notes = []
        self.current_source_path = None
        self.history_path = os.path.join(os.path.dirname(__file__), "history.json")
        self.stats_path = os.path.join(os.path.dirname(__file__), "word_stats.json")
        self.dictation_stats_path = os.path.join(os.path.dirname(__file__), "dictation_stats.json")

    def clear(self):
        self.words = []
        self.notes = []
        self.current_source_path = None

    def set_words(self, words, notes=None, preserve_source=False):
        self.words = list(words)
        base_notes = list(notes or [])
        if len(base_notes) < len(self.words):
            base_notes.extend([""] * (len(self.words) - len(base_notes)))
        self.notes = base_notes[: len(self.words)]
        if not preserve_source:
            self.current_source_path = None

    def load_from_file(self, path):
        words = []
        notes = []
        lower_path = str(path or "").lower()
        if lower_path.endswith(".txt"):
            with open(path, "r", encoding="utf-8") as f:
                for line in f:
                    raw = str(line or "").rstrip("\r\n")
                    if "\t" in raw:
                        word, note = raw.split("\t", 1)
                    else:
                        word, note = raw, ""
                    word = word.strip()
                    if word:
                        words.append(word)
                        notes.append(str(note or "").strip())
        elif lower_path.endswith(".csv"):
            with open(path, "r", encoding="utf-8") as f:
                reader = csv.reader(f)
                for row in reader:
                    if row:
                        word = row[0].strip()
                        if word:
                            words.append(word)
                            note = row[1].strip() if len(row) > 1 else ""
                            notes.append(note)
        self.words = words
        self.notes = notes
        self.current_source_path = os.path.abspath(path)
        self.add_history(path)

        # Update word statistics, such as apple: 1, banana: 2
        # FIXME add word split to handle a phrase
        stats = self.load_stats()
        for word in words:
            if word in stats:
                stats[word] += 1
            else:
                stats[word] = 1
        self.save_stats(stats)

        return words

    def get_current_source_path(self):
        return self.current_source_path

    def has_current_source_file(self):
        return bool(self.current_source_path and os.path.exists(self.current_source_path))

    def detach_current_source(self):
        self.current_source_path = None

    def save_to_current_file(self):
        if not self.current_source_path:
            return False
        return self.save_to_file(self.current_source_path)

    def save_to_file(self, path):
        target = os.path.abspath(str(path or "").strip())
        if not target:
            return False

        if target.lower().endswith(".txt"):
            with open(target, "w", encoding="utf-8", newline="") as f:
                for idx, word in enumerate(self.words):
                    value = str(word or "").strip()
                    if value:
                        note = ""
                        if idx < len(self.notes):
                            note = str(self.notes[idx] or "").strip()
                        if note:
                            f.write(f"{value}\t{note}\n")
                        else:
                            f.write(value + "\n")
        elif target.lower().endswith(".csv"):
            with open(target, "w", encoding="utf-8", newline="") as f:
                writer = csv.writer(f)
                for idx, word in enumerate(self.words):
                    value = str(word or "").strip()
                    if value:
                        note = ""
                        if idx < len(self.notes):
                            note = str(self.notes[idx] or "").strip()
                        if note:
                            writer.writerow([value, note])
                        else:
                            writer.writerow([value])
        else:
            raise ValueError("Unsupported file type for save.")

        self.current_source_path = target
        return True

    def load_history(self):
        if not os.path.exists(self.history_path):
            return []
        try:
            with open(self.history_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list):
                return data
        except Exception:
            # Corrupted history file; back it up and reset
            try:
                bad_path = self.history_path + ".bad"
                if os.path.exists(bad_path):
                    os.remove(bad_path)
                os.rename(self.history_path, bad_path)
            except Exception:
                pass
            return []
        return []

    def load_stats(self)->dict[str, int]:
        if not os.path.exists(self.stats_path):
            return {}
        try:
            with open(self.stats_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                return data
        except Exception:
            pass
        return {}

    def save_stats(self, stats: dict[str, int]):
        try:
            with open(self.stats_path, "w", encoding="utf-8") as f:
                json.dump(stats, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def load_dictation_stats(self):
        if not os.path.exists(self.dictation_stats_path):
            return {}
        try:
            with open(self.dictation_stats_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                return data
        except Exception:
            pass
        return {}

    def save_dictation_stats(self, stats):
        try:
            with open(self.dictation_stats_path, "w", encoding="utf-8") as f:
                json.dump(stats or {}, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _analyze_spelling_error(self, word, user_input):
        target = str(word or "").strip().casefold()
        typed = str(user_input or "").strip().casefold()
        if not target or typed == target:
            return ""
        compact_target = target.replace(" ", "")
        compact_typed = typed.replace(" ", "")
        if compact_target == compact_typed:
            return "空格错误"
        if len(compact_typed) < len(compact_target):
            base = "缺字母"
        elif len(compact_typed) > len(compact_target):
            base = "多字母"
        else:
            base = "拼写错误"
        if sorted(compact_typed) == sorted(compact_target):
            return "顺序错误"
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
            return "发音型错误"
        similarity = SequenceMatcher(None, compact_typed, compact_target).ratio()
        if similarity >= 0.75 and base == "拼写错误":
            return "近似拼写"
        return base

    def record_dictation_result(self, word, user_input, correct):
        token = str(word or "").strip()
        if not token:
            return
        stats = self.load_dictation_stats()
        entry = stats.get(token)
        if not isinstance(entry, dict):
            entry = {
                "wrong_count": 0,
                "correct_count": 0,
                "last_wrong_input": "",
                "last_wrong_type": "",
                "last_result": "",
                "last_seen": "",
            }
        if correct:
            entry["correct_count"] = int(entry.get("correct_count", 0) or 0) + 1
            entry["last_result"] = "correct"
        else:
            entry["wrong_count"] = int(entry.get("wrong_count", 0) or 0) + 1
            entry["last_wrong_input"] = str(user_input or "").strip()
            entry["last_wrong_type"] = self._analyze_spelling_error(token, user_input)
            entry["last_result"] = "wrong"
        entry["last_seen"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        stats[token] = entry
        self.save_dictation_stats(stats)

    def get_dictation_word_stats(self, word):
        token = str(word or "").strip()
        if not token:
            return {
                "wrong_count": 0,
                "correct_count": 0,
                "last_wrong_input": "",
                "last_wrong_type": "",
                "last_result": "",
                "last_seen": "",
            }
        stats = self.load_dictation_stats()
        entry = stats.get(token)
        if not isinstance(entry, dict):
            return {
                "wrong_count": 0,
                "correct_count": 0,
                "last_wrong_input": "",
                "last_wrong_type": "",
                "last_result": "",
                "last_seen": "",
            }
        return {
            "wrong_count": int(entry.get("wrong_count", 0) or 0),
            "correct_count": int(entry.get("correct_count", 0) or 0),
            "last_wrong_input": str(entry.get("last_wrong_input") or ""),
            "last_wrong_type": str(entry.get("last_wrong_type") or ""),
            "last_result": str(entry.get("last_result") or ""),
            "last_seen": str(entry.get("last_seen") or ""),
        }

    def get_last_dictation_accuracy(self):
        stats = self.load_dictation_stats()
        meta = stats.get(DICTATION_META_KEY)
        if not isinstance(meta, dict):
            return None
        value = meta.get("last_session_accuracy")
        try:
            return float(value)
        except Exception:
            return None

    def save_last_dictation_accuracy(self, accuracy):
        stats = self.load_dictation_stats()
        meta = stats.get(DICTATION_META_KEY)
        if not isinstance(meta, dict):
            meta = {}
        try:
            meta["last_session_accuracy"] = float(accuracy)
        except Exception:
            meta["last_session_accuracy"] = 0.0
        meta["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        stats[DICTATION_META_KEY] = meta
        self.save_dictation_stats(stats)

    def recent_wrong_words(self, words=None, limit=50):
        stats = self.load_dictation_stats()
        allowed = {str(word or "").strip() for word in (words or []) if str(word or "").strip()}
        items = []
        for word, entry in stats.items():
            if str(word).startswith("__"):
                continue
            if allowed and word not in allowed:
                continue
            if not isinstance(entry, dict):
                continue
            wrong_count = int(entry.get("wrong_count", 0) or 0)
            if wrong_count <= 0:
                continue
            items.append(
                {
                    "word": word,
                    "wrong_count": wrong_count,
                    "correct_count": int(entry.get("correct_count", 0) or 0),
                    "last_wrong_input": str(entry.get("last_wrong_input") or ""),
                    "last_wrong_type": str(entry.get("last_wrong_type") or ""),
                    "last_result": str(entry.get("last_result") or ""),
                    "last_seen": str(entry.get("last_seen") or ""),
                }
            )
        items.sort(key=lambda item: (item.get("wrong_count", 0), item.get("last_seen", "")), reverse=True)
        return items[: max(1, int(limit or 50))]

    def clear_wrong_word(self, word):
        token = str(word or "").strip()
        if not token:
            return False
        stats = self.load_dictation_stats()
        entry = stats.get(token)
        if not isinstance(entry, dict):
            return False
        entry["wrong_count"] = 0
        entry["last_wrong_input"] = ""
        entry["last_wrong_type"] = ""
        entry["last_result"] = "correct"
        entry["last_seen"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        stats[token] = entry
        self.save_dictation_stats(stats)
        return True

    def add_history(self, path):
        history = self.load_history()
        path = os.path.abspath(path)
        name = os.path.basename(path)
        now = datetime.now().strftime("%Y-%m-%d %H:%M")

        # remove existing same path
        history = [h for h in history if h.get("path") != path]
        history.insert(0, {"path": path, "name": name, "time": now})

        try:
            with open(self.history_path, "w", encoding="utf-8") as f:
                json.dump(history, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

        return history

    def remove_history(self, path):
        target = os.path.abspath(str(path or "").strip())
        if not target:
            return []
        history = [h for h in self.load_history() if str(h.get("path") or "").strip() != target]
        try:
            with open(self.history_path, "w", encoding="utf-8") as f:
                json.dump(history, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
        return history

    def rename_history_path(self, old_path, new_path):
        old_target = os.path.abspath(str(old_path or "").strip())
        new_target = os.path.abspath(str(new_path or "").strip())
        if not old_target or not new_target:
            return []
        history = self.load_history()
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        for item in history:
            path = os.path.abspath(str(item.get("path") or "").strip()) if str(item.get("path") or "").strip() else ""
            if path != old_target:
                continue
            item["path"] = new_target
            item["name"] = os.path.basename(new_target)
            item["time"] = now
            break
        try:
            with open(self.history_path, "w", encoding="utf-8") as f:
                json.dump(history, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
        if self.current_source_path and os.path.abspath(self.current_source_path) == old_target:
            self.current_source_path = new_target
        return history

    def add_wrong_word(self, word, user_input=""):
        token = str(word or "").strip()
        if not token:
            return False
        stats = self.load_dictation_stats()
        entry = stats.get(token)
        if not isinstance(entry, dict):
            entry = {
                "wrong_count": 0,
                "correct_count": 0,
                "last_wrong_input": "",
                "last_wrong_type": "",
                "last_result": "",
                "last_seen": "",
            }
        entry["wrong_count"] = int(entry.get("wrong_count", 0) or 0) + 1
        entry["last_wrong_input"] = str(user_input or "").strip()
        entry["last_wrong_type"] = self._analyze_spelling_error(token, user_input) if user_input else "手动加入"
        entry["last_result"] = "wrong"
        entry["last_seen"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        stats[token] = entry
        self.save_dictation_stats(stats)
        return True
