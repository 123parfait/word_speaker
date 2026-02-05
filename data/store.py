# -*- coding: utf-8 -*-
import csv
import json
import os
from datetime import datetime


class WordStore:
    def __init__(self):
        self.words = []
        self.history_path = os.path.join(os.path.dirname(__file__), "history.json")
        self.stats_path = os.path.join(os.path.dirname(__file__), "word_stats.json")

    def clear(self):
        self.words = []

    def set_words(self, words):
        self.words = list(words)

    def load_from_file(self, path):
        words = []
        if path.endswith(".txt"):
            with open(path, "r", encoding="utf-8") as f:
                for line in f:
                    word = line.strip()
                    if word:
                        words.append(word)
        elif path.endswith(".csv"):
            with open(path, "r", encoding="utf-8") as f:
                reader = csv.reader(f)
                for row in reader:
                    if row:
                        word = row[0].strip()
                        if word:
                            words.append(word)
        self.words = words
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
