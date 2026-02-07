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

        # Update word statistics, split sentences into words for counting
        stats = self.load_stats()
        for sentence in words:
            # Split each sentence into words for statistics
            sentence_words = self._split_into_words(sentence)
            for word in sentence_words:
                if word in stats:
                    stats[word] += 1
                else:
                    stats[word] = 1
        self.save_stats(stats)

        return words
    
    def _split_into_words(self, text):
        """
        Split text into individual words, handling multiple delimiters and removing punctuation
        """
        import re
        # Split on whitespace, commas, semicolons, and other common delimiters
        # Also remove punctuation and convert to lowercase
        words = re.findall(r'\b\w+\b', text.lower())
        # Filter out empty words
        return [word.strip() for word in words if word.strip()]

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

    def _sort_stats(self)->dict[int, dict[str, int]]:
        """
        rank the word frequency to show the unfamiliar words
        :return: a dict, key:value = ranking:word_statistics
        """
        stats = self.load_stats()
        stats_list = []
        for word, freq in stats.items(): stats_list.append((word, freq))
        # do a bubble sort with descending order
        stats_length = len(stats_list)
        order_flag = True
        for i in range(0, stats_length - 1):
            if i > 0 and order_flag == True: break
            for j in range(0, stats_length - i - 1):
                if stats_list[j][1] < stats_list[j + 1][1]:
                    stats_list[j], stats_list[j + 1] = stats_list[j + 1], stats_list[j]
                    order_flag = False
        result = {}
        for i in range(0, stats_length): result[i + 1] = stats_list[i]
        return result
