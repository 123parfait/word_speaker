# -*- coding: utf-8 -*-


class MainPlaybackController:
    def __init__(self):
        self.play_state = "stopped"
        self.queue = []
        self.pos = -1
        self.current_word = None

    def build_queue(self, words):
        return list(range(len(words or [])))

    def rebuild(self, *, words, selected_idx=None):
        if not words:
            self.queue = []
            self.pos = -1
            self.current_word = None
            return None
        self.queue = self.build_queue(words)
        if self.current_word in words:
            self.pos = words.index(self.current_word)
        else:
            self.pos = int(selected_idx) if selected_idx is not None else 0
        self.pos = max(0, min(self.pos, len(self.queue) - 1))
        self.current_word = words[self.queue[self.pos]]
        return self.current_word

    def start_or_resume(self, *, words, selected_idx=None):
        if not words:
            self.play_state = "stopped"
            self.queue = []
            self.pos = -1
            self.current_word = None
            return None
        if not self.queue or self.pos < 0:
            self.rebuild(words=words, selected_idx=selected_idx)
        self.play_state = "playing"
        return self.current_word

    def pause(self):
        self.play_state = "paused"

    def reset(self):
        self.play_state = "stopped"
        self.queue = []
        self.pos = -1
        self.current_word = None

    def advance(self, words):
        if not words:
            self.reset()
            return None
        if not self.queue:
            self.queue = self.build_queue(words)
            self.pos = 0
        else:
            self.pos += 1
            if self.pos >= len(self.queue):
                self.queue = self.build_queue(words)
                self.pos = 0
        self.current_word = words[self.queue[self.pos]]
        return self.current_word

    def set_current_by_selection(self, *, words, selected_idx=None):
        if not words:
            self.reset()
            return None
        self.queue = self.build_queue(words)
        self.pos = int(selected_idx) if selected_idx is not None else 0
        self.pos = max(0, min(self.pos, len(self.queue) - 1))
        self.current_word = words[self.queue[self.pos]]
        return self.current_word
