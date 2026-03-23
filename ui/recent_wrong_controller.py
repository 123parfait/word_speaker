# -*- coding: utf-8 -*-
from dataclasses import dataclass

from services.phonetics import get_cached_phonetics, set_cached_phonetic
from services.translation import get_cached_translations, set_cached_translation
from services.word_analysis import get_cached_pos, set_cached_pos
from services.tts import (
    cleanup_word_audio_cache as tts_cleanup_word_audio_cache,
    queue_word_audio_generation as tts_queue_word_audio_generation,
)


@dataclass
class RecentWrongNoteResult:
    word: str
    changed: bool


@dataclass
class RecentWrongWordResult:
    old_word: str
    new_word: str
    changed: bool


@dataclass
class ManualWrongWordResult:
    word: str
    added_to_word_list: bool


class RecentWrongController:
    def __init__(self, store):
        self.store = store

    def update_note(self, word, raw_value):
        token = str(word or "").strip()
        if not token:
            raise ValueError("Word cannot be empty.")
        new_note = " ".join(str(raw_value or "").strip().split())
        old_note = str(self.store.get_dictation_word_stats(token).get("note") or "").strip()
        if new_note == old_note:
            return RecentWrongNoteResult(word=token, changed=False)
        self.store.set_recent_wrong_note(token, new_note)
        return RecentWrongNoteResult(word=token, changed=True)

    def update_word(self, old_word, raw_value, *, translations=None, word_pos=None):
        source_word = str(old_word or "").strip()
        new_word = str(raw_value or "").strip()
        if not source_word:
            raise ValueError("Source word cannot be empty.")
        if not new_word:
            raise ValueError("Word cannot be empty.")
        if new_word == source_word:
            return RecentWrongWordResult(old_word=source_word, new_word=new_word, changed=False)
        current_entry = self.store.get_dictation_word_stats(source_word)
        current_translation = str(
            (translations or {}).get(source_word) or get_cached_translations([source_word]).get(source_word) or ""
        ).strip()
        current_pos = str((word_pos or {}).get(source_word) or get_cached_pos([source_word]).get(source_word) or "").strip()
        current_phonetic = str(
            current_entry.get("phonetic")
            or get_cached_phonetics([source_word]).get(source_word)
            or ""
        ).strip()
        if current_translation:
            set_cached_translation(new_word, current_translation)
            if isinstance(translations, dict):
                translations[new_word] = current_translation
        if current_pos:
            set_cached_pos(new_word, current_pos)
            if isinstance(word_pos, dict):
                word_pos[new_word] = current_pos
        if current_phonetic:
            set_cached_phonetic(new_word, current_phonetic)
        if current_entry.get("note"):
            self.store.set_recent_wrong_note(new_word, current_entry.get("note"))
        self.store.rename_recent_wrong_word(source_word, new_word)
        return RecentWrongWordResult(old_word=source_word, new_word=new_word, changed=True)

    def add_manual_wrong_word(self, word, *, recent_wrong_source_path=None):
        token = str(word or "").strip()
        if not token:
            raise ValueError("Word cannot be empty.")
        added_to_word_list = False
        if token not in self.store.words:
            self.store.words.append(token)
            self.store.notes.append("")
            added_to_word_list = True
        self.store.add_wrong_word(token)
        tts_queue_word_audio_generation(token, source_path=recent_wrong_source_path)
        return ManualWrongWordResult(word=token, added_to_word_list=added_to_word_list)

    def clear_wrong_word(self, word, *, recent_wrong_source_path=None):
        token = str(word or "").strip()
        if not token:
            raise ValueError("Word cannot be empty.")
        self.store.clear_wrong_word(token)
        if recent_wrong_source_path:
            tts_cleanup_word_audio_cache(token, source_path=recent_wrong_source_path)
        return token
