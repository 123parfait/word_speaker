# -*- coding: utf-8 -*-
import os
from dataclasses import dataclass

from services.tts import (
    cleanup_cache_for_source_path as tts_cleanup_cache_for_source_path,
    rebind_manual_session_cache_to_source as tts_rebind_manual_session_cache_to_source,
    rename_cache_source_path as tts_rename_cache_source_path,
    set_preferred_pending_source as tts_set_preferred_pending_source,
)


@dataclass
class SaveWordsResult:
    path: str
    history_updated: bool = True


@dataclass
class DeleteWordResult:
    word: str
    saved_to_source: bool


@dataclass
class AddWordResult:
    word: str
    note: str
    index: int
    saved_to_source: bool


@dataclass
class UpdateNoteResult:
    word: str
    note: str
    saved_to_source: bool
    changed: bool


@dataclass
class UpdateWordResult:
    old_word: str
    new_word: str
    note: str
    saved_to_source: bool
    changed: bool


@dataclass
class ApplyManualWordsResult:
    words: list[str]
    notes: list[str]
    status_text: str
    saved_to_source: bool
    source_bound: bool


@dataclass
class DeleteHistoryResult:
    name: str
    removed_count: int
    detached_current_source: bool


@dataclass
class RenameHistoryResult:
    path: str
    migrated: int
    queued: int


class WordListController:
    def __init__(self, store):
        self.store = store

    def _update_word_stats(self, words):
        try:
            stats = self.store.load_stats()
            for word in words:
                stats[word] = int(stats.get(word, 0)) + 1
            self.store.save_stats(stats)
        except Exception:
            return

    def create_blank_list(self):
        self.store.clear()

    def load_words(self, path):
        return self.store.load_from_file(path)

    def open_history_path(self, path):
        return self.store.load_from_file(path)

    def save_words_as(self, path, *, words_snapshot=None, was_manual_session=False):
        self.store.save_to_file(path)
        self.store.add_history(path)
        if was_manual_session and words_snapshot:
            tts_rebind_manual_session_cache_to_source(words_snapshot, path)
        tts_set_preferred_pending_source(path)
        return SaveWordsResult(path=path)

    def save_back_to_source(self):
        if not self.store.has_current_source_file():
            return False
        self.store.save_to_current_file()
        return True

    def apply_manual_words(self, words, notes, *, mode="replace", ui_language="zh"):
        normalized_words = list(words or [])
        normalized_notes = list(notes or [])
        keep_source_binding = self.store.has_current_source_file()
        if mode == "append":
            merged_words = list(self.store.words) + normalized_words
            merged_notes = list(self.store.notes) + normalized_notes
            self.store.set_words(merged_words, merged_notes, preserve_source=keep_source_binding)
            self._update_word_stats(normalized_words)
            status_text = f"Appended {len(normalized_words)} words."
            result_words = merged_words
            result_notes = merged_notes
        else:
            self.store.set_words(normalized_words, normalized_notes, preserve_source=keep_source_binding)
            self._update_word_stats(normalized_words)
            status_text = f"Loaded {len(normalized_words)} words."
            result_words = normalized_words
            result_notes = normalized_notes

        saved_to_source = False
        if keep_source_binding:
            saved_to_source = self.save_back_to_source()
            if saved_to_source:
                suffix = " Saved to source file." if ui_language == "en" else " 已保存到源文件。"
            else:
                suffix = (
                    " Save to source file failed; changes are kept in memory."
                    if ui_language == "en"
                    else " 保存源文件失败，修改暂时只保留在内存中。"
                )
            status_text += suffix

        return ApplyManualWordsResult(
            words=result_words,
            notes=result_notes,
            status_text=status_text,
            saved_to_source=saved_to_source,
            source_bound=keep_source_binding,
        )

    def delete_word(self, index, translations=None, word_pos=None):
        if index is None or index < 0 or index >= len(self.store.words):
            raise IndexError("word index out of range")
        word = self.store.words.pop(index)
        if index < len(self.store.notes):
            self.store.notes.pop(index)
        if isinstance(translations, dict):
            translations.pop(word, None)
        if isinstance(word_pos, dict):
            word_pos.pop(word, None)
        saved_to_source = self.save_back_to_source()
        return DeleteWordResult(word=word, saved_to_source=saved_to_source)

    def add_word(self, raw_word, raw_note=""):
        word = " ".join(str(raw_word or "").strip().split())
        note = " ".join(str(raw_note or "").strip().split())
        if not word:
            raise ValueError("Word cannot be empty.")
        self.store.words.append(word)
        self.store.notes.append(note)
        self._update_word_stats([word])
        saved_to_source = self.save_back_to_source() if self.store.has_current_source_file() else False
        return AddWordResult(
            word=word,
            note=note,
            index=len(self.store.words) - 1,
            saved_to_source=saved_to_source,
        )

    def update_note(self, index, raw_value):
        if index is None or index < 0 or index >= len(self.store.words):
            raise IndexError("word index out of range")
        new_note = " ".join(str(raw_value or "").strip().split())
        old_note = self.store.notes[index] if index < len(self.store.notes) else ""
        if index >= len(self.store.notes):
            self.store.notes.extend([""] * (index - len(self.store.notes) + 1))
        if new_note == old_note:
            return UpdateNoteResult(
                word=str(self.store.words[index] or ""),
                note=new_note,
                saved_to_source=False,
                changed=False,
            )
        self.store.notes[index] = new_note
        saved_to_source = False
        source_path = str(self.store.get_current_source_path() or "").lower()
        if source_path.endswith(".csv"):
            saved_to_source = self.save_back_to_source()
        return UpdateNoteResult(
            word=str(self.store.words[index] or ""),
            note=new_note,
            saved_to_source=saved_to_source,
            changed=True,
        )

    def update_word(self, index, raw_value, *, translations=None, word_pos=None):
        if index is None or index < 0 or index >= len(self.store.words):
            raise IndexError("word index out of range")
        new_word = " ".join(str(raw_value or "").strip().split())
        old_word = str(self.store.words[index] or "")
        if not new_word:
            raise ValueError("Word cannot be empty.")
        if new_word == old_word:
            return UpdateWordResult(
                old_word=old_word,
                new_word=new_word,
                note=str(self.store.notes[index] if index < len(self.store.notes) else ""),
                saved_to_source=False,
                changed=False,
            )
        old_note = self.store.notes[index] if index < len(self.store.notes) else ""
        self.store.words[index] = new_word
        if isinstance(translations, dict) and old_word in translations and new_word not in translations:
            translations.pop(old_word, None)
        if isinstance(word_pos, dict):
            word_pos.pop(old_word, None)
        saved_to_source = self.save_back_to_source()
        return UpdateWordResult(
            old_word=old_word,
            new_word=new_word,
            note=str(old_note or ""),
            saved_to_source=saved_to_source,
            changed=True,
        )

    def delete_history_item(self, path):
        target = str(path or "").strip()
        name = str(os.path.basename(target) or target)
        current_source = os.path.abspath(str(self.store.get_current_source_path() or "").strip()) if self.store.get_current_source_path() else ""
        self.store.remove_history(target)
        removed_count = tts_cleanup_cache_for_source_path(target)
        detached_current_source = bool(current_source and current_source == os.path.abspath(target))
        if detached_current_source:
            self.store.detach_current_source()
        return DeleteHistoryResult(
            name=name,
            removed_count=removed_count,
            detached_current_source=detached_current_source,
        )

    def rename_history_item(self, old_path, new_path):
        old_target = os.path.abspath(str(old_path or "").strip())
        new_target = os.path.abspath(str(new_path or "").strip())
        current_source = os.path.abspath(str(self.store.get_current_source_path() or "").strip()) if self.store.get_current_source_path() else ""
        os.rename(old_target, new_target)
        migration = tts_rename_cache_source_path(old_target, new_target) or {}
        self.store.rename_history_path(old_target, new_target)
        if current_source and current_source == old_target:
            tts_set_preferred_pending_source(new_target)
        return RenameHistoryResult(
            path=new_target,
            migrated=int(migration.get("migrated", 0) or 0),
            queued=int(migration.get("queued", 0) or 0),
        )
