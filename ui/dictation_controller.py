# -*- coding: utf-8 -*-
from dataclasses import dataclass

from services.tts import (
    cleanup_word_audio_cache as tts_cleanup_word_audio_cache,
    promote_word_audio_to_recent_wrong as tts_promote_word_audio_to_recent_wrong,
)


@dataclass
class DictationAttemptResult:
    attempt_entry: dict
    appended_wrong_item: dict | None
    cleared_recent_wrong: bool


@dataclass
class DictationSessionSummary:
    accuracy: float
    total: int


@dataclass
class DictationSessionState:
    pool: list[str]
    index: int
    current_word: str
    wrong_items: list[dict]
    attempts: list[dict]
    correct_count: int
    answer_revealed: bool
    running: bool
    paused: bool
    session_source_path: str | None
    session_list_mode: str
    previous_accuracy: float | None
    summary_text: str


@dataclass
class DictationResetState:
    pool: list[str]
    index: int
    current_word: str
    session_source_path: str | None
    session_list_mode: str
    attempts: list[dict]
    wrong_items: list[dict]
    correct_count: int
    answer_revealed: bool
    running: bool
    paused: bool
    progress_text: str
    timer_text: str


class DictationController:
    def __init__(self, store):
        self.store = store

    def accuracy_so_far(self, attempts):
        attempted = len(list(attempts or []))
        if attempted <= 0:
            return 0.0
        correct = sum(1 for item in attempts if item.get("correct"))
        return (correct / float(attempted)) * 100.0

    def finish_session(self, *, correct_count, total):
        accuracy = (float(correct_count) / float(total)) * 100.0 if total else 0.0
        self.store.save_last_dictation_accuracy(accuracy)
        return DictationSessionSummary(accuracy=accuracy, total=int(total or 0))

    def build_session_state(self, *, pool, list_mode, session_source_path=None, start_index=0):
        words = list(pool or [])
        previous_accuracy = self.store.get_last_dictation_accuracy()
        safe_start = max(0, min(int(start_index or 0), max(0, len(words) - 1)))
        if safe_start > 0 and words:
            words = words[safe_start:] + words[:safe_start]
        return DictationSessionState(
            pool=words,
            index=-1,
            current_word="",
            wrong_items=[],
            attempts=[],
            correct_count=0,
            answer_revealed=False,
            running=True,
            paused=False,
            session_source_path=session_source_path,
            session_list_mode=str(list_mode or "all").strip().lower() or "all",
            previous_accuracy=previous_accuracy,
            summary_text="",
        )

    def build_reset_state(self):
        return DictationResetState(
            pool=[],
            index=-1,
            current_word="",
            session_source_path=None,
            session_list_mode="all",
            attempts=[],
            wrong_items=[],
            correct_count=0,
            answer_revealed=False,
            running=False,
            paused=False,
            progress_text="Spelling (0/0)",
            timer_text="5s",
        )

    def build_review_rows(
        self,
        attempts,
        *,
        translations=None,
        word_pos=None,
        blank_answer_label="",
        wrong_only=False,
    ):
        rows = []
        cached_translations = translations or {}
        cached_pos = word_pos or {}
        ordered_attempts = sorted(list(attempts or []), key=lambda item: int(item.get("position", 0)))
        for item in ordered_attempts:
            word = str(item.get("word") or "").strip()
            if not word:
                continue
            pos = str(cached_pos.get(word) or "").strip()
            translation = str(cached_translations.get(word) or "").strip()
            subtitle = " ".join(part for part in [pos, translation] if part).strip()
            typed = str(item.get("input") or "").strip() or str(blank_answer_label or "")
            rows.append(
                {
                    "position": int(item.get("position", 0)),
                    "word": word,
                    "subtitle": subtitle,
                    "input": typed,
                    "correct": bool(item.get("correct")),
                    "wrong_count": int(item.get("wrong_count", 0) or 0),
                }
            )
        if wrong_only:
            rows = [row for row in rows if not row.get("correct")]
        return rows

    def record_attempt(
        self,
        *,
        target,
        user_text,
        is_correct,
        position,
        list_mode,
        recent_wrong_source_path=None,
        session_source_path=None,
    ):
        token = str(target or "").strip()
        typed = str(user_text or "").strip()
        previous_stats = self.store.snapshot_dictation_word_stats(token)
        self.store.record_dictation_result(token, typed, is_correct)
        stat_entry = self.store.get_dictation_word_stats(token)
        attempt_entry = {
            "position": int(position),
            "word": token,
            "input": typed,
            "correct": bool(is_correct),
            "wrong_count": int(stat_entry.get("wrong_count", 0) or 0),
            "previous_stats": previous_stats,
        }
        appended_wrong_item = None
        cleared_recent_wrong = False
        if is_correct:
            if str(list_mode or "").strip().lower() == "recent":
                self.store.clear_wrong_word(token)
                if recent_wrong_source_path:
                    tts_cleanup_word_audio_cache(token, source_path=recent_wrong_source_path)
                attempt_entry["wrong_count"] = 0
                cleared_recent_wrong = True
        else:
            appended_wrong_item = {"position": int(position), "word": token, "input": typed}
            tts_promote_word_audio_to_recent_wrong(
                token,
                source_path=session_source_path or self.store.get_current_source_path(),
            )
        return DictationAttemptResult(
            attempt_entry=attempt_entry,
            appended_wrong_item=appended_wrong_item,
            cleared_recent_wrong=cleared_recent_wrong,
        )

    def revert_attempt(self, attempt_entry, *, recent_wrong_source_path=None):
        entry = dict(attempt_entry or {})
        token = str(entry.get("word") or "").strip()
        if not token:
            return False
        previous_stats = entry.get("previous_stats")
        self.store.restore_dictation_word_stats(token, previous_stats)
        if not previous_stats or int(previous_stats.get("wrong_count", 0) or 0) <= 0:
            if recent_wrong_source_path:
                tts_cleanup_word_audio_cache(token, source_path=recent_wrong_source_path)
        return True
