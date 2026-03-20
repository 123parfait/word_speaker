# -*- coding: utf-8 -*-
from dataclasses import dataclass


@dataclass
class DetailViewState:
    detail_word: str
    detail_translation: str
    detail_note: str
    detail_meta: str
    review_focus: str
    review_source: str
    review_stats: str | None
    review_open_source_enabled: bool
    has_selection: bool
    has_words: bool


def _compose_detail_translation(*, selected_pos="", selected_translation="", selected_phonetic=""):
    parts = []
    if str(selected_pos or "").strip():
        parts.append(str(selected_pos).strip())
    if str(selected_translation or "").strip():
        parts.append(str(selected_translation).strip())
    detail_translation = " ".join(parts).strip()
    phonetic = str(selected_phonetic or "").strip()
    if phonetic:
        detail_translation = f"{detail_translation}    {phonetic}".strip() if detail_translation else phonetic
    return detail_translation or "词性 / 中文 / 音标：加载中或不可用"


def _source_name(current_source_path, has_unsaved_manual_words):
    if current_source_path:
        import os

        return os.path.basename(current_source_path)
    if has_unsaved_manual_words:
        return "manual list (unsaved)"
    return "manual list"


def build_detail_view_state(
    *,
    total_words,
    selected_idx,
    current_word,
    selected_word=None,
    selected_note="",
    selected_translation="",
    selected_pos="",
    selected_phonetic="",
    current_source_path=None,
    has_current_source_file=False,
    has_unsaved_manual_words=False,
    order_mode="order",
    play_state="stopped",
):
    source_label = f"Source file: {_source_name(current_source_path, has_unsaved_manual_words)}"
    if selected_idx is None or selected_idx < 0 or not selected_word:
        return DetailViewState(
            detail_word="No word selected",
            detail_translation="",
            detail_note="Select a word to see notes and translation.",
            detail_meta=f"Words loaded: {total_words}",
            review_focus="Focus: select a word or start playback.",
            review_source=source_label,
            review_stats=f"Words: {total_words} | Current mode: {order_mode} | Current status: {play_state}",
            review_open_source_enabled=bool(has_current_source_file),
            has_selection=False,
            has_words=bool(total_words),
        )

    detail_translation = _compose_detail_translation(
        selected_pos=selected_pos,
        selected_translation=selected_translation,
        selected_phonetic=selected_phonetic,
    )
    meta_parts = [f"Position: {int(selected_idx) + 1}/{int(total_words)}"]
    if current_word == selected_word:
        meta_parts.append("Playback focus")
    return DetailViewState(
        detail_word=str(selected_word),
        detail_translation=detail_translation,
        detail_note=f"Notes: {selected_note}" if str(selected_note or "").strip() else "Notes: none",
        detail_meta=" | ".join(meta_parts),
        review_focus=f"Focus: {selected_word}. Use Dictation for recall, or Tools for sentence and corpus lookup.",
        review_source=source_label,
        review_stats=f"Words: {total_words} | Current mode: {order_mode} | Current status: {play_state}",
        review_open_source_enabled=bool(has_current_source_file),
        has_selection=True,
        has_words=bool(total_words),
    )


def build_recent_wrong_detail_view_state(
    *,
    context_word,
    wrong_count,
    note="",
    translation="",
    pos_label="",
    phonetic="",
    current_source_path=None,
    has_current_source_file=False,
    has_unsaved_manual_words=False,
):
    source_label = f"Source file: {_source_name(current_source_path, has_unsaved_manual_words)}"
    detail_translation = _compose_detail_translation(
        selected_pos=pos_label,
        selected_translation=translation,
        selected_phonetic=phonetic,
    )
    return DetailViewState(
        detail_word=str(context_word),
        detail_translation=detail_translation,
        detail_note=f"Notes: {note}" if str(note or "").strip() else "Notes: none",
        detail_meta=f"Recent wrong word | Wrong count: {int(wrong_count or 0)}",
        review_focus=f"Focus: {context_word}. Edit or review from Recent Wrong.",
        review_source=source_label,
        review_stats=None,
        review_open_source_enabled=bool(has_current_source_file),
        has_selection=False,
        has_words=True,
    )
