# -*- coding: utf-8 -*-
from dataclasses import dataclass

from services.translation import get_cached_translations
from services.word_analysis import get_cached_pos


@dataclass(frozen=True)
class DictationListState:
    all_items: list[dict]
    recent_items: list[dict]
    all_tab_label: str
    recent_tab_label: str
    rows: list[tuple[str, tuple[str, str, str], str]]
    empty_row: tuple[str, str, str] | None
    select_first: bool


def _resolve_word_pos(word, word_pos):
    pos_label = str(word_pos.get(word) or "").strip()
    if pos_label:
        return pos_label
    try:
        return str(get_cached_pos([word]).get(word) or "").strip()
    except Exception:
        return ""


def _resolve_translation(word, translations):
    zh_text = str(translations.get(word) or "").strip()
    if zh_text:
        return zh_text
    try:
        return str(get_cached_translations([word]).get(word) or "").strip()
    except Exception:
        return ""


def format_word_subline(word, *, word_pos, translations):
    pos_label = _resolve_word_pos(word, word_pos)
    zh_text = _resolve_translation(word, translations)
    parts = [part for part in (pos_label, zh_text) if part]
    if parts:
        return " ".join(parts)
    return "..."


def build_word_table_values(idx, word, *, note, word_pos, translations):
    display_text = f"{word}\n{format_word_subline(word, word_pos=word_pos, translations=translations)}"
    return (f"{idx + 1}.", display_text, note or "")


def build_dictation_table_values(idx, item, *, note_by_word, word_pos, translations, is_recent_mode):
    word = str(item.get("word") or "").strip()
    pos_text = _resolve_word_pos(word, word_pos)
    zh_text = _resolve_translation(word, translations)
    subtitle = format_word_subline(word, word_pos=word_pos, translations=translations)
    if pos_text or zh_text:
        subtitle = f"{pos_text}. {zh_text}".strip(". ").strip() if pos_text else zh_text

    note_value = str(note_by_word.get(word) or "").strip()
    if is_recent_mode and not note_value:
        note_value = str(item.get("note") or "").strip()

    wrong_count = int(item.get("wrong_count", 0) or 0)
    if wrong_count:
        subtitle = f"{subtitle}  |  错过{wrong_count}次"

    if is_recent_mode:
        wrong_input = str(item.get("last_wrong_input") or "").strip()
        wrong_type = str(item.get("last_wrong_type") or "").strip()
        note_parts = []
        if wrong_type:
            note_parts.append(wrong_type)
        if wrong_input:
            note_parts.append(f"错写: {wrong_input}")
        note_value = " | ".join(note_parts)

    return (f"{idx + 1}.", f"{word}\n{subtitle}", note_value)


def build_dictation_list_state(*, words, notes, recent_items, mode, word_pos, translations, tr):
    all_items = [
        {
            "word": word,
            "wrong_count": 0,
            "correct_count": 0,
            "last_wrong_input": "",
            "last_result": "",
            "last_seen": "",
        }
        for word in words
    ]
    recent_items = list(recent_items or [])
    target_mode = "recent" if str(mode or "").strip().lower() == "recent" else "all"
    items = recent_items if target_mode == "recent" else all_items
    note_by_word = {
        str(word or "").strip(): str(note or "").strip()
        for word, note in zip(words or [], notes or [])
        if str(word or "").strip()
    }

    rows = []
    for idx, item in enumerate(items):
        tag = "even" if idx % 2 == 0 else "odd"
        rows.append(
            (
                str(idx),
                build_dictation_table_values(
                    idx,
                    item,
                    note_by_word=note_by_word,
                    word_pos=word_pos,
                    translations=translations,
                    is_recent_mode=(target_mode == "recent"),
                ),
                tag,
            )
        )

    empty_row = None
    if not rows:
        if target_mode == "recent":
            empty_text = tr("dictation_empty_recent")
        else:
            empty_text = tr("dictation_empty_list") if words else tr("import_words_first")
        empty_row = ("empty", ("", empty_text, ""))

    return DictationListState(
        all_items=all_items,
        recent_items=recent_items,
        all_tab_label=f"全部({len(all_items)})",
        recent_tab_label=f"近期错词({len(recent_items)})",
        rows=rows,
        empty_row=empty_row,
        select_first=bool(rows),
    )
