# -*- coding: utf-8 -*-
from services.phonetics import get_cached_phonetics
from services.translation import get_cached_translations
from services.word_analysis import get_cached_pos
from ui.word_metadata_async import (
    start_analysis_task,
    start_phonetic_task,
    start_single_phonetic_task,
    start_single_translation_task,
    start_translation_task,
)
from ui.word_metadata_presenter import (
    build_render_words_state,
    can_apply_batch_metadata,
    can_apply_single_translation,
    normalize_requested_words,
)
from ui.word_table_helper import refresh_word_table_rows


def start_analysis_job(host, words, token):
    requested_words = normalize_requested_words(words)
    if not requested_words:
        return
    host.pending_analysis_words.update(requested_words)
    start_analysis_task(
        requested_words=requested_words,
        token=token,
        after=host.after,
        on_complete=host._apply_pos_analysis,
    )


def start_translation_job(host, words, token):
    requested_words = normalize_requested_words(words)
    if not requested_words:
        return
    host.pending_translation_words.update(requested_words)
    start_translation_task(
        requested_words=requested_words,
        token=token,
        after=host.after,
        on_complete=host._apply_translations,
    )


def start_phonetic_job(host, words, token):
    requested_words = normalize_requested_words(words)
    if not requested_words:
        return
    host.pending_phonetic_words.update(requested_words)
    start_phonetic_task(
        requested_words=requested_words,
        token=token,
        after=host.after,
        on_complete=host._apply_phonetics,
    )


def start_single_translation(host, row_idx, word):
    token = host.translation_token
    start_single_translation_task(
        word=word,
        row_idx=row_idx,
        token=token,
        after=host.after,
        on_complete=host._apply_single_translation,
    )


def start_single_phonetic(host, row_idx, word):
    token = host.phonetic_token
    start_single_phonetic_task(
        word=word,
        row_idx=row_idx,
        token=token,
        after=host.after,
        on_complete=host._apply_single_phonetic,
    )


def apply_pos_analysis(host, token, requested_words, analyzed):
    for word in requested_words or []:
        host.pending_analysis_words.discard(word)
    if not can_apply_batch_metadata(
        token=token,
        active_token=host.analysis_token,
        has_word_table=bool(host.word_table),
    ):
        return
    host.word_pos.update(analyzed)
    refresh_word_table_rows(
        table=host.word_table,
        words=host.store.words,
        notes=host.store.notes,
        build_values=host._build_word_table_values,
    )
    host._refresh_selection_details()


def apply_phonetics(host, token, requested_words, phonetics):
    for word in requested_words or []:
        host.pending_phonetic_words.discard(word)
    if not can_apply_batch_metadata(
        token=token,
        active_token=host.phonetic_token,
        has_word_table=bool(host.word_table),
    ):
        return
    host.word_phonetics.update(phonetics)
    host._refresh_selection_details()


def apply_translations(host, token, requested_words, translated):
    for word in requested_words or []:
        host.pending_translation_words.discard(word)
    if not can_apply_batch_metadata(
        token=token,
        active_token=host.translation_token,
        has_word_table=bool(host.word_table),
    ):
        return
    host.translations.update(translated)
    refresh_word_table_rows(
        table=host.word_table,
        words=host.store.words,
        notes=host.store.notes,
        build_values=host._build_word_table_values,
    )
    host._refresh_selection_details()


def apply_single_translation(host, token, row_idx, word, zh_text):
    if not can_apply_single_translation(
        token=token,
        active_token=host.translation_token,
        row_idx=row_idx,
        word=word,
        current_words=host.store.words,
        has_word_table=bool(host.word_table),
    ):
        return
    host.translations[word] = zh_text
    iid = str(row_idx)
    if host.word_table.exists(iid):
        note = host.store.notes[row_idx] if row_idx < len(host.store.notes) else ""
        tag = "even" if row_idx % 2 == 0 else "odd"
        host.word_table.item(iid, values=host._build_word_table_values(row_idx, word, note), tags=(tag,))
    host._refresh_selection_details()


def apply_single_phonetic(host, token, row_idx, word, phonetic_text):
    if not can_apply_single_translation(
        token=token,
        active_token=host.phonetic_token,
        row_idx=row_idx,
        word=word,
        current_words=host.store.words,
        has_word_table=bool(host.word_table),
    ):
        return
    host.word_phonetics[word] = phonetic_text
    host._refresh_selection_details()


def render_words(host, words):
    if not host.word_table:
        return
    host.cancel_word_edit()
    host.translation_token += 1
    host.analysis_token += 1
    host.phonetic_token += 1
    token = host.translation_token
    analysis_token = host.analysis_token
    phonetic_token = host.phonetic_token
    state = build_render_words_state(
        words=words,
        cached_translations=get_cached_translations(words),
        cached_pos=get_cached_pos(words),
        cached_phonetics=get_cached_phonetics(words),
    )
    host.translations = dict(state["translations"])
    host.word_pos = dict(state["word_pos"])
    host.word_phonetics = dict(state["word_phonetics"])
    host.pending_translation_words.clear()
    host.pending_analysis_words.clear()
    host.pending_phonetic_words.clear()
    host.word_table.delete(*host.word_table.get_children())
    for idx, word in enumerate(words):
        note = host.store.notes[idx] if idx < len(host.store.notes) else ""
        tag = "even" if idx % 2 == 0 else "odd"
        host.word_table.insert(
            "",
            "end",
            iid=str(idx),
            values=host._build_word_table_values(idx, word, note),
            tags=(tag,),
        )
    host.update_empty_state()
    host._refresh_selection_details()
    host.refresh_dictation_recent_list()
    if state["missing_translations"]:
        start_translation_job(host, state["missing_translations"], token)
    if state["missing_pos"]:
        start_analysis_job(host, state["missing_pos"], analysis_token)
    if state["missing_phonetics"]:
        start_phonetic_job(host, state["missing_phonetics"], phonetic_token)
    if words:
        host._start_audio_precache_job(words)


def ensure_word_metadata(host, word):
    target = str(word or "").strip()
    if not target:
        return
    if not str(host.word_pos.get(target) or "").strip() and target not in host.pending_analysis_words:
        start_analysis_job(host, [target], host.analysis_token)
    if not str(host.translations.get(target) or "").strip() and target not in host.pending_translation_words:
        start_translation_job(host, [target], host.translation_token)
    if not str(host.word_phonetics.get(target) or "").strip() and target not in host.pending_phonetic_words:
        start_phonetic_job(host, [target], host.phonetic_token)
