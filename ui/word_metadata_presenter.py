# -*- coding: utf-8 -*-


def normalize_requested_words(words):
    return [str(word or "").strip() for word in (words or []) if str(word or "").strip()]


def can_apply_single_translation(*, token, active_token, row_idx, word, current_words, has_word_table):
    if token != active_token or not has_word_table:
        return False
    if row_idx < 0 or row_idx >= len(current_words or []):
        return False
    return current_words[row_idx] == word


def build_render_words_state(*, words, cached_translations, cached_pos, cached_phonetics):
    word_list = list(words or [])
    cached_translations = dict(cached_translations or {})
    cached_pos = dict(cached_pos or {})
    cached_phonetics = dict(cached_phonetics or {})
    return {
        "translations": cached_translations,
        "word_pos": cached_pos,
        "word_phonetics": cached_phonetics,
        "missing_translations": [word for word in word_list if word not in cached_translations],
        "missing_pos": [word for word in word_list if word not in cached_pos],
        "missing_phonetics": [word for word in word_list if word not in cached_phonetics],
    }


def can_apply_batch_metadata(*, token, active_token, has_word_table):
    return token == active_token and has_word_table
