# -*- coding: utf-8 -*-
import threading

from services.phonetics import get_phonetics
from services.translation import translate_words as translate_words_en_zh
from services.word_analysis import analyze_words
from ui.async_event_helper import emit_event


def start_analysis_task(*, requested_words, token, target_queue):
    def _run():
        analyzed = analyze_words(requested_words)
        emit_event(
            target_queue,
            "analysis_done",
            token,
            {"requested_words": requested_words, "analyzed": analyzed},
        )

    threading.Thread(target=_run, daemon=True).start()


def start_translation_task(*, requested_words, token, target_queue):
    def _run():
        translated = translate_words_en_zh(requested_words)
        emit_event(
            target_queue,
            "translation_done",
            token,
            {"requested_words": requested_words, "translated": translated},
        )

    threading.Thread(target=_run, daemon=True).start()


def start_phonetic_task(*, requested_words, token, target_queue):
    def _run():
        phonetics = get_phonetics(requested_words)
        emit_event(
            target_queue,
            "phonetic_done",
            token,
            {"requested_words": requested_words, "phonetics": phonetics},
        )

    threading.Thread(target=_run, daemon=True).start()


def start_single_translation_task(*, word, row_idx, token, target_queue):
    def _run():
        translated = translate_words_en_zh([word])
        emit_event(
            target_queue,
            "single_translation_done",
            token,
            {
                "row_idx": row_idx,
                "word": word,
                "zh_text": translated.get(word) or "",
            },
        )

    threading.Thread(target=_run, daemon=True).start()


def start_single_phonetic_task(*, word, row_idx, token, target_queue):
    def _run():
        phonetics = get_phonetics([word])
        emit_event(
            target_queue,
            "single_phonetic_done",
            token,
            {
                "row_idx": row_idx,
                "word": word,
                "phonetic_text": phonetics.get(word) or "",
            },
        )

    threading.Thread(target=_run, daemon=True).start()
