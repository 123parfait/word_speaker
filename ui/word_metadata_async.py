# -*- coding: utf-8 -*-
import threading

from services.phonetics import get_phonetics
from services.translation import translate_words as translate_words_en_zh
from services.word_analysis import analyze_words


def start_analysis_task(*, requested_words, token, after, on_complete):
    def _run():
        analyzed = analyze_words(requested_words)
        after(0, lambda: on_complete(token, requested_words, analyzed))

    threading.Thread(target=_run, daemon=True).start()


def start_translation_task(*, requested_words, token, after, on_complete):
    def _run():
        translated = translate_words_en_zh(requested_words)
        after(0, lambda: on_complete(token, requested_words, translated))

    threading.Thread(target=_run, daemon=True).start()


def start_phonetic_task(*, requested_words, token, after, on_complete):
    def _run():
        phonetics = get_phonetics(requested_words)
        after(0, lambda: on_complete(token, requested_words, phonetics))

    threading.Thread(target=_run, daemon=True).start()


def start_single_translation_task(*, word, row_idx, token, after, on_complete):
    def _run():
        translated = translate_words_en_zh([word])
        after(0, lambda: on_complete(token, row_idx, word, translated.get(word) or ""))

    threading.Thread(target=_run, daemon=True).start()


def start_single_phonetic_task(*, word, row_idx, token, after, on_complete):
    def _run():
        phonetics = get_phonetics([word])
        after(0, lambda: on_complete(token, row_idx, word, phonetics.get(word) or ""))

    threading.Thread(target=_run, daemon=True).start()
