# -*- coding: utf-8 -*-
import threading

from services.app_config import get_llm_api_key
from services.gemini_writer import generate_english_passage_with_gemini, generate_example_sentence_with_gemini
from services.ielts_passage import build_ielts_listening_passage
from services.synonyms import get_synonyms as get_synonyms_for_word


def start_passage_generation_task(*, token, words, model_name, emit_event):
    def _run():
        result = None
        fallback_error = None
        try:
            result = generate_english_passage_with_gemini(
                words,
                api_key=get_llm_api_key(),
                max_words=24,
                timeout=90,
                model=model_name,
            )
            emit_event("partial", token, result.get("passage", ""))
        except Exception as exc:
            fallback_error = str(exc)
            try:
                result = build_ielts_listening_passage(words, max_words=24)
                result["source"] = "template"
                result["fallback_reason"] = fallback_error
            except Exception as exc2:
                template_error = str(exc2)
                emit_event(
                    "error",
                    token,
                    f"Failed to generate passage.\nGemini: {fallback_error}\nTemplate: {template_error}",
                )
                emit_event("done", token, None)
                return
        emit_event("result", token, result)
        emit_event("done", token, None)

    threading.Thread(target=_run, daemon=True).start()


def start_sentence_generation_task(*, token, word, model_name, fallback_sentence, emit_event):
    def _run():
        try:
            sentence = generate_example_sentence_with_gemini(
                word=word,
                api_key=get_llm_api_key(),
                model=model_name,
                timeout=45,
            )
            source = f"Gemini ({model_name}, IELTS)"
        except Exception:
            sentence = fallback_sentence
            source = "Fallback"
        emit_event(
            "result",
            token,
            {"word": word, "sentence": sentence, "source": source},
        )
        emit_event("done", token, None)

    threading.Thread(target=_run, daemon=True).start()


def start_synonym_lookup_task(*, token, word, emit_event):
    def _run():
        try:
            result = get_synonyms_for_word(word, limit=12)
            emit_event("result", token, result)
        except Exception as exc:
            emit_event("error", token, str(exc))
        emit_event("done", token, None)

    threading.Thread(target=_run, daemon=True).start()
