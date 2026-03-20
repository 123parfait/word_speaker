# -*- coding: utf-8 -*-
import threading
from tkinter import messagebox

from services.tts import cancel_all as tts_cancel_all, get_runtime_label as tts_get_runtime_label, speak_async, speak_stream_async
from ui.async_event_helper import clear_event_queue, drain_event_queue, emit_event
from ui.passage_panel import build_passage_window
from ui.passage_presenter import build_generated_passage_state, build_passage_audio_status, build_partial_passage_state
from ui.word_tools_async import start_sentence_generation_task, start_synonym_lookup_task
from ui.word_tools_panel import build_sentence_window, build_synonym_window
from ui.word_tools_presenter import build_sentence_view_state, build_synonym_view_state


def open_passage_window(host, tooltip_type):
    if host.passage_window and host.passage_window.winfo_exists():
        host.passage_window.lift()
        return
    build_passage_window(host, tooltip_type)
    host._set_passage_text(host.current_passage or "")
    host.refresh_gemini_models()


def generate_passage(host):
    if not host.store.words:
        messagebox.showinfo("Info", "Please import words first.")
        return
    if not host._require_gemini_ready():
        return

    words = host._get_selected_words_for_passage()
    model_name = host._get_selected_gemini_model()
    host.passage_generation_token += 1
    token = host.passage_generation_token
    selected_count = len(host._get_selected_indices())
    source_text = f"{len(words)} selected words" if selected_count else f"{len(words)} words"
    host.passage_status_var.set(f"Generating with Gemini ({model_name}) from {source_text}...")
    host.current_passage = ""
    host.current_passage_original = ""
    host.current_passage_words = []
    host.passage_is_practice = False
    host.passage_cloze_text = ""
    host.passage_answers = []
    host._set_passage_text("")
    host._clear_passage_practice_input()
    host._clear_passage_practice_result()
    host.passage_generation_active_token = token
    clear_passage_event_queue(host)

    threading.Thread(
        target=lambda: host._run_passage_generation(token, words, model_name),
        daemon=True,
    ).start()
    host.after(80, lambda t=token: host._poll_passage_generation_events(t))


def clear_passage_event_queue(host):
    clear_event_queue(host.passage_event_queue)


def emit_passage_event(host, event_type, token, payload=None):
    emit_event(host.passage_event_queue, event_type, token, payload)


def poll_passage_generation_events(host, token):
    done = drain_event_queue(
        target_queue=host.passage_event_queue,
        token=token,
        active_token=host.passage_generation_active_token,
        handlers={
            "partial": lambda payload: update_partial_passage(host, token, payload),
            "result": lambda payload: apply_generated_passage(host, token, payload),
            "error": lambda payload: messagebox.showerror("Generate Error", str(payload or "Unknown error")),
        },
    )
    if not done and token == host.passage_generation_active_token:
        host.after(80, lambda t=token: host._poll_passage_generation_events(t))


def update_partial_passage(host, token, text):
    if token != host.passage_generation_token:
        return
    state = build_partial_passage_state(text)
    host.current_passage = state["passage"]
    host.current_passage_original = host.current_passage
    if state["has_passage"]:
        host._set_passage_text(host.current_passage)


def apply_generated_passage(host, token, result):
    if token != host.passage_generation_token:
        return
    state = build_generated_passage_state(result, default_model=host.DEFAULT_GEMINI_MODEL if hasattr(host, "DEFAULT_GEMINI_MODEL") else "gemini-2.5-flash")
    host.current_passage = state["passage"]
    host.current_passage_original = host.current_passage
    host.current_passage_words = list(state["used_words"])
    host.passage_is_practice = False
    host.passage_cloze_text = ""
    host.passage_answers = []
    host._set_passage_text(host.current_passage)
    host.passage_status_var.set(state["status_text"])


def pause_word_playback_for_passage(host):
    host.cancel_schedule()
    host.play_token += 1
    if host.play_state == "playing":
        host.play_state = "paused"
        host.status_var.set("Paused (reading passage)")
        host.update_play_button()
    tts_cancel_all()


def play_generated_passage(host):
    text = host.current_passage_original.strip() or host.current_passage.strip() or host._get_passage_text()
    if not text:
        messagebox.showinfo("Info", "Generate a passage first.")
        return

    speech_text = host._speech_text_from_passage(text)
    if not speech_text:
        messagebox.showinfo("Info", "Passage is empty.")
        return

    pause_word_playback_for_passage(host)
    runtime = tts_get_runtime_label()
    token = speak_stream_async(
        speech_text,
        host.volume_var.get() / 100.0,
        rate_ratio=host.speech_rate_var.get(),
        cancel_before=False,
        chunk_chars=90,
    )
    host.passage_status_var.set(build_passage_audio_status(runtime))
    host._watch_tts_backend(token, target="passage", text_label="passage")


def stop_passage_playback(host):
    tts_cancel_all()
    host.passage_status_var.set("Stopped.")


def clear_sentence_event_queue(host):
    clear_event_queue(host.sentence_event_queue)


def emit_sentence_event(host, event_type, token, payload=None):
    emit_event(host.sentence_event_queue, event_type, token, payload)


def clear_synonym_event_queue(host):
    clear_event_queue(host.synonym_event_queue)


def emit_synonym_event(host, event_type, token, payload=None):
    emit_event(host.synonym_event_queue, event_type, token, payload)


def poll_synonym_events(host, token):
    done = drain_event_queue(
        target_queue=host.synonym_event_queue,
        token=token,
        active_token=host.synonym_lookup_active_token,
        handlers={
            "result": lambda payload: show_synonym_window(
                host,
                (payload or {}).get("word", ""),
                (payload or {}).get("focus", ""),
                (payload or {}).get("synonyms") or [],
                (payload or {}).get("source", ""),
            ),
            "error": lambda payload: messagebox.showerror(host.tr("synonyms_error"), str(payload or "Unknown error")),
        },
    )
    if not done and token == host.synonym_lookup_active_token:
        host.after(80, lambda t=token: host._poll_synonym_events(t))


def poll_sentence_events(host, token):
    done = drain_event_queue(
        target_queue=host.sentence_event_queue,
        token=token,
        active_token=host.sentence_generation_active_token,
        handlers={
            "result": lambda payload: show_sentence_window(
                host,
                (payload or {}).get("word", ""),
                (payload or {}).get("sentence", ""),
                (payload or {}).get("source", "Unknown"),
            ),
            "error": lambda payload: messagebox.showerror("Sentence Error", str(payload or "Unknown error")),
        },
    )
    if not done and token == host.sentence_generation_active_token:
        host.after(80, lambda t=token: host._poll_sentence_events(t))


def make_sentence_for_selected_word(host):
    word = host._get_context_word()
    if not word:
        host.show_info("select_word_first")
        return
    if not host._require_gemini_ready():
        return
    model_name = host._get_selected_gemini_model()
    host.status_var.set(f"Generating IELTS sentence for '{word}' with {model_name}...")
    host.sentence_generation_token += 1
    token = host.sentence_generation_token
    host.sentence_generation_active_token = token
    clear_sentence_event_queue(host)
    start_sentence_generation_task(
        token=token,
        word=word,
        model_name=model_name,
        fallback_sentence=host._fallback_sentence(word),
        emit_event=lambda event_type, event_token, payload=None: emit_sentence_event(host, event_type, event_token, payload),
    )
    host.after(80, lambda t=token: host._poll_sentence_events(t))


def lookup_synonyms_for_selected_word(host):
    word = host._get_context_word()
    if not word:
        host.show_info("select_word_first")
        return
    host.status_var.set(f"Looking up synonyms for '{word}'...")
    host.synonym_lookup_token += 1
    token = host.synonym_lookup_token
    host.synonym_lookup_active_token = token
    clear_synonym_event_queue(host)
    start_synonym_lookup_task(
        token=token,
        word=word,
        emit_event=lambda event_type, event_token, payload=None: emit_synonym_event(host, event_type, event_token, payload),
    )
    host.after(80, lambda t=token: host._poll_synonym_events(t))


def show_sentence_window(host, word, sentence, source):
    state = build_sentence_view_state(word, sentence, source)
    host.status_var.set(state["status_text"])
    build_sentence_window(
        host,
        state=state,
        on_read=lambda s=sentence: speak_async(
            s,
            host.volume_var.get() / 100.0,
            rate_ratio=host.speech_rate_var.get(),
            cancel_before=True,
            source_path=host.store.get_current_source_path(),
        ),
    )


def show_synonym_window(host, word, focus, synonyms, source=None):
    state = build_synonym_view_state(
        tr=host.tr,
        trf=host.trf,
        word=word,
        focus=focus,
        synonyms=synonyms,
        source=source,
    )
    host.status_var.set(state["status_text"])
    build_synonym_window(host, state=state)
