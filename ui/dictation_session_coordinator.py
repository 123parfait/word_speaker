# -*- coding: utf-8 -*-
import re
import unicodedata

from services.tts import cancel_all as tts_cancel_all, speak_async


def normalize_compare_text(text):
    raw = unicodedata.normalize("NFKC", str(text or "").casefold())
    raw = raw.replace("£", " ").replace("$", " ").replace("€", " ").replace("¥", " ")
    raw = raw.replace("'", "").replace('"', "")
    raw = raw.replace("-", " ").replace("/", " ").replace("\\", " ")
    raw = re.sub(r"[.,:;!?()\[\]{}]", " ", raw)
    raw = re.sub(r"\s+", " ", raw).strip()
    compact = re.sub(r"[^0-9a-z ]+", "", raw)
    return re.sub(r"\s+", " ", compact).strip()


def start_session(host, start_index=0):
    host._stop_main_word_playback()
    pool = host._get_dictation_pool()
    if not pool:
        host.show_info("no_words_for_dictation")
        reset_view(host)
        return
    state = host.dictation_controller.build_session_state(
        pool=pool,
        list_mode=host.dictation_list_mode_var.get(),
        session_source_path=host._get_dictation_preview_source_path(),
        start_index=start_index,
        order_mode=host.dictation_order_var.get(),
    )
    host.dictation_pool = list(state.pool)
    host.dictation_session_list_mode = state.session_list_mode
    host.dictation_session_source_path = state.session_source_path
    host.dictation_previous_session_accuracy = state.previous_accuracy
    host.dictation_index = state.index
    host.dictation_wrong_items = list(state.wrong_items)
    host.dictation_session_attempts = list(state.attempts)
    host.dictation_correct_count = state.correct_count
    host.dictation_current_word = state.current_word
    host.dictation_answer_revealed = state.answer_revealed
    host.dictation_running = state.running
    host.dictation_paused = state.paused
    host.dictation_summary_var.set(state.summary_text)
    host.update_dictation_play_button()
    host._show_dictation_frame(host.dictation_session_frame)
    advance_word(host, initial=True)


def play_current_word(host):
    if not host.dictation_running or not host.dictation_current_word:
        return
    host._cancel_dictation_play_start()
    host.dictation_paused = False
    host.update_dictation_play_button()
    host.status_var.set(host.trf("dictation_playing", word=host.dictation_current_word))
    delay_ms = 120

    def _start_playback():
        host.dictation_play_after = None
        if not host.dictation_running or host.dictation_paused or not host.dictation_current_word:
            return
        speak_async(
            host.dictation_current_word,
            host._dictation_playback_volume_ratio(),
            rate_ratio=1.0 if host.dictation_speed_var.get() == "adaptive" else float(host.dictation_speed_var.get()),
            cancel_before=True,
            source_path=host.dictation_session_source_path or host._get_dictation_preview_source_path(),
            pre_silence_ms=90,
        )
        restart_timer(host)
        host._focus_dictation_input()

    host.dictation_play_after = host.after(delay_ms, _start_playback)


def replay_current_word(host):
    if not host.dictation_current_word:
        return
    play_current_word(host)


def toggle_play_pause(host):
    if not host.dictation_running or not host.dictation_current_word:
        return
    if host.dictation_paused:
        play_current_word(host)
    else:
        pause_session(host)


def pause_session(host):
    host.dictation_paused = True
    host._cancel_dictation_play_start()
    tts_cancel_all()
    host._cancel_dictation_timer()
    host.update_dictation_play_button()
    host.dictation_status_var.set(host.tr("dictation_paused"))
    host._focus_dictation_input()


def previous_word(host):
    if not host.dictation_running or not host.dictation_pool:
        return
    host._cancel_dictation_play_start()
    host._cancel_dictation_feedback_reset()
    host._cancel_dictation_timer()
    tts_cancel_all()
    host.dictation_paused = False

    target_position = host.dictation_index if host.dictation_answer_revealed else (host.dictation_index - 1)
    if target_position < 0:
        target_position = 0

    invalidated_attempt = None
    for idx, item in enumerate(host.dictation_session_attempts):
        if int(item.get("position", -1)) == int(target_position):
            invalidated_attempt = host.dictation_session_attempts.pop(idx)
            break
    if invalidated_attempt:
        host.dictation_controller.revert_attempt(
            invalidated_attempt,
            recent_wrong_source_path=host._get_recent_wrong_cache_source_path(),
        )
        if invalidated_attempt.get("correct"):
            host.dictation_correct_count = max(0, host.dictation_correct_count - 1)
        host.dictation_wrong_items = [
            item
            for item in host.dictation_wrong_items
            if int(item.get("position", -1)) != int(target_position)
        ]
        host.refresh_dictation_recent_list()
        host._refresh_dictation_answer_review_popup()

    host.dictation_index = max(-1, target_position - 1)
    advance_word(host, initial=True)


def restart_timer(host):
    host._cancel_dictation_timer()
    host.dictation_seconds_left = host._dictation_seconds_for_speed()
    if host.dictation_seconds_left <= 0:
        host.dictation_timer_var.set("")
        return
    host.dictation_timer_var.set(f"{host.dictation_seconds_left}s")
    host.dictation_timer_after = host.after(1000, host._tick_dictation_timer)


def tick_timer(host):
    if host.dictation_paused or not host.dictation_running:
        return
    host.dictation_seconds_left -= 1
    if host.dictation_seconds_left <= 0:
        host.dictation_timer_var.set("0s")
        finalize_attempt(host, trigger="timeout")
        return
    host.dictation_timer_var.set(f"{host.dictation_seconds_left}s")
    host.dictation_timer_after = host.after(1000, host._tick_dictation_timer)


def on_input_change(host):
    if not host.dictation_running or not host.dictation_current_word or not host.dictation_input:
        return
    value = str(host.dictation_input.get() or "").strip()
    target = str(host.dictation_current_word or "").strip()
    if host.dictation_feedback_var.get() != "live":
        host._set_dictation_input_color("neutral")
        return
    if not value:
        host._set_dictation_input_color("neutral")
        return
    value_key = normalize_compare_text(value)
    target_key = normalize_compare_text(target)
    if value_key == target_key:
        host._set_dictation_input_color("correct")
        finalize_attempt(host, trigger="input")
        return
    if target_key.startswith(value_key):
        host._set_dictation_input_color("neutral")
        host.dictation_status_var.set(host.tr("dictation_keep_spelling"))
        return
    host._set_dictation_input_color("wrong")
    host.dictation_status_var.set(host.tr("dictation_wrong_live"))


def finalize_attempt(host, trigger="manual"):
    if not host.dictation_running or host.dictation_answer_revealed or not host.dictation_current_word:
        return
    host._cancel_dictation_timer()
    host.dictation_answer_revealed = True
    user_text = str(host.dictation_input.get() or "").strip() if host.dictation_input else ""
    target = str(host.dictation_current_word or "").strip()
    is_correct = normalize_compare_text(user_text) == normalize_compare_text(target)
    if is_correct:
        host.dictation_correct_count += 1
        host._set_dictation_input_color("correct")
        host.dictation_status_var.set(host.tr("dictation_correct"))
    else:
        if host.dictation_input and host.dictation_input.winfo_exists():
            try:
                host.dictation_input.delete(0, "end")
                host.dictation_input.insert(0, target)
                host.dictation_input.select_range(0, "end")
                host.dictation_input.icursor("end")
            except Exception:
                pass
        host._set_dictation_input_color("wrong")
        host.dictation_status_var.set(host.trf("dictation_wrong_answer", word=target))
    if host.dictation_session_frame and host.dictation_session_frame.winfo_exists():
        try:
            host.dictation_session_frame.update_idletasks()
        except Exception:
            pass
    result = host.dictation_controller.record_attempt(
        target=target,
        user_text=user_text,
        is_correct=is_correct,
        position=host.dictation_index,
        list_mode=host.dictation_session_list_mode,
        recent_wrong_source_path=host._get_recent_wrong_cache_source_path(),
        session_source_path=host.dictation_session_source_path,
    )
    attempt_entry = dict(result.attempt_entry)
    replaced = False
    for idx, item in enumerate(host.dictation_session_attempts):
        if int(item.get("position", -1)) == int(host.dictation_index):
            host.dictation_session_attempts[idx] = attempt_entry
            replaced = True
            break
    if not replaced:
        host.dictation_session_attempts.append(attempt_entry)
    if result.cleared_recent_wrong:
        host.refresh_dictation_recent_list()
    if result.appended_wrong_item:
        host.dictation_wrong_items.append(dict(result.appended_wrong_item))
    host._refresh_dictation_answer_review_popup()

    if not is_correct:
        delay = 2200
    elif trigger == "input":
        delay = 1150
    else:
        delay = 1450
    host._cancel_dictation_feedback_reset()
    host.dictation_feedback_after = host.after(delay, host._go_to_next_dictation_word)


def advance_word(host, initial=False):
    if not host.dictation_running and not initial:
        return
    if not initial and host.dictation_current_word and not host.dictation_answer_revealed:
        finalize_attempt(host, trigger="manual")
        return

    host._cancel_dictation_play_start()
    host._cancel_dictation_feedback_reset()
    host._cancel_dictation_timer()
    host.dictation_index += 1
    total = len(host.dictation_pool)
    if host.dictation_index >= total:
        finish_session(host)
        return

    host.dictation_current_word = host.dictation_pool[host.dictation_index]
    host.dictation_answer_revealed = False
    host.dictation_progress_var.set(f"Spelling ({host.dictation_index + 1}/{total})")
    host.dictation_progress["value"] = ((host.dictation_index + 1) / max(1, total)) * 100.0
    host.dictation_status_var.set(host.tr("dictation_listen_type"))
    seconds_for_speed = host._dictation_seconds_for_speed()
    host.dictation_timer_var.set(f"{seconds_for_speed}s" if seconds_for_speed else "")
    if host.dictation_input:
        host.dictation_input.delete(0, "end")
        host.dictation_input.focus_set()
    host._set_dictation_input_color("neutral")
    host.update_dictation_play_button()
    play_current_word(host)


def finish_session(host):
    host.dictation_running = False
    host.dictation_paused = False
    host._cancel_dictation_play_start()
    host._cancel_dictation_timer()
    host._cancel_dictation_feedback_reset()
    host.update_dictation_play_button()
    summary = host.dictation_controller.finish_session(
        correct_count=host.dictation_correct_count,
        total=len(host.dictation_pool),
    )
    host.dictation_summary_var.set(f"{summary.accuracy:.2f}%")
    host.dictation_status_var.set(host.tr("dictation_session_complete"))
    host._render_dictation_answer_review_views()
    host.refresh_dictation_recent_list()
    host._show_dictation_frame(host.dictation_result_frame)


def reset_view(host):
    state = host.dictation_controller.build_reset_state()
    host.dictation_running = state.running
    host.dictation_paused = state.paused
    host.dictation_pool = list(state.pool)
    host.dictation_index = state.index
    host.dictation_current_word = state.current_word
    host.dictation_session_source_path = state.session_source_path
    host.dictation_session_list_mode = state.session_list_mode
    host.dictation_session_attempts = list(state.attempts)
    host.dictation_wrong_items = list(state.wrong_items)
    host.dictation_correct_count = state.correct_count
    host.dictation_answer_revealed = state.answer_revealed
    host._cancel_dictation_play_start()
    host._cancel_dictation_timer()
    host._cancel_dictation_feedback_reset()
    host.update_dictation_play_button()
    host.dictation_progress_var.set(state.progress_text)
    host.dictation_timer_var.set(state.timer_text)
    host.dictation_status_var.set(host.tr("dictation_recent_title"))
    host.dictation_progress["value"] = 0
    if host.dictation_input:
        host.dictation_input.delete(0, "end")
    host._set_dictation_input_color("neutral")
    if host.dictation_result_accuracy_var is not None:
        host.dictation_result_accuracy_var.set("0.00%")
    if host.dictation_result_last_var is not None:
        host.dictation_result_last_var.set("-")
    if host.dictation_result_filter_var is not None:
        host.dictation_result_filter_var.set(host.tr("show_wrong_only"))
    host._render_dictation_answer_review_tree(host.dictation_result_review_tree)
    host.close_dictation_answer_review_popup()
    host._show_dictation_frame(host.dictation_setup_frame)
    host.refresh_dictation_recent_list()
