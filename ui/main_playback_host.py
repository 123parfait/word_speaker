# -*- coding: utf-8 -*-
from services.tts import (
    cancel_all as tts_cancel_all,
    get_backend_status as tts_get_backend_status,
    get_runtime_label as tts_get_runtime_label,
    has_cached_word_audio as tts_has_cached_word_audio,
    speak_async,
)
from services.voice_manager import SOURCE_GEMINI, get_voice_source


def sync_state(host):
    host.play_state = host.main_playback_controller.play_state
    host.queue = list(host.main_playback_controller.queue)
    host.pos = int(host.main_playback_controller.pos)
    host.current_word = host.main_playback_controller.current_word


def rebuild_on_mode_change(host):
    host.main_playback_controller.current_word = host.current_word
    host.main_playback_controller.play_state = host.play_state
    host.main_playback_controller.rebuild(
        words=host.store.words,
        selected_idx=host._get_selected_index(),
    )
    sync_state(host)
    if host.current_word is None:
        return
    set_current_word(host)
    if host.play_state == "playing":
        play_current(host)


def toggle_play(host):
    if not host.store.words:
        host.show_info("import_words_first")
        return
    if host.play_state == "playing":
        host.main_playback_controller.pause()
        sync_state(host)
        host.cancel_schedule()
        tts_cancel_all()
        host.play_token += 1
        host.status_var.set("已暂停顺序播放。")
        update_play_button(host)
        return

    host.main_playback_controller.current_word = host.current_word
    host.main_playback_controller.start_or_resume(
        words=host.store.words,
        selected_idx=host._get_selected_index(),
    )
    sync_state(host)
    host.play_token += 1
    update_play_button(host)
    play_current(host)


def play_current(host):
    if not host.current_word:
        return
    runtime = tts_get_runtime_label()
    source_path = host.store.get_current_source_path()
    cached = get_voice_source() == SOURCE_GEMINI and tts_has_cached_word_audio(
        host.current_word,
        source_path=source_path,
    )
    token = speak_async(
        host.current_word,
        host.volume_var.get() / 100.0,
        rate_ratio=host.speech_rate_var.get(),
        cancel_before=True,
        source_path=source_path,
    )
    index_text = f"{host.pos + 1}/{len(host.store.words)}" if host.store.words and host.pos >= 0 else "0/0"
    if cached:
        host.status_var.set(f"顺序播放中：{index_text}  {host.current_word}（缓存音频）")
    else:
        host.status_var.set(f"顺序播放中：{index_text}  {host.current_word}（{runtime}）")
    host._watch_tts_backend(token, target="playback", text_label=host.current_word)
    schedule_next(host, token)


def schedule_next(host, playback_token):
    host.cancel_schedule()
    if host.play_state != "playing":
        return
    host.playback_schedule_token += 1
    schedule_token = host.playback_schedule_token

    def _poll_duration(attempt=0):
        if host.play_state != "playing" or schedule_token != host.playback_schedule_token:
            return
        status = tts_get_backend_status(playback_token)
        duration_seconds = float((status or {}).get("duration_seconds") or 0.0)
        if duration_seconds > 0:
            interval_seconds = max(0.2, float(host.interval_var.get()))
            token = host.play_token
            host.after_id = host.after(int((duration_seconds + interval_seconds) * 1000), lambda: next_word(host, token))
            return
        if attempt < 80:
            host.after(100, lambda: _poll_duration(attempt + 1))
            return
        interval_seconds = max(0.2, float(host.interval_var.get()))
        token = host.play_token
        host.after_id = host.after(int(interval_seconds * 1000), lambda: next_word(host, token))

    _poll_duration()


def next_word(host, token):
    if host.play_state != "playing" or token != host.play_token:
        return
    host.main_playback_controller.advance(host.store.words)
    sync_state(host)
    set_current_word(host)
    play_current(host)


def set_current_word(host):
    if not host.queue or host.pos < 0 or host.pos >= len(host.queue):
        return
    idx = host.queue[host.pos]
    host.current_word = host.store.words[idx]
    if host.word_table and host.word_table.exists(str(idx)):
        try:
            host.suppress_word_select_action = True
            row_id = str(idx)
            host.word_table.selection_set(row_id)
            host.word_table.focus(row_id)
            host.word_table.see(row_id)
        except Exception:
            pass
    host.status_var.set(f"顺序播放：{idx + 1}/{len(host.store.words)}  {host.current_word}")
    host._refresh_selection_details()


def update_play_button(host):
    if host.play_state == "playing":
        host.play_btn.config(text=("Pause" if host.ui_language_var.get() == "en" else "暂停"))
    else:
        host.play_btn.config(text=host.tr("play"))
    host._refresh_selection_details()


def reset_state(host):
    host.cancel_schedule()
    tts_cancel_all()
    host.play_token += 1
    host.main_playback_controller.reset()
    sync_state(host)
    if host.word_table:
        host.word_table.selection_remove(*host.word_table.selection())
    host.status_var.set(host.tr("stopped"))
    update_play_button(host)
    host._refresh_selection_details()


def stop(host):
    host.cancel_schedule()
    host.play_token += 1
    host.play_state = "stopped"
    update_play_button(host)
    tts_cancel_all()
