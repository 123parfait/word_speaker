# -*- coding: utf-8 -*-
import queue

from services.app_config import get_llm_api_key, get_tts_api_key, get_tts_api_provider
from services.tts import get_online_tts_queue_status as tts_get_online_tts_queue_status
from ui.api_key_panel import build_api_key_window
from ui.settings_panel import build_settings_window
from ui.settings_presenter import build_tts_runtime_status


def clear_validation_queue(host):
    try:
        while True:
            host.gemini_validation_queue.get_nowait()
    except queue.Empty:
        return


def emit_validation_event(host, event_type, token, payload=None):
    try:
        host.gemini_validation_queue.put_nowait((event_type, token, payload))
    except Exception:
        return


def poll_validation_events(host, token):
    if token != host.gemini_validation_active_token:
        return
    done = False
    try:
        while True:
            event_type, event_token, payload = host.gemini_validation_queue.get_nowait()
            if event_token != token:
                continue
            if event_type == "success":
                host._finish_gemini_validation_success(payload or {})
            elif event_type == "success_tts":
                host._finish_tts_validation_success(payload or {})
            elif event_type == "success_api_setup":
                host._finish_combined_api_validation(payload or {})
            elif event_type == "error":
                host._finish_gemini_validation_error(str(payload or "Unknown error"))
            elif event_type == "error_tts":
                host._finish_tts_validation_error(str(payload or "Unknown error"))
            elif event_type == "done":
                done = True
    except queue.Empty:
        pass
    if not done and token == host.gemini_validation_active_token:
        host.after(80, lambda t=token: host._poll_gemini_validation_events(t))


def open_api_key_window(host, force_llm=False, force_tts=False, initial_section="llm"):
    host.gemini_verified = host.gemini_verified and not force_llm
    host.api_key_force_llm = host.api_key_force_llm or force_llm
    host.api_key_force_tts = host.api_key_force_tts or force_tts
    if host.api_key_window and host.api_key_window.winfo_exists():
        host.api_key_window.lift()
        host.api_key_window.focus_force()
        return

    host.gemini_key_status_var.set("Paste your LLM API key, then test it.")
    host.tts_key_status_var.set("Paste your TTS API key, then test it.")
    host.gemini_key_var.set(get_llm_api_key())
    host.tts_key_var.set(get_tts_api_key())
    host.llm_api_provider_var.set(host.tr("provider_gemini"))
    host.tts_api_provider_var.set(host._tts_provider_label(get_tts_api_provider()))
    build_api_key_window(host, initial_section=initial_section)


def close_api_key_window(host):
    if host.api_key_window and host.api_key_window.winfo_exists():
        try:
            host.api_key_window.grab_release()
        except Exception:
            pass
        host.api_key_window.destroy()
    host.api_key_window = None
    host.api_key_test_btn = None
    host.api_llm_entry = None
    host.api_tts_entry = None
    host.gemini_key_test_btn = None
    host.tts_key_test_btn = None
    llm_missing = host.api_key_force_llm and not str(get_llm_api_key() or "").strip()
    tts_missing = host.api_key_force_tts and not str(get_tts_api_key() or "").strip()
    host.api_key_force_llm = False
    host.api_key_force_tts = False
    if llm_missing or tts_missing:
        host.winfo_toplevel().destroy()


def set_api_entry_error(host, field, has_error):
    widget = host.api_llm_entry if field == "llm" else host.api_tts_entry
    if not widget or not widget.winfo_exists():
        return
    if has_error:
        widget.configure(bg="#fff1f2", highlightbackground="#ef4444", highlightcolor="#ef4444")
    else:
        widget.configure(bg="white", highlightbackground="#cbd5e1", highlightcolor="#2563eb")


def maybe_close_api_key_window(host):
    llm_ready = bool(str(get_llm_api_key() or "").strip())
    tts_ready = bool(str(get_tts_api_key() or "").strip())
    if host.api_key_force_llm and not llm_ready:
        return
    if host.api_key_force_tts and not tts_ready:
        return
    if host.api_key_window and host.api_key_window.winfo_exists():
        host._close_api_key_window()


def close_settings_window(host):
    if host.gemini_status_after:
        try:
            host.after_cancel(host.gemini_status_after)
        except Exception:
            pass
    host.gemini_status_after = None
    if host.settings_window and host.settings_window.winfo_exists():
        host.settings_window.destroy()
    host.settings_window = None


def refresh_settings_runtime_status(host):
    status = tts_get_online_tts_queue_status()
    provider_label = host._tts_provider_label(get_tts_api_provider())
    view_state = build_tts_runtime_status(
        tr=host.tr,
        trf=host.trf,
        provider_label=provider_label,
        status=status,
    )
    host.gemini_runtime_status_var.set(view_state["runtime_status"])
    host.gemini_retry_status_var.set(view_state["retry_status"])

    if host.settings_window and host.settings_window.winfo_exists():
        host.gemini_status_after = host.after(1000, host._refresh_settings_gemini_status)
    else:
        host.gemini_status_after = None


def open_settings_window(host):
    build_settings_window(host)
    host.update_speed_buttons()
    host.update_speech_rate_buttons()
    host.refresh_voice_list()
    refresh_settings_runtime_status(host)
