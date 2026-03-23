# -*- coding: utf-8 -*-
import time
import tkinter as tk
from tkinter import ttk

from services.tts import (
    get_recent_wrong_cache_source as tts_get_recent_wrong_cache_source,
    get_runtime_label as tts_get_runtime_label,
    has_cached_word_audio as tts_has_cached_word_audio,
    speak_async,
)
from services.voice_manager import SOURCE_GEMINI, get_voice_source
from ui.dictation_panel import build_dictation_panel
from ui.list_presenter import build_dictation_list_state


def _center_window(win):
    if not win or not win.winfo_exists():
        return
    try:
        win.update_idletasks()
        width = max(int(win.winfo_width() or 0), 1)
        height = max(int(win.winfo_height() or 0), 1)
        screen_w = max(1, int(win.winfo_screenwidth() or 1))
        screen_h = max(1, int(win.winfo_screenheight() or 1))
        x = max(0, (screen_w - width) // 2)
        y = max(0, (screen_h - height) // 2)
        win.geometry(f"+{x}+{y}")
    except Exception:
        pass


def open_window(host):
    if host.dictation_window and host.dictation_window.winfo_exists():
        host.dictation_window.deiconify()
        host.dictation_window.lift()
        refresh_recent_list(host)
        return

    host._stop_main_word_playback()
    host.dictation_window = tk.Toplevel(host)
    host.dictation_window.title("Dictation")
    host.dictation_window.configure(bg="#f6f7fb")
    host.dictation_window.minsize(640, 620)
    host.dictation_window.geometry("720x700")
    host.dictation_window.protocol("WM_DELETE_WINDOW", host.close_dictation_window)

    host.check_panel = ttk.Frame(host.dictation_window, style="Card.TFrame")
    host.check_panel.pack(fill="both", expand=True, padx=12, pady=12)
    host.check_panel.grid_columnconfigure(0, weight=1)
    host.check_panel.grid_rowconfigure(0, weight=1)
    build_dictation_panel(host, host.check_panel)
    refresh_recent_list(host)
    show_frame(host, host.dictation_setup_frame)
    _center_window(host.dictation_window)


def close_window(host):
    host.close_dictation_mode_picker()
    host.close_dictation_answer_review_popup()
    host.close_dictation_volume_popup()
    host._cancel_dictation_timer()
    host._cancel_dictation_feedback_reset()
    if host.dictation_window and host.dictation_window.winfo_exists():
        host.dictation_window.destroy()
    host.dictation_window = None
    host.check_panel = None
    host.dictation_setup_frame = None
    host.dictation_session_frame = None
    host.dictation_result_frame = None
    host.dictation_recent_list = None
    host.dictation_mode_hint_label = None
    host.dictation_input = None
    host.dictation_result_label = None
    host.dictation_progress = None
    host.dictation_answer_review_tree = None
    host.play_btn_check = None
    host.dictation_volume_btn = None
    host.dictation_speed_buttons = []
    host.dictation_order_buttons = []
    host.dictation_feedback_buttons = []
    host.dictation_play_after = None


def refresh_recent_list(host):
    if not host.dictation_recent_list:
        return
    state = build_dictation_list_state(
        words=host.store.words,
        notes=host.store.notes,
        recent_items=host.store.recent_wrong_words(limit=100),
        mode=host.dictation_list_mode_var.get(),
        word_pos=host.word_pos,
        translations=host.translations,
        tr=host.tr,
    )
    host.dictation_all_items = state.all_items
    host.dictation_recent_items = state.recent_items
    host.dictation_all_tab_var.set(state.all_tab_label)
    host.dictation_recent_tab_var.set(state.recent_tab_label)
    host.dictation_recent_list.delete(*host.dictation_recent_list.get_children())
    if state.rows:
        for row_id, values, tag in state.rows:
            host.dictation_recent_list.insert("", tk.END, iid=row_id, values=values, tags=(tag,))
        host.suppress_dictation_select_action = True
        host.dictation_recent_list.selection_set("0")
        host.dictation_recent_list.focus("0")
    else:
        host.dictation_recent_list.insert("", tk.END, iid=state.empty_row[0], values=state.empty_row[1])
    set_list_mode(host, host.dictation_list_mode_var.get(), refresh=False)


def get_source_items(host):
    if host.dictation_list_mode_var.get() == "recent":
        return list(host.dictation_recent_items)
    return list(host.dictation_all_items)


def set_list_mode(host, mode, refresh=True):
    target = "recent" if str(mode or "").strip().lower() == "recent" else "all"
    host.dictation_list_mode_var.set(target)
    if getattr(host, "dictation_all_tab_btn", None):
        host.dictation_all_tab_btn.config(style="SelectedSpeed.TButton" if target == "all" else "Speed.TButton")
    if getattr(host, "dictation_recent_tab_btn", None):
        host.dictation_recent_tab_btn.config(
            style="SelectedSpeed.TButton" if target == "recent" else "Speed.TButton"
        )
    if host.dictation_recent_list:
        host.dictation_recent_list.heading(
            "meta",
            text=(host.tr("error_type") if target == "recent" else host.tr("notes")),
        )
    host.dictation_setup_status_var.set(
        host.tr("dictation_recent_hint") if target == "recent" else host.tr("dictation_all_hint")
    )
    if refresh:
        refresh_recent_list(host)


def show_frame(host, target):
    for frame in (
        host.dictation_setup_frame,
        host.dictation_session_frame,
        host.dictation_result_frame,
    ):
        if not frame:
            continue
        if frame is target:
            frame.grid()
        else:
            frame.grid_remove()


def open_mode_picker(host, auto_start=True):
    has_recent_wrong = bool(host.store.recent_wrong_words(limit=1))
    if not host.store.words and not has_recent_wrong:
        host.show_info("import_words_first")
        return
    if host.dictation_mode_popup and host.dictation_mode_popup.winfo_exists():
        host.dictation_mode_popup.lift()
        return

    parent = host.dictation_window if host.dictation_window and host.dictation_window.winfo_exists() else host
    win = tk.Toplevel(parent)
    host.dictation_mode_popup = win
    win.title(host.tr("mode_picker_title"))
    win.configure(bg="#f6f7fb")
    win.resizable(False, False)
    win.transient(parent)
    win.grab_set()

    wrap = ttk.Frame(win, style="Card.TFrame")
    wrap.pack(fill="both", expand=True, padx=12, pady=12)
    wrap.grid_columnconfigure(0, weight=1)

    ttk.Label(wrap, text=host.tr("mode_picker_title"), style="Card.TLabel").grid(row=0, column=0, sticky="w")

    mode_row = ttk.Frame(wrap, style="Card.TFrame")
    mode_row.grid(row=1, column=0, sticky="ew", pady=(10, 12))
    host.dictation_mode_buttons = []
    for idx, (value, label_key) in enumerate(
        (
            ("quiz", "mode_quiz"),
            ("word_mode", "mode_word_mode"),
            ("answer_review", "mode_answer_review"),
            ("online_spelling", "mode_online_spelling"),
        )
    ):
        btn = ttk.Button(mode_row, text=host.tr(label_key), command=lambda v=value: host.set_dictation_mode(v))
        btn.grid(row=0, column=idx, padx=(0 if idx == 0 else 6, 0), sticky="ew")
        mode_row.grid_columnconfigure(idx, weight=1)
        host.dictation_mode_buttons.append((value, btn))

    options_card = ttk.Frame(wrap, style="Card.TFrame")
    options_card.grid(row=2, column=0, sticky="ew")
    options_card.grid_columnconfigure(0, weight=1)

    ttk.Label(options_card, text=host.tr("playback_speed"), style="Card.TLabel").grid(row=0, column=0, sticky="w")
    speed_row = ttk.Frame(options_card, style="Card.TFrame")
    speed_row.grid(row=1, column=0, sticky="w", pady=(6, 10))
    host.dictation_speed_buttons = []
    for idx, value in enumerate(("1.0", "1.2", "1.4", "1.6", "adaptive")):
        text = host.tr("adaptive_speed") if value == "adaptive" else f"x{value}"
        btn = ttk.Button(speed_row, text=text, command=lambda v=value: host.set_dictation_speed(v))
        btn.grid(row=0, column=idx, padx=(0 if idx == 0 else 6, 0))
        host.dictation_speed_buttons.append((value, btn))

    ttk.Label(options_card, text=host.tr("dictation_order"), style="Card.TLabel").grid(row=2, column=0, sticky="w", pady=(10, 0))
    order_row = ttk.Frame(options_card, style="Card.TFrame")
    order_row.grid(row=3, column=0, sticky="w", pady=(6, 10))
    host.dictation_order_buttons = []
    for idx, (value, text_key) in enumerate((("order", "dictation_order_sequential"), ("random", "dictation_order_random"))):
        btn = ttk.Button(order_row, text=host.tr(text_key), command=lambda v=value: host.set_dictation_order(v))
        btn.grid(row=0, column=idx, padx=(0 if idx == 0 else 6, 0))
        host.dictation_order_buttons.append((value, btn))

    feedback_head = ttk.Frame(options_card, style="Card.TFrame")
    feedback_head.grid(row=4, column=0, sticky="ew", pady=(10, 0))
    feedback_head.grid_columnconfigure(0, weight=1)
    ttk.Label(feedback_head, text=host.tr("feedback"), style="Card.TLabel").grid(row=0, column=0, sticky="w")
    tk.Checkbutton(
        feedback_head,
        text=host.tr("live_feedback"),
        variable=host.dictation_live_feedback_var,
        command=lambda: host.set_dictation_feedback(host.dictation_live_feedback_var.get()),
        bg="#f6f7fb",
        activebackground="#f6f7fb",
        anchor="w",
        relief="flat",
        highlightthickness=0,
        bd=0,
        font=("Segoe UI", 11),
    ).grid(row=0, column=1, sticky="e")

    ttk.Label(options_card, text=host.tr("dictation_feedback_display"), style="Card.TLabel").grid(
        row=5, column=0, sticky="w", pady=(8, 0)
    )
    display_row = ttk.Frame(options_card, style="Card.TFrame")
    display_row.grid(row=6, column=0, sticky="w", pady=(6, 0))
    tk.Checkbutton(
        display_row,
        text=host.tr("dictation_feedback_show_answer"),
        variable=host.dictation_show_answer_var,
        bg="#f6f7fb",
        activebackground="#f6f7fb",
        anchor="w",
        relief="flat",
        highlightthickness=0,
        bd=0,
        font=("Segoe UI", 11),
    ).grid(
        row=0, column=0, sticky="w"
    )
    tk.Checkbutton(
        display_row,
        text=host.tr("dictation_feedback_show_note"),
        variable=host.dictation_show_note_var,
        command=host.update_dictation_feedback_layout,
        bg="#f6f7fb",
        activebackground="#f6f7fb",
        anchor="w",
        relief="flat",
        highlightthickness=0,
        bd=0,
        font=("Segoe UI", 11),
    ).grid(
        row=0, column=1, padx=(10, 0), sticky="w"
    )
    tk.Checkbutton(
        display_row,
        text=host.tr("dictation_feedback_show_phonetic"),
        variable=host.dictation_show_phonetic_var,
        bg="#f6f7fb",
        activebackground="#f6f7fb",
        anchor="w",
        relief="flat",
        highlightthickness=0,
        bd=0,
        font=("Segoe UI", 11),
    ).grid(row=0, column=2, padx=(10, 0), sticky="w")

    duration_row = ttk.Frame(options_card, style="Card.TFrame")
    duration_row.grid(row=7, column=0, sticky="w", pady=(8, 0))
    ttk.Label(duration_row, text=host.tr("dictation_feedback_duration"), style="Card.TLabel").grid(row=0, column=0, sticky="w")
    ttk.Entry(duration_row, textvariable=host.dictation_feedback_seconds_var, width=6).grid(
        row=0, column=1, padx=(8, 6), sticky="w"
    )
    ttk.Label(duration_row, text=host.tr("seconds_short"), style="Card.TLabel", foreground="#667085").grid(
        row=0, column=2, sticky="w"
    )

    host.dictation_mode_hint_label = ttk.Label(
        wrap,
        text="",
        style="Card.TLabel",
        foreground="#667085",
        wraplength=420,
        justify="left",
    )
    host.dictation_mode_hint_label.grid(row=3, column=0, sticky="w", pady=(12, 0))

    btn_row = ttk.Frame(wrap, style="Card.TFrame")
    btn_row.grid(row=4, column=0, sticky="ew", pady=(14, 0))
    btn_row.grid_columnconfigure(0, weight=1)
    btn_row.grid_columnconfigure(1, weight=1)
    ttk.Button(btn_row, text=host.tr("cancel"), command=host.close_dictation_mode_picker).grid(
        row=0,
        column=0,
        padx=(0, 6),
        sticky="ew",
    )
    ttk.Button(
        btn_row,
        text=host.tr("confirm"),
        style="Primary.TButton",
        command=lambda a=auto_start: host.confirm_dictation_mode_picker(auto_start=a),
    ).grid(row=0, column=1, padx=(6, 0), sticky="ew")

    host.set_dictation_mode(host.dictation_mode_var.get())
    host.set_dictation_speed(host.dictation_speed_var.get())
    host.set_dictation_order(host.dictation_order_var.get())
    host.set_dictation_feedback(host.dictation_live_feedback_var.get())
    host.update_dictation_feedback_layout()


def close_mode_picker(host):
    if host.dictation_mode_popup and host.dictation_mode_popup.winfo_exists():
        try:
            host.dictation_mode_popup.grab_release()
        except Exception:
            pass
        host.dictation_mode_popup.destroy()
    host.dictation_mode_popup = None


def set_mode(host, mode):
    selected = str(mode or "online_spelling").strip().lower()
    host.dictation_mode_var.set(selected)
    for value, btn in getattr(host, "dictation_mode_buttons", []):
        btn.config(style="SelectedSpeed.TButton" if value == selected else "Speed.TButton")
    hint = ""
    if selected != "online_spelling":
        hint = host.tr("mode_not_ready")
    if getattr(host, "dictation_mode_hint_label", None):
        host.dictation_mode_hint_label.config(text=hint)


def set_order(host, order_mode):
    selected = "random" if str(order_mode or "").strip().lower() == "random" else "order"
    host.dictation_order_var.set(selected)
    for value, btn in getattr(host, "dictation_order_buttons", []):
        btn.config(style="SelectedSpeed.TButton" if value == selected else "Speed.TButton")


def confirm_mode_picker(host, auto_start=True):
    selected = host.dictation_mode_var.get()
    if selected != "online_spelling":
        host.set_dictation_mode("online_spelling")
    host.close_dictation_mode_picker()
    if auto_start:
        show_frame(host, host.dictation_session_frame)
        host.start_online_spelling_session()


def start_from_selected_word(host):
    if not host.dictation_recent_list:
        return
    selection = host.dictation_recent_list.selection()
    if not selection or selection[0] == "empty":
        host.show_info("select_word_first")
        return
    try:
        start_index = int(selection[0])
    except Exception:
        start_index = 0
    show_frame(host, host.dictation_session_frame)
    host.set_dictation_speed(host.dictation_speed_var.get())
    host.set_dictation_feedback(host.dictation_live_feedback_var.get())
    host.start_online_spelling_session(start_index=start_index)


def get_pool(host):
    items = get_source_items(host)
    if items:
        words = [str(item.get("word") or "").strip() for item in items if str(item.get("word") or "").strip()]
        if words:
            return words
    return list(host.store.words)


def get_preview_source_path(host):
    if host.dictation_list_mode_var.get() == "recent":
        return tts_get_recent_wrong_cache_source()
    return host.store.get_current_source_path()


def on_list_selected(host, _event=None):
    if host.suppress_dictation_select_action:
        host.suppress_dictation_select_action = False
        return
    if not host.dictation_recent_list:
        return
    selection = host.dictation_recent_list.selection()
    if not selection or selection[0] == "empty":
        return
    store_index = host._dictation_row_to_store_index(host.dictation_recent_list, row_id=selection[0])
    selected_word = ""
    try:
        view_index = int(selection[0])
        items = get_source_items(host)
        if 0 <= view_index < len(items):
            selected_word = str(items[view_index].get("word") or "").strip()
    except Exception:
        selected_word = ""
    if store_index is not None and store_index < len(host.store.words):
        host._set_word_action_context(store_index, origin="dictation", word=selected_word)
    else:
        host._set_word_action_context(None, origin="dictation", word=selected_word)
    host._refresh_selection_details()


def on_list_click_play(host, event=None):
    tree = host.dictation_recent_list
    if not tree:
        return "break"
    row_id = str(tree.identify_row(event.y) or "").strip() if event is not None else ""
    if not row_id or row_id == "empty":
        return "break"
    store_index = host._dictation_row_to_store_index(tree, row_id=row_id)
    if store_index is not None and 0 <= store_index < len(host.store.words):
        word = host.store.words[store_index]
        host._set_word_action_context(store_index, origin="dictation")
    else:
        try:
            view_index = int(row_id)
        except Exception:
            return "break"
        items = get_source_items(host)
        if view_index < 0 or view_index >= len(items):
            return "break"
        word = str(items[view_index].get("word") or "").strip()
        if not word:
            return "break"
    speak_preview(host, word=word, store_index=store_index)
    return "break"


def on_review_tree_click(host, event=None):
    tree = getattr(event, "widget", None) if event is not None else None
    if tree is None:
        return "break"
    row_id = str(tree.identify_row(event.y) or "").strip() if event is not None else ""
    if not row_id or row_id == "empty":
        return "break"
    try:
        item = tree.item(row_id)
    except Exception:
        return "break"
    values = item.get("values") or ()
    if not values:
        return "break"
    word_text = str(values[0] or "").splitlines()[0].strip()
    word = re.sub(r"^\d+\.\s*", "", word_text, count=1).strip()
    if not word:
        return "break"
    store_index = None
    try:
        store_index = host.store.words.index(word)
    except Exception:
        store_index = None
    speak_preview(host, word=word, store_index=store_index)
    return "break"


def speak_preview(host, word=None, store_index=None):
    preview_word = str(word or "").strip()
    if not preview_word and store_index is not None and 0 <= store_index < len(host.store.words):
        preview_word = str(host.store.words[store_index] or "").strip()
    if not preview_word:
        return
    source_path = get_preview_source_path(host)
    preview_key = (str(source_path or "").strip(), str(store_index if store_index is not None else preview_word).strip().lower())
    now = time.time()
    if host.last_dictation_preview_key == preview_key and (now - host.last_dictation_preview_at) < 0.35:
        return
    host.last_dictation_preview_key = preview_key
    host.last_dictation_preview_at = now
    runtime = tts_get_runtime_label()
    cached = get_voice_source() == SOURCE_GEMINI and tts_has_cached_word_audio(
        preview_word,
        source_path=source_path,
    )
    token = speak_async(
        preview_word,
        host._dictation_playback_volume_ratio(),
        rate_ratio=host.speech_rate_var.get(),
        cancel_before=True,
        source_path=source_path,
    )
    if cached:
        host.status_var.set(f"Playing cached audio for '{preview_word}'.")
    else:
        host.status_var.set(f"Generating '{preview_word}' with {runtime}...")
    host._watch_tts_backend(token, target="status", text_label=preview_word)


def set_speed(host, value):
    host.dictation_speed_var.set(str(value))
    for current, btn in getattr(host, "dictation_speed_buttons", []):
        btn.config(style="SelectedSpeed.TButton" if current == host.dictation_speed_var.get() else "Speed.TButton")


def set_feedback(host, value):
    enabled = bool(value) if not isinstance(value, str) else str(value).strip().lower() == "live"
    host.dictation_live_feedback_var.set(enabled)
    host.dictation_feedback_var.set("live" if enabled else "none")


def update_feedback_layout(host):
    frame = getattr(host, "dictation_session_frame", None)
    if not frame:
        return
    reserve_two_lines = bool(getattr(host, "dictation_show_note_var", None) and host.dictation_show_note_var.get())
    frame.grid_rowconfigure(3, minsize=68 if reserve_two_lines else 38)


def seconds_for_speed(host):
    mapping = {"1.0": 5, "1.2": 4, "1.4": 3, "1.6": 2, "adaptive": 0}
    return int(mapping.get(host.dictation_speed_var.get(), 5))


def on_volume_change(host, _value=None):
    value = int(host.dictation_volume_var.get())
    if host.dictation_volume_value_label and host.dictation_volume_value_label.winfo_exists():
        host.dictation_volume_value_label.config(text=host.trf("dictation_volume_level", value=value))


def close_volume_popup(host):
    if host.dictation_volume_popup and host.dictation_volume_popup.winfo_exists():
        host.dictation_volume_popup.destroy()
    host.dictation_volume_popup = None
    host.dictation_volume_scale = None
    host.dictation_volume_value_label = None


def toggle_volume_popup(host):
    if not host._dictation_window_active() or not host.dictation_volume_btn:
        return
    if host.dictation_volume_popup and host.dictation_volume_popup.winfo_exists():
        close_volume_popup(host)
        return

    popup = tk.Toplevel(host.dictation_window)
    popup.title(host.tr("dictation_volume"))
    popup.configure(bg="#f6f7fb")
    popup.resizable(False, False)
    popup.transient(host.dictation_window)
    popup.protocol("WM_DELETE_WINDOW", host.close_dictation_volume_popup)
    host.dictation_volume_popup = popup

    wrap = ttk.Frame(popup, style="Card.TFrame", padding=12)
    wrap.pack(fill="both", expand=True)
    ttk.Label(wrap, text=host.tr("dictation_volume"), style="Card.TLabel").pack(anchor="w")
    ttk.Label(
        wrap,
        text=host.tr("dictation_volume_tip"),
        style="Card.TLabel",
        foreground="#667085",
        wraplength=260,
        justify="left",
    ).pack(anchor="w", pady=(4, 10))
    host.dictation_volume_scale = tk.Scale(
        wrap,
        from_=0,
        to=600,
        orient=tk.HORIZONTAL,
        length=260,
        resolution=10,
        showvalue=False,
        variable=host.dictation_volume_var,
        highlightthickness=0,
        command=lambda value=None: on_volume_change(host, value),
    )
    host.dictation_volume_scale.pack(anchor="w")
    host.dictation_volume_value_label = ttk.Label(
        wrap,
        text=host.trf("dictation_volume_level", value=int(host.dictation_volume_var.get())),
        style="Card.TLabel",
    )
    host.dictation_volume_value_label.pack(anchor="w", pady=(6, 0))
    on_volume_change(host)

    try:
        host.dictation_window.update_idletasks()
        popup.update_idletasks()
        x = host.dictation_volume_btn.winfo_rootx() - 210
        y = host.dictation_volume_btn.winfo_rooty() + host.dictation_volume_btn.winfo_height() + 6
        popup.geometry(f"+{max(0, x)}+{max(0, y)}")
    except Exception:
        pass
