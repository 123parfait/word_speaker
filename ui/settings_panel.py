# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk


def build_settings_window(host):
    host._sync_provider_vars()
    win = tk.Toplevel(host)
    host.settings_window = win
    win.title(host.tr("settings_title"))
    win.configure(bg="#f6f7fb")
    win.resizable(False, False)
    win.protocol("WM_DELETE_WINDOW", host._close_settings_window)

    container = ttk.Frame(win, style="Card.TFrame")
    container.pack(padx=10, pady=10)

    left_menu = ttk.Frame(container, style="Card.TFrame", width=120)
    left_menu.grid(row=0, column=0, sticky="n")
    right_panel = ttk.Frame(container, style="Card.TFrame", width=360, height=260)
    right_panel.grid(row=0, column=1, padx=(10, 0), sticky="n")
    right_panel.grid_propagate(False)

    host.settings_sections_visible = {
        "source": True,
        "speed": True,
        "language": False,
    }
    sections = []

    def rebuild_sections():
        for item in sections:
            item["frame"].pack_forget()
            item["sep"].pack_forget()

        visible_keys = [k for k, v in host.settings_sections_visible.items() if v]
        if not visible_keys:
            right_panel.grid_remove()
            return
        right_panel.grid()

        for idx, item in enumerate(sections):
            key = item["key"]
            if not host.settings_sections_visible.get(key, False):
                continue
            item["frame"].pack(fill=tk.X, pady=(0, 6))
            has_next = False
            for later in sections[idx + 1 :]:
                if host.settings_sections_visible.get(later["key"], False):
                    has_next = True
                    break
            if has_next:
                item["sep"].pack(fill=tk.X, pady=(0, 6))

    def toggle_section(key):
        host.settings_sections_visible[key] = not host.settings_sections_visible.get(key, False)
        rebuild_sections()

    for label, key in [
        (host.tr("settings_toggle_source"), "source"),
        (host.tr("settings_toggle_speed"), "speed"),
        (host.tr("ui_language"), "language"),
    ]:
        btn = ttk.Button(left_menu, text=label, command=lambda k=key: toggle_section(k))
        btn.pack(fill=tk.X, pady=4)

    source_section = ttk.Frame(right_panel, style="Card.TFrame")
    ttk.Label(source_section, text=host.tr("source"), style="Card.TLabel").pack(anchor="w")

    host.voice_combo = ttk.Combobox(
        source_section,
        textvariable=host.voice_var,
        state="readonly",
        width=32,
    )
    host.voice_combo.pack(anchor="w")
    host.voice_combo.bind("<<ComboboxSelected>>", host.on_voice_change)

    llm_row = ttk.Frame(source_section, style="Card.TFrame")
    llm_row.pack(anchor="w", pady=(8, 0), fill="x")
    ttk.Label(llm_row, text=f"{host.tr('llm_api')}:", style="Card.TLabel").pack(side=tk.LEFT)
    ttk.Combobox(
        llm_row,
        textvariable=host.llm_api_provider_var,
        values=[host.tr("provider_gemini")],
        state="readonly",
        width=12,
    ).pack(side=tk.LEFT, padx=(6, 6))
    ttk.Button(llm_row, text=host.tr("llm_api"), command=host.open_gemini_key_window).pack(side=tk.LEFT)

    tts_row = ttk.Frame(source_section, style="Card.TFrame")
    tts_row.pack(anchor="w", pady=(8, 0), fill="x")
    ttk.Label(tts_row, text=f"{host.tr('tts_api')}:", style="Card.TLabel").pack(side=tk.LEFT)
    tts_provider_combo = ttk.Combobox(
        tts_row,
        textvariable=host.tts_api_provider_var,
        values=list(host._tts_provider_options().keys()),
        state="readonly",
        width=12,
    )
    tts_provider_combo.pack(side=tk.LEFT, padx=(6, 6))
    tts_provider_combo.bind("<<ComboboxSelected>>", host._on_tts_provider_selected)
    ttk.Button(tts_row, text=host.tr("tts_api"), command=host.open_tts_key_window).pack(side=tk.LEFT)
    ttk.Label(
        source_section,
        textvariable=host.gemini_runtime_status_var,
        style="Card.TLabel",
        foreground="#4b5563",
    ).pack(anchor="w", pady=(8, 0))
    ttk.Label(
        source_section,
        textvariable=host.gemini_retry_status_var,
        style="Card.TLabel",
        foreground="#667085",
    ).pack(anchor="w", pady=(2, 0))
    sections.append({"key": "source", "frame": source_section, "sep": ttk.Separator(right_panel, orient="horizontal")})

    speed_section = ttk.Frame(right_panel, style="Card.TFrame")
    ttk.Label(speed_section, text=host.tr("speed"), style="Card.TLabel").pack(anchor="w")
    ttk.Label(
        speed_section,
        text=host.tr("speed_desc"),
        style="Card.TLabel",
        foreground="#666",
    ).pack(anchor="w", pady=(0, 4))
    host.speed_buttons = []
    speed_row = ttk.Frame(speed_section, style="Card.TFrame")
    speed_row.pack(anchor="w")
    for v in [1, 2, 3, 5, 10]:
        btn = ttk.Button(speed_row, text=f"{v}s", command=lambda val=v: host.set_interval(val))
        btn.pack(side=tk.LEFT, padx=3)
        host.speed_buttons.append((v, btn))

    custom_row = ttk.Frame(speed_section, style="Card.TFrame")
    custom_row.pack(anchor="w", pady=(4, 0))
    ttk.Label(custom_row, text=host.tr("custom_seconds"), style="Card.TLabel").pack(side=tk.LEFT)
    host.custom_interval = ttk.Entry(custom_row, width=6)
    host.custom_interval.pack(side=tk.LEFT, padx=4)
    host.custom_interval.bind("<Return>", lambda _e: host.apply_custom_interval())
    ttk.Button(custom_row, text=host.tr("apply"), command=host.apply_custom_interval).pack(side=tk.LEFT)

    ttk.Label(
        speed_section,
        text=host.tr("pronunciation_speed"),
        style="Card.TLabel",
        foreground="#666",
    ).pack(anchor="w", pady=(8, 4))
    host.speech_rate_buttons = []
    speech_row = ttk.Frame(speed_section, style="Card.TFrame")
    speech_row.pack(anchor="w")
    for v in [0.6, 0.8, 1.0, 1.2]:
        btn = ttk.Button(speech_row, text=f"{v:.1f}x", command=lambda val=v: host.set_speech_rate(val))
        btn.pack(side=tk.LEFT, padx=3)
        host.speech_rate_buttons.append((v, btn))
    sections.append({"key": "speed", "frame": speed_section, "sep": ttk.Separator(right_panel, orient="horizontal")})

    language_section = ttk.Frame(right_panel, style="Card.TFrame")
    ttk.Label(language_section, text=host.tr("ui_language"), style="Card.TLabel").pack(anchor="w")
    language_combo = ttk.Combobox(
        language_section,
        textvariable=host.ui_language_var,
        state="readonly",
        width=18,
        values=("zh", "en"),
    )
    language_combo.pack(anchor="w", pady=(4, 0))
    language_combo.bind("<<ComboboxSelected>>", host.on_ui_language_change)
    ttk.Label(
        language_section,
        text=f"zh = {host.tr('language_zh')}   |   en = {host.tr('language_en')}",
        style="Card.TLabel",
        foreground="#666",
    ).pack(anchor="w", pady=(6, 0))
    sections.append({"key": "language", "frame": language_section, "sep": ttk.Separator(right_panel, orient="horizontal")})

    rebuild_sections()
    return win
