# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk


def build_api_key_window(host, *, initial_section="llm"):
    win = tk.Toplevel(host)
    host.api_key_window = win
    win.title(host.tr("api_setup"))
    win.configure(bg="#f6f7fb")
    win.resizable(False, False)
    win.transient(host.winfo_toplevel())

    wrap = ttk.Frame(win, style="Card.TFrame")
    wrap.pack(fill="both", expand=True, padx=12, pady=12)

    ttk.Label(wrap, text=host.tr("api_setup"), style="Card.TLabel").pack(anchor="w")
    ttk.Label(
        wrap,
        text=host.tr("api_setup_desc"),
        style="Card.TLabel",
        foreground="#666",
    ).pack(anchor="w", pady=(0, 10))

    llm_section = ttk.Frame(wrap, style="Card.TFrame")
    llm_section.pack(fill="x")
    ttk.Label(llm_section, text=host.tr("llm_api_setup"), style="Card.TLabel").pack(anchor="w")
    ttk.Label(
        llm_section,
        text=host.tr("llm_key_desc"),
        style="Card.TLabel",
        foreground="#666",
    ).pack(anchor="w", pady=(0, 8))
    llm_provider_row = ttk.Frame(llm_section, style="Card.TFrame")
    llm_provider_row.pack(anchor="w", pady=(0, 8), fill="x")
    ttk.Label(llm_provider_row, text=f"{host.tr('api_provider')}:", style="Card.TLabel").pack(side=tk.LEFT)
    llm_provider_combo = ttk.Combobox(
        llm_provider_row,
        textvariable=host.llm_api_provider_var,
        values=[host.tr("provider_gemini")],
        state="readonly",
        width=18,
    )
    llm_provider_combo.pack(side=tk.LEFT, padx=(6, 0))
    llm_provider_combo.bind("<<ComboboxSelected>>", lambda _e: host._on_llm_provider_selected())

    llm_entry = tk.Entry(
        llm_section,
        textvariable=host.gemini_key_var,
        width=54,
        show="*",
        relief="solid",
        bd=1,
        highlightthickness=1,
        highlightbackground="#cbd5e1",
        highlightcolor="#2563eb",
        bg="white",
    )
    llm_entry.pack(fill="x")
    llm_entry.icursor(tk.END)
    llm_entry.bind("<Return>", lambda _event: host.test_and_save_api_keys())
    llm_entry.bind("<KeyRelease>", lambda _event: host._set_api_entry_error("llm", False))

    ttk.Label(
        llm_section,
        text=host.tr("gemini_model_desc"),
        style="Card.TLabel",
        foreground="#666",
    ).pack(anchor="w", pady=(8, 4))
    combo = ttk.Combobox(
        llm_section,
        textvariable=host.gemini_model_var,
        values=host.gemini_model_values or host.list_gemini_models(),
        state="readonly",
        width=24,
    )
    combo.pack(anchor="w")
    combo.bind("<<ComboboxSelected>>", host.on_gemini_model_change)
    ttk.Label(llm_section, textvariable=host.gemini_key_status_var, style="Card.TLabel", foreground="#444").pack(
        anchor="w", pady=(10, 0)
    )
    ttk.Separator(wrap, orient="horizontal").pack(fill="x", pady=12)

    tts_section = ttk.Frame(wrap, style="Card.TFrame")
    tts_section.pack(fill="x")
    ttk.Label(tts_section, text=host.tr("tts_api_setup"), style="Card.TLabel").pack(anchor="w")
    ttk.Label(
        tts_section,
        text=host.tr("tts_key_desc"),
        style="Card.TLabel",
        foreground="#666",
    ).pack(anchor="w", pady=(0, 8))
    tts_provider_row = ttk.Frame(tts_section, style="Card.TFrame")
    tts_provider_row.pack(anchor="w", pady=(0, 8), fill="x")
    ttk.Label(tts_provider_row, text=f"{host.tr('api_provider')}:", style="Card.TLabel").pack(side=tk.LEFT)
    tts_provider_combo = ttk.Combobox(
        tts_provider_row,
        textvariable=host.tts_api_provider_var,
        values=list(host._tts_provider_options().keys()),
        state="readonly",
        width=18,
    )
    tts_provider_combo.pack(side=tk.LEFT, padx=(6, 0))
    tts_provider_combo.bind("<<ComboboxSelected>>", host._on_tts_provider_selected)

    tts_entry = tk.Entry(
        tts_section,
        textvariable=host.tts_key_var,
        width=54,
        show="*",
        relief="solid",
        bd=1,
        highlightthickness=1,
        highlightbackground="#cbd5e1",
        highlightcolor="#2563eb",
        bg="white",
    )
    tts_entry.pack(fill="x")
    tts_entry.icursor(tk.END)
    tts_entry.bind("<Return>", lambda _event: host.test_and_save_api_keys())
    tts_entry.bind("<KeyRelease>", lambda _event: host._set_api_entry_error("tts", False))

    ttk.Label(tts_section, textvariable=host.tts_key_status_var, style="Card.TLabel", foreground="#444").pack(
        anchor="w", pady=(10, 0)
    )
    footer = ttk.Frame(wrap, style="Card.TFrame")
    footer.pack(fill="x", pady=(12, 0))
    api_key_test_btn = ttk.Button(footer, text=host.tr("test_and_save"), command=host.test_and_save_api_keys)
    api_key_test_btn.pack(side=tk.LEFT)
    ttk.Button(footer, text=host.tr("close"), command=host._close_api_key_window).pack(side=tk.RIGHT)

    host.api_key_test_btn = api_key_test_btn
    host.api_llm_entry = llm_entry
    host.api_tts_entry = tts_entry
    host.gemini_model_combo = combo
    host._set_api_entry_error("llm", False)
    host._set_api_entry_error("tts", False)

    if initial_section == "tts":
        tts_entry.focus_set()
    else:
        llm_entry.focus_set()

    win.grab_set()
    win.protocol("WM_DELETE_WINDOW", host._close_api_key_window)
    return win
