# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk


def build_history_tab(host, tooltip_cls):
    host.history_panel = host.history_tab

    hist_title = ttk.Label(host.history_panel, text=host.tr("history"), style="Card.TLabel")
    hist_title.grid(row=0, column=0, padx=10, pady=(10, 4), sticky="w")

    host.history_list = tk.Listbox(
        host.history_panel,
        width=30,
        height=16,
        bg="#ffffff",
        fg="#222222",
        selectbackground="#cce1ff",
        selectforeground="#111111",
        highlightthickness=1,
        highlightbackground="#d9dbe1",
    )
    host.history_list.grid(row=1, column=0, columnspan=2, padx=10, pady=(4, 6))
    host.history_list.bind("<Double-1>", host.on_history_open)
    host.history_list.bind("<Button-3>", host.on_history_right_click)

    host.history_context_menu = tk.Menu(host, tearoff=0)
    host.history_context_menu.add_command(label=host.tr("rename_history_file"), command=host.rename_selected_history_item)
    host.history_context_menu.add_command(label=host.tr("delete_history"), command=host.delete_selected_history_item)

    host.history_empty = ttk.Label(
        host.history_panel, text=host.tr("no_history"), style="Card.TLabel", foreground="#666"
    )
    host.history_empty.grid(row=2, column=0, columnspan=2, padx=10, pady=(0, 6), sticky="w")

    btn_open = ttk.Button(host.history_panel, text="⭳", command=host.on_history_open)
    btn_open.grid(row=3, column=0, padx=10, pady=(0, 10), sticky="w")
    tooltip_cls(btn_open, "Import Selected")

    host.history_path = ttk.Label(host.history_panel, text="", style="Card.TLabel")
    host.history_path.grid(row=3, column=1, padx=10, pady=(0, 10), sticky="e")


def build_tools_tab(host):
    host.tools_tab.grid_columnconfigure(0, weight=1)
    tools_wrap = ttk.Frame(host.tools_tab, style="Card.TFrame")
    tools_wrap.pack(fill="both", expand=True, padx=10, pady=10)
    tools_wrap.grid_columnconfigure(0, weight=1)
    tools_wrap.grid_columnconfigure(1, weight=1)
    ttk.Label(tools_wrap, text=host.tr("learning_tools"), style="Card.TLabel").grid(
        row=0, column=0, columnspan=2, sticky="w"
    )
    host.tools_sentence_btn = ttk.Button(
        tools_wrap,
        text=host.tr("generate_sentence"),
        command=host.make_sentence_for_selected_word,
        takefocus=False,
    )
    host.tools_sentence_btn.grid(row=1, column=0, padx=(0, 6), pady=(8, 4), sticky="ew")
    host.tools_find_btn = ttk.Button(
        tools_wrap,
        text=host.tr("find_corpus_sentences"),
        command=host.open_find_window,
        takefocus=False,
    )
    host.tools_find_btn.grid(row=1, column=1, padx=(6, 0), pady=(8, 4), sticky="ew")
    host.tools_passage_btn = ttk.Button(
        tools_wrap,
        text=host.tr("generate_ielts_passage"),
        command=host.open_passage_window,
        takefocus=False,
    )
    host.tools_passage_btn.grid(row=2, column=0, padx=(0, 6), pady=4, sticky="ew")
    host.tools_settings_btn = ttk.Button(
        tools_wrap,
        text=host.tr("voice_model_settings"),
        command=host.toggle_settings,
        takefocus=False,
    )
    host.tools_settings_btn.grid(row=2, column=1, padx=(6, 0), pady=4, sticky="ew")
    host.tools_update_btn = ttk.Button(
        tools_wrap,
        text=host.tr("update_app"),
        command=host.open_update_dialog,
        takefocus=False,
    )
    host.tools_update_btn.grid(row=3, column=0, padx=(0, 6), pady=(8, 4), sticky="ew")
    host.tools_sync_cache_btn = ttk.Button(
        tools_wrap,
        text=host.tr("sync_shared_cache"),
        command=host.sync_shared_cache_package_online,
        takefocus=False,
    )
    host.tools_sync_cache_btn.grid(row=3, column=1, padx=(6, 0), pady=(8, 4), sticky="ew")
    ttk.Label(tools_wrap, text=host.tr("shared_cache_tools"), style="Card.TLabel").grid(
        row=4, column=0, columnspan=2, sticky="w", pady=(12, 0)
    )
    host.tools_export_cache_btn = ttk.Button(
        tools_wrap,
        text=host.tr("export_shared_cache"),
        command=host.export_shared_cache_package,
        takefocus=False,
    )
    host.tools_export_cache_btn.grid(row=5, column=0, padx=(0, 6), pady=(8, 4), sticky="ew")
    host.tools_import_cache_btn = ttk.Button(
        tools_wrap,
        text=host.tr("import_shared_cache"),
        command=host.import_shared_cache_package,
        takefocus=False,
    )
    host.tools_import_cache_btn.grid(row=5, column=1, padx=(6, 0), pady=(8, 4), sticky="ew")
    next_row = 6
    ttk.Label(tools_wrap, text=host.tr("resource_pack_tools"), style="Card.TLabel").grid(
        row=next_row, column=0, columnspan=2, sticky="w", pady=(12, 0)
    )
    host.tools_export_resource_pack_btn = ttk.Button(
        tools_wrap,
        text=host.tr("export_resource_pack"),
        command=host.export_word_resource_pack_tool,
        takefocus=False,
    )
    host.tools_export_resource_pack_btn.grid(row=next_row + 1, column=0, padx=(0, 6), pady=(8, 4), sticky="ew")
    host.tools_import_resource_pack_btn = ttk.Button(
        tools_wrap,
        text=host.tr("import_resource_pack"),
        command=host.import_word_resource_pack_tool,
        takefocus=False,
    )
    host.tools_import_resource_pack_btn.grid(row=next_row + 1, column=1, padx=(6, 0), pady=(8, 4), sticky="ew")
