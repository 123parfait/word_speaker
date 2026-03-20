# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk


def build_main_shell(host):
    header = ttk.Frame(host)
    header.pack(fill=tk.X, pady=(0, 6))
    title = ttk.Label(header, text="Word Speaker", font=("Segoe UI", 14, "bold"))
    title.pack(side=tk.LEFT)

    ttk.Separator(host, orient="horizontal").pack(fill=tk.X, pady=(0, 10))

    host.main = ttk.Frame(host)
    host.main.pack(fill="both", expand=True)
    host.main.grid_columnconfigure(0, weight=5)
    host.main.grid_columnconfigure(2, weight=4)
    host.main.grid_rowconfigure(0, weight=1)

    host.left = ttk.Frame(host.main, style="Card.TFrame")
    host.left.grid(row=0, column=0, sticky="nsew")
    host.mid_sep = ttk.Separator(host.main, orient="vertical")
    host.mid_sep.grid(row=0, column=1, sticky="ns", padx=10)
    host.right = ttk.Frame(host.main, style="Card.TFrame")
    host.right.grid(row=0, column=2, sticky="nsew")


def build_word_list_panel(host, tooltip_cls):
    left_title = ttk.Label(host.left, text=host.tr("word_list"), style="Card.TLabel")
    left_title.grid(row=0, column=0, columnspan=2, padx=12, pady=(12, 2), sticky="w")
    ttk.Label(
        host.left,
        text=host.tr("word_list_desc"),
        style="Card.TLabel",
        foreground="#667085",
    ).grid(row=1, column=0, columnspan=2, padx=12, pady=(0, 6), sticky="w")

    top_btn_row = ttk.Frame(host.left, style="Card.TFrame")
    top_btn_row.grid(row=2, column=0, columnspan=2, padx=12, pady=6, sticky="ew")
    top_btn_row.grid_columnconfigure(0, weight=1)
    top_btn_row.grid_columnconfigure(1, weight=1)
    top_btn_row.grid_columnconfigure(2, weight=1)
    top_btn_row.grid_columnconfigure(3, weight=1)

    btn_load = ttk.Button(top_btn_row, text=host.tr("import"), style="Primary.TButton", command=host.load_words)
    btn_load.grid(row=0, column=0, padx=(0, 6), sticky="ew")
    tooltip_cls(btn_load, "Import Words")

    btn_manual = ttk.Button(top_btn_row, text=host.tr("paste_type"), command=host.open_manual_words_window)
    btn_manual.grid(row=0, column=1, padx=3, sticky="ew")
    tooltip_cls(btn_manual, "Type/Paste words or a two-column table")

    host.save_as_btn = ttk.Button(top_btn_row, text=host.tr("save_as"), command=host.save_words_as)
    host.save_as_btn.grid(row=0, column=2, padx=3, sticky="ew")
    tooltip_cls(host.save_as_btn, "Save the current list to a txt or csv file")

    host.new_list_btn = ttk.Button(top_btn_row, text=host.tr("new_list"), command=host.new_blank_list)
    host.new_list_btn.grid(row=0, column=3, padx=(6, 0), sticky="ew")
    tooltip_cls(host.new_list_btn, "Create a new empty list")

    table_wrap = ttk.Frame(host.left, style="Card.TFrame")
    table_wrap.grid(row=3, column=0, columnspan=2, padx=12, pady=(4, 6), sticky="nsew")
    host.left.grid_rowconfigure(3, weight=1)
    host.left.grid_columnconfigure(0, weight=1)
    host.left.grid_columnconfigure(1, weight=1)
    host.word_table = ttk.Treeview(
        table_wrap,
        columns=("idx", "word", "note"),
        show="headings",
        height=18,
        selectmode="extended",
        style="WordList.Treeview",
    )
    host.word_table.heading("idx", text="#")
    host.word_table.heading("word", text=host.tr("word"))
    host.word_table.heading("note", text=host.tr("notes"))
    host.word_table.column("idx", width=70, anchor="center", stretch=False)
    host.word_table.column("word", width=500, anchor="w")
    host.word_table.column("note", width=240, anchor="w")

    host.word_table_scroll = ttk.Scrollbar(table_wrap, orient="vertical", command=host.word_table.yview)
    host.word_table.configure(yscrollcommand=host.word_table_scroll.set)
    host.word_table.grid(row=0, column=0, sticky="nsew")
    host.word_table_scroll.grid(row=0, column=1, sticky="ns")
    host.word_table.tag_configure("even", background="#ffffff")
    host.word_table.tag_configure("odd", background="#fbfcfe")
    table_wrap.grid_rowconfigure(0, weight=1)
    table_wrap.grid_columnconfigure(0, weight=1)
    host.word_table.bind("<<TreeviewSelect>>", host.on_word_selected)
    host.word_table.bind("<Double-1>", host.on_word_double_click)
    host.word_table.bind("<Button-3>", host.on_word_right_click)
    host.word_table.bind("<F2>", host.start_edit_selected_word)
    host.word_table.bind("<Escape>", host.cancel_word_edit)
    host.word_table.bind("<Control-v>", host.on_word_table_paste)
    host.word_table.bind("<Control-V>", host.on_word_table_paste)

    host.word_context_menu = tk.Menu(host, tearoff=0)
    host.word_context_menu.add_command(label=host.tr("add_word"), command=host.prompt_add_word)
    host.word_context_menu.add_command(label=host.tr("edit_word"), command=host.start_edit_selected_word)
    host.word_context_menu.add_command(label=host.tr("edit_note"), command=host.start_edit_selected_note)
    host.word_context_menu.add_command(label=host.tr("edit_pos_translation"), command=host.edit_selected_word_meta)
    host.word_context_menu.add_command(label="Find", command=host.search_selected_word_in_corpus)
    host.word_context_menu.add_command(label=host.tr("generate_sentence"), command=host.make_sentence_for_selected_word)
    host.word_context_menu.add_command(label=host.tr("lookup_synonyms"), command=host.lookup_synonyms_for_selected_word)
    host.word_context_menu.add_command(label=host.tr("inspect_audio_cache"), command=host.inspect_selected_word_audio_cache)
    host.word_context_menu.add_command(label=host.tr("replace_audio_with_piper"), command=host.replace_selected_word_audio_with_piper)
    host.word_context_menu.add_command(label=host.tr("clear_word_audio_override"), command=host.clear_selected_word_audio_override)
    host.word_context_menu.add_separator()
    host.word_context_menu.add_command(label=host.tr("delete_word"), command=host.delete_selected_word)

    host.dictation_context_menu = tk.Menu(host, tearoff=0)
    host.dictation_context_menu.add_command(label=host.tr("add_word"), command=host.prompt_add_word)
    host.dictation_context_menu.add_command(label=host.tr("edit_word"), command=host.prompt_edit_context_word)
    host.dictation_context_menu.add_command(label=host.tr("edit_note"), command=host.prompt_edit_context_note)
    host.dictation_context_menu.add_command(label=host.tr("edit_pos_translation"), command=host.edit_selected_word_meta)
    host.dictation_context_menu.add_command(label="Find", command=host.search_selected_word_in_corpus)
    host.dictation_context_menu.add_command(label=host.tr("generate_sentence"), command=host.make_sentence_for_selected_word)
    host.dictation_context_menu.add_command(label=host.tr("lookup_synonyms"), command=host.lookup_synonyms_for_selected_word)
    host.dictation_context_menu.add_command(label=host.tr("inspect_audio_cache"), command=host.inspect_selected_word_audio_cache)
    host.dictation_context_menu.add_command(label=host.tr("replace_audio_with_piper"), command=host.replace_selected_word_audio_with_piper)
    host.dictation_context_menu.add_command(label=host.tr("clear_word_audio_override"), command=host.clear_selected_word_audio_override)
    host.dictation_context_menu.add_separator()
    host.dictation_context_menu.add_command(label=host.tr("delete_word"), command=host.delete_selected_word)

    host.empty_label = ttk.Label(
        host.left,
        text=host.tr("no_words"),
        style="Card.TLabel",
        foreground="#666",
    )
    host.empty_label.grid(row=4, column=0, columnspan=2, padx=12, pady=(0, 8), sticky="w")

    host.player_frame = ttk.Frame(host.left, style="Card.TFrame")
    host.player_frame.grid(row=5, column=0, columnspan=2, padx=12, pady=(4, 6), sticky="ew")
    host.player_frame.grid_columnconfigure(0, weight=1)
    host.player_frame.grid_columnconfigure(1, weight=1)
    host.player_frame.grid_columnconfigure(2, weight=1)

    host.play_btn = ttk.Button(
        host.player_frame, text=host.tr("play"), style="Icon.TButton", command=host.toggle_play
    )
    host.play_btn.grid(row=0, column=0, padx=4, sticky="ew")
    tooltip_cls(host.play_btn, "Start / Pause")

    host.settings_btn = ttk.Button(
        host.player_frame, text=host.tr("settings"), style="Icon.TButton", command=host.toggle_settings
    )
    host.settings_btn.grid(row=0, column=1, padx=4, sticky="ew")
    tooltip_cls(host.settings_btn, "Settings")

    host.dictation_btn = ttk.Button(
        host.player_frame, text=host.tr("dictation"), style="Icon.TButton", command=host.open_dictation_window
    )
    host.dictation_btn.grid(row=0, column=2, padx=4, sticky="ew")
    tooltip_cls(host.dictation_btn, "Open Dictation Window")

    host.status_label = ttk.Label(
        host.left, textvariable=host.status_var, style="Card.TLabel", foreground="#444"
    )
    host.status_label.grid(row=6, column=0, columnspan=2, padx=12, pady=(0, 12), sticky="w")
