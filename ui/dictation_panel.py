# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk


def build_dictation_panel(host, parent):
    host.dictation_setup_frame = ttk.Frame(parent, style="Card.TFrame")
    host.dictation_setup_frame.grid(row=0, column=0, sticky="nsew")
    host.dictation_setup_frame.grid_columnconfigure(0, weight=1)
    host.dictation_setup_frame.grid_rowconfigure(2, weight=1)
    tab_row = ttk.Frame(host.dictation_setup_frame, style="Card.TFrame")
    tab_row.grid(row=0, column=0, sticky="ew", pady=(0, 8))
    tab_row.grid_columnconfigure(0, weight=1)
    tab_row.grid_columnconfigure(1, weight=1)
    host.dictation_all_tab_btn = ttk.Button(
        tab_row,
        textvariable=host.dictation_all_tab_var,
        command=lambda: host.set_dictation_list_mode("all"),
    )
    host.dictation_all_tab_btn.grid(row=0, column=0, padx=(0, 6), sticky="ew")
    host.dictation_recent_tab_btn = ttk.Button(
        tab_row,
        textvariable=host.dictation_recent_tab_var,
        command=lambda: host.set_dictation_list_mode("recent"),
    )
    host.dictation_recent_tab_btn.grid(row=0, column=1, padx=(6, 0), sticky="ew")

    ttk.Label(
        host.dictation_setup_frame,
        textvariable=host.dictation_status_var,
        style="Card.TLabel",
        foreground="#667085",
        wraplength=640,
        justify="left",
    ).grid(row=1, column=0, sticky="w", pady=(0, 8))

    recent_wrap = ttk.Frame(host.dictation_setup_frame, style="Card.TFrame")
    recent_wrap.grid(row=2, column=0, sticky="nsew")
    recent_wrap.grid_columnconfigure(0, weight=1)
    recent_wrap.grid_rowconfigure(0, weight=1)
    host.dictation_recent_list = ttk.Treeview(
        recent_wrap,
        columns=("idx", "word", "meta"),
        show="headings",
        selectmode="browse",
        height=14,
        style="WordList.Treeview",
    )
    host.dictation_recent_list.heading("idx", text="#")
    host.dictation_recent_list.heading("word", text=host.tr("word"))
    host.dictation_recent_list.heading("meta", text=host.tr("notes"))
    host.dictation_recent_list.column("idx", width=60, anchor="center", stretch=False)
    host.dictation_recent_list.column("word", width=360, minwidth=220, anchor="w", stretch=False)
    host.dictation_recent_list.column("meta", width=300, minwidth=160, anchor="w", stretch=True)
    host.dictation_recent_list.grid(row=0, column=0, sticky="nsew")
    host.dictation_recent_list.tag_configure("even", background="#ffffff")
    host.dictation_recent_list.tag_configure("odd", background="#fbfcfe")
    host.dictation_recent_list.bind("<<TreeviewSelect>>", host.on_dictation_list_selected)
    host.dictation_recent_list.bind("<ButtonRelease-1>", host.on_dictation_list_click_play)
    recent_scroll = ttk.Scrollbar(recent_wrap, orient="vertical", command=host.dictation_recent_list.yview)
    recent_scroll.grid(row=0, column=1, sticky="ns")
    host.dictation_recent_list.configure(yscrollcommand=recent_scroll.set)
    host.dictation_recent_list.bind("<Button-3>", host.on_dictation_word_right_click)

    setup_row = ttk.Frame(host.dictation_setup_frame, style="Card.TFrame")
    setup_row.grid(row=3, column=0, sticky="ew", pady=(10, 0))
    setup_row.grid_columnconfigure(0, weight=1)
    setup_row.grid_columnconfigure(1, weight=1)
    ttk.Button(setup_row, text=host.tr("start_from_word"), command=host.start_dictation_from_selected_word).grid(
        row=0, column=0, padx=(0, 6), sticky="ew"
    )
    ttk.Button(setup_row, text=host.tr("start_learning"), command=host.open_dictation_mode_picker).grid(
        row=0, column=1, padx=(6, 0), sticky="ew"
    )

    host.dictation_session_frame = ttk.Frame(parent, style="Card.TFrame")
    host.dictation_session_frame.grid(row=0, column=0, sticky="nsew")
    host.dictation_session_frame.grid_remove()
    host.dictation_session_frame.grid_columnconfigure(0, weight=1)

    session_header = ttk.Frame(host.dictation_session_frame, style="Card.TFrame")
    session_header.grid(row=0, column=0, sticky="ew")
    session_header.grid_columnconfigure(0, weight=1)
    session_header.grid_columnconfigure(1, weight=0)
    session_header.grid_columnconfigure(2, weight=0)
    ttk.Label(session_header, textvariable=host.dictation_progress_var, style="Card.TLabel").grid(
        row=0, column=0, sticky="w"
    )
    ttk.Label(session_header, text=host.tr("answer"), style="Card.TLabel", foreground="#667085").grid(
        row=0, column=1, sticky="e", padx=(0, 8)
    )
    host.dictation_volume_btn = ttk.Button(
        session_header,
        text=host.tr("dictation_volume_button"),
        width=3,
        command=host.toggle_dictation_volume_popup,
    )
    host.dictation_volume_btn.grid(row=0, column=2, sticky="e")

    host.dictation_progress = ttk.Progressbar(
        host.dictation_session_frame,
        orient="horizontal",
        mode="determinate",
        maximum=100,
    )
    host.dictation_progress.grid(row=1, column=0, sticky="ew", pady=(8, 10))

    input_card = ttk.Frame(host.dictation_session_frame, style="Card.TFrame")
    input_card.grid(row=2, column=0, sticky="ew", pady=(12, 0))
    input_card.grid_columnconfigure(0, weight=1)
    input_card.grid_columnconfigure(1, weight=0, minsize=78)
    host.dictation_input = tk.Entry(
        input_card,
        font=("Segoe UI", 24, "bold"),
        relief="solid",
        bd=1,
        bg="#f6f6f8",
        fg="#202020",
        insertbackground="#202020",
    )
    host.dictation_input.grid(row=0, column=0, sticky="ew", padx=(0, 8), ipady=20)
    host.dictation_input.bind("<KeyRelease>", host.on_dictation_input_change)
    host.dictation_input.bind("<Return>", host.on_dictation_enter)
    host.dictation_input.bind("<FocusIn>", lambda _e: host.close_dictation_volume_popup())
    host.dictation_input.bind("<Button-1>", lambda _e: host.close_dictation_volume_popup())
    timer_wrap = ttk.Frame(input_card, style="Card.TFrame")
    timer_wrap.grid(row=0, column=1, sticky="ne")
    ttk.Label(
        timer_wrap,
        textvariable=host.dictation_timer_var,
        font=("Segoe UI", 18),
        style="Card.TLabel",
        foreground="#6b7280",
        width=6,
        anchor="e",
        justify="right",
    ).grid(row=0, column=0, sticky="e", pady=(4, 0))

    host.dictation_result_label = ttk.Label(
        host.dictation_session_frame,
        textvariable=host.dictation_status_var,
        style="Card.TLabel",
        foreground="#667085",
    )
    host.dictation_result_label.grid(row=3, column=0, sticky="w", pady=(8, 10))

    ttk.Button(
        host.dictation_session_frame,
        text=host.tr("next_word"),
        style="Primary.TButton",
        command=host.advance_dictation_word,
    ).grid(row=4, column=0, sticky="ew")

    control_row = ttk.Frame(host.dictation_session_frame, style="Card.TFrame")
    control_row.grid(row=5, column=0, sticky="ew", pady=(14, 0))
    control_row.grid_columnconfigure(0, weight=1)
    control_row.grid_columnconfigure(1, weight=1)
    control_row.grid_columnconfigure(2, weight=1)
    control_row.grid_columnconfigure(3, weight=1)
    ttk.Button(
        control_row,
        text=host.tr("dictation_settings"),
        command=lambda: host.open_dictation_mode_picker(auto_start=False),
    ).grid(row=0, column=0, padx=(0, 6), sticky="ew")
    ttk.Button(control_row, text=host.tr("previous_word"), command=host.previous_dictation_word).grid(
        row=0, column=1, padx=3, sticky="ew"
    )
    host.play_btn_check = ttk.Button(control_row, text=host.tr("play"), command=host.toggle_dictation_play_pause)
    host.play_btn_check.grid(row=0, column=2, padx=3, sticky="ew")
    ttk.Button(control_row, text=host.tr("replay"), command=host.replay_dictation_word).grid(
        row=0, column=3, padx=(6, 0), sticky="ew"
    )
    ttk.Button(
        host.dictation_session_frame,
        text=host.tr("answer_review"),
        command=host.open_dictation_answer_review_popup,
    ).grid(row=6, column=0, sticky="ew", pady=(10, 0))

    host.dictation_result_frame = ttk.Frame(parent, style="Card.TFrame")
    host.dictation_result_frame.grid(row=0, column=0, sticky="nsew")
    host.dictation_result_frame.grid_remove()
    host.dictation_result_frame.grid_columnconfigure(0, weight=1)
    host.dictation_result_accuracy_var = tk.StringVar(value="0.00%")
    host.dictation_result_last_var = tk.StringVar(value="-")
    host.dictation_result_filter_var = tk.StringVar(value=host.tr("show_wrong_only"))

    ttk.Label(
        host.dictation_result_frame,
        textvariable=host.dictation_result_accuracy_var,
        font=("Segoe UI", 28, "bold"),
        style="Card.TLabel",
        foreground="#5b5cf0",
    ).grid(row=0, column=0, sticky="n", pady=(18, 2))
    ttk.Label(host.dictation_result_frame, text=host.tr("answer_review_so_far"), style="Card.TLabel").grid(
        row=1, column=0, sticky="n"
    )
    result_previous_row = ttk.Frame(host.dictation_result_frame, style="Card.TFrame")
    result_previous_row.grid(row=2, column=0, sticky="n", pady=(8, 12))
    ttk.Label(
        result_previous_row,
        text=f"{host.tr('last_session_accuracy')}: ",
        style="Card.TLabel",
        foreground="#667085",
    ).grid(row=0, column=0, sticky="e")
    ttk.Label(
        result_previous_row,
        textvariable=host.dictation_result_last_var,
        style="Card.TLabel",
        foreground="#667085",
    ).grid(row=0, column=1, sticky="w")

    result_action_row = ttk.Frame(host.dictation_result_frame, style="Card.TFrame")
    result_action_row.grid(row=3, column=0, sticky="ew", pady=(0, 10))
    result_action_row.grid_columnconfigure(0, weight=1)
    result_action_row.grid_columnconfigure(1, weight=1)
    ttk.Button(
        result_action_row,
        textvariable=host.dictation_result_filter_var,
        style="Primary.TButton",
        command=host._toggle_dictation_answer_review_filter,
    ).grid(row=0, column=0, sticky="ew", padx=(80, 10))
    ttk.Button(
        result_action_row,
        text=host.tr("back_to_list"),
        command=host.reset_dictation_view,
    ).grid(row=0, column=1, sticky="ew", padx=(10, 80))

    result_table_wrap = ttk.Frame(host.dictation_result_frame, style="Card.TFrame")
    result_table_wrap.grid(row=4, column=0, sticky="nsew")
    result_table_wrap.grid_columnconfigure(0, weight=1)
    result_table_wrap.grid_rowconfigure(0, weight=1)
    host.dictation_result_frame.grid_rowconfigure(4, weight=1)

    result_tree = ttk.Treeview(
        result_table_wrap,
        columns=("word", "input", "count"),
        show="headings",
        style="WordList.Treeview",
        height=12,
    )
    result_tree.heading("word", text=host.tr("word"))
    result_tree.heading("input", text=host.tr("your_answer"))
    result_tree.heading("count", text=host.tr("wrong_times"))
    result_tree.column("word", width=390, minwidth=240, anchor="w", stretch=True)
    result_tree.column("input", width=220, minwidth=150, anchor="w", stretch=True)
    result_tree.column("count", width=110, minwidth=90, anchor="center", stretch=False)
    result_tree.tag_configure("correct", foreground="#15803d")
    result_tree.tag_configure("wrong", foreground="#dc2626")
    result_tree.grid(row=0, column=0, sticky="nsew")
    result_scroll = ttk.Scrollbar(result_table_wrap, orient="vertical", command=result_tree.yview)
    result_scroll.grid(row=0, column=1, sticky="ns")
    result_tree.configure(yscrollcommand=result_scroll.set)
    host.dictation_result_review_tree = result_tree


def build_dictation_answer_review_popup(host):
    popup = tk.Toplevel(host.dictation_window or host)
    popup.title(host.tr("answer_review_title"))
    popup.configure(bg="#f6f7fb")
    popup.geometry("760x560")
    popup.minsize(680, 460)
    popup.protocol("WM_DELETE_WINDOW", host.close_dictation_answer_review_popup)
    host.dictation_answer_review_popup = popup

    host.dictation_answer_review_accuracy_var = tk.StringVar(value="0.00%")
    host.dictation_answer_review_last_var = tk.StringVar(value="-")
    host.dictation_answer_review_filter_var = tk.StringVar(value=host.tr("show_wrong_only"))

    wrap = ttk.Frame(popup, style="Card.TFrame")
    wrap.pack(fill="both", expand=True, padx=14, pady=14)
    wrap.grid_columnconfigure(0, weight=1)
    wrap.grid_rowconfigure(4, weight=1)

    ttk.Label(
        wrap,
        textvariable=host.dictation_answer_review_accuracy_var,
        font=("Segoe UI", 28, "bold"),
        style="Card.TLabel",
        foreground="#5b5cf0",
    ).grid(row=0, column=0, sticky="n", pady=(8, 2))
    ttk.Label(wrap, text=host.tr("answer_review_so_far"), style="Card.TLabel").grid(row=1, column=0, sticky="n")
    previous_row = ttk.Frame(wrap, style="Card.TFrame")
    previous_row.grid(row=2, column=0, sticky="n", pady=(8, 12))
    ttk.Label(
        previous_row,
        text=f"{host.tr('last_session_accuracy')}: ",
        style="Card.TLabel",
        foreground="#667085",
    ).grid(row=0, column=0, sticky="e")
    ttk.Label(
        previous_row,
        textvariable=host.dictation_answer_review_last_var,
        style="Card.TLabel",
        foreground="#667085",
    ).grid(row=0, column=1, sticky="w")

    action_row = ttk.Frame(wrap, style="Card.TFrame")
    action_row.grid(row=3, column=0, sticky="ew", pady=(0, 10))
    action_row.grid_columnconfigure(0, weight=1)
    action_row.grid_columnconfigure(1, weight=1)
    ttk.Button(
        action_row,
        textvariable=host.dictation_answer_review_filter_var,
        style="Primary.TButton",
        command=host._toggle_dictation_answer_review_filter,
    ).grid(row=0, column=0, sticky="ew", padx=(80, 10))
    ttk.Button(
        action_row,
        text=host.tr("back_to_list"),
        command=host._return_from_dictation_answer_review,
    ).grid(row=0, column=1, sticky="ew", padx=(10, 80))

    table_wrap = ttk.Frame(wrap, style="Card.TFrame")
    table_wrap.grid(row=4, column=0, sticky="nsew")
    table_wrap.grid_columnconfigure(0, weight=1)
    table_wrap.grid_rowconfigure(0, weight=1)

    tree = ttk.Treeview(
        table_wrap,
        columns=("word", "input", "count"),
        show="headings",
        style="WordList.Treeview",
        height=12,
    )
    tree.heading("word", text=host.tr("word"))
    tree.heading("input", text=host.tr("your_answer"))
    tree.heading("count", text=host.tr("wrong_times"))
    tree.column("word", width=390, minwidth=240, anchor="w", stretch=True)
    tree.column("input", width=220, minwidth=150, anchor="w", stretch=True)
    tree.column("count", width=110, minwidth=90, anchor="center", stretch=False)
    tree.tag_configure("correct", foreground="#15803d")
    tree.tag_configure("wrong", foreground="#dc2626")
    tree.grid(row=0, column=0, sticky="nsew")
    scroll = ttk.Scrollbar(table_wrap, orient="vertical", command=tree.yview)
    scroll.grid(row=0, column=1, sticky="ns")
    tree.configure(yscrollcommand=scroll.set)
    host.dictation_answer_review_tree = tree
