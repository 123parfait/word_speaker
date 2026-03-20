# -*- coding: utf-8 -*-
from tkinter import ttk


def build_detail_card(host):
    host.detail_card = ttk.Frame(host.right, style="Card.TFrame")
    host.detail_card.configure(width=396)
    host.detail_card.grid(row=0, column=0, padx=12, pady=(12, 6), sticky="ew")
    host.detail_card.grid_columnconfigure(0, weight=1)

    ttk.Label(host.detail_card, text=host.tr("current_word"), style="Card.TLabel").grid(
        row=0, column=0, padx=12, pady=(12, 4), sticky="w"
    )
    ttk.Label(
        host.detail_card,
        textvariable=host.detail_word_var,
        style="Card.TLabel",
        font=("Segoe UI", 16, "bold"),
        wraplength=360,
        justify="left",
    ).grid(row=1, column=0, padx=12, sticky="w")
    ttk.Label(
        host.detail_card,
        textvariable=host.detail_translation_var,
        style="Card.TLabel",
        foreground="#1e3a8a",
        wraplength=360,
        justify="left",
    ).grid(row=2, column=0, padx=12, pady=(2, 0), sticky="w")
    ttk.Label(
        host.detail_card,
        textvariable=host.detail_note_var,
        style="Card.TLabel",
        foreground="#4b5563",
        wraplength=360,
        justify="left",
    ).grid(row=3, column=0, padx=12, pady=(4, 0), sticky="w")
    ttk.Label(
        host.detail_card,
        textvariable=host.detail_meta_var,
        style="Card.TLabel",
        foreground="#667085",
        wraplength=360,
        justify="left",
    ).grid(row=4, column=0, padx=12, pady=(6, 12), sticky="w")


def build_review_tab(host):
    host.review_tab.grid_columnconfigure(0, weight=1)
    review_card = ttk.Frame(host.review_tab, style="Card.TFrame")
    review_card.configure(width=396)
    review_card.pack(fill="both", expand=True, padx=10, pady=10)
    review_card.grid_columnconfigure(0, weight=1)
    ttk.Label(review_card, text=host.tr("study_focus"), style="Card.TLabel").grid(
        row=0, column=0, sticky="w"
    )
    ttk.Label(
        review_card,
        textvariable=host.review_focus_var,
        style="Card.TLabel",
        foreground="#4b5563",
        wraplength=360,
        justify="left",
    ).grid(row=1, column=0, sticky="w", pady=(4, 10))
    ttk.Label(
        review_card,
        textvariable=host.review_source_var,
        style="Card.TLabel",
        foreground="#667085",
        wraplength=360,
        justify="left",
    ).grid(row=2, column=0, sticky="w", pady=(0, 4))
    ttk.Label(
        review_card,
        textvariable=host.review_stats_var,
        style="Card.TLabel",
        foreground="#667085",
        wraplength=360,
        justify="left",
    ).grid(row=3, column=0, sticky="w", pady=(0, 12))
    review_actions = ttk.Frame(review_card, style="Card.TFrame")
    review_actions.grid(row=4, column=0, sticky="ew")
    review_actions.grid_columnconfigure(0, weight=1)
    review_actions.grid_columnconfigure(1, weight=1)
    host.review_open_source_btn = ttk.Button(
        review_actions,
        text=host.tr("open_history"),
        command=host.toggle_history,
    )
    host.review_open_source_btn.grid(row=0, column=0, padx=(0, 6), sticky="ew")
    ttk.Button(
        review_actions,
        text=host.tr("open_tools"),
        command=lambda: host._select_sidebar_tab("tools"),
    ).grid(row=0, column=1, padx=(6, 0), sticky="ew")
