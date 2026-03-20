# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk


def build_manual_words_window(host):
    host.manual_words_window = tk.Toplevel(host)
    host.manual_words_window.title("Manual Words")
    host.manual_words_window.configure(bg="#f6f7fb")
    host.manual_words_window.resizable(False, False)

    wrap = ttk.Frame(host.manual_words_window, style="Card.TFrame")
    wrap.pack(fill="both", expand=True, padx=10, pady=10)
    ttk.Label(wrap, text=host.tr("paste_preview"), style="Card.TLabel").pack(anchor="w")
    ttk.Label(
        wrap,
        text=host.tr("paste_preview_desc"),
        style="Card.TLabel",
        foreground="#666",
    ).pack(anchor="w", pady=(0, 6))

    table_wrap = ttk.Frame(wrap, style="Card.TFrame")
    table_wrap.pack(fill="both", expand=True)
    host.manual_words_table = ttk.Treeview(
        table_wrap,
        columns=("en", "note"),
        show="headings",
        height=12,
        selectmode="browse",
    )
    host.manual_words_table.heading("en", text=host.tr("english"))
    host.manual_words_table.heading("note", text=host.tr("notes"))
    host.manual_words_table.column("en", width=260, anchor="w")
    host.manual_words_table.column("note", width=420, anchor="w")
    host.manual_words_table.grid(row=0, column=0, sticky="nsew")
    host.manual_words_table_scroll = ttk.Scrollbar(
        table_wrap,
        orient="vertical",
        command=host.manual_words_table.yview,
    )
    host.manual_words_table_scroll.grid(row=0, column=1, sticky="ns")
    host.manual_words_table.configure(yscrollcommand=host.manual_words_table_scroll.set)
    table_wrap.grid_rowconfigure(0, weight=1)
    table_wrap.grid_columnconfigure(0, weight=1)
    host.manual_words_table.bind("<Control-v>", host.on_manual_preview_paste)
    host.manual_words_table.bind("<Control-V>", host.on_manual_preview_paste)
    host.manual_words_table.bind("<Double-1>", host._start_manual_preview_edit)
    host.manual_words_table.bind("<F2>", host._start_manual_preview_edit)
    host.manual_words_table.bind("<Delete>", lambda _e: host._delete_selected_manual_preview_rows())
    host.manual_words_table.bind("<Escape>", host._cancel_manual_preview_edit)
    host.manual_words_table.focus_set()

    helper = ttk.Label(
        wrap,
        text="Tip: press Ctrl+V to paste, or add rows and edit cells directly.",
        style="Card.TLabel",
        foreground="#666",
    )
    helper.pack(anchor="w", pady=(6, 0))

    row = ttk.Frame(wrap, style="Card.TFrame")
    row.pack(fill="x", pady=(8, 0))
    ttk.Button(
        row,
        text=host.tr("paste_clipboard"),
        command=host.on_manual_preview_paste,
    ).pack(side=tk.LEFT, padx=(0, 6))
    ttk.Button(
        row,
        text=host.tr("add_row"),
        command=host._add_manual_preview_row,
    ).pack(side=tk.LEFT, padx=(0, 6))
    ttk.Button(
        row,
        text=host.tr("delete_selected"),
        command=host._delete_selected_manual_preview_rows,
    ).pack(side=tk.LEFT, padx=(0, 6))
    ttk.Button(
        row,
        text=host.tr("clear"),
        command=host._clear_manual_preview,
    ).pack(side=tk.LEFT, padx=(0, 12))
    ttk.Button(
        row,
        text=host.tr("replace_list"),
        command=lambda: host._apply_manual_words_from_editor("replace"),
    ).pack(side=tk.LEFT, padx=(0, 6))
    ttk.Button(
        row,
        text=host.tr("append"),
        command=lambda: host._apply_manual_words_from_editor("append"),
    ).pack(side=tk.LEFT, padx=(0, 6))
    ttk.Button(
        row,
        text=host.tr("close"),
        command=host._close_manual_words_window,
    ).pack(side=tk.LEFT)

    host.manual_words_window.protocol("WM_DELETE_WINDOW", host._close_manual_words_window)
