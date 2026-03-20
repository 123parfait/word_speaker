# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk


def build_find_window(host):
    host.find_window = tk.Toplevel(host)
    host.find_window.title("Find Corpus Sentences")
    host.find_window.configure(bg="#f6f7fb")
    host.find_window.minsize(900, 620)

    wrap = ttk.Frame(host.find_window, style="Card.TFrame")
    wrap.pack(fill="both", expand=True, padx=10, pady=10)
    wrap.grid_columnconfigure(0, weight=3)
    wrap.grid_columnconfigure(1, weight=2)
    wrap.grid_rowconfigure(2, weight=1)

    ttk.Label(wrap, text=host.tr("find_corpus_sentences"), style="Card.TLabel").grid(
        row=0, column=0, columnspan=2, sticky="w"
    )
    ttk.Label(
        wrap,
        text=host.tr("find_desc"),
        style="Card.TLabel",
        foreground="#666",
    ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 8))

    top = ttk.Frame(wrap, style="Card.TFrame")
    top.grid(row=2, column=0, sticky="nsew", padx=(0, 10))
    top.grid_columnconfigure(0, weight=1)
    top.grid_rowconfigure(2, weight=1)
    top.grid_rowconfigure(4, weight=0)

    search_row = ttk.Frame(top, style="Card.TFrame")
    search_row.grid(row=0, column=0, sticky="ew", pady=(0, 8))
    search_row.grid_columnconfigure(0, weight=1)
    entry = ttk.Entry(search_row, textvariable=host.find_search_var)
    entry.grid(row=0, column=0, sticky="ew", padx=(0, 6))
    entry.bind("<Return>", lambda _e: host.run_find_search())
    ttk.Button(search_row, text=host.tr("search"), command=host.run_find_search).grid(row=0, column=1, padx=(0, 6))
    ttk.Label(search_row, text=host.tr("show"), style="Card.TLabel").grid(row=0, column=2, padx=(0, 4))
    limit_combo = ttk.Combobox(
        search_row,
        textvariable=host.find_limit_var,
        state="readonly",
        width=5,
        values=("20", "50", "100"),
    )
    limit_combo.grid(row=0, column=3, padx=(0, 6))
    ttk.Label(search_row, text=host.tr("results"), style="Card.TLabel").grid(row=0, column=4, padx=(0, 6))
    ttk.Button(search_row, text=host.tr("use_selected_word"), command=host.search_selected_word_in_corpus).grid(
        row=0, column=5, padx=(0, 6)
    )
    host.find_import_btn = ttk.Button(search_row, text=host.tr("import_docs"), command=host.import_find_documents)
    host.find_import_btn.grid(row=0, column=6)

    ttk.Label(top, textvariable=host.find_status_var, style="Card.TLabel", foreground="#444").grid(
        row=1, column=0, sticky="w", pady=(0, 8)
    )

    host.find_results_table = ttk.Treeview(
        top,
        columns=("sentence", "source"),
        show="headings",
        height=18,
        selectmode="browse",
    )
    host.find_results_table.heading("sentence", text=host.tr("sentence"))
    host.find_results_table.heading("source", text=host.tr("source"))
    host.find_results_table.column("sentence", width=600, anchor="w")
    host.find_results_table.column("source", width=260, anchor="w")
    find_scroll = ttk.Scrollbar(top, orient="vertical", command=host.find_results_table.yview)
    host.find_results_table.configure(yscrollcommand=find_scroll.set)
    host.find_results_table.grid(row=2, column=0, sticky="nsew")
    find_scroll.grid(row=2, column=1, sticky="ns")
    host.find_results_table.bind("<<TreeviewSelect>>", host._on_find_result_select)

    ttk.Label(top, text=host.tr("preview"), style="Card.TLabel").grid(row=3, column=0, sticky="w", pady=(8, 4))
    host.find_preview_text = tk.Text(
        top,
        wrap="word",
        height=7,
        bg="#ffffff",
        fg="#222222",
        highlightthickness=1,
        highlightbackground="#d9dbe1",
    )
    host.find_preview_text.grid(row=4, column=0, columnspan=2, sticky="ew")
    host.find_preview_text.tag_configure("hit", background="#ffe58f", foreground="#111111")
    host.find_preview_text.configure(state="disabled")

    side = ttk.Frame(wrap, style="Card.TFrame")
    side.grid(row=2, column=1, sticky="nsew")
    side.grid_rowconfigure(1, weight=1)
    side.grid_columnconfigure(0, weight=1)

    side_header = ttk.Frame(side, style="Card.TFrame")
    side_header.grid(row=0, column=0, sticky="ew")
    side_header.grid_columnconfigure(0, weight=1)
    ttk.Label(side_header, text=host.tr("indexed_documents"), style="Card.TLabel").grid(row=0, column=0, sticky="w")
    ttk.Button(side_header, text=host.tr("clear_filter"), command=host.clear_find_document_filter).grid(
        row=0, column=1, sticky="e"
    )
    host.find_docs_list = tk.Listbox(
        side,
        bg="#ffffff",
        fg="#222222",
        selectbackground="#cce1ff",
        selectforeground="#111111",
        highlightthickness=1,
        highlightbackground="#d9dbe1",
    )
    host.find_docs_list.grid(row=1, column=0, sticky="nsew", pady=(6, 0))
    host.find_docs_list.bind("<Button-3>", host.on_find_docs_right_click)
    host.find_docs_context_menu = tk.Menu(host.find_window, tearoff=0)
    host.find_docs_context_menu.add_command(
        label=host.tr("delete_corpus_doc"),
        command=host.delete_selected_corpus_document,
    )

    def _on_close():
        host.find_window.destroy()
        host.find_window = None
        host.find_results_table = None
        host.find_preview_text = None
        host.find_docs_list = None
        host.find_docs_context_menu = None
        host.find_import_btn = None
        host.find_doc_items = []
        host.find_result_items = {}

    host.find_window.protocol("WM_DELETE_WINDOW", _on_close)
