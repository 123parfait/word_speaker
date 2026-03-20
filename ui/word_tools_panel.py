# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk


def build_sentence_window(host, *, state, on_read):
    if host.sentence_window and host.sentence_window.winfo_exists():
        host.sentence_window.destroy()

    host.sentence_window = tk.Toplevel(host)
    host.sentence_window.title(state["title"])
    host.sentence_window.configure(bg="#f6f7fb")
    host.sentence_window.resizable(False, False)

    wrap = ttk.Frame(host.sentence_window, style="Card.TFrame")
    wrap.pack(fill="both", expand=True, padx=10, pady=10)
    ttk.Label(wrap, text=state["word_label"], style="Card.TLabel").pack(anchor="w")
    ttk.Label(wrap, text=state["source_label"], style="Card.TLabel", foreground="#666").pack(anchor="w", pady=(0, 6))

    text = tk.Text(
        wrap,
        wrap="word",
        height=4,
        width=56,
        bg="#ffffff",
        fg="#222222",
        highlightthickness=1,
        highlightbackground="#d9dbe1",
    )
    text.pack(fill="x")
    text.insert("1.0", state["sentence_text"])
    text.config(state="disabled")

    row = ttk.Frame(wrap, style="Card.TFrame")
    row.pack(fill="x", pady=(8, 0))
    ttk.Button(row, text=host.tr("read_sentence"), command=on_read).pack(side=tk.LEFT, padx=(0, 6))
    ttk.Button(row, text=host.tr("close"), command=host.sentence_window.destroy).pack(side=tk.LEFT)


def build_synonym_window(host, *, state):
    if host.synonym_window and host.synonym_window.winfo_exists():
        host.synonym_window.destroy()

    host.synonym_window = tk.Toplevel(host)
    host.synonym_window.title(state["title"])
    host.synonym_window.configure(bg="#f6f7fb")
    host.synonym_window.resizable(False, False)

    wrap = ttk.Frame(host.synonym_window, style="Card.TFrame")
    wrap.pack(fill="both", expand=True, padx=10, pady=10)
    ttk.Label(wrap, text=state["word_label"], style="Card.TLabel").pack(anchor="w")
    if state["focus_label"]:
        ttk.Label(
            wrap,
            text=state["focus_label"],
            style="Card.TLabel",
            foreground="#666",
        ).pack(anchor="w", pady=(0, 4))
    ttk.Label(
        wrap,
        text=state["source_label"],
        style="Card.TLabel",
        foreground="#666",
    ).pack(anchor="w", pady=(0, 8))

    text = tk.Text(
        wrap,
        wrap="word",
        height=8,
        width=56,
        bg="#ffffff",
        fg="#222222",
        highlightthickness=1,
        highlightbackground="#d9dbe1",
    )
    text.pack(fill="both", expand=True)
    text.insert("1.0", state["synonym_text"])
    text.config(state="disabled")

    row = ttk.Frame(wrap, style="Card.TFrame")
    row.pack(fill="x", pady=(8, 0))
    ttk.Button(row, text=host.tr("close"), command=host.synonym_window.destroy).pack(side=tk.LEFT)
