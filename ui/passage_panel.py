# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk


def build_passage_window(host, tooltip_cls):
    host.passage_window = tk.Toplevel(host)
    host.passage_window.title("IELTS Passage Builder")
    host.passage_window.configure(bg="#f6f7fb")
    host.passage_window.resizable(True, True)
    host.passage_window.minsize(680, 460)

    wrap = ttk.Frame(host.passage_window, style="Card.TFrame")
    wrap.pack(fill="both", expand=True, padx=10, pady=10)

    ttk.Label(wrap, text=host.tr("passage_title"), style="Card.TLabel").pack(anchor="w")
    ttk.Label(
        wrap,
        text=host.tr("passage_desc"),
        style="Card.TLabel",
        foreground="#666",
    ).pack(anchor="w", pady=(0, 8))

    ctrl = ttk.Frame(wrap, style="Card.TFrame")
    ctrl.pack(fill="x", pady=(0, 8))

    btn_generate = ttk.Button(ctrl, text=host.tr("generate"), command=host.generate_ielts_passage)
    btn_generate.pack(side=tk.LEFT, padx=(0, 6))
    tooltip_cls(btn_generate, "Build passage from selected words, or the full list if nothing is selected")

    btn_play = ttk.Button(ctrl, text=host.tr("read_with_gemini"), command=host.play_generated_passage)
    btn_play.pack(side=tk.LEFT, padx=(0, 6))
    tooltip_cls(btn_play, "Speak current passage")

    btn_stop = ttk.Button(ctrl, text=host.tr("stop"), command=host.stop_passage_playback)
    btn_stop.pack(side=tk.LEFT, padx=(0, 6))
    tooltip_cls(btn_stop, "Stop speaking")

    btn_practice = ttk.Button(ctrl, text=host.tr("practice"), command=host.start_passage_practice)
    btn_practice.pack(side=tk.LEFT, padx=(0, 6))
    tooltip_cls(btn_practice, "Hide keywords as blanks")

    btn_check_practice = ttk.Button(ctrl, text=host.tr("check"), command=host.check_passage_practice)
    btn_check_practice.pack(side=tk.LEFT)
    tooltip_cls(btn_check_practice, "Check your filled answers")

    model_wrap = ttk.Frame(ctrl, style="Card.TFrame")
    model_wrap.pack(side=tk.RIGHT)
    ttk.Label(model_wrap, text=host.tr("model"), style="Card.TLabel").pack(side=tk.LEFT, padx=(0, 4))
    host.gemini_model_combo = ttk.Combobox(
        model_wrap,
        textvariable=host.gemini_model_var,
        state="readonly",
        width=20,
    )
    host.gemini_model_combo.pack(side=tk.LEFT)
    host.gemini_model_combo.bind("<<ComboboxSelected>>", host.on_gemini_model_change)

    host.passage_text = tk.Text(
        wrap,
        wrap="word",
        height=18,
        bg="#ffffff",
        fg="#222222",
        highlightthickness=1,
        highlightbackground="#d9dbe1",
    )
    host.passage_text.pack(fill="both", expand=True)

    practice_wrap = ttk.Frame(wrap, style="Card.TFrame")
    practice_wrap.pack(fill="x", pady=(8, 0))
    ttk.Label(
        practice_wrap,
        text=host.tr("practice_tip"),
        style="Card.TLabel",
        foreground="#666",
    ).pack(anchor="w")

    host.passage_practice_input = tk.Text(
        practice_wrap,
        wrap="word",
        height=4,
        bg="#ffffff",
        fg="#222222",
        highlightthickness=1,
        highlightbackground="#d9dbe1",
    )
    host.passage_practice_input.pack(fill="x", pady=(4, 6))

    host.passage_practice_result = tk.Text(
        practice_wrap,
        wrap="word",
        height=5,
        bg="#ffffff",
        fg="#222222",
        highlightthickness=1,
        highlightbackground="#d9dbe1",
    )
    host.passage_practice_result.pack(fill="x")
    host.passage_practice_result.tag_configure("wrong", foreground="#c62828")
    host.passage_practice_result.tag_configure("missing", foreground="#c62828", underline=1)
    host.passage_practice_result.tag_configure("extra", foreground="#ef6c00")
    host.passage_practice_result.config(state="disabled")

    ttk.Label(wrap, textvariable=host.passage_status_var, style="Card.TLabel", foreground="#444").pack(
        anchor="w", pady=(8, 0)
    )

    def _on_close():
        host.passage_window.destroy()
        host.passage_window = None
        host.passage_text = None
        host.passage_practice_input = None
        host.passage_practice_result = None

    host.passage_window.protocol("WM_DELETE_WINDOW", _on_close)
