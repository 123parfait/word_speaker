# -*- coding: utf-8 -*-
import tkinter as tk

from services.diff_view import build_diff_view


def apply_diff(text_widget, expected, actual):
    state = build_diff_view(expected, actual)

    text_widget.config(state="normal")
    text_widget.delete("1.0", tk.END)

    text_widget.insert(tk.END, "Expected: ")
    _insert_segments(text_widget, state["expected_segments"])
    text_widget.insert(tk.END, "\nYour input: ")
    _insert_segments(text_widget, state["actual_segments"])
    text_widget.insert(tk.END, "\nLegend: ")
    _insert_segments(text_widget, state["legend"])

    text_widget.config(state="disabled")


def _insert_segments(text_widget, segments):
    for text, tag in segments:
        if tag:
            text_widget.insert(tk.END, text, tag)
        else:
            text_widget.insert(tk.END, text)
