# -*- coding: utf-8 -*-
import difflib
import tkinter as tk


def apply_diff(text_widget, expected, actual):
    expected_lower = expected.lower()
    actual_lower = actual.lower()
    matcher = difflib.SequenceMatcher(a=expected_lower, b=actual_lower)
    opcodes = matcher.get_opcodes()

    text_widget.config(state="normal")
    text_widget.delete("1.0", tk.END)

    text_widget.insert(tk.END, "Expected: ")
    _insert_expected(text_widget, expected, opcodes)
    text_widget.insert(tk.END, "\nYour input: ")
    _insert_actual(text_widget, actual, opcodes)
    text_widget.insert(tk.END, "\nLegend: ")
    text_widget.insert(tk.END, "missing", "missing")
    text_widget.insert(tk.END, " / ")
    text_widget.insert(tk.END, "extra", "extra")
    text_widget.insert(tk.END, " / ")
    text_widget.insert(tk.END, "wrong", "wrong")

    text_widget.config(state="disabled")


def _insert_expected(text_widget, expected, opcodes):
    for tag, i1, i2, _j1, _j2 in opcodes:
        if tag == "insert":
            continue
        segment = expected[i1:i2]
        if tag == "equal":
            text_widget.insert(tk.END, segment)
        elif tag == "replace":
            text_widget.insert(tk.END, segment, "wrong")
        elif tag == "delete":
            text_widget.insert(tk.END, segment, "missing")


def _insert_actual(text_widget, actual, opcodes):
    for tag, _i1, _i2, j1, j2 in opcodes:
        if tag == "delete":
            continue
        segment = actual[j1:j2]
        if tag == "equal":
            text_widget.insert(tk.END, segment)
        elif tag == "replace":
            text_widget.insert(tk.END, segment, "wrong")
        elif tag == "insert":
            text_widget.insert(tk.END, segment, "extra")
