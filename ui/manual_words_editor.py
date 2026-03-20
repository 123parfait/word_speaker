# -*- coding: utf-8 -*-
import re
import tkinter as tk
from tkinter import ttk


def append_manual_preview_rows(table, rows, *, replace=False):
    if not table:
        return
    if replace:
        table.delete(*table.get_children())
    start = len(table.get_children())
    for offset, row in enumerate(rows):
        word = str(row.get("word") or "").strip()
        note = str(row.get("note") or "").strip()
        if not word:
            continue
        table.insert("", tk.END, iid=f"manual_{start + offset}", values=(word, note))


def collect_manual_rows_from_table(table):
    rows = []
    if not table:
        return rows
    for item_id in table.get_children():
        values = list(table.item(item_id, "values") or [])
        word = str(values[0] if len(values) > 0 else "").strip()
        note = str(values[1] if len(values) > 1 else "").strip()
        if not word:
            continue
        rows.append({"word": word, "note": note})
    return rows


def clear_manual_preview(table, cancel_edit):
    cancel_edit()
    if table:
        table.delete(*table.get_children())


def close_manual_words_window(window, cancel_edit):
    cancel_edit()
    if window and window.winfo_exists():
        window.destroy()


def cancel_manual_preview_edit(entry):
    if entry and entry.winfo_exists():
        entry.destroy()


def finish_manual_preview_edit(*, entry, table, item_id, column_id, cancel_edit):
    if not entry or not table:
        return "break"
    new_value = re.sub(r"\s+", " ", str(entry.get() or "").strip())
    cancel_edit()
    if not item_id or not table.exists(item_id):
        return "break"
    values = list(table.item(item_id, "values") or ["", ""])
    while len(values) < 2:
        values.append("")
    if column_id == "#1" and not new_value:
        return "break"
    if column_id == "#1":
        values[0] = new_value
    elif column_id == "#2":
        values[1] = new_value
    table.item(item_id, values=values)
    return "break"


def resolve_manual_preview_target(table, *, event=None, item_id=None, column_id=None):
    if not table:
        return None, None
    target_item = item_id
    target_column = column_id or "#1"
    if event is not None:
        target_item = table.identify_row(event.y)
        target_column = table.identify_column(event.x)
    if not target_item:
        target_item = str(table.focus() or "").strip()
    if not target_item:
        selection = table.selection()
        if selection:
            target_item = selection[0]
    if target_column not in ("#1", "#2"):
        target_column = "#1"
    if not target_item or not table.exists(target_item):
        return None, None
    return target_item, target_column


def start_manual_preview_edit(*, table, event=None, item_id=None, column_id=None, cancel_edit, on_finish):
    target_item, target_column = resolve_manual_preview_target(
        table,
        event=event,
        item_id=item_id,
        column_id=column_id,
    )
    if not target_item:
        return None
    cancel_edit()
    bbox = table.bbox(target_item, target_column)
    if not bbox:
        return None
    x, y, width, height = bbox
    values = list(table.item(target_item, "values") or ["", ""])
    while len(values) < 2:
        values.append("")
    current_value = values[0] if target_column == "#1" else values[1]
    entry = ttk.Entry(table)
    entry.insert(0, current_value)
    entry.select_range(0, tk.END)
    entry.focus_set()
    entry.place(x=x, y=y, width=width, height=height)
    entry.bind("<Return>", on_finish)
    entry.bind("<Escape>", lambda _event=None: cancel_edit())
    entry.bind("<FocusOut>", on_finish)
    return {
        "entry": entry,
        "item_id": target_item,
        "column_id": target_column,
    }


def add_manual_preview_row(table):
    if not table:
        return None
    item_id = f"manual_{len(table.get_children())}"
    table.insert("", tk.END, iid=item_id, values=("", ""))
    table.selection_set(item_id)
    table.focus(item_id)
    return item_id


def delete_selected_manual_preview_rows(table, cancel_edit):
    if not table:
        return
    selected = list(table.selection())
    if not selected:
        return
    cancel_edit()
    for item_id in selected:
        if table.exists(item_id):
            table.delete(item_id)
