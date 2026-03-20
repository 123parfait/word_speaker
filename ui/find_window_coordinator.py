# -*- coding: utf-8 -*-
import os
import tkinter as tk
from tkinter import filedialog, messagebox

from services.corpus_search import (
    corpus_stats,
    get_nlp_status,
    list_documents as list_corpus_documents,
    remove_document as remove_corpus_document,
)
from ui.async_event_helper import clear_event_queue, drain_event_queue, emit_event
from ui.find_async import start_find_import_task, start_find_search_task
from ui.find_controller import (
    build_find_clear_filter_status,
    build_find_import_start_state,
    build_find_search_start_state,
)
from ui.find_panel import build_find_window
from ui.find_presenter import (
    build_find_corpus_summary_state,
    build_find_import_completion_message,
    build_find_import_status,
    build_find_preview_state,
    build_find_search_result_state,
    build_find_search_status,
    get_selected_find_document,
)


def open_window(host):
    if host.find_window and host.find_window.winfo_exists():
        host.find_window.lift()
        set_query_from_selection(host)
        return
    build_find_window(host)
    set_query_from_selection(host)
    refresh_corpus_summary(host)


def set_query_from_selection(host):
    word = host._get_context_word()
    if not word:
        return
    host.find_search_var.set(word)


def refresh_corpus_summary(host):
    try:
        stats = corpus_stats()
        docs = list_corpus_documents(limit=200)
    except Exception as e:
        if host.find_docs_list:
            host.find_docs_list.delete(0, tk.END)
        host.find_status_var.set(str(e))
        return
    state = build_find_corpus_summary_state(stats, docs)
    if host.find_docs_list:
        host.find_docs_list.delete(0, tk.END)
        host.find_doc_items = list(docs)
        for label in state.doc_labels:
            host.find_docs_list.insert(tk.END, label)
    host.find_status_var.set(state.status_text)


def on_docs_right_click(host, event):
    if not host.find_docs_list or not host.find_docs_context_menu:
        return
    index = host.find_docs_list.nearest(event.y)
    if index < 0 or index >= len(host.find_doc_items):
        return
    try:
        host.find_docs_list.selection_clear(0, tk.END)
        host.find_docs_list.selection_set(index)
        host.find_docs_list.activate(index)
    except Exception:
        pass
    try:
        host.find_docs_context_menu.tk_popup(event.x_root, event.y_root)
    finally:
        host.find_docs_context_menu.grab_release()
    return "break"


def delete_selected_document(host):
    if not host.find_docs_list or not host.find_doc_items:
        return
    selection = host.find_docs_list.curselection()
    if not selection:
        return
    idx = int(selection[0])
    if idx < 0 or idx >= len(host.find_doc_items):
        return
    item = host.find_doc_items[idx]
    path = str(item.get("path") or "").strip()
    name = str(item.get("name") or os.path.basename(path) or path)
    if not path:
        return
    if not messagebox.askyesno(host.tr("find_corpus_sentences"), host.trf("delete_corpus_doc_confirm", name=name)):
        return
    removed = remove_corpus_document(path)
    clear_document_filter(host)
    refresh_corpus_summary(host)
    run_search(host)
    if removed:
        messagebox.showinfo(host.tr("find_corpus_sentences"), host.trf("corpus_doc_deleted", name=name))


def clear_task_queue(host):
    clear_event_queue(host.find_task_queue)


def emit_task_event(host, event_type, token, payload=None):
    emit_event(host.find_task_queue, event_type, token, payload)


def poll_task_events(host, token):
    done = drain_event_queue(
        target_queue=host.find_task_queue,
        token=token,
        active_token=host.find_active_token,
        handlers={
            "import_done": lambda payload: apply_import_result(host, payload or {}),
            "search_done": lambda payload: apply_search_result(host, payload or {}),
            "error": lambda payload: handle_task_error(host, str(payload or "Unknown error")),
        },
    )
    if not done and token == host.find_active_token:
        host.after(80, lambda t=token: poll_task_events(host, t))


def handle_task_error(host, message):
    if host.find_import_btn:
        host.find_import_btn.state(["!disabled"])
    host.find_status_var.set(message)
    messagebox.showerror("Find Error", message)


def import_documents(host):
    if not host.find_window:
        open_window(host)
    try:
        get_nlp_status()
    except Exception as e:
        messagebox.showerror("Find Setup Error", str(e))
        host.find_status_var.set(str(e))
        return
    paths = filedialog.askopenfilenames(
        title="Choose documents",
        filetypes=[("Supported files", "*.txt *.docx *.pdf"), ("Text", "*.txt"), ("Word", "*.docx"), ("PDF", "*.pdf")],
    )
    state = build_find_import_start_state(paths)
    if not state:
        return
    host.find_task_token += 1
    token = host.find_task_token
    host.find_active_token = token
    clear_task_queue(host)
    if host.find_import_btn:
        host.find_import_btn.state(["disabled"])
    host.find_status_var.set(state.status_text)
    start_find_import_task(paths=state.paths, token=token, emit_event=lambda et, tk, payload=None: emit_task_event(host, et, tk, payload))
    host.after(80, lambda t=token: poll_task_events(host, t))


def apply_import_result(host, payload):
    refresh_corpus_summary(host)
    status, errors = build_find_import_status(payload)
    host.find_status_var.set(status)
    if host.find_import_btn:
        host.find_import_btn.state(["!disabled"])
    messagebox.showinfo("导入完成", build_find_import_completion_message(payload))
    if errors:
        messagebox.showerror("Import Warning", "\n".join(errors[:10]))


def run_search(host):
    selected_doc = get_selected_document(host)
    try:
        state = build_find_search_start_state(
            query_text=host.find_search_var.get(),
            limit_text=host.find_limit_var.get(),
            selected_doc=selected_doc,
            status_builder=build_find_search_status,
        )
    except ValueError:
        messagebox.showinfo("Info", "Enter a word or phrase first.")
        return
    host.find_limit_var.set(state.limit_text)
    try:
        get_nlp_status()
    except Exception as e:
        messagebox.showerror("Find Setup Error", str(e))
        host.find_status_var.set(str(e))
        return
    host.find_task_token += 1
    token = host.find_task_token
    host.find_active_token = token
    clear_task_queue(host)
    host.find_status_var.set(state.status_text)
    start_find_search_task(
        query=state.query,
        limit=state.limit,
        document_path=state.document_path,
        token=token,
        emit_event=lambda et, tk, payload=None: emit_task_event(host, et, tk, payload),
    )
    host.after(80, lambda t=token: poll_task_events(host, t))


def search_selected_word(host):
    if not host.find_window or not host.find_window.winfo_exists():
        open_window(host)
    set_query_from_selection(host)
    run_search(host)


def get_selected_document(host):
    if not host.find_docs_list:
        return None
    return get_selected_find_document(host.find_doc_items, host.find_docs_list.curselection())


def clear_document_filter(host):
    if host.find_docs_list:
        host.find_docs_list.selection_clear(0, tk.END)
    host.find_status_var.set(build_find_clear_filter_status())


def apply_search_result(host, payload):
    state = build_find_search_result_state(payload=payload, doc_items=host.find_doc_items)
    host.find_result_items = state.result_items
    if host.find_results_table:
        host.find_results_table.delete(*host.find_results_table.get_children())
        for row_id, values in state.result_rows:
            host.find_results_table.insert("", tk.END, iid=row_id, values=values)
        if state.first_row_id:
            host.find_results_table.selection_set(state.first_row_id)
            host.find_results_table.focus(state.first_row_id)
            show_result_preview(host, state.first_row_id)
        else:
            clear_preview(host)
    else:
        clear_preview(host)
    host.find_status_var.set(state.status_text)


def clear_preview(host):
    if not host.find_preview_text:
        return
    host.find_preview_text.configure(state="normal")
    host.find_preview_text.delete("1.0", tk.END)
    host.find_preview_text.configure(state="disabled")


def on_result_select(host, _event=None):
    if not host.find_results_table:
        return
    selection = host.find_results_table.selection()
    if not selection:
        clear_preview(host)
        return
    show_result_preview(host, selection[0])


def show_result_preview(host, row_id):
    state = build_find_preview_state(host.find_result_items.get(row_id))
    if not state or not host.find_preview_text:
        clear_preview(host)
        return
    text = host.find_preview_text
    text.configure(state="normal")
    text.delete("1.0", tk.END)
    text.insert("1.0", state.sentence)
    for start, end in state.highlight_ranges:
        if start < end:
            text.tag_add("hit", f"1.0+{int(start)}c", f"1.0+{int(end)}c")
    if state.source:
        text.insert(tk.END, "\n\n")
        source_start = text.index(tk.END)
        text.insert(tk.END, state.source)
        text.tag_add("source", source_start, tk.END)
    text.tag_configure("source", foreground="#666666")
    text.configure(state="disabled")
