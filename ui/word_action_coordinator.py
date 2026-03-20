# -*- coding: utf-8 -*-
from tkinter import ttk


def set_context(host, index, origin="main", word=None):
    try:
        idx = int(index)
    except Exception:
        idx = None
    token = host._normalize_import_word_text(word or "")
    if idx is None or idx < 0 or idx >= len(host.store.words):
        host.word_action_index = None
        host.word_action_word = token
        host.word_action_origin = "main"
        if token:
            host.word_action_origin = str(origin or "main")
        return None
    host.word_action_index = idx
    host.word_action_word = str(host.store.words[idx] or "").strip()
    host.word_action_origin = str(origin or "main")
    return idx


def clear_context(host):
    host.word_action_index = None
    host.word_action_word = ""
    host.word_action_origin = "main"


def get_context_or_selected_index(host):
    if host.word_action_index is not None and 0 <= host.word_action_index < len(host.store.words):
        return host.word_action_index
    return host._get_selected_index()


def get_context_word(host):
    if host.word_action_index is not None and 0 <= host.word_action_index < len(host.store.words):
        return str(host.store.words[host.word_action_index] or "").strip()
    token = str(host.word_action_word or "").strip()
    if token:
        return token
    selected_idx = host._get_selected_index()
    if selected_idx is None or selected_idx >= len(host.store.words):
        return ""
    return str(host.store.words[selected_idx] or "").strip()


def get_context_audio_source_path(host):
    if host.word_action_origin == "dictation" and host.dictation_list_mode_var.get() == "recent":
        return host._get_recent_wrong_cache_source_path()
    return host.store.get_current_source_path()


def get_word_audio_override_source_path(host):
    return host.store.get_current_source_path() or get_context_audio_source_path(host)


def dictation_row_to_store_index(host, tree, row_id=None):
    if not tree:
        return None
    item_id = str(row_id or tree.focus() or "").strip()
    if not item_id or item_id == "empty":
        selection = tree.selection()
        if selection:
            item_id = str(selection[0] or "").strip()
    if not item_id or item_id == "empty":
        return None
    try:
        view_index = int(item_id)
    except Exception:
        return None
    items = host._get_dictation_source_items()
    if view_index < 0 or view_index >= len(items):
        return None
    word = str(items[view_index].get("word") or "").strip()
    if not word:
        return None
    try:
        return host.store.words.index(word)
    except ValueError:
        return None


def on_word_right_click(host, event):
    if not host.word_table or not host.word_context_menu:
        return
    row_id = host.word_table.identify_row(event.y)
    if not row_id:
        return
    try:
        row_idx = int(row_id)
    except Exception:
        return
    try:
        host.suppress_word_select_action = True
        host.word_table.selection_set(row_id)
        host.word_table.focus(row_id)
    except Exception:
        pass
    set_context(host, row_idx, origin="main")
    try:
        host.word_context_menu.tk_popup(event.x_root, event.y_root)
    finally:
        host.word_context_menu.grab_release()
    return "break"


def on_dictation_word_right_click(host, event):
    tree = event.widget if isinstance(event.widget, ttk.Treeview) else None
    if not tree or not host.dictation_context_menu:
        return
    row_id = str(tree.identify_row(event.y) or "").strip()
    if not row_id or row_id == "empty":
        return
    try:
        host.suppress_dictation_select_action = True
        tree.selection_set(row_id)
        tree.focus(row_id)
    except Exception:
        pass
    selected_idx = dictation_row_to_store_index(host, tree, row_id=row_id)
    selected_word = ""
    try:
        view_index = int(row_id)
        items = host._get_dictation_source_items()
        if 0 <= view_index < len(items):
            selected_word = str(items[view_index].get("word") or "").strip()
    except Exception:
        selected_word = ""
    if selected_idx is not None:
        set_context(host, selected_idx, origin="dictation", word=selected_word)
    elif selected_word:
        set_context(host, None, origin="dictation", word=selected_word)
    else:
        return "break"
    try:
        host.dictation_context_menu.tk_popup(event.x_root, event.y_root)
    finally:
        host.dictation_context_menu.grab_release()
    return "break"
