# -*- coding: utf-8 -*-
import os
from dataclasses import dataclass


@dataclass
class HistoryListState:
    labels: list[str]
    current_path: str
    is_empty: bool


@dataclass
class SelectedHistoryItem:
    path: str
    name: str
    index: int


@dataclass
class RenameHistoryTarget:
    old_path: str
    new_path: str
    new_name: str
    changed: bool


def build_history_list_state(history):
    items = list(history or [])
    labels = [f"{item.get('name', '')}  ({item.get('time', '')})" for item in items]
    current_path = str(items[0].get("path", "")) if items else ""
    return HistoryListState(labels=labels, current_path=current_path, is_empty=not items)


def get_selected_history_item(history, selection):
    items = list(history or [])
    if not selection:
        return None
    idx = int(selection[0])
    if idx < 0 or idx >= len(items):
        return None
    item = items[idx]
    path = str(item.get("path") or "").strip()
    name = str(item.get("name") or os.path.basename(path) or path)
    return SelectedHistoryItem(path=path, name=name, index=idx)


def build_rename_history_target(old_path, requested_name):
    old_target = os.path.abspath(str(old_path or "").strip())
    old_name = os.path.basename(old_target)
    new_name = str(requested_name or "").strip()
    if not new_name:
        raise ValueError("empty_name")
    if os.path.basename(new_name) != new_name:
        raise ValueError("invalid_name")
    _old_root, old_ext = os.path.splitext(old_name)
    _new_root, new_ext = os.path.splitext(new_name)
    if not new_ext and old_ext:
        new_name = f"{new_name}{old_ext}"
    new_path = os.path.join(os.path.dirname(old_target), new_name)
    changed = os.path.abspath(new_path) != old_target
    return RenameHistoryTarget(
        old_path=old_target,
        new_path=os.path.abspath(new_path),
        new_name=new_name,
        changed=changed,
    )
