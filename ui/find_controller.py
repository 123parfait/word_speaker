# -*- coding: utf-8 -*-
from dataclasses import dataclass


@dataclass(frozen=True)
class FindImportStartState:
    paths: tuple[str, ...]
    status_text: str
    display_name: str


@dataclass(frozen=True)
class FindSearchStartState:
    query: str
    limit: int
    limit_text: str
    document_path: str | None
    status_text: str


def build_find_import_start_state(paths):
    import os

    normalized_paths = tuple(str(path or "").strip() for path in (paths or []) if str(path or "").strip())
    if not normalized_paths:
        return None
    first_name = os.path.basename(normalized_paths[0])
    if len(normalized_paths) == 1:
        status_text = f"正在导入 {first_name}，请稍候……"
    else:
        status_text = f"正在导入 {first_name} 等 {len(normalized_paths)} 个文档，请稍候……"
    return FindImportStartState(
        paths=normalized_paths,
        status_text=status_text,
        display_name=first_name,
    )


def build_find_search_start_state(*, query_text, limit_text, selected_doc, status_builder):
    query = str(query_text or "").strip()
    if not query:
        raise ValueError("empty_query")
    try:
        limit = int(str(limit_text or "20").strip())
    except Exception:
        limit = 20
    document_path = None
    selected_doc_name = ""
    if selected_doc:
        document_path = str(selected_doc.get("path") or "").strip() or None
        selected_doc_name = str(selected_doc.get("name") or "").strip()
    return FindSearchStartState(
        query=query,
        limit=limit,
        limit_text=str(limit),
        document_path=document_path,
        status_text=status_builder(query=query, limit=limit, selected_doc_name=selected_doc_name),
    )


def build_find_clear_filter_status():
    return "Document filter cleared. Search will use all indexed documents."
