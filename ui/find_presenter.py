# -*- coding: utf-8 -*-
from dataclasses import dataclass


@dataclass(frozen=True)
class FindCorpusSummaryState:
    doc_labels: list[str]
    status_text: str


@dataclass(frozen=True)
class FindSearchResultState:
    result_items: dict[str, dict]
    result_rows: list[tuple[str, tuple[str, str]]]
    first_row_id: str | None
    status_text: str


@dataclass(frozen=True)
class FindPreviewState:
    sentence: str
    source: str
    highlight_ranges: list[tuple[int, int]]


def build_find_corpus_summary_state(stats, docs):
    doc_labels = []
    for item in docs or []:
        doc_labels.append(
            f"{item.get('name')} "
            f"({int(item.get('chunk_count') or 0)} chunks / {int(item.get('sentence_count') or 0)} sentences)"
        )
    status_text = (
        f"Indexed {stats.get('documents', 0)} documents / "
        f"{stats.get('chunks', 0)} chunks / "
        f"{stats.get('sentences', 0)} sentences. "
        f"NLP: {stats.get('nlp_mode')}"
    )
    return FindCorpusSummaryState(doc_labels=doc_labels, status_text=status_text)


def get_selected_find_document(doc_items, selection):
    if not doc_items or not selection:
        return None
    idx = int(selection[0])
    if idx < 0 or idx >= len(doc_items):
        return None
    return doc_items[idx]


def build_find_import_status(payload):
    files_count = int(payload.get("files") or 0)
    chunk_count = int(payload.get("chunks") or 0)
    sent_count = int(payload.get("sentences") or 0)
    errors = list(payload.get("errors") or [])
    status = f"Imported {files_count} files and indexed {chunk_count} chunks / {sent_count} sentences."
    if errors:
        status += f" Errors: {len(errors)}"
    return status, errors


def build_find_import_completion_message(payload):
    files_count = int(payload.get("files") or 0)
    chunk_count = int(payload.get("chunks") or 0)
    sent_count = int(payload.get("sentences") or 0)
    errors = list(payload.get("errors") or [])
    if files_count <= 0:
        message = "没有成功导入任何文档。"
    else:
        message = f"已导入 {files_count} 个文件，{sent_count} 句，{chunk_count} 个片段。"
    if errors:
        message += f"\n\n另有 {len(errors)} 个错误。"
    return message


def build_find_search_status(*, query, limit, selected_doc_name):
    if selected_doc_name:
        return f"Searching '{query}' in {selected_doc_name} (up to {limit} results)..."
    return f"Searching corpus for '{query}' (up to {limit} results)..."


def build_find_search_result_state(*, payload, doc_items):
    results = list(payload.get("results") or [])
    query = str(payload.get("query") or "").strip()
    lemmas = list(payload.get("lemmas") or [])
    document_path = str(payload.get("document_path") or "").strip()
    filtered_name = ""
    for item in doc_items or []:
        if str(item.get("path") or "").strip() == document_path:
            filtered_name = str(item.get("name") or "").strip()
            break

    result_items = {}
    result_rows = []
    for idx, item in enumerate(results):
        source_bits = [item.get("source_file") or ""]
        for key in ("test_label", "section_label", "part_label", "speaker_label", "question_label"):
            value = str(item.get(key) or "").strip()
            if value:
                source_bits.append(value)
        if item.get("page_num"):
            source_bits.append(f"p.{item.get('page_num')}")
        source = " · ".join(bit for bit in source_bits if bit)
        row_item = dict(item)
        row_item["source_text"] = source
        row_id = f"find_{idx}"
        result_items[row_id] = row_item
        result_rows.append((row_id, (row_item.get("sentence_text") or "", source)))

    if filtered_name:
        status_text = (
            f"Found {len(results)} results for '{query}' in {filtered_name}. "
            f"Lemmas: {', '.join(lemmas) if lemmas else 'n/a'}"
        )
    else:
        status_text = f"Found {len(results)} results for '{query}'. Lemmas: {', '.join(lemmas) if lemmas else 'n/a'}"

    return FindSearchResultState(
        result_items=result_items,
        result_rows=result_rows,
        first_row_id=(result_rows[0][0] if result_rows else None),
        status_text=status_text,
    )


def build_find_preview_state(item):
    if not item:
        return None
    return FindPreviewState(
        sentence=str(item.get("sentence_text") or ""),
        source=str(item.get("source_text") or ""),
        highlight_ranges=list(item.get("highlight_ranges") or []),
    )
