# -*- coding: utf-8 -*-
import hashlib
import os
from pathlib import Path

from services.corpus_ingest import (
    clean_line as _clean_line,
    doc_sentences as _doc_sentences,
    get_nlp,
    get_nlp_status,
    highlight_ranges as _highlight_ranges,
    iter_file_blocks as _iter_file_blocks,
    lemma_doc as _lemma_doc,
    parse_structured_blocks as _parse_structured_blocks,
)
from services.corpus_index_store import (
    create_import_record as _create_import_record,
    ensure_schema,
    fetch_documents as _fetch_documents,
    fetch_stats as _fetch_stats,
    finish_import_record as _finish_import_record,
    get_connection as _get_connection,
    replace_document as _replace_document,
    remove_document_by_path as _remove_document_by_path,
    search_sentence_rows as _search_sentence_rows,
)


def _file_hash(path):
    sha1 = hashlib.sha1()
    with open(path, "rb") as fp:
        while True:
            chunk = fp.read(1024 * 1024)
            if not chunk:
                break
            sha1.update(chunk)
    return sha1.hexdigest()

def import_corpus_files(paths):
    ensure_schema()
    conn = _get_connection()
    summary = {
        "files": 0,
        "chunks": 0,
        "sentences": 0,
        "errors": [],
        "nlp_mode": get_nlp_status(),
    }
    try:
        for input_path in paths:
            path = os.path.abspath(str(input_path or "").strip())
            if not path or not os.path.exists(path):
                continue
            import_id = None
            try:
                file_hash = _file_hash(path)
                document_id = _replace_document(
                    conn,
                    path=path,
                    name=os.path.basename(path),
                    file_type=Path(path).suffix.lower().lstrip("."),
                    file_hash=file_hash,
                )
                import_id = _create_import_record(conn, path, document_id)
                raw_blocks = _iter_file_blocks(path)
                structured_blocks = _parse_structured_blocks(raw_blocks)
                sort_key = 0
                document_chunks = 0
                document_sentences = 0
                for block in structured_blocks:
                    chunk_text = _clean_line(block.get("text"))
                    if len(chunk_text) < 2:
                        continue
                    cur_chunk = conn.execute(
                        """
                        INSERT INTO chunks(
                            document_id, chunk_text, page_num, test_label, section_label,
                            part_label, speaker_label, question_label, sort_key
                        ) VALUES(?,?,?,?,?,?,?,?,?)
                        """,
                        (
                            document_id,
                            chunk_text,
                            block.get("page_num"),
                            block.get("test_label") or "",
                            block.get("section_label") or "",
                            block.get("part_label") or "",
                            block.get("speaker_label") or "",
                            block.get("question_label") or "",
                            sort_key,
                        ),
                    )
                    chunk_id = int(cur_chunk.lastrowid)
                    document_chunks += 1
                    sentence_order = 0
                    for sentence in _doc_sentences(chunk_text):
                        clean_sentence = _clean_line(sentence)
                        if len(clean_sentence) < 2:
                            continue
                        lemmas = _lemma_doc(clean_sentence)
                        lemma_text = " ".join(lemmas)
                        cur = conn.execute(
                            """
                            INSERT INTO sentences(
                                document_id, chunk_id, sentence_text, lemma_text, source_file, page_num,
                                test_label, section_label, part_label, speaker_label, question_label,
                                sentence_order, sort_key
                            ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)
                            """,
                            (
                                document_id,
                                chunk_id,
                                clean_sentence,
                                lemma_text,
                                os.path.basename(path),
                                block.get("page_num"),
                                block.get("test_label") or "",
                                block.get("section_label") or "",
                                block.get("part_label") or "",
                                block.get("speaker_label") or "",
                                block.get("question_label") or "",
                                sentence_order,
                                sort_key,
                            ),
                        )
                        sentence_id = int(cur.lastrowid)
                        for lemma in sorted(set(lemmas)):
                            if not lemma:
                                continue
                            conn.execute(
                                "INSERT OR IGNORE INTO sentence_lemmas(sentence_id, lemma) VALUES(?, ?)",
                                (sentence_id, lemma),
                            )
                        sentence_order += 1
                        sort_key += 1
                        document_sentences += 1
                        summary["sentences"] += 1
                    summary["chunks"] += 1
                summary["files"] += 1
                _finish_import_record(
                    conn,
                    import_id,
                    "completed",
                    chunks_indexed=document_chunks,
                    sentences_indexed=document_sentences,
                )
            except Exception as e:
                summary["errors"].append(f"{os.path.basename(path)}: {e}")
                try:
                    if import_id:
                        _finish_import_record(conn, import_id, "failed", error_message=str(e))
                except Exception:
                    pass
        conn.commit()
        return summary
    finally:
        conn.close()


def corpus_stats():
    stats = _fetch_stats()
    stats["nlp_mode"] = get_nlp_status()
    return stats


def list_documents(limit=200):
    return _fetch_documents(limit=limit)


def remove_document(document_path):
    return _remove_document_by_path(document_path)


def search_corpus(query, limit=50, document_path=None):
    ensure_schema()
    q = _clean_line(query)
    if not q:
        return {"query": "", "lemmas": [], "results": [], "nlp_mode": get_nlp_status()}

    lemmas = [lemma for lemma in _lemma_doc(q) if lemma]
    if not lemmas:
        return {"query": q, "lemmas": [], "results": [], "nlp_mode": get_nlp_status()}

    search_result = _search_sentence_rows(lemmas, limit=limit, document_path=_clean_line(document_path))
    results = []
    for item in search_result["rows"]:
        item["match_type"] = "lemma"
        item["highlight_ranges"] = _highlight_ranges(item.get("sentence_text") or "", q, lemmas)
        results.append(item)
    return {
        "query": q,
        "lemmas": lemmas,
        "results": results,
        "nlp_mode": get_nlp_status(),
        "document_path": search_result["document_path"],
    }
