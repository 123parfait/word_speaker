# -*- coding: utf-8 -*-
import os
import sqlite3
from datetime import datetime
from pathlib import Path


DB_PATH = Path(__file__).resolve().parent.parent / "data" / "corpus_index.db"


def get_connection():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def _table_columns(conn, table_name):
    rows = conn.execute(f"PRAGMA table_info({table_name})").fetchall()
    return {str(row["name"]) for row in rows}


def ensure_schema():
    conn = get_connection()
    try:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS documents (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                path TEXT NOT NULL UNIQUE,
                name TEXT NOT NULL,
                file_type TEXT NOT NULL,
                source_type TEXT NOT NULL DEFAULT '',
                file_hash TEXT,
                imported_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS chunks (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                document_id INTEGER NOT NULL,
                chunk_text TEXT NOT NULL,
                page_num INTEGER,
                test_label TEXT,
                section_label TEXT,
                part_label TEXT,
                speaker_label TEXT,
                question_label TEXT,
                sort_key INTEGER NOT NULL,
                FOREIGN KEY(document_id) REFERENCES documents(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS sentences (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                document_id INTEGER NOT NULL,
                chunk_id INTEGER,
                sentence_text TEXT NOT NULL,
                lemma_text TEXT NOT NULL,
                source_file TEXT NOT NULL,
                page_num INTEGER,
                test_label TEXT,
                section_label TEXT,
                part_label TEXT,
                speaker_label TEXT,
                question_label TEXT,
                sentence_order INTEGER NOT NULL DEFAULT 0,
                sort_key INTEGER NOT NULL,
                FOREIGN KEY(document_id) REFERENCES documents(id) ON DELETE CASCADE,
                FOREIGN KEY(chunk_id) REFERENCES chunks(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS sentence_lemmas (
                sentence_id INTEGER NOT NULL,
                lemma TEXT NOT NULL,
                FOREIGN KEY(sentence_id) REFERENCES sentences(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS imports (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                document_id INTEGER,
                source_path TEXT NOT NULL,
                status TEXT NOT NULL,
                error_message TEXT,
                chunks_indexed INTEGER NOT NULL DEFAULT 0,
                sentences_indexed INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                finished_at TEXT,
                FOREIGN KEY(document_id) REFERENCES documents(id) ON DELETE SET NULL
            );
            """
        )

        document_columns = _table_columns(conn, "documents")
        if "source_type" not in document_columns:
            conn.execute("ALTER TABLE documents ADD COLUMN source_type TEXT NOT NULL DEFAULT ''")

        sentence_columns = _table_columns(conn, "sentences")
        if "chunk_id" not in sentence_columns:
            conn.execute("ALTER TABLE sentences ADD COLUMN chunk_id INTEGER")
        if "sentence_order" not in sentence_columns:
            conn.execute("ALTER TABLE sentences ADD COLUMN sentence_order INTEGER NOT NULL DEFAULT 0")

        conn.executescript(
            """
            CREATE INDEX IF NOT EXISTS idx_sentences_document_id ON sentences(document_id);
            CREATE INDEX IF NOT EXISTS idx_sentences_chunk_id ON sentences(chunk_id);
            CREATE INDEX IF NOT EXISTS idx_chunks_document_id ON chunks(document_id);
            CREATE INDEX IF NOT EXISTS idx_sentence_lemmas_lemma ON sentence_lemmas(lemma);
            CREATE UNIQUE INDEX IF NOT EXISTS idx_sentence_lemmas_unique ON sentence_lemmas(sentence_id, lemma);
            CREATE INDEX IF NOT EXISTS idx_imports_document_id ON imports(document_id);
            """
        )

        conn.commit()
    finally:
        conn.close()


def infer_source_type(name, file_type):
    label = f"{str(name or '').lower()} {str(file_type or '').lower()}"
    if "audio" in label or "transcript" in label or "script" in label:
        return "ielts_transcript"
    if "reading" in label or "passage" in label:
        return "ielts_reading"
    return "generic"


def replace_document(conn, path, name, file_type, file_hash):
    row = conn.execute("SELECT id FROM documents WHERE path = ?", (path,)).fetchone()
    if row:
        document_id = int(row["id"])
        conn.execute(
            "DELETE FROM sentence_lemmas WHERE sentence_id IN (SELECT id FROM sentences WHERE document_id = ?)",
            (document_id,),
        )
        conn.execute("DELETE FROM sentences WHERE document_id = ?", (document_id,))
        conn.execute("DELETE FROM chunks WHERE document_id = ?", (document_id,))
        conn.execute(
            "UPDATE documents SET name = ?, file_type = ?, source_type = ?, file_hash = ?, imported_at = ? WHERE id = ?",
            (
                name,
                file_type,
                infer_source_type(name, file_type),
                file_hash,
                datetime.now().isoformat(timespec="seconds"),
                document_id,
            ),
        )
        return document_id

    cur = conn.execute(
        "INSERT INTO documents(path, name, file_type, source_type, file_hash, imported_at) VALUES(?,?,?,?,?,?)",
        (
            path,
            name,
            file_type,
            infer_source_type(name, file_type),
            file_hash,
            datetime.now().isoformat(timespec="seconds"),
        ),
    )
    return int(cur.lastrowid)


def create_import_record(conn, path, document_id):
    cur = conn.execute(
        """
        INSERT INTO imports(document_id, source_path, status, created_at)
        VALUES(?, ?, ?, ?)
        """,
        (
            document_id,
            path,
            "running",
            datetime.now().isoformat(timespec="seconds"),
        ),
    )
    return int(cur.lastrowid)


def finish_import_record(conn, import_id, status, chunks_indexed=0, sentences_indexed=0, error_message=""):
    conn.execute(
        """
        UPDATE imports
        SET status = ?, error_message = ?, chunks_indexed = ?, sentences_indexed = ?, finished_at = ?
        WHERE id = ?
        """,
        (
            status,
            error_message,
            int(chunks_indexed or 0),
            int(sentences_indexed or 0),
            datetime.now().isoformat(timespec="seconds"),
            int(import_id),
        ),
    )


def fetch_stats():
    ensure_schema()
    conn = get_connection()
    try:
        docs = conn.execute("SELECT COUNT(*) AS c FROM documents").fetchone()["c"]
        chunks = conn.execute("SELECT COUNT(*) AS c FROM chunks").fetchone()["c"]
        sentences = conn.execute("SELECT COUNT(*) AS c FROM sentences").fetchone()["c"]
        return {
            "documents": int(docs or 0),
            "chunks": int(chunks or 0),
            "sentences": int(sentences or 0),
        }
    finally:
        conn.close()


def fetch_documents(limit=200):
    ensure_schema()
    conn = get_connection()
    try:
        rows = conn.execute(
            """
            SELECT
                d.name,
                d.path,
                d.file_type,
                d.source_type,
                d.imported_at,
                COUNT(DISTINCT c.id) AS chunk_count,
                COUNT(DISTINCT s.id) AS sentence_count
            FROM documents d
            LEFT JOIN chunks c ON c.document_id = d.id
            LEFT JOIN sentences s ON s.document_id = d.id
            GROUP BY d.id
            ORDER BY d.imported_at DESC, d.name ASC
            LIMIT ?
            """,
            (int(limit),),
        ).fetchall()
        return [dict(row) for row in rows]
    finally:
        conn.close()


def remove_document_by_path(document_path):
    ensure_schema()
    target = os.path.abspath(str(document_path or "").strip())
    if not target:
        return 0
    conn = get_connection()
    try:
        row = conn.execute("SELECT id FROM documents WHERE path = ?", (target,)).fetchone()
        if not row:
            return 0
        document_id = int(row["id"])
        conn.execute(
            "DELETE FROM sentence_lemmas WHERE sentence_id IN (SELECT id FROM sentences WHERE document_id = ?)",
            (document_id,),
        )
        conn.execute("DELETE FROM sentences WHERE document_id = ?", (document_id,))
        conn.execute("DELETE FROM chunks WHERE document_id = ?", (document_id,))
        conn.execute("DELETE FROM documents WHERE id = ?", (document_id,))
        conn.commit()
        return 1
    finally:
        conn.close()


def search_sentence_rows(lemmas, limit=50, document_path=None):
    ensure_schema()
    conn = get_connection()
    try:
        doc_path = os.path.abspath(str(document_path or "").strip()) if document_path else ""
        if len(lemmas) == 1:
            sql = """
                SELECT
                    s.*,
                    c.chunk_text,
                    d.source_type,
                    d.path AS document_path,
                    d.imported_at
                FROM sentence_lemmas sl
                JOIN sentences s ON s.id = sl.sentence_id
                LEFT JOIN chunks c ON c.id = s.chunk_id
                LEFT JOIN documents d ON d.id = s.document_id
                WHERE sl.lemma = ?
            """
            params = [lemmas[0]]
            if doc_path:
                sql += " AND d.path = ?"
                params.append(doc_path)
            sql += """
                ORDER BY d.imported_at DESC, d.name ASC, s.page_num, s.sort_key, s.sentence_order
                LIMIT ?
            """
            params.append(int(limit))
            rows = conn.execute(sql, tuple(params)).fetchall()
        else:
            placeholders = ",".join("?" for _ in lemmas)
            sql = f"""
                WITH matched AS (
                    SELECT s.id
                    FROM sentences s
                    JOIN sentence_lemmas sl ON sl.sentence_id = s.id
                    LEFT JOIN documents d ON d.id = s.document_id
                    WHERE sl.lemma IN ({placeholders})
            """
            params = list(lemmas)
            if doc_path:
                sql += " AND d.path = ?"
                params.append(doc_path)
            sql += """
                    GROUP BY s.id
                    HAVING COUNT(DISTINCT sl.lemma) >= ?
                )
                SELECT
                    s.*,
                    c.chunk_text,
                    d.source_type,
                    d.path AS document_path,
                    d.imported_at
                FROM matched m
                JOIN sentences s ON s.id = m.id
                LEFT JOIN chunks c ON c.id = s.chunk_id
                LEFT JOIN documents d ON d.id = s.document_id
                ORDER BY d.imported_at DESC, d.name ASC, s.page_num, s.sort_key, s.sentence_order
                LIMIT ?
            """
            params.extend([len(set(lemmas)), int(limit)])
            rows = conn.execute(sql, tuple(params)).fetchall()
        return {
            "rows": [dict(row) for row in rows],
            "document_path": doc_path,
        }
    finally:
        conn.close()
