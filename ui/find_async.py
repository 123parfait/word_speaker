# -*- coding: utf-8 -*-
import threading

from services.corpus_search import import_corpus_files, search_corpus


def start_find_import_task(*, paths, token, emit_event):
    def _run():
        try:
            result = import_corpus_files(paths)
            emit_event("import_done", token, result)
        except Exception as exc:
            emit_event("error", token, str(exc))
        emit_event("done", token, None)

    threading.Thread(target=_run, daemon=True).start()


def start_find_search_task(*, query, limit, document_path, token, emit_event):
    def _run():
        try:
            result = search_corpus(
                query,
                limit=limit,
                document_path=document_path,
            )
            emit_event("search_done", token, result)
        except Exception as exc:
            emit_event("error", token, str(exc))
        emit_event("done", token, None)

    threading.Thread(target=_run, daemon=True).start()
