# -*- coding: utf-8 -*-


def refresh_word_table_rows(*, table, words, notes, build_values):
    if not table:
        return
    for idx, word in enumerate(words or []):
        row_id = str(idx)
        if not table.exists(row_id):
            continue
        note = notes[idx] if idx < len(notes or []) else ""
        tag = "even" if idx % 2 == 0 else "odd"
        table.item(row_id, values=build_values(idx, word, note), tags=(tag,))
