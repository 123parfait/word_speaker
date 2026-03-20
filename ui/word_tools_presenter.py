# -*- coding: utf-8 -*-


def build_sentence_view_state(word, sentence, source):
    return {
        "title": f"Sentence - {word}",
        "status_text": f"Sentence ready for '{word}'.",
        "word_label": f"Word: {word}",
        "source_label": f"Source: {source}",
        "sentence_text": sentence or "",
    }


def build_synonym_view_state(*, tr, trf, word, focus, synonyms, source=None):
    source_key = str(source or "").strip().lower()
    if source_key == "gemini":
        source_label = tr("synonyms_source_gemini")
    else:
        source_label = tr("synonyms_source_local")
    return {
        "title": f"{tr('synonyms_title')} - {word}",
        "status_text": trf("synonyms_ready", word=word),
        "word_label": f"Word: {word}",
        "focus_label": (trf("synonyms_focus", word=focus) if focus else ""),
        "source_label": trf("synonyms_source", source=source_label),
        "synonym_text": ("\n".join(f"- {item}" for item in synonyms) if synonyms else tr("no_synonyms_found")),
    }
