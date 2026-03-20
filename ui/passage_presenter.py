# -*- coding: utf-8 -*-
import re


def speech_text_from_passage(text):
    lines = [line.strip() for line in str(text or "").splitlines()]
    if lines and lines[0].lower().startswith("ielts listening practice -"):
        lines = lines[1:]
    compact = "\n".join(line for line in lines if line)
    return compact.strip()


def normalize_answer(text):
    return re.sub(r"\s+", " ", str(text or "").strip().casefold())


def replace_first_case_insensitive(text, target, repl):
    if not target:
        return text, False
    pattern = re.compile(re.escape(str(target)), re.IGNORECASE)
    match = pattern.search(text)
    if not match:
        return text, False
    return text[: match.start()] + repl + text[match.end() :], True


def build_cloze_passage(passage, keywords, max_blanks=12):
    text = str(passage or "").strip()
    if not text:
        return "", []
    cloze = text
    answers = []
    for word in keywords:
        if len(answers) >= max_blanks:
            break
        key = str(word or "").strip()
        if not key:
            continue
        cloze, replaced = replace_first_case_insensitive(cloze, key, "____")
        if replaced:
            answers.append(key)
    return cloze, answers


def build_passage_practice_state(*, current_passage_original, current_passage, current_passage_words, store_words):
    source_text = str(current_passage_original or "").strip() or str(current_passage or "").strip()
    if not source_text:
        return {"error": "missing_passage"}
    keywords = list(current_passage_words) if current_passage_words else list(store_words or [])
    cloze, answers = build_cloze_passage(source_text, keywords, max_blanks=12)
    if not answers:
        return {"error": "no_keywords"}
    status_text = f"Practice ready: {len(answers)} blanks. Fill one answer per line, then click Check."
    return {
        "cloze": cloze,
        "answers": answers,
        "status_text": status_text,
    }


def build_passage_practice_check_state(*, answers, user_lines):
    correct = 0
    for idx, answer in enumerate(answers or []):
        user_value = user_lines[idx] if idx < len(user_lines) else ""
        if normalize_answer(user_value) == normalize_answer(answer):
            correct += 1
    return {
        "expected_text": "\n".join(answers or []),
        "actual_text": "\n".join(user_lines or []),
        "status_text": f"Practice check: {correct}/{len(answers or [])} correct.",
    }


def build_partial_passage_state(text):
    passage = str(text or "").strip()
    return {
        "passage": passage,
        "has_passage": bool(passage),
    }


def build_generated_passage_state(result, *, default_model):
    passage = str((result or {}).get("passage", "")).strip()
    used_words = list((result or {}).get("used_words") or [])
    source = (result or {}).get("source")

    if source == "gemini":
        coverage = int(float((result or {}).get("coverage", 0.0)) * 100)
        model = (result or {}).get("model") or default_model
        missed_words = list((result or {}).get("missed_words") or [])
        suffix = " (low coverage)" if (result or {}).get("low_coverage") else ""
        repaired_suffix = " | repaired" if (result or {}).get("repaired") else ""
        missed_suffix = f" | missed: {', '.join(missed_words[:3])}" if missed_words else ""
        status_text = (
            f"Gemini ({model}) | used {len(used_words)} words | "
            f"coverage {coverage}%{suffix}{repaired_suffix}{missed_suffix}."
        )
    else:
        skipped_count = len((result or {}).get("skipped_words") or [])
        scenario = (result or {}).get("scenario") or "General"
        reason = (result or {}).get("fallback_reason") or "Gemini not available."
        extra = f", skipped {skipped_count}" if skipped_count else ""
        status_text = f"Template fallback ({scenario}) | used {len(used_words)} words{extra}. Reason: {reason}"

    return {
        "passage": passage,
        "used_words": used_words,
        "status_text": status_text,
    }


def build_passage_audio_status(runtime):
    return f"Generating passage audio with {runtime}..."
