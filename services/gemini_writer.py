# -*- coding: utf-8 -*-
import json
import re
import urllib.error
import urllib.request


DEFAULT_GEMINI_MODEL = "gemini-2.5-flash"
GEMINI_VALIDATION_MODEL = "gemini-2.5-flash-lite"
GEMINI_MODELS = [
    "gemini-2.5-flash",
    "gemini-2.5-flash-lite",
]
GEMINI_OPENAI_URL = "https://generativelanguage.googleapis.com/v1beta/openai/chat/completions"


def list_available_gemini_models():
    return list(GEMINI_MODELS)


def choose_preferred_generation_model(models, fallback=DEFAULT_GEMINI_MODEL):
    values = [str(m).strip() for m in models if str(m).strip()]
    if not values:
        return str(fallback)
    if fallback in values:
        return str(fallback)
    for preferred in GEMINI_MODELS:
        if preferred in values:
            return preferred
    return values[0]


def _normalize_words(words, max_words=24):
    seen = set()
    result = []
    for word in words:
        w = re.sub(r"\s+", " ", str(word or "").strip())
        if not w:
            continue
        key = w.casefold()
        if key in seen:
            continue
        seen.add(key)
        result.append(w)
        if len(result) >= max_words:
            break
    return result


def _build_passage_prompt(words):
    terms = ", ".join(words)
    return (
        "You are generating an IELTS Listening style passage.\n\n"
        "Requirements:\n"
        "- English only\n"
        "- 120-170 words\n"
        "- Plain text only (no title, no markdown, no explanation)\n"
        "- Semi-formal spoken tone\n"
        "- Must sound like a real conversation or academic discussion\n"
        "- Include 8-12 sentences\n"
        "- Must include at least one number or specific detail\n"
        "- Must include at least 3 checkable details (date/time/room/fee/quantity/deadline)\n"
        "- Must involve a practical task (e.g., research project, data analysis, course planning, client discussion)\n\n"
        "Content rules:\n"
        "- Avoid abstract philosophical explanations\n"
        "- Avoid essay-style arguments\n"
        "- Focus on concrete information and decisions\n"
        "- Information should be suitable for IELTS listening questions (names, figures, processes, problems, solutions)\n\n"
        "Vocabulary rule:\n"
        "- Use at least 70% of the provided words naturally\n"
        "- Do not output a word list or explain usage\n\n"
        f"Provided words: {terms}\n"
    )


def _build_repair_prompt(passage, missing_words):
    missing = ", ".join(str(w) for w in missing_words if str(w).strip())
    return (
        "Revise the passage below.\n"
        "Requirements:\n"
        "- Keep it English only, plain text only.\n"
        "- Keep 120-170 words.\n"
        "- Keep the IELTS listening practical tone.\n"
        "- Keep it natural (not a list).\n"
        "- You MUST include all missing words exactly as written at least once.\n"
        "- Do not output explanation.\n\n"
        f"Missing words (must include all): {missing}\n\n"
        "Original passage:\n"
        f"{passage}\n"
    )


def _build_sentence_prompt(word):
    return (
        "You are an IELTS English coach.\n"
        "Write exactly one idiomatic, natural English sentence for the target word.\n"
        "Rules:\n"
        "- English only\n"
        "- Plain text only\n"
        "- One sentence only\n"
        "- 12-24 words\n"
        "- Target Band 7+ IELTS quality\n"
        "- Practical context (study/work/daily task)\n"
        "- Use the target word or phrase naturally and exactly once\n"
        "- Avoid dictionary definitions and generic filler\n"
        "- Prefer useful collocations and realistic wording\n"
        "- No quotation marks unless required by grammar\n\n"
        f"Target word: {word}\n"
    )


def _strip_code_fence(text):
    raw = str(text or "").strip()
    if raw.startswith("```"):
        raw = re.sub(r"^```[a-zA-Z0-9_-]*\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)
    return raw.strip()


def _trim_to_word_limit(text, max_words=190):
    tokens = re.findall(r"\S+", str(text or ""))
    if len(tokens) <= max_words:
        return str(text or "").strip()
    trimmed = " ".join(tokens[:max_words]).rstrip(" ,;:")
    if trimmed and trimmed[-1] not in ".!?":
        trimmed += "."
    return trimmed


def _expand_contractions(text):
    t = str(text or "").casefold().replace("\u2019", "'")
    rules = [
        (r"\bi'm\b", "i am"),
        (r"\byou're\b", "you are"),
        (r"\bwe're\b", "we are"),
        (r"\bthey're\b", "they are"),
        (r"\bit's\b", "it is"),
        (r"\bthat's\b", "that is"),
        (r"\bthere's\b", "there is"),
        (r"\bcan't\b", "can not"),
        (r"\bwon't\b", "will not"),
        (r"n't\b", " not"),
        (r"'re\b", " are"),
        (r"'ve\b", " have"),
        (r"'ll\b", " will"),
        (r"'d\b", " would"),
        (r"'m\b", " am"),
    ]
    for pattern, repl in rules:
        t = re.sub(pattern, repl, t)
    return t


def _normalize_match_text(text):
    t = _expand_contractions(text)
    t = re.sub(r"[^a-z0-9]+", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def _simple_stem(token):
    tok = str(token or "").strip()
    if len(tok) <= 3:
        return tok
    for suf in ("ingly", "edly", "ing", "ed", "ies"):
        if tok.endswith(suf):
            base = tok[: -len(suf)]
            if len(base) >= 3:
                if suf == "ies":
                    return base + "y"
                return base
    if tok.endswith("es") and len(tok) > 4:
        if tok.endswith(("ses", "xes", "zes", "ches", "shes", "oes")):
            base = tok[:-2]
            if len(base) >= 3:
                return base
    if tok.endswith("s") and len(tok) > 3 and not tok.endswith(("ss", "us", "is")):
        base = tok[:-1]
        if len(base) >= 3:
            return base
    return tok


def _keyword_match(keyword, text_norm, text_stems, text_stem_set):
    key_norm = _normalize_match_text(keyword)
    if not key_norm:
        return False
    if f" {key_norm} " in f" {text_norm} ":
        return True
    key_tokens = key_norm.split()
    key_stems = [_simple_stem(tok) for tok in key_tokens if tok]
    if not key_stems:
        return False
    n = len(key_stems)
    if n <= len(text_stems):
        for i in range(0, len(text_stems) - n + 1):
            if text_stems[i : i + n] == key_stems:
                return True
    return all(st in text_stem_set for st in key_stems)


def _coverage_details(text, words):
    if not words:
        return 0.0, [], []
    text_norm = _normalize_match_text(text)
    text_tokens = text_norm.split()
    text_stems = [_simple_stem(tok) for tok in text_tokens]
    text_stem_set = set(text_stems)
    matched = []
    missed = []
    for word in words:
        if _keyword_match(word, text_norm, text_stems, text_stem_set):
            matched.append(word)
        else:
            missed.append(word)
    ratio = len(matched) / float(len(words))
    return ratio, matched, missed


def _format_missing_term(term):
    t = str(term or "").strip()
    if not t:
        return ""
    return f'"{t}"' if " " in t else t


def _join_terms(terms):
    cleaned = [t for t in terms if t]
    if not cleaned:
        return ""
    if len(cleaned) == 1:
        return cleaned[0]
    if len(cleaned) == 2:
        return f"{cleaned[0]} and {cleaned[1]}"
    return f"{', '.join(cleaned[:-1])}, and {cleaned[-1]}"


def _inject_missing_words(text, missing_words):
    terms = [_format_missing_term(w) for w in missing_words if str(w).strip()]
    terms_text = _join_terms(terms)
    if not terms_text:
        return str(text or "").strip()
    sentence = f" For accuracy, the final checklist mentions {terms_text}."
    base = str(text or "").strip()
    if not base:
        return sentence.strip()
    if base[-1] not in ".!?":
        base += "."
    return base + sentence


def _extract_error_message(http_error):
    try:
        raw = http_error.read().decode("utf-8", errors="ignore")
    except Exception:
        return str(http_error)
    try:
        data = json.loads(raw)
    except Exception:
        return raw or str(http_error)
    error = data.get("error") or {}
    message = error.get("message") or data.get("message")
    return str(message or raw or http_error)


def _request_gemini(prompt, api_key, model, timeout, temperature, max_tokens):
    if not str(api_key or "").strip():
        raise RuntimeError("Gemini API key is empty.")
    payload = {
        "model": str(model),
        "messages": [
            {
                "role": "system",
                "content": "You write natural IELTS-quality English and must follow the user's format exactly.",
            },
            {"role": "user", "content": str(prompt)},
        ],
        "temperature": float(temperature),
        "max_tokens": int(max_tokens),
    }
    req = urllib.request.Request(
        GEMINI_OPENAI_URL,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {str(api_key).strip()}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as response:
            data = json.loads(response.read().decode("utf-8", errors="ignore"))
    except urllib.error.HTTPError as e:
        raise RuntimeError(_extract_error_message(e)) from e
    except urllib.error.URLError as e:
        raise RuntimeError(f"Gemini request failed: {e.reason}") from e
    except Exception as e:
        raise RuntimeError(f"Gemini request failed: {e}") from e

    choices = data.get("choices") or []
    if not choices:
        raise RuntimeError("Gemini returned no choices.")
    message = choices[0].get("message") or {}
    content = message.get("content")
    if isinstance(content, list):
        parts = []
        for item in content:
            if isinstance(item, dict):
                piece = item.get("text") or ""
                if piece:
                    parts.append(str(piece))
        content = "".join(parts)
    text = _strip_code_fence(content)
    if not text:
        raise RuntimeError("Gemini returned empty text.")
    return text


def validate_gemini_api_key(api_key, model=GEMINI_VALIDATION_MODEL, timeout=25):
    text = _request_gemini(
        prompt="Reply with OK only.",
        api_key=api_key,
        model=model,
        timeout=timeout,
        temperature=0.0,
        max_tokens=8,
    )
    if "ok" not in text.casefold():
        raise RuntimeError("Gemini key test did not return OK.")
    return True


def generate_english_passage_with_gemini(words, api_key, max_words=24, timeout=90, model=DEFAULT_GEMINI_MODEL):
    used_words = _normalize_words(words, max_words=max_words)
    if not used_words:
        raise ValueError("No valid words to generate passage.")

    text = _request_gemini(
        prompt=_build_passage_prompt(used_words),
        api_key=api_key,
        model=model,
        timeout=timeout,
        temperature=0.7,
        max_tokens=320,
    )
    text = _trim_to_word_limit(text, max_words=190)
    ratio, matched_words, missed_words = _coverage_details(text, used_words)
    repaired = False

    if ratio < 0.7 and missed_words:
        try:
            repaired_text = _request_gemini(
                prompt=_build_repair_prompt(text, missed_words),
                api_key=api_key,
                model=model,
                timeout=timeout,
                temperature=0.45,
                max_tokens=320,
            )
            repaired_text = _trim_to_word_limit(repaired_text, max_words=190)
            if repaired_text:
                ratio2, matched2, missed2 = _coverage_details(repaired_text, used_words)
                if ratio2 >= ratio:
                    text = repaired_text
                    ratio = ratio2
                    matched_words = matched2
                    missed_words = missed2
                    repaired = True
        except Exception:
            pass

    if ratio < 0.7 and missed_words:
        text = _inject_missing_words(text, missed_words)
        text = _trim_to_word_limit(text, max_words=190)
        ratio, matched_words, missed_words = _coverage_details(text, used_words)
        repaired = True

    return {
        "passage": text,
        "used_words": used_words,
        "model": str(model),
        "source": "gemini",
        "coverage": ratio,
        "matched_words": matched_words,
        "missed_words": missed_words,
        "low_coverage": ratio < 0.5,
        "repaired": repaired,
    }


def generate_example_sentence_with_gemini(word, api_key, model=DEFAULT_GEMINI_MODEL, timeout=45):
    target = re.sub(r"\s+", " ", str(word or "").strip())
    if not target:
        raise ValueError("Word is empty.")
    text = _request_gemini(
        prompt=_build_sentence_prompt(target),
        api_key=api_key,
        model=model,
        timeout=timeout,
        temperature=0.45,
        max_tokens=90,
    )
    text = re.sub(r"\s+", " ", text).strip()
    text = text.splitlines()[0].strip()
    parts = re.split(r"(?<=[.!?])\s+", text, maxsplit=1)
    sentence = parts[0].strip()
    if sentence and sentence[-1] not in ".!?":
        sentence += "."
    ratio, matched, _missed = _coverage_details(sentence, [target])
    if ratio <= 0.0 or not matched:
        raise RuntimeError("Generated sentence does not include the target word.")
    return sentence
