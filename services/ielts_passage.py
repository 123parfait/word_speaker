# -*- coding: utf-8 -*-
import random
import re


_SCENARIOS = [
    {
        "title": "University Orientation",
        "intro": "You will hear a student talking to an adviser during orientation week.",
        "detail": "They discuss registration, study support, and practical arrangements for the first month.",
        "closing": "At the end, the adviser repeats the key points to make sure nothing is missed.",
        "speaker_1": "Adviser",
        "speaker_2": "Student",
    },
    {
        "title": "Travel And Accommodation",
        "intro": "You will hear two students planning travel and accommodation for a short course.",
        "detail": "They compare transport choices, room options, and payment details.",
        "closing": "Before finishing, they summarize the final plan and confirm the checklist.",
        "speaker_1": "Student A",
        "speaker_2": "Student B",
    },
    {
        "title": "Community Project Briefing",
        "intro": "You will hear a briefing for new volunteers joining a local community project.",
        "detail": "The coordinator explains responsibilities, timing, and how information will be shared.",
        "closing": "The session ends with a short recap of priorities for the coming week.",
        "speaker_1": "Coordinator",
        "speaker_2": "Volunteer",
    },
]

_GROUP_SENTENCE_TEMPLATES = [
    "At this point, they connect {terms} to the timetable and room arrangement for the day.",
    "They then compare notes on {terms} so the final plan matches the booking details.",
    "In the middle of the discussion, {terms} come up when they check what has already been completed.",
    "The speaker revisits {terms} while explaining what should be written on the form.",
    "Later, {terms} are used to clarify the difference between the first option and the backup option.",
    "Before moving on, they verify {terms} to avoid confusion during the listening task.",
    "They also mention {terms} when outlining the steps for the next session.",
]


def _clean_word(word):
    text = str(word or "").strip()
    text = re.sub(r"\s+", " ", text)
    return text


def _unique_words(words):
    seen = set()
    result = []
    for word in words:
        w = _clean_word(word)
        if not w:
            continue
        key = w.casefold()
        if key in seen:
            continue
        seen.add(key)
        result.append(w)
    return result


def _pick_seed(words):
    # Stable seed for same word list, so the generated passage feels consistent.
    return sum(ord(ch) for ch in "|".join(words)) % 10_000_019


def _format_term(word):
    w = str(word).strip()
    if " " in w:
        return f'"{w}"'
    return w


def _join_terms(words):
    terms = [_format_term(w) for w in words if str(w).strip()]
    if not terms:
        return ""
    if len(terms) == 1:
        return terms[0]
    if len(terms) == 2:
        return f"{terms[0]} and {terms[1]}"
    return f"{', '.join(terms[:-1])}, and {terms[-1]}"


def _chunk_words(words, size):
    out = []
    for i in range(0, len(words), size):
        out.append(words[i : i + size])
    return out


def build_ielts_listening_passage(words, max_words=20):
    """
    Build a short IELTS-listening-style passage from imported words.

    Returns:
        dict with keys:
        - passage: str
        - used_words: list[str]
        - skipped_words: list[str]
        - scenario: str
    """
    unique = _unique_words(words)
    if not unique:
        raise ValueError("No valid words to build a passage.")

    used_words = unique[: max(1, int(max_words))]
    skipped_words = unique[len(used_words) :]
    rng = random.Random(_pick_seed(used_words))
    scenario = rng.choice(_SCENARIOS)

    groups = _chunk_words(used_words, 3)
    mid_lines = []
    for idx, group in enumerate(groups):
        terms = _join_terms(group)
        template = _GROUP_SENTENCE_TEMPLATES[idx % len(_GROUP_SENTENCE_TEMPLATES)]
        mid_lines.append(template.format(terms=terms))

    first_group = _join_terms(groups[0]) if groups else ""
    second_group = _join_terms(groups[1]) if len(groups) > 1 else ""
    third_group = _join_terms(groups[2]) if len(groups) > 2 else ""
    s1 = scenario.get("speaker_1", "Speaker 1")
    s2 = scenario.get("speaker_2", "Speaker 2")

    dialogue = [
        f"{s1}: Thanks for coming in early. We need to finalize today's plan.",
        f"{s2}: Sure, I have the notes ready and can update the checklist as we talk.",
    ]
    if first_group:
        dialogue.append(
            f"{s1}: The first items we should settle are {first_group}, because they affect the schedule."
        )
    if second_group:
        dialogue.append(
            f"{s2}: After that, we can move on to {second_group} and confirm the final details."
        )
    if third_group:
        dialogue.append(
            f"{s1}: Good. In the last part we should review {third_group} before we close the meeting."
        )
    dialogue.append(f"{s2}: Perfect. That sequence sounds clear and easy to follow.")

    passage = "\n\n".join(
        [
            f"IELTS Listening Practice - {scenario['title']}",
            scenario["intro"],
            scenario["detail"],
            "\n".join(dialogue),
            *mid_lines,
            scenario["closing"],
        ]
    )

    return {
        "passage": passage.strip(),
        "used_words": used_words,
        "skipped_words": skipped_words,
        "scenario": scenario["title"],
    }
