# -*- coding: utf-8 -*-
import re


def normalize_import_word_text(text):
    value = str(text or "").strip()
    if not value:
        return ""
    value = value.translate(
        {
            0x2018: ord("'"),
            0x2019: ord("'"),
            0x201B: ord("'"),
            0x2032: ord("'"),
            0x00B4: ord("'"),
            0x2010: ord("-"),
            0x2011: ord("-"),
            0x2012: ord("-"),
            0x2013: ord("-"),
            0x2014: ord("-"),
            0x2015: ord("-"),
            0x2212: ord("-"),
        }
    )
    return re.sub(r"\s+", " ", value).strip()


def looks_like_contextual_phrase_word_line(text, next_line=""):
    value = str(text or "").strip()
    if not value:
        return False
    next_value = str(next_line or "").strip()
    if not next_value or not re.search(r"[\u4e00-\u9fff]", next_value):
        return False
    if re.search(r"[.!?;:，；。！？：]", value):
        return False
    allowed_non_ascii = {0x00A3, 0x20AC, 0x00A5}
    if any(ord(ch) > 127 and ord(ch) not in allowed_non_ascii for ch in value):
        return False
    sanitized = value.translate(
        {
            0x00A3: ord(" "),
            0x20AC: ord(" "),
            0x00A5: ord(" "),
            0x0024: ord(" "),
        }
    )
    if re.search(r"[^A-Za-z0-9 '\-()/&]", sanitized):
        return False
    words = re.findall(r"[A-Za-z]+(?:['-][A-Za-z]+)?", value)
    if not words or len(words) > 8:
        return False
    letters = sum(len(token) for token in words)
    if letters <= 0:
        return False
    if len(value) > max(letters + 18, 56):
        return False
    return True


def looks_like_word_line(text, next_line=""):
    value = str(text or "").strip()
    if not value:
        return False
    if re.search(r"[\u4e00-\u9fff]", value):
        return False
    if any(ord(ch) > 127 for ch in value):
        return looks_like_contextual_phrase_word_line(value, next_line=next_line)
    if re.search(r"[.!?;:，；。！？：]", value):
        return False
    if re.search(r"[^A-Za-z0-9 '\-()/&]", value):
        return looks_like_contextual_phrase_word_line(value, next_line=next_line)
    words = re.findall(r"[A-Za-z]+(?:['-][A-Za-z]+)?", value)
    if not words:
        return False
    if len(words) > 5:
        return looks_like_contextual_phrase_word_line(value, next_line=next_line)
    letters = sum(len(token) for token in words)
    if letters <= 0:
        return False
    if len(value) > max(letters + 10, 36):
        return looks_like_contextual_phrase_word_line(value, next_line=next_line)
    return True


def parse_manual_rows(raw_text):
    text = str(raw_text or "")
    lines = text.splitlines()
    rows = []
    pending_word = None
    pending_note_lines = []
    total_lines = len(lines)
    for idx, raw_line in enumerate(lines):
        raw_line_text = str(raw_line or "")
        line = raw_line_text.strip()
        normalized_line = normalize_import_word_text(line)
        has_leading_indent = bool(raw_line_text[: len(raw_line_text) - len(raw_line_text.lstrip())])
        next_nonempty_line = ""
        for follow_idx in range(idx + 1, total_lines):
            candidate = str(lines[follow_idx] or "").strip()
            if candidate:
                next_nonempty_line = candidate
                break
        line = line.strip()
        if not line:
            continue
        if "\t" in line:
            if pending_word:
                rows.append({"word": pending_word, "note": " | ".join(pending_note_lines).strip()})
                pending_word = None
                pending_note_lines = []
            parts = [part.strip() for part in line.split("\t")]
            word = normalize_import_word_text(parts[0] or "")
            note = " | ".join(part for part in parts[1:] if str(part).strip())
            if word:
                rows.append({"word": word, "note": note})
            continue
        if pending_word and has_leading_indent:
            pending_note_lines.append(line)
            continue
        if looks_like_word_line(normalized_line, next_line=next_nonempty_line):
            if pending_word:
                rows.append({"word": pending_word, "note": " | ".join(pending_note_lines).strip()})
            pending_word = normalized_line
            pending_note_lines = []
            continue

        if pending_word:
            pending_note_lines.append(line)
        else:
            parts = re.split(r"[,;；，]+", line)
            if len(parts) > 1 and all(str(part or "").strip() for part in parts):
                for part in parts:
                    token = normalize_import_word_text(part or "")
                    if token:
                        rows.append({"word": token, "note": ""})
            else:
                rows.append({"word": normalized_line or line, "note": ""})

    if pending_word:
        rows.append({"word": pending_word, "note": " | ".join(pending_note_lines).strip()})
    return rows


def extract_clipboard_html_fragment(raw_html):
    text = str(raw_html or "")
    if not text:
        return ""
    start_match = re.search(r"StartFragment:(\d+)", text)
    end_match = re.search(r"EndFragment:(\d+)", text)
    if start_match and end_match:
        try:
            start = int(start_match.group(1))
            end = int(end_match.group(1))
            fragment = text[start:end]
            if fragment.strip():
                return fragment
        except Exception:
            pass
    marker_start = "<!--StartFragment-->"
    marker_end = "<!--EndFragment-->"
    if marker_start in text and marker_end in text:
        start = text.index(marker_start) + len(marker_start)
        end = text.index(marker_end, start)
        fragment = text[start:end]
        if fragment.strip():
            return fragment
    return text


def parse_clipboard_html_rows(raw_html, *, table_parser_cls):
    fragment = extract_clipboard_html_fragment(raw_html)
    if "<table" not in fragment.lower():
        return []
    parser = table_parser_cls()
    try:
        parser.feed(fragment)
        parser.close()
    except Exception:
        return []
    rows = []
    for cells in parser.rows:
        if len(cells) < 2:
            continue
        word = normalize_import_word_text(cells[0] or "")
        if not word:
            continue
        note_parts = []
        for cell in cells[1:]:
            for line in str(cell or "").replace("\r", "\n").split("\n"):
                clean_line = re.sub(r"[ \t]+", " ", str(line or "")).strip()
                if clean_line:
                    note_parts.append(clean_line)
        note = " | ".join(note_parts).strip()
        rows.append({"word": word, "note": note})
    return rows


def parse_tabular_text_rows(raw_text):
    text = str(raw_text or "").replace("\r\n", "\n").replace("\r", "\n")
    if "\t" not in text:
        return []
    rows = []
    current = None
    for raw_line in text.split("\n"):
        line = str(raw_line or "")
        if not line.strip():
            continue
        if "\t" in line:
            parts = [part.strip() for part in line.split("\t")]
            word = normalize_import_word_text(parts[0] or "")
            if not word:
                current = None
                continue
            note = " | ".join(str(part or "").strip() for part in parts[1:] if str(part or "").strip()).strip()
            current = {"word": word, "note": note}
            rows.append(current)
            continue
        if current is not None:
            extra = str(line or "").strip()
            if extra:
                current["note"] = " | ".join(part for part in [current.get("note") or "", extra] if part).strip()
    return rows


def read_clipboard_import_rows(*, html_text, raw_text, table_parser_cls):
    html_rows = parse_clipboard_html_rows(html_text, table_parser_cls=table_parser_cls)
    if html_rows:
        return html_rows
    table_rows = parse_tabular_text_rows(raw_text)
    if table_rows:
        return table_rows
    return parse_manual_rows(raw_text)


def normalize_manual_input_rows(rows):
    normalized_words = []
    normalized_notes = []
    for row in rows:
        if isinstance(row, dict):
            word = normalize_import_word_text(row.get("word") or "")
            note = re.sub(r"\s+", " ", str(row.get("note") or "").strip())
        else:
            word = normalize_import_word_text(row or "")
            note = ""
        if not word:
            continue
        normalized_words.append(word)
        normalized_notes.append(note)
    return normalized_words, normalized_notes
