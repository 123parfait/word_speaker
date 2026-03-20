# -*- coding: utf-8 -*-
import re


_NLP = None
_NLP_MODE = "en_core_web_sm"


def get_nlp():
    global _NLP, _NLP_MODE
    if _NLP is not None:
        return _NLP, _NLP_MODE

    try:
        import spacy
    except Exception as e:
        raise RuntimeError(
            "spaCy is not installed. Run: pip install spacy spacy-lookups-data"
        ) from e

    try:
        _NLP = spacy.load("en_core_web_sm")
        return _NLP, _NLP_MODE
    except Exception as load_error:
        try:
            import en_core_web_sm

            _NLP = en_core_web_sm.load()
            return _NLP, _NLP_MODE
        except Exception as import_error:
            raise RuntimeError(
                "spaCy model 'en_core_web_sm' is missing. Run: python -m spacy download en_core_web_sm"
            ) from import_error if import_error else load_error


def get_nlp_status():
    _nlp, mode = get_nlp()
    return mode


def clean_line(text):
    return re.sub(r"\s+", " ", str(text or "").strip())


def looks_like_margin_label(text, x0, page_width):
    text = clean_line(text)
    upper = text.upper()
    if not text or page_width in (None, 0):
        return False
    if x0 is None or float(x0) < float(page_width) * 0.72:
        return False
    return bool(
        re.fullmatch(r"Q\d{1,2}", upper)
        or upper == "EXAMPLE"
        or re.fullmatch(r"PAGE\s+\d+", upper)
    )


def looks_like_speaker_label(text):
    return bool(re.fullmatch(r"([A-Z][A-Z '\-]{1,30}|[A-Za-z][A-Za-z '\-]{1,30}):", clean_line(text)))


def clean_pdf_block_text(text, x0, page_width):
    parts = []
    for raw_line in str(text or "").splitlines():
        line = clean_line(raw_line)
        if not line:
            continue
        if looks_like_margin_label(line, x0, page_width):
            continue
        if looks_like_speaker_label(line):
            continue
        if re.fullmatch(r"\d{1,4}", line):
            continue
        parts.append(line)
    clean_text = clean_line(" ".join(parts))
    clean_text = re.sub(r"\s+\bQ\d{1,2}\b\s*$", "", clean_text, flags=re.IGNORECASE)
    clean_text = re.sub(r"\s+\bEXAMPLE\b\s*$", "", clean_text, flags=re.IGNORECASE)
    return clean_line(clean_text)


def iter_docx_blocks(path):
    try:
        from docx import Document
        from docx.document import Document as DocxDocument
        from docx.table import Table
        from docx.text.paragraph import Paragraph
    except Exception as e:
        raise RuntimeError("python-docx is not installed. Run: pip install python-docx") from e

    document = Document(path)

    def _iter_block_items(parent):
        parent_elm = parent.element.body if isinstance(parent, DocxDocument) else parent._tc
        for child in parent_elm.iterchildren():
            if child.tag.endswith("}p"):
                yield Paragraph(child, parent)
            elif child.tag.endswith("}tbl"):
                yield Table(child, parent)

    for block in _iter_block_items(document):
        if isinstance(block, Paragraph):
            text = clean_line(block.text)
            if text:
                yield {"text": text, "page_num": None}
        elif isinstance(block, Table):
            for row in block.rows:
                cells = [clean_line(cell.text) for cell in row.cells]
                cells = [c for c in cells if c]
                if not cells:
                    continue
                if len(cells) == 1:
                    yield {"text": cells[0], "page_num": None}
                else:
                    yield {"text": "\t".join(cells[:2]), "page_num": None}


def iter_pdf_blocks(path):
    try:
        import fitz
    except Exception as e:
        raise RuntimeError("PyMuPDF is not installed. Run: pip install pymupdf") from e

    with fitz.open(path) as doc:
        for page_index, page in enumerate(doc, start=1):
            page_width = float(page.rect.width or 0)
            text_data = page.get_text("dict")
            text_blocks = []
            for block in text_data.get("blocks", []):
                if block.get("type") != 0:
                    continue
                bbox = block.get("bbox") or (0, 0, 0, 0)
                x0, y0 = float(bbox[0] or 0), float(bbox[1] or 0)
                lines = []
                for line in block.get("lines", []):
                    spans = [span.get("text", "") for span in line.get("spans", [])]
                    line_text = clean_line("".join(spans))
                    if line_text:
                        lines.append(line_text)
                cleaned = clean_pdf_block_text("\n".join(lines), x0, page_width)
                if not cleaned:
                    continue
                text_blocks.append(
                    {
                        "text": cleaned,
                        "page_num": page_index,
                        "x0": x0,
                        "page_width": page_width,
                        "y0": y0,
                    }
                )
            text_blocks.sort(key=lambda item: (-item.get("y0", 0), item.get("x0", 0)))
            for item in text_blocks:
                yield item


def iter_text_blocks(path):
    with open(path, "r", encoding="utf-8", errors="ignore") as fp:
        for raw_line in fp:
            line = clean_line(raw_line)
            if line:
                yield {"text": line, "page_num": None}


def iter_file_blocks(path):
    lower = str(path).lower()
    if lower.endswith(".docx"):
        return list(iter_docx_blocks(path))
    if lower.endswith(".pdf"):
        return list(iter_pdf_blocks(path))
    return list(iter_text_blocks(path))


def is_continuation_line(raw_text, block, last_item, context):
    if not last_item:
        return False
    if block.get("page_num") != last_item.get("page_num"):
        return False
    if context["test_label"] != last_item.get("test_label"):
        return False
    if context["section_label"] != last_item.get("section_label"):
        return False
    if re.match(r"^([A-Z][A-Z '\-]{1,30}|[A-Za-z][A-Za-z '\-]{1,30}):\s*(.+)$", raw_text):
        return False
    if looks_like_speaker_label(raw_text):
        return False

    prev_text = clean_line(last_item.get("text"))
    if not prev_text:
        return False
    starts_lower = raw_text[:1].islower()
    prev_incomplete = prev_text[-1:] not in ".!?"
    if not (starts_lower or prev_incomplete):
        return False

    x0 = block.get("x0")
    page_width = block.get("page_width")
    last_x0 = last_item.get("_x0")
    if x0 is not None and page_width:
        if float(x0) < float(page_width) * 0.10:
            return False
        if last_x0 is not None and abs(float(x0) - float(last_x0)) > float(page_width) * 0.18 and not starts_lower:
            return False

    return True


def parse_structured_blocks(blocks):
    parsed = []
    context = {
        "test_label": "",
        "section_label": "",
        "part_label": "",
        "speaker_label": "",
        "question_label": "",
    }
    last_item = None
    pending_speaker = ""
    for block in blocks:
        raw_text = clean_line(block.get("text"))
        page_num = block.get("page_num")
        if not raw_text:
            continue

        upper = raw_text.upper()
        if re.fullmatch(r"TEST\s+\d+", upper):
            context["test_label"] = raw_text
            context["part_label"] = ""
            context["speaker_label"] = ""
            pending_speaker = ""
            continue
        if re.fullmatch(r"SECTION\s+\d+", upper) or re.fullmatch(r"PASSAGE\s+\d+", upper):
            context["section_label"] = raw_text
            context["speaker_label"] = ""
            pending_speaker = ""
            continue
        if len(raw_text) <= 40 and not re.search(r"[.!?]", raw_text) and raw_text == raw_text.title():
            context["part_label"] = raw_text
            continue

        question_match = re.search(r"\bQ\d+\b", raw_text, flags=re.IGNORECASE)
        if question_match:
            context["question_label"] = question_match.group(0).upper()
            raw_text = clean_line(raw_text.replace(question_match.group(0), ""))
            if not raw_text:
                continue

        speaker_only_match = re.fullmatch(r"([A-Z][A-Z '\-]{1,30}|[A-Za-z][A-Za-z '\-]{1,30}):", raw_text)
        if speaker_only_match:
            pending_speaker = clean_line(speaker_only_match.group(1))
            continue

        speaker_match = re.match(r"^([A-Z][A-Z '\-]{1,30}|[A-Za-z][A-Za-z '\-]{1,30}):\s*(.+)$", raw_text)
        if speaker_match:
            speaker = clean_line(speaker_match.group(1))
            text = clean_line(speaker_match.group(2))
            item = {
                "text": text,
                "page_num": page_num,
                "test_label": context["test_label"],
                "section_label": context["section_label"],
                "part_label": context["part_label"],
                "speaker_label": speaker,
                "question_label": context["question_label"],
                "_x0": block.get("x0"),
            }
            parsed.append(item)
            last_item = item
            pending_speaker = ""
            continue

        if is_continuation_line(raw_text, block, last_item, context):
            last_item["text"] = clean_line(f"{last_item['text']} {raw_text}")
            continue

        item = {
            "text": raw_text,
            "page_num": page_num,
            "test_label": context["test_label"],
            "section_label": context["section_label"],
            "part_label": context["part_label"],
            "speaker_label": pending_speaker,
            "question_label": context["question_label"],
            "_x0": block.get("x0"),
        }
        parsed.append(item)
        last_item = item
        pending_speaker = ""
    return parsed


def doc_sentences(text):
    nlp, _mode = get_nlp()
    doc = nlp(text)
    raw_sents = [sent.text.strip() for sent in doc.sents if sent.text and sent.text.strip()]
    if raw_sents:
        merged = []
        for sent in raw_sents:
            if merged:
                prev = merged[-1]
                if sent[:1].islower() or prev[-1:] not in ".!?":
                    merged[-1] = clean_line(f"{prev} {sent}")
                    continue
            merged.append(sent)
        return merged
    return [text.strip()] if str(text or "").strip() else []


def lemma_doc(text):
    nlp, _mode = get_nlp()
    doc = nlp(text)
    lemmas = []
    for token in doc:
        if token.is_space or token.is_punct:
            continue
        lemma = str(getattr(token, "lemma_", "") or "").strip().lower()
        if lemma in ("-pron-", ""):
            lemma = str(token.text or "").strip().lower()
        lemmas.append(lemma)
    return lemmas


def highlight_ranges(text, query, lemmas):
    clean_text = str(text or "")
    clean_query = clean_line(query)
    target_lemmas = {str(lemma or "").strip().lower() for lemma in (lemmas or []) if str(lemma or "").strip()}
    ranges = []

    if clean_query:
        for match in re.finditer(re.escape(clean_query), clean_text, flags=re.IGNORECASE):
            ranges.append((match.start(), match.end()))

    if target_lemmas:
        nlp, _mode = get_nlp()
        doc = nlp(clean_text)
        for token in doc:
            if token.is_space or token.is_punct:
                continue
            lemma = str(getattr(token, "lemma_", "") or "").strip().lower()
            if lemma in ("-pron-", ""):
                lemma = str(token.text or "").strip().lower()
            if lemma in target_lemmas:
                ranges.append((int(token.idx), int(token.idx) + len(token.text)))

    if not ranges:
        return []

    ranges.sort(key=lambda item: (item[0], item[1]))
    merged = [list(ranges[0])]
    for start, end in ranges[1:]:
        last = merged[-1]
        if start <= last[1]:
            last[1] = max(last[1], end)
        else:
            merged.append([start, end])
    return [(start, end) for start, end in merged]
