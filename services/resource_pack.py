# -*- coding: utf-8 -*-
import json
import os
import zipfile
from datetime import datetime, timezone
from pathlib import Path


RESOURCE_PACK_KIND = "wordspeaker.word_resource_pack"
RESOURCE_PACK_VERSION = 1
MANIFEST_FILE = "manifest.json"
WORDS_FILE = "words.json"


def _clean_text(value):
    return str(value or "").strip()


def _normalize_entry(raw_entry):
    if not isinstance(raw_entry, dict):
        return None
    word = _clean_text(raw_entry.get("word"))
    if not word:
        return None
    return {
        "word": word,
        "note": _clean_text(raw_entry.get("note")),
        "translation": _clean_text(raw_entry.get("translation")),
        "pos": _clean_text(raw_entry.get("pos")),
    }


def _normalize_entries(entries):
    normalized = []
    for item in entries or []:
        entry = _normalize_entry(item)
        if entry is not None:
            normalized.append(entry)
    return normalized


def export_word_resource_pack(package_path, entries, metadata=None):
    target_path = Path(os.path.abspath(str(package_path or "").strip()))
    if not str(target_path):
        raise ValueError("Resource pack output path is empty.")
    if target_path.parent:
        target_path.parent.mkdir(parents=True, exist_ok=True)

    normalized_entries = _normalize_entries(entries)
    manifest = {
        "kind": RESOURCE_PACK_KIND,
        "version": RESOURCE_PACK_VERSION,
        "exported_at": datetime.now(timezone.utc).isoformat(),
        "entry_count": len(normalized_entries),
    }
    meta = metadata if isinstance(metadata, dict) else {}
    cleaned_meta = {}
    for key, value in meta.items():
        name = _clean_text(key)
        if not name:
            continue
        text = _clean_text(value)
        if text:
            cleaned_meta[name] = text
    if cleaned_meta:
        manifest["metadata"] = cleaned_meta

    with zipfile.ZipFile(target_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(MANIFEST_FILE, json.dumps(manifest, ensure_ascii=False, indent=2))
        zf.writestr(WORDS_FILE, json.dumps(normalized_entries, ensure_ascii=False, indent=2))

    translation_count = sum(1 for entry in normalized_entries if entry.get("translation"))
    pos_count = sum(1 for entry in normalized_entries if entry.get("pos"))
    note_count = sum(1 for entry in normalized_entries if entry.get("note"))
    return {
        "path": str(target_path),
        "manifest": manifest,
        "entries": normalized_entries,
        "entry_count": len(normalized_entries),
        "translation_count": translation_count,
        "pos_count": pos_count,
        "note_count": note_count,
    }


def import_word_resource_pack(package_path):
    source_path = Path(os.path.abspath(str(package_path or "").strip()))
    if not source_path.exists() or not source_path.is_file():
        raise FileNotFoundError("Resource pack not found.")

    with zipfile.ZipFile(source_path, "r") as zf:
        try:
            manifest = json.loads(zf.read(MANIFEST_FILE).decode("utf-8"))
        except KeyError as exc:
            raise RuntimeError("Resource pack is missing manifest.json.") from exc
        except Exception as exc:
            raise RuntimeError("Failed to read resource pack manifest.") from exc
        if not isinstance(manifest, dict):
            raise RuntimeError("Resource pack manifest is invalid.")
        if str(manifest.get("kind") or "").strip() != RESOURCE_PACK_KIND:
            raise RuntimeError("This file is not a Word Speaker resource pack.")
        try:
            payload = json.loads(zf.read(WORDS_FILE).decode("utf-8"))
        except KeyError as exc:
            raise RuntimeError("Resource pack is missing words.json.") from exc
        except Exception as exc:
            raise RuntimeError("Failed to read resource pack words.") from exc

    if isinstance(payload, dict):
        payload = payload.get("entries")
    entries = _normalize_entries(payload if isinstance(payload, list) else [])
    translation_count = sum(1 for entry in entries if entry.get("translation"))
    pos_count = sum(1 for entry in entries if entry.get("pos"))
    note_count = sum(1 for entry in entries if entry.get("note"))
    return {
        "path": str(source_path),
        "manifest": manifest,
        "entries": entries,
        "entry_count": len(entries),
        "translation_count": translation_count,
        "pos_count": pos_count,
        "note_count": note_count,
    }
