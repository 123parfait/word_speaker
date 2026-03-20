import os
import shutil


def _cache_group_key(cache_path, metadata, *, normalize_text):
    stem = os.path.splitext(os.path.basename(str(cache_path or "").strip()))[0]
    base = stem.rsplit("_", 1)[0] if "_" in stem else stem
    filename_guess = normalize_text(str(base or "").replace("_", " "), ensure_sentence_end=False)
    if filename_guess:
        return filename_guess.rstrip(".!?;:").casefold()
    text_value = normalize_text((metadata or {}).get("text"), ensure_sentence_end=False)
    if text_value:
        return text_value.rstrip(".!?;:").casefold()
    return str(base or "").rstrip(".!?;:").casefold()


def _infer_text_from_cache_filename(cache_path, metadata=None, *, normalize_text):
    text_value = normalize_text((metadata or {}).get("text"), ensure_sentence_end=False)
    if text_value:
        return text_value
    stem = os.path.splitext(os.path.basename(str(cache_path or "").strip()))[0]
    base = stem.rsplit("_", 1)[0] if "_" in stem else stem
    guess = str(base or "").replace("_", " ").strip()
    return normalize_text(guess, ensure_sentence_end=False)


def _cache_group_source_bucket(cache_path, *, source_word_cache_root_dir):
    try:
        rel_path = os.path.relpath(str(cache_path or "").strip(), source_word_cache_root_dir)
    except Exception:
        return ""
    parts = rel_path.split(os.sep)
    return parts[0] if parts else ""


def _cache_sort_key(
    cache_path,
    metadata,
    *,
    is_pending_gemini,
    resolve_cache_audio_path,
    cache_meta_path,
    is_online_backend,
):
    metadata = dict(metadata or {})
    desired_backend = str(metadata.get("desired_backend") or "").strip().lower()
    backend = str(metadata.get("backend") or "").strip().lower()
    updated_at = int(metadata.get("updated_at") or 0)
    pending = 1 if is_pending_gemini(cache_path) else 0
    playable = 1 if resolve_cache_audio_path(cache_path) else 0
    actual_wav = 1 if os.path.exists(cache_path) else 0
    try:
        latest_mtime = max(
            os.path.getmtime(path)
            for path in (cache_path, cache_meta_path(cache_path))
            if os.path.exists(path)
        )
    except Exception:
        latest_mtime = 0
    return (
        1 if is_online_backend(desired_backend) else 0,
        1 if backend == desired_backend and is_online_backend(backend) else 0,
        pending,
        playable,
        actual_wav,
        updated_at,
        latest_mtime,
    )


def migrate_flat_root_cache_layout(
    *,
    legacy_word_cache_wrapper_dir,
    global_word_cache_dir,
    shared_word_cache_dir,
    source_word_cache_root_dir,
    load_json_file,
    cache_meta_path,
    normalize_source_path,
    filename_letter_bucket,
    source_bucket_name,
    write_json_file,
    pending_gemini_queue_path,
    log_warning,
):
    renamed_paths = {}
    scan_roots = []
    for root in (legacy_word_cache_wrapper_dir, global_word_cache_dir):
        if os.path.isdir(root) and root not in scan_roots:
            scan_roots.append(root)

    for scan_root in scan_roots:
        entries = list(os.listdir(scan_root))
        for name in entries:
            full_path = os.path.join(scan_root, name)
            if os.path.isdir(full_path):
                continue
            if name.lower().endswith(".wav"):
                cache_path = full_path
            elif name.lower().endswith(".wav.json"):
                cache_path = full_path[:-5]
            else:
                continue

            cache_name = os.path.basename(cache_path)
            metadata = load_json_file(cache_meta_path(cache_path), {})
            source_hint = normalize_source_path((metadata or {}).get("source_path"))
            if source_hint == "shared":
                target_dir = os.path.join(shared_word_cache_dir, filename_letter_bucket(cache_name))
            else:
                target_dir = os.path.join(
                    source_word_cache_root_dir,
                    source_bucket_name(source_hint),
                    filename_letter_bucket(cache_name),
                )
            target_cache_path = os.path.join(target_dir, cache_name)
            if os.path.abspath(target_cache_path) == os.path.abspath(cache_path):
                continue

            os.makedirs(target_dir, exist_ok=True)
            wav_src = cache_path
            wav_dst = target_cache_path
            meta_src = cache_meta_path(cache_path)
            meta_dst = cache_meta_path(target_cache_path)

            try:
                if os.path.exists(wav_src):
                    if os.path.exists(wav_dst):
                        os.remove(wav_src)
                    else:
                        shutil.move(wav_src, wav_dst)
                if os.path.exists(meta_src):
                    meta_payload = load_json_file(meta_src, {})
                    if isinstance(meta_payload, dict):
                        meta_payload["cache_path"] = target_cache_path
                    if os.path.exists(meta_dst):
                        os.remove(meta_src)
                        if isinstance(meta_payload, dict) and meta_payload:
                            write_json_file(meta_dst, meta_payload)
                    else:
                        write_json_file(meta_dst, meta_payload if isinstance(meta_payload, dict) else {})
                        os.remove(meta_src)
                renamed_paths[cache_path] = target_cache_path
            except Exception as exc:
                log_warning(
                    "tts_rename_cache_source_migrate_entry_failed",
                    old_cache_path=cache_path,
                    new_cache_path=target_cache_path,
                    error=exc,
                )
                continue

    if not renamed_paths:
        return

    pending_items = load_json_file(pending_gemini_queue_path, [])
    if isinstance(pending_items, list):
        changed = False
        for item in pending_items:
            if not isinstance(item, dict):
                continue
            old_cache_path = str(item.get("cache_path") or "").strip()
            if old_cache_path in renamed_paths:
                item["cache_path"] = renamed_paths[old_cache_path]
                changed = True
        if changed:
            write_json_file(pending_gemini_queue_path, pending_items)


def migrate_legacy_word_wrapper_layout(
    *,
    legacy_word_cache_wrapper_dir,
    source_word_cache_root_dir,
    shared_word_cache_dir,
    load_json_file,
    write_json_file,
    pending_gemini_queue_path,
):
    if not os.path.isdir(legacy_word_cache_wrapper_dir):
        return

    renamed_paths = {}
    move_plan = [
        (os.path.join(legacy_word_cache_wrapper_dir, "sources"), source_word_cache_root_dir),
        (os.path.join(legacy_word_cache_wrapper_dir, "global"), shared_word_cache_dir),
    ]
    for legacy_dir, target_dir in move_plan:
        if not os.path.isdir(legacy_dir):
            continue
        os.makedirs(target_dir, exist_ok=True)
        for root, dirs, files in os.walk(legacy_dir, topdown=False):
            rel_root = os.path.relpath(root, legacy_dir)
            current_target_root = target_dir if rel_root == "." else os.path.join(target_dir, rel_root)
            os.makedirs(current_target_root, exist_ok=True)
            for name in files:
                src = os.path.join(root, name)
                dst = os.path.join(current_target_root, name)
                try:
                    if os.path.exists(dst):
                        os.remove(src)
                    else:
                        shutil.move(src, dst)
                    if name.lower().endswith(".wav"):
                        renamed_paths[src] = dst
                except Exception:
                    continue
            for name in dirs:
                old_dir = os.path.join(root, name)
                try:
                    if os.path.isdir(old_dir) and not os.listdir(old_dir):
                        os.rmdir(old_dir)
                except Exception:
                    pass
        try:
            if os.path.isdir(legacy_dir) and not os.listdir(legacy_dir):
                os.rmdir(legacy_dir)
        except Exception:
            pass

    pending_items = load_json_file(pending_gemini_queue_path, [])
    if isinstance(pending_items, list) and renamed_paths:
        changed = False
        for item in pending_items:
            if not isinstance(item, dict):
                continue
            old_cache_path = str(item.get("cache_path") or "").strip()
            if old_cache_path in renamed_paths:
                item["cache_path"] = renamed_paths[old_cache_path]
                changed = True
        if changed:
            write_json_file(pending_gemini_queue_path, pending_items)

    try:
        if os.path.isdir(legacy_word_cache_wrapper_dir) and not os.listdir(legacy_word_cache_wrapper_dir):
            os.rmdir(legacy_word_cache_wrapper_dir)
    except Exception:
        pass


def collapse_existing_lightweight_source_caches(
    *,
    source_word_cache_root_dir,
    load_json_file,
    normalize_source_path,
    infer_text_from_cache_filename,
    current_online_provider,
    resolve_cache_audio_path,
    alias_source_cache_to_shared,
):
    if not os.path.isdir(source_word_cache_root_dir):
        return
    for root, _, names in os.walk(source_word_cache_root_dir):
        for name in names:
            if not name.lower().endswith(".wav.json"):
                continue
            meta_path = os.path.join(root, name)
            cache_path = meta_path[:-5]
            metadata = load_json_file(meta_path, {})
            if not isinstance(metadata, dict):
                continue
            source_path = normalize_source_path(metadata.get("source_path"))
            if source_path == "shared":
                continue
            text_value = infer_text_from_cache_filename(cache_path, metadata)
            if not text_value:
                continue
            linked_shared = str(metadata.get("linked_shared_path") or "").strip()
            backend = str(metadata.get("backend") or "").strip().lower()
            desired_backend = str(metadata.get("desired_backend") or "").strip().lower()
            if linked_shared:
                alias_source_cache_to_shared(
                    text_value,
                    source_path=source_path,
                    shared_path=linked_shared,
                    backend=backend or current_online_provider(),
                    desired_backend=desired_backend or backend or current_online_provider(),
                    metadata=metadata,
                    cache_path=cache_path,
                )
                continue
            playable_path = resolve_cache_audio_path(cache_path)
            if not playable_path:
                continue
            try:
                alias_source_cache_to_shared(
                    text_value,
                    source_path=source_path,
                    backend=backend or current_online_provider(),
                    desired_backend=desired_backend or backend or current_online_provider(),
                    metadata=metadata,
                    cache_path=cache_path,
                )
            except Exception:
                continue


def collapse_all_source_cache_entities_to_aliases(
    *,
    source_word_cache_root_dir,
    load_cache_metadata,
    normalize_source_path,
    infer_text_from_cache_filename,
    current_online_provider,
    alias_source_cache_to_shared,
):
    if not os.path.isdir(source_word_cache_root_dir):
        return 0
    collapsed = 0
    for root, _, names in os.walk(source_word_cache_root_dir):
        for name in names:
            if not name.lower().endswith(".wav"):
                continue
            cache_path = os.path.join(root, name)
            metadata = load_cache_metadata(cache_path)
            if not isinstance(metadata, dict):
                continue
            source_path = normalize_source_path(metadata.get("source_path"))
            if source_path == "shared":
                continue
            text_value = infer_text_from_cache_filename(cache_path, metadata)
            if not text_value:
                continue
            backend = str(metadata.get("backend") or "").strip().lower()
            desired_backend = str(metadata.get("desired_backend") or backend or current_online_provider()).strip().lower()
            try:
                alias_path = alias_source_cache_to_shared(
                    text_value,
                    source_path=source_path,
                    backend=backend or current_online_provider(),
                    desired_backend=desired_backend or backend or current_online_provider(),
                    metadata=metadata,
                    cache_path=cache_path,
                )
                if alias_path:
                    collapsed += 1
            except Exception:
                continue
    return collapsed


def cleanup_duplicate_source_cache_entries(
    *,
    source_word_cache_root_dir,
    shared_word_cache_dir,
    load_cache_metadata,
    normalize_text,
    normalize_source_path,
    is_pending_gemini,
    resolve_cache_audio_path,
    cache_meta_path,
    is_online_backend,
    log_warning,
    remove_cache_metadata,
    remove_pending_gemini,
    save_cache_metadata,
    enqueue_existing_cache_for_online_replacement,
):
    def collect_grouped(root_dir, root_kind):
        grouped = {}
        if not os.path.isdir(root_dir):
            return grouped
        for root, _, names in os.walk(root_dir):
            for name in names:
                lower_name = name.lower()
                if lower_name.endswith(".wav"):
                    cache_path = os.path.join(root, name)
                elif lower_name.endswith(".wav.json"):
                    cache_path = os.path.join(root, name[:-5])
                else:
                    continue
                metadata = load_cache_metadata(cache_path)
                bucket = _cache_group_source_bucket(cache_path, source_word_cache_root_dir=source_word_cache_root_dir) if root_kind == "source" else "shared"
                key = _cache_group_key(cache_path, metadata, normalize_text=normalize_text)
                if not (bucket and key):
                    continue
                grouped.setdefault((bucket, key), {})[cache_path] = metadata
        return grouped

    def remove_cache_entry(cache_path):
        removed_local = 0
        try:
            if os.path.exists(cache_path):
                os.remove(cache_path)
                removed_local += 1
        except Exception as exc:
            log_warning("tts_remove_cache_entry_audio_failed", cache_path=cache_path, error=exc)
        meta_path = cache_meta_path(cache_path)
        try:
            if os.path.exists(meta_path):
                os.remove(meta_path)
                removed_local += 1
        except Exception as exc:
            log_warning("tts_remove_cache_entry_meta_failed", meta_path=meta_path, error=exc)
        remove_cache_metadata(cache_path)
        remove_pending_gemini(cache_path)
        return removed_local

    if not os.path.isdir(source_word_cache_root_dir) and not os.path.isdir(shared_word_cache_dir):
        return 0

    removed = 0
    shared_path_rewrites = {}
    shared_grouped = collect_grouped(shared_word_cache_dir, "shared")
    for (_bucket, _key), item_map in shared_grouped.items():
        cache_paths = list(item_map.keys())
        if len(cache_paths) <= 1:
            continue
        keep_path = max(
            cache_paths,
            key=lambda path: _cache_sort_key(
                path,
                item_map.get(path),
                is_pending_gemini=is_pending_gemini,
                resolve_cache_audio_path=resolve_cache_audio_path,
                cache_meta_path=cache_meta_path,
                is_online_backend=is_online_backend,
            ),
        )
        for cache_path in cache_paths:
            if cache_path == keep_path:
                continue
            shared_path_rewrites[cache_path] = keep_path
            removed += remove_cache_entry(cache_path)

    if shared_path_rewrites and os.path.isdir(source_word_cache_root_dir):
        for root, _, names in os.walk(source_word_cache_root_dir):
            for name in names:
                if not name.lower().endswith(".wav.json"):
                    continue
                cache_path = os.path.join(root, name[:-5])
                metadata = load_cache_metadata(cache_path)
                if not isinstance(metadata, dict):
                    continue
                linked_shared = str(metadata.get("linked_shared_path") or "").strip()
                rewritten = shared_path_rewrites.get(linked_shared)
                if not rewritten:
                    continue
                metadata["linked_shared_path"] = rewritten
                save_cache_metadata(cache_path, metadata)

    source_grouped = collect_grouped(source_word_cache_root_dir, "source")
    for (_bucket, _key), item_map in source_grouped.items():
        cache_paths = list(item_map.keys())
        if len(cache_paths) <= 1:
            continue
        keep_path = max(
            cache_paths,
            key=lambda path: _cache_sort_key(
                path,
                item_map.get(path),
                is_pending_gemini=is_pending_gemini,
                resolve_cache_audio_path=resolve_cache_audio_path,
                cache_meta_path=cache_meta_path,
                is_online_backend=is_online_backend,
            ),
        )
        keep_metadata = item_map.get(keep_path) or {}
        keep_source = normalize_source_path(keep_metadata.get("source_path"))
        keep_text = _infer_text_from_cache_filename(keep_path, keep_metadata, normalize_text=normalize_text)
        pending_seen = False
        for cache_path in cache_paths:
            if cache_path == keep_path:
                continue
            pending_seen = pending_seen or is_pending_gemini(cache_path)
            removed += remove_cache_entry(cache_path)
        if pending_seen and keep_text and keep_source and not is_pending_gemini(keep_path):
            keep_backend = str(keep_metadata.get("backend") or "").strip().lower()
            keep_desired = str(keep_metadata.get("desired_backend") or "").strip().lower()
            if keep_backend != keep_desired and is_online_backend(keep_desired):
                enqueue_existing_cache_for_online_replacement(keep_text, keep_path, source_path=keep_source)
    return removed


def normalize_cache_metadata_texts(
    *,
    shared_word_cache_dir,
    source_word_cache_root_dir,
    load_cache_metadata,
    save_cache_metadata,
    normalize_text,
):
    roots = []
    if os.path.isdir(shared_word_cache_dir):
        roots.append(("shared", shared_word_cache_dir))
    if os.path.isdir(source_word_cache_root_dir):
        roots.append(("source", source_word_cache_root_dir))
    if not roots:
        return 0

    updated = 0
    for root_kind, root_dir in roots:
        for root, _, names in os.walk(root_dir):
            for name in names:
                if not name.lower().endswith(".wav.json"):
                    continue
                cache_path = os.path.join(root, name[:-5])
                metadata = load_cache_metadata(cache_path)
                if not isinstance(metadata, dict):
                    continue
                normalized_guess = _infer_text_from_cache_filename(cache_path, {}, normalize_text=normalize_text)
                if not normalized_guess:
                    continue
                payload = dict(metadata)
                current_text = normalize_text(payload.get("text"), ensure_sentence_end=False)
                if current_text == normalized_guess and payload.get("text") == normalized_guess:
                    continue
                payload["text"] = normalized_guess
                if root_kind == "shared":
                    payload["source_path"] = "shared"
                save_cache_metadata(cache_path, payload)
                updated += 1
    return updated
