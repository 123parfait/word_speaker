# -*- coding: utf-8 -*-
import os
import shutil


def resolve_short_text_cache(
    *,
    text,
    source_path,
    selected_backend,
    request_token,
    volume,
    current_online_provider,
    word_cache_path,
    resolve_cache_audio_path,
    load_cache_metadata,
    set_backend_status,
    backend_label_from_key,
    wav_duration_seconds,
    clone_to_temp,
    ensure_source_online_cache,
    is_online_backend,
    has_valid_online_cache,
    online_provider_label,
    enqueue_existing_cache_for_online_replacement,
    legacy_word_cache_path,
    save_cache_metadata,
):
    cache_path = None

    if selected_backend in {"kokoro", "piper"}:
        cache_path = word_cache_path(text, source_path=source_path)
        playable_cache = resolve_cache_audio_path(cache_path)
        metadata = load_cache_metadata(cache_path)
        backend_key = str(metadata.get("backend") or "").strip().lower()
        if playable_cache and backend_key == selected_backend:
            if request_token is not None:
                set_backend_status(
                    request_token,
                    backend_label_from_key(backend_key),
                    from_cache=True,
                    fallback=False,
                    duration_seconds=wav_duration_seconds(playable_cache),
                )
            return {
                "cache_path": cache_path,
                "hit": True,
                "wav_path": clone_to_temp(playable_cache, volume=volume),
            }
        return {
            "cache_path": cache_path,
            "hit": False,
            "wav_path": "",
        }

    cache_path = ensure_source_online_cache(text, source_path=source_path)
    if is_online_backend(selected_backend) and has_valid_online_cache(cache_path):
        playable_cache = resolve_cache_audio_path(cache_path) or cache_path
        if request_token is not None:
            set_backend_status(
                request_token,
                online_provider_label(),
                from_cache=True,
                fallback=False,
                duration_seconds=wav_duration_seconds(playable_cache),
            )
        return {
            "cache_path": cache_path,
            "hit": True,
            "wav_path": clone_to_temp(playable_cache, volume=volume),
        }

    if is_online_backend(selected_backend) and os.path.exists(cache_path):
        metadata = load_cache_metadata(cache_path)
        backend_key = str(metadata.get("backend") or "").strip().lower()
        desired_backend = str(metadata.get("desired_backend") or current_online_provider()).strip().lower()
        playable_cache = resolve_cache_audio_path(cache_path) or cache_path
        if backend_key in {"piper", "kokoro"}:
            enqueue_existing_cache_for_online_replacement(text, cache_path, source_path=source_path)
            if request_token is not None:
                set_backend_status(
                    request_token,
                    backend_label_from_key(backend_key),
                    from_cache=True,
                    fallback=True,
                    duration_seconds=wav_duration_seconds(playable_cache),
                )
            return {
                "cache_path": cache_path,
                "hit": True,
                "wav_path": clone_to_temp(playable_cache, volume=volume),
            }
        if is_online_backend(backend_key) and desired_backend != backend_key:
            enqueue_existing_cache_for_online_replacement(text, cache_path, source_path=source_path)
            if request_token is not None:
                set_backend_status(
                    request_token,
                    backend_label_from_key(backend_key),
                    from_cache=True,
                    fallback=True,
                    duration_seconds=wav_duration_seconds(playable_cache),
                )
            return {
                "cache_path": cache_path,
                "hit": True,
                "wav_path": clone_to_temp(playable_cache, volume=volume),
            }

    legacy_cache = legacy_word_cache_path(text, source_path=source_path)
    if legacy_cache and os.path.exists(legacy_cache) and not os.path.exists(cache_path):
        try:
            os.makedirs(os.path.dirname(cache_path), exist_ok=True)
            shutil.move(legacy_cache, cache_path)
        except Exception:
            try:
                shutil.copyfile(legacy_cache, cache_path)
            except Exception:
                cache_path = legacy_cache
        if not load_cache_metadata(cache_path):
            save_cache_metadata(
                cache_path,
                {
                    "backend": "unknown",
                    "desired_backend": current_online_provider(),
                    "source_path": str(source_path or "").strip() or None,
                },
            )

    if not has_valid_online_cache(cache_path):
        enqueue_existing_cache_for_online_replacement(text, cache_path, source_path=source_path)

    return {
        "cache_path": cache_path,
        "hit": False,
        "wav_path": "",
    }
