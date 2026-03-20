# -*- coding: utf-8 -*-


def execute_local_backend(
    *,
    selected_backend,
    normalized,
    volume,
    rate_ratio,
    short_text,
    cache_path,
    source_path,
    request_token,
    synthesize_with_kokoro,
    synthesize_with_piper,
    save_word_cache_file,
    set_backend_status,
    wav_duration_seconds,
):
    if selected_backend == "kokoro":
        wav_path, backend_label, _can_cache = synthesize_with_kokoro(
            normalized,
            volume=volume,
            rate_ratio=rate_ratio,
        )
        if short_text and cache_path:
            try:
                save_word_cache_file(
                    cache_path,
                    wav_path,
                    text=normalized,
                    source_path=source_path,
                    backend="kokoro",
                    desired_backend="kokoro",
                )
            except Exception:
                pass
        if request_token is not None:
            set_backend_status(
                request_token,
                backend_label,
                from_cache=False,
                fallback=False,
                duration_seconds=wav_duration_seconds(wav_path),
            )
        return {"handled": True, "wav_path": wav_path}

    if selected_backend == "piper":
        wav_path, backend_label, _can_cache = synthesize_with_piper(
            normalized,
            volume=volume,
            rate_ratio=rate_ratio,
        )
        if short_text and cache_path:
            try:
                save_word_cache_file(
                    cache_path,
                    wav_path,
                    text=normalized,
                    source_path=source_path,
                    backend="piper",
                    desired_backend="piper",
                )
            except Exception:
                pass
        if request_token is not None:
            set_backend_status(
                request_token,
                backend_label,
                from_cache=False,
                fallback=False,
                duration_seconds=wav_duration_seconds(wav_path),
            )
        return {"handled": True, "wav_path": wav_path}

    return {"handled": False, "wav_path": ""}


def execute_online_with_fallback(
    *,
    text,
    normalized,
    volume,
    rate_ratio,
    short_text,
    cache_path,
    source_path,
    request_token,
    online_timeout_seconds,
    current_online_provider,
    primary_online_provider,
    backend_key_from_label,
    set_backend_status,
    wav_duration_seconds,
    synthesize_with_user_online,
    save_word_cache_file,
    enqueue_existing_cache_for_online_replacement,
    synthesize_with_local_placeholder,
    kokoro_ready,
    synthesize_with_kokoro_voice,
):
    try:
        wav_path, backend_label, _can_cache = synthesize_with_user_online(
            normalized,
            volume=volume,
            short_text=short_text,
            timeout_seconds=online_timeout_seconds,
        )
        actual_online_backend = backend_key_from_label(backend_label)
        desired_online_backend = primary_online_provider()
        if request_token is not None:
            set_backend_status(
                request_token,
                backend_label,
                from_cache=False,
                fallback=False,
                duration_seconds=wav_duration_seconds(wav_path),
            )
        if short_text:
            try:
                save_word_cache_file(
                    cache_path,
                    wav_path,
                    text=normalized,
                    source_path=source_path,
                    backend=actual_online_backend,
                    desired_backend=desired_online_backend,
                )
                if actual_online_backend != desired_online_backend:
                    enqueue_existing_cache_for_online_replacement(text, cache_path, source_path=source_path)
            except Exception:
                pass
        return wav_path
    except Exception:
        if short_text:
            try:
                (wav_path, backend_label, _can_cache), placeholder_backend = synthesize_with_local_placeholder(
                    normalized,
                    volume=volume,
                    rate_ratio=rate_ratio,
                )
                try:
                    save_word_cache_file(
                        cache_path,
                        wav_path,
                        text=normalized,
                        source_path=source_path,
                        backend=placeholder_backend,
                        desired_backend=current_online_provider(),
                    )
                except Exception:
                    pass
                enqueue_existing_cache_for_online_replacement(text, cache_path, source_path=source_path)
                if request_token is not None:
                    set_backend_status(
                        request_token,
                        backend_label,
                        from_cache=False,
                        fallback=True,
                        duration_seconds=wav_duration_seconds(wav_path),
                    )
                return wav_path
            except Exception:
                enqueue_existing_cache_for_online_replacement(text, cache_path, source_path=source_path)
        if kokoro_ready():
            wav_path, backend_label, _can_cache = synthesize_with_kokoro_voice(
                normalized,
                volume=volume,
                rate_ratio=rate_ratio,
                voice_id="bf_emma",
                lang="en-gb",
            )
            if request_token is not None:
                set_backend_status(
                    request_token,
                    backend_label,
                    from_cache=False,
                    fallback=True,
                    duration_seconds=wav_duration_seconds(wav_path),
                )
            return wav_path
        raise
