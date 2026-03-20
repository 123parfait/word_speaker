import os
import threading
import time


def run_pending_online_worker(
    *,
    runtime_state,
    next_pending_item,
    remove_pending_item,
    normalize_text,
    current_online_provider,
    wait_for_queue_slot,
    log_info,
    log_warning,
    log_error,
    synthesize_gemini,
    synthesize_elevenlabs,
    record_queue_success,
    record_queue_rate_limit,
    record_queue_soft_failure,
    set_queue_status,
    save_word_cache_file,
    secondary_online_provider,
    get_fallback_key,
    synthesize_with_online_provider,
    rate_limit_cooldown_for_provider,
    defer_queue,
    queue_interval_for_provider,
    is_rate_limited_error,
    synthesize_local_placeholder,
    online_queue_manager,
    refresh_queue_status_counts,
):
    try:
        while True:
            next_item = next_pending_item()
            if not next_item:
                return
            cache_path_local, item = next_item
            if not isinstance(item, dict):
                remove_pending_item(cache_path_local)
                continue
            normalized_text = normalize_text(item.get("text"), ensure_sentence_end=False)
            source_path = str(item.get("source_path") or "").strip() or None
            if not normalized_text:
                remove_pending_item(cache_path_local)
                continue
            desired_backend = str(item.get("desired_backend") or current_online_provider()).strip().lower()
            if desired_backend not in {"gemini", "elevenlabs"}:
                desired_backend = current_online_provider()
            try:
                wait_for_queue_slot(provider=desired_backend)
                log_info(
                    "tts_queue_item_start",
                    provider=desired_backend,
                    cache_path=cache_path_local,
                    source_path=source_path or "",
                    text=normalized_text[:120],
                )
                wav_path, _label, _can_cache = (
                    synthesize_elevenlabs(normalized_text, volume=1.0, short_text=True)
                    if desired_backend == "elevenlabs"
                    else synthesize_gemini(normalized_text, volume=1.0, short_text=True)
                )
                record_queue_success(desired_backend)
                set_queue_status(
                    state="ok",
                    next_retry_at=0.0,
                    last_success_at=time.time(),
                    last_error="",
                )
                try:
                    save_word_cache_file(
                        cache_path_local,
                        wav_path,
                        text=normalized_text,
                        source_path=source_path,
                        backend=desired_backend,
                        desired_backend=desired_backend,
                    )
                finally:
                    try:
                        os.remove(wav_path)
                    except Exception:
                        pass
                remove_pending_item(cache_path_local)
                log_info("tts_queue_item_done", provider=desired_backend, cache_path=cache_path_local)
            except Exception as exc:
                log_warning(
                    "tts_queue_item_failed",
                    provider=desired_backend,
                    cache_path=cache_path_local,
                    error=exc,
                )
                secondary_backend = secondary_online_provider(desired_backend)
                primary_rate_limited = is_rate_limited_error(exc)
                if primary_rate_limited:
                    record_queue_rate_limit(desired_backend)
                else:
                    record_queue_soft_failure(desired_backend)
                if secondary_backend:
                    try:
                        fallback_key = get_fallback_key(secondary_backend)
                        wav_path, _label, _can_cache = synthesize_with_online_provider(
                            normalized_text,
                            volume=1.0,
                            short_text=True,
                            provider=secondary_backend,
                            api_key=fallback_key,
                        )
                        try:
                            save_word_cache_file(
                                cache_path_local,
                                wav_path,
                                text=normalized_text,
                                source_path=source_path,
                                backend=secondary_backend,
                                desired_backend=desired_backend,
                            )
                        finally:
                            try:
                                os.remove(wav_path)
                            except Exception:
                                pass
                        record_queue_success(secondary_backend)
                        log_info(
                            "tts_queue_item_fallback_success",
                            provider=desired_backend,
                            fallback_provider=secondary_backend,
                            cache_path=cache_path_local,
                        )
                        if primary_rate_limited:
                            cooldown_seconds = rate_limit_cooldown_for_provider(desired_backend)
                            defer_queue(
                                cooldown_seconds,
                                state="rate_limited",
                                provider=desired_backend,
                            )
                            set_queue_status(
                                last_error=str(exc),
                                next_retry_at=time.time() + cooldown_seconds,
                            )
                        else:
                            set_queue_status(
                                state="ok",
                                next_retry_at=time.time() + queue_interval_for_provider(desired_backend),
                                last_success_at=time.time(),
                                last_error="",
                            )
                        continue
                    except Exception as fallback_exc:
                        record_queue_soft_failure(secondary_backend)
                        log_error(
                            "tts_queue_item_fallback_failed",
                            provider=desired_backend,
                            fallback_provider=secondary_backend,
                            cache_path=cache_path_local,
                            error=fallback_exc,
                        )
                if primary_rate_limited:
                    cooldown_seconds = rate_limit_cooldown_for_provider(desired_backend)
                    set_queue_status(
                        state="rate_limited",
                        next_retry_at=time.time() + cooldown_seconds,
                        last_error=str(exc),
                    )
                    if not os.path.exists(cache_path_local):
                        try:
                            (wav_path, _label, _can_cache), placeholder_backend = synthesize_local_placeholder(
                                normalized_text,
                                volume=1.0,
                                rate_ratio=1.0,
                            )
                            try:
                                save_word_cache_file(
                                    cache_path_local,
                                    wav_path,
                                    text=normalized_text,
                                    source_path=source_path,
                                    backend=placeholder_backend,
                                    desired_backend=desired_backend,
                                )
                            finally:
                                try:
                                    os.remove(wav_path)
                                except Exception as cleanup_exc:
                                    log_warning("tts_queue_placeholder_cleanup_failed", path=wav_path, error=cleanup_exc)
                        except Exception as placeholder_exc:
                            log_warning(
                                "tts_queue_placeholder_fallback_failed",
                                provider=desired_backend,
                                cache_path=cache_path_local,
                                error=placeholder_exc,
                            )
                    log_warning(
                        "tts_queue_item_rate_limited",
                        provider=desired_backend,
                        cache_path=cache_path_local,
                        cooldown_seconds=cooldown_seconds,
                    )
                    threading.Event().wait(cooldown_seconds)
                    continue
                set_queue_status(
                    state="error",
                    next_retry_at=0.0,
                    last_error=str(exc),
                )
                log_error(
                    "tts_queue_item_abandoned",
                    provider=desired_backend,
                    cache_path=cache_path_local,
                    error=exc,
                )
                remove_pending_item(cache_path_local)
    finally:
        runtime_state.pending_gemini_worker_running = False
        refresh_queue_status_counts()
        with runtime_state.pending_gemini_lock:
            still_pending = bool(runtime_state.pending_gemini_replacements)
        if not still_pending:
            current_status = online_queue_manager.get_status()
            updates = {
                "next_retry_at": 0.0,
                "worker_running": False,
            }
            if current_status.get("state") != "rate_limited":
                updates["state"] = "idle"
            online_queue_manager.set_status(**updates)
