# -*- coding: utf-8 -*-


def provider_key(provider):
    return "elevenlabs" if str(provider or "").strip().lower() == "elevenlabs" else "gemini"


def current_online_provider(tts_api_provider):
    return provider_key(tts_api_provider)


def online_provider_label(provider=None, *, current_provider="gemini"):
    backend = provider_key(provider or current_provider)
    return "ElevenLabs TTS" if backend == "elevenlabs" else "Gemini TTS"


def primary_online_provider(tts_api_provider):
    return current_online_provider(tts_api_provider)


def secondary_online_provider(primary=None, *, has_llm_api_key=False):
    first = provider_key(primary)
    if first == "elevenlabs" and has_llm_api_key:
        return "gemini"
    return None


def online_provider_candidates(primary=None, *, has_llm_api_key=False):
    first = provider_key(primary)
    providers = [first]
    secondary = secondary_online_provider(first, has_llm_api_key=has_llm_api_key)
    if secondary and provider_key(secondary) not in providers:
        providers.append(provider_key(secondary))
    return providers


def rate_limit_cooldown_for_provider(provider, *, gemini_seconds, elevenlabs_seconds):
    return elevenlabs_seconds if provider_key(provider) == "elevenlabs" else gemini_seconds


def manual_request_cooldown_for_provider(provider, *, gemini_seconds, elevenlabs_seconds):
    return elevenlabs_seconds if provider_key(provider) == "elevenlabs" else gemini_seconds


def local_fallback_ready(*, piper_ready=False, kokoro_ready=False):
    return bool(piper_ready or kokoro_ready)


def interactive_online_timeout(
    short_text,
    *,
    local_fallback_ready=False,
    online_timeout_seconds,
    fast_short_timeout_seconds,
    fast_long_timeout_seconds,
):
    if not local_fallback_ready:
        return online_timeout_seconds
    return fast_short_timeout_seconds if short_text else fast_long_timeout_seconds


def backend_key(source=None, fallback_backend=None, *, source_kokoro, source_piper, current_online_provider):
    source_name = str(source or "").strip().lower()
    fallback_name = str(fallback_backend or "").strip().lower()
    if fallback_name in {"gemini", "elevenlabs", "kokoro", "piper"}:
        return fallback_name
    if source_name == source_kokoro:
        return "kokoro"
    if source_name == source_piper:
        return "piper"
    return provider_key(current_online_provider)


def backend_label_from_key(backend_key):
    backend = str(backend_key or "").strip().lower()
    if backend == "kokoro":
        return "Kokoro (Offline)"
    if backend == "piper":
        return "Piper (Local)"
    if backend == "elevenlabs":
        return "ElevenLabs TTS"
    return "Gemini TTS"


def backend_key_from_label(label):
    value = str(label or "").strip().lower()
    if "elevenlabs" in value:
        return "elevenlabs"
    if "kokoro" in value:
        return "kokoro"
    if "piper" in value:
        return "piper"
    return "gemini"


def synthesize_with_online(
    text,
    volume,
    *,
    short_text,
    timeout_seconds,
    primary_provider,
    secondary_provider,
    get_fallback_key,
    synthesize_with_online_provider,
):
    primary = provider_key(primary_provider)
    try:
        return (
            synthesize_with_online_provider(
                text,
                volume=volume,
                short_text=short_text,
                provider=primary,
                timeout_seconds=timeout_seconds,
            ),
            primary,
        )
    except Exception:
        secondary = provider_key(secondary_provider) if secondary_provider else None
        if secondary:
            fallback_key = get_fallback_key(secondary)
            return (
                synthesize_with_online_provider(
                    text,
                    volume=volume,
                    short_text=short_text,
                    provider=secondary,
                    api_key=fallback_key,
                    timeout_seconds=timeout_seconds,
                ),
                secondary,
            )
        raise


def synthesize_with_selected_source(
    text,
    volume,
    rate_ratio,
    *,
    short_text,
    timeout_seconds,
    source,
    source_kokoro,
    source_piper,
    synthesize_with_kokoro,
    synthesize_with_piper,
    synthesize_with_user_online,
    local_fallback_ready,
    synthesize_with_local_placeholder,
):
    if source == source_kokoro:
        return synthesize_with_kokoro(text, volume=volume, rate_ratio=rate_ratio), False
    if source == source_piper:
        return synthesize_with_piper(text, volume=volume, rate_ratio=rate_ratio), False
    try:
        return (
            synthesize_with_user_online(
                text,
                volume=volume,
                short_text=short_text,
                timeout_seconds=timeout_seconds,
            ),
            False,
        )
    except Exception:
        if local_fallback_ready:
            result, _placeholder_backend = synthesize_with_local_placeholder(
                text,
                volume=volume,
                rate_ratio=rate_ratio,
            )
            return result, True
        raise
