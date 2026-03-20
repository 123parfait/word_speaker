# -*- coding: utf-8 -*-


def build_combined_api_validation_request(*, llm_key, tts_key, tts_provider, model_name, force_llm=False, force_tts=False):
    llm_key = str(llm_key or "").strip()
    tts_key = str(tts_key or "").strip()
    model_name = str(model_name or "").strip()
    tts_provider = str(tts_provider or "gemini").strip().lower()
    llm_required = bool(force_llm or llm_key)
    tts_required = bool(force_tts or tts_key)
    return {
        "llm_key": llm_key,
        "tts_key": tts_key,
        "tts_provider": tts_provider,
        "model_name": model_name,
        "llm_required": llm_required,
        "tts_required": tts_required,
    }


def build_combined_api_local_validation_state(request):
    payload = dict(request or {})
    llm_required = bool(payload.get("llm_required"))
    tts_required = bool(payload.get("tts_required"))
    llm_key = str(payload.get("llm_key") or "").strip()
    tts_key = str(payload.get("tts_key") or "").strip()

    llm_error = llm_required and not llm_key
    tts_error = tts_required and not tts_key
    has_local_error = bool(llm_error or tts_error)

    return {
        "has_local_error": has_local_error,
        "llm_error": llm_error,
        "tts_error": tts_error,
        "llm_status": "Please enter an LLM API key." if llm_error else "Paste your LLM API key, then test it.",
        "tts_status": "Please enter a TTS API key." if tts_error else "Paste your TTS API key, then test it.",
        "has_any_request": bool(llm_required or tts_required),
    }


def build_combined_api_apply_state(payload):
    data = dict(payload or {})
    llm_required = bool(data.get("llm_required"))
    tts_required = bool(data.get("tts_required"))
    llm_ok = bool(data.get("llm_ok"))
    tts_ok = bool(data.get("tts_ok"))
    llm_error = str(data.get("llm_error") or "").strip()
    tts_error = str(data.get("tts_error") or "").strip()
    provider = str(data.get("tts_provider") or "gemini").strip().lower()

    return {
        "llm_required": llm_required,
        "tts_required": tts_required,
        "llm_ok": llm_ok,
        "tts_ok": tts_ok,
        "llm_error_message": llm_error or "LLM API key test failed. Please paste another key.",
        "tts_error_message": tts_error or "TTS API key test failed. Please paste another key.",
        "llm_api_key": str(data.get("llm_api_key") or "").strip(),
        "tts_api_key": str(data.get("tts_api_key") or "").strip(),
        "llm_model": str(data.get("llm_model") or "").strip(),
        "tts_provider": provider,
        "all_ok": (not llm_required or llm_ok) and (not tts_required or tts_ok),
    }


def build_single_llm_success_state(payload, *, default_model):
    data = dict(payload or {})
    return {
        "api_key": str(data.get("api_key") or "").strip(),
        "model_name": str(data.get("model") or default_model).strip(),
        "status_text": "LLM API key is valid.",
        "main_status": "LLM API ready.",
    }


def build_single_tts_success_state(payload):
    data = dict(payload or {})
    return {
        "api_key": str(data.get("api_key") or "").strip(),
        "provider": str(data.get("provider") or "gemini").strip().lower(),
        "status_text": "TTS API key is valid.",
        "main_status": "TTS API ready.",
    }


def build_single_api_error_state(*, kind):
    if str(kind or "").strip().lower() == "tts":
        return {
            "status_text": "TTS API key test failed. Please paste another key.",
        }
    return {
        "status_text": "LLM API key test failed. Please paste another key.",
    }
