# -*- coding: utf-8 -*-
import threading

from services.gemini_writer import validate_gemini_api_key
from services.tts import validate_tts_api_key


def start_gemini_validation_task(*, token, api_key, model_name, emit_event):
    def _run():
        try:
            validate_gemini_api_key(api_key, model=model_name, timeout=25)
            emit_event(
                "success",
                token,
                {"api_key": api_key, "model": model_name},
            )
        except Exception as exc:
            emit_event("error", token, str(exc))
        emit_event("done", token, None)

    threading.Thread(target=_run, daemon=True).start()


def start_tts_validation_task(*, token, api_key, provider, emit_event):
    def _run():
        try:
            validate_tts_api_key(api_key, provider, timeout=30)
            emit_event(
                "success_tts",
                token,
                {"api_key": api_key, "provider": provider},
            )
        except Exception as exc:
            emit_event("error_tts", token, str(exc))
        emit_event("done", token, None)

    threading.Thread(target=_run, daemon=True).start()


def start_combined_api_validation_task(
    *,
    token,
    llm_required,
    tts_required,
    llm_key,
    tts_key,
    model_name,
    tts_provider,
    emit_event,
):
    def _run():
        result = {
            "llm_required": bool(llm_required),
            "tts_required": bool(tts_required),
            "llm_ok": False,
            "tts_ok": False,
            "llm_error": "",
            "tts_error": "",
            "llm_api_key": llm_key,
            "tts_api_key": tts_key,
            "llm_model": model_name,
            "tts_provider": tts_provider,
        }
        try:
            if llm_required:
                try:
                    validate_gemini_api_key(llm_key, model=model_name, timeout=25)
                    result["llm_ok"] = True
                except Exception as exc:
                    result["llm_error"] = str(exc)
            if tts_required:
                try:
                    validate_tts_api_key(tts_key, tts_provider, timeout=30)
                    result["tts_ok"] = True
                except Exception as exc:
                    result["tts_error"] = str(exc)
            emit_event("success_api_setup", token, result)
        finally:
            emit_event("done", token, None)

    threading.Thread(target=_run, daemon=True).start()
