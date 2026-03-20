# -*- coding: utf-8 -*-
import time


def build_tts_runtime_status(*, tr, trf, provider_label, status, now=None):
    payload = dict(status or {})
    state = str(payload.get("state") or "idle").strip().lower()
    queue_count = int(payload.get("queue_count") or 0)
    now_ts = time.time() if now is None else float(now)

    if state == "rate_limited":
        status_text = trf("tts_status_limited", provider=provider_label)
    elif state == "ok":
        status_text = trf("tts_status_normal", provider=provider_label)
    elif state == "error":
        status_text = trf("tts_status_error", provider=provider_label)
    else:
        status_text = trf("tts_status_idle", provider=provider_label)

    if state == "rate_limited":
        queue_text = trf("tts_status_queue_waiting", count=queue_count)
    elif queue_count > 0:
        queue_text = trf("tts_status_queue_processing", count=queue_count)
    else:
        queue_text = trf("tts_status_queue", count=queue_count)

    next_retry_at = float(payload.get("next_retry_at") or 0.0)
    if next_retry_at > 0:
        remaining_seconds = max(0, int(round(next_retry_at - now_ts)))
        retry_text = time.strftime("%H:%M:%S", time.localtime(next_retry_at))
        retry_status = trf("tts_status_retry_at_in", time=retry_text, seconds=remaining_seconds)
    else:
        retry_status = tr("tts_status_retry_none")

    return {
        "runtime_status": f"{status_text} | {queue_text}",
        "retry_status": retry_status,
    }


def parse_custom_interval(text, *, minimum=0.2):
    value = float(str(text or "").strip())
    if value < float(minimum):
        raise ValueError("interval_too_small")
    return value
