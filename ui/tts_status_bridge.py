# -*- coding: utf-8 -*-
from services.tts import get_backend_status as tts_get_backend_status


def watch_tts_backend(host, token, target, text_label):
    host.tts_status_request += 1
    request_id = host.tts_status_request

    def _poll(attempt=0):
        if request_id != host.tts_status_request:
            return
        status = tts_get_backend_status(token)
        if status:
            label = status.get("label") or "TTS"
            fallback = bool(status.get("fallback"))
            from_cache = bool(status.get("from_cache"))
            if target == "passage":
                if fallback:
                    host.passage_status_var.set(f"Playing passage via {label} after Gemini fallback.")
                elif from_cache:
                    host.passage_status_var.set(f"Playing passage via cached {label} audio.")
                else:
                    host.passage_status_var.set(f"Playing passage via {label}.")
            elif target == "playback":
                index_text = f"{host.pos + 1}/{len(host.store.words)}" if host.store.words and host.pos >= 0 else "0/0"
                if fallback:
                    host.status_var.set(f"顺序播放中：{index_text}（{label} 回退）")
                elif from_cache:
                    host.status_var.set(f"顺序播放中：{index_text}（{label} 缓存）")
                else:
                    host.status_var.set(f"顺序播放中：{index_text}（{label}）")
            else:
                if fallback:
                    host.status_var.set(f"Playing '{text_label}' via {label} after Gemini fallback.")
                elif from_cache:
                    host.status_var.set(f"Playing cached audio for '{text_label}' via {label}.")
                else:
                    host.status_var.set(f"Playing '{text_label}' via {label}.")
            return
        if attempt < 120:
            host.after(250, lambda: _poll(attempt + 1))

    host.after(250, _poll)
