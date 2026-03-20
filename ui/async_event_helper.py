# -*- coding: utf-8 -*-
import queue


def clear_event_queue(target_queue):
    try:
        while True:
            target_queue.get_nowait()
    except queue.Empty:
        return


def emit_event(target_queue, event_type, token, payload=None):
    try:
        target_queue.put_nowait((event_type, token, payload))
    except Exception:
        return


def drain_event_queue(*, target_queue, token, active_token, handlers):
    if token != active_token:
        return True
    done = False
    try:
        while True:
            event_type, event_token, payload = target_queue.get_nowait()
            if event_token != token:
                continue
            if event_type == "done":
                done = True
                continue
            handler = handlers.get(event_type)
            if handler:
                handler(payload)
    except queue.Empty:
        pass
    return done
