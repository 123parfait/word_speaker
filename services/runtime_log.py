# -*- coding: utf-8 -*-
import json
import os
import threading
import time


BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOG_PATH = os.path.join(BASE_DIR, "data", "runtime.log")
_LOG_LOCK = threading.Lock()


def get_runtime_log_path():
    return LOG_PATH


def _normalize_details(details):
    payload = {}
    for key, value in (details or {}).items():
        try:
            payload[str(key)] = value if isinstance(value, (str, int, float, bool, type(None))) else str(value)
        except Exception:
            payload[str(key)] = "<unserializable>"
    return payload


def write_runtime_log(level, event, **details):
    try:
        os.makedirs(os.path.dirname(LOG_PATH), exist_ok=True)
        record = {
            "ts": time.strftime("%Y-%m-%d %H:%M:%S"),
            "level": str(level or "INFO").upper(),
            "event": str(event or "unknown"),
            "details": _normalize_details(details),
        }
        line = json.dumps(record, ensure_ascii=False)
        with _LOG_LOCK:
            with open(LOG_PATH, "a", encoding="utf-8") as fp:
                fp.write(line + "\n")
    except Exception:
        return


def log_info(event, **details):
    write_runtime_log("INFO", event, **details)


def log_warning(event, **details):
    write_runtime_log("WARN", event, **details)


def log_error(event, **details):
    write_runtime_log("ERROR", event, **details)
