# -*- coding: utf-8 -*-
import threading
import time


class OnlineTtsQueueManager:
    def __init__(self, *, throttle_config):
        self._throttle_config = dict(throttle_config or {})
        self._status_lock = threading.Lock()
        self._status = {
            "state": "idle",
            "next_retry_at": 0.0,
            "last_success_at": 0.0,
            "last_error": "",
            "worker_running": False,
            "queue_count": 0,
        }
        self._attempt_lock = threading.Lock()
        self._last_attempt_at = 0.0
        self._throttle_state_lock = threading.Lock()
        self._throttle_state = {
            provider: {
                "current_interval": float(config["base_interval"]),
                "success_streak": 0,
                "last_event": "idle",
            }
            for provider, config in self._throttle_config.items()
        }

    def provider_key(self, provider):
        return "elevenlabs" if str(provider or "").strip().lower() == "elevenlabs" else "gemini"

    def throttle_config(self, provider):
        return self._throttle_config[self.provider_key(provider)]

    def set_status(self, **updates):
        with self._status_lock:
            self._status.update(updates)

    def get_status(self):
        with self._status_lock:
            return dict(self._status)

    def get_queue_throttle_snapshot(self, provider):
        key = self.provider_key(provider)
        config = self.throttle_config(key)
        with self._throttle_state_lock:
            state = dict(self._throttle_state.get(key) or {})
        if not state:
            state = {
                "current_interval": float(config["base_interval"]),
                "success_streak": 0,
                "last_event": "idle",
            }
        return state

    def queue_interval_for_provider(self, provider):
        state = self.get_queue_throttle_snapshot(provider)
        return float(state.get("current_interval") or self.throttle_config(provider)["base_interval"])

    def record_queue_success(self, provider):
        key = self.provider_key(provider)
        config = self.throttle_config(key)
        with self._throttle_state_lock:
            state = self._throttle_state.setdefault(
                key,
                {
                    "current_interval": float(config["base_interval"]),
                    "success_streak": 0,
                    "last_event": "idle",
                },
            )
            state["success_streak"] = int(state.get("success_streak") or 0) + 1
            if state["success_streak"] >= int(config["success_streak"]):
                state["current_interval"] = max(
                    float(config["min_interval"]),
                    float(state.get("current_interval") or config["base_interval"]) - float(config["success_step"]),
                )
                state["success_streak"] = 0
            state["last_event"] = "success"
        self.set_status(last_success_at=time.time())

    def record_queue_soft_failure(self, provider):
        key = self.provider_key(provider)
        config = self.throttle_config(key)
        with self._throttle_state_lock:
            state = self._throttle_state.setdefault(
                key,
                {
                    "current_interval": float(config["base_interval"]),
                    "success_streak": 0,
                    "last_event": "idle",
                },
            )
            state["success_streak"] = 0
            state["current_interval"] = min(
                float(config["max_interval"]),
                max(
                    float(config["base_interval"]),
                    float(state.get("current_interval") or config["base_interval"]) + float(config["soft_fail_step"]),
                ),
            )
            state["last_event"] = "soft_failure"

    def record_queue_rate_limit(self, provider):
        key = self.provider_key(provider)
        config = self.throttle_config(key)
        with self._throttle_state_lock:
            state = self._throttle_state.setdefault(
                key,
                {
                    "current_interval": float(config["base_interval"]),
                    "success_streak": 0,
                    "last_event": "idle",
                },
            )
            state["success_streak"] = 0
            state["current_interval"] = min(
                float(config["max_interval"]),
                max(
                    float(config["base_interval"]),
                    float(state.get("current_interval") or config["base_interval"]) + float(config["rate_limit_step"]),
                ),
            )
            state["last_event"] = "rate_limited"

    def refresh_counts(self, *, queue_count, worker_running):
        self.set_status(queue_count=int(queue_count or 0), worker_running=bool(worker_running))

    def defer(self, wait_seconds, *, state=None, provider=None):
        wait_seconds = max(0.0, float(wait_seconds or 0.0))
        interval_seconds = self.queue_interval_for_provider(provider)
        with self._attempt_lock:
            now = time.time()
            current_next = self._last_attempt_at + interval_seconds
            target_next = max(current_next, now + wait_seconds)
            self._last_attempt_at = target_next - interval_seconds
        updates = {"next_retry_at": target_next}
        if state:
            updates["state"] = state
        self.set_status(**updates)

    def wait_for_slot(self, provider=None):
        interval_seconds = self.queue_interval_for_provider(provider)
        while True:
            with self._attempt_lock:
                now = time.time()
                wait_seconds = interval_seconds - (now - self._last_attempt_at)
                if wait_seconds <= 0:
                    self._last_attempt_at = now
                    return
                next_retry_at = now + wait_seconds
            self.set_status(state="ok", next_retry_at=next_retry_at)
            threading.Event().wait(wait_seconds)
