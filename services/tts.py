# -*- coding: utf-8 -*-
import threading
import pyttsx3
from tkinter import messagebox

from services.voice_manager import select_voice

_lock = threading.Lock()
_current_engine = None
_token = 0


def stop():
    global _current_engine
    try:
        with _lock:
            if _current_engine:
                _current_engine.stop()
    except Exception:
        pass


def speak_async(text, volume=1.0, cancel_before=False):
    global _current_engine, _token
    if cancel_before:
        stop()
    _token += 1
    my_token = _token

    def _run():
        global _current_engine
        try:
            with _lock:
                if my_token != _token:
                    return
                engine = pyttsx3.init(driverName="sapi5")
                select_voice(engine)
                _current_engine = engine
                engine.setProperty("volume", max(0.0, min(1.0, volume)))
                engine.say(text)
                engine.runAndWait()
        except Exception as e:
            messagebox.showerror("Speech Error", f"Error: {e}")
        finally:
            with _lock:
                if _current_engine:
                    try:
                        _current_engine.stop()
                    except Exception:
                        pass
                    _current_engine = None

    threading.Thread(target=_run, daemon=True).start()


def cancel_all():
    global _token
    _token += 1
    stop()
