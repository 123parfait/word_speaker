# -*- coding: utf-8 -*-
import threading
import pyttsx3
from tkinter import messagebox

_speak_lock = threading.Lock()


def _speak_text(text):
    try:
        with _speak_lock:
            engine = pyttsx3.init(driverName="sapi5")
            engine.say(text)
            engine.runAndWait()
    except Exception as e:
        messagebox.showerror("Speech Error", f"Error: {e}")


def speak_async(text):
    threading.Thread(target=_speak_text, args=(text,), daemon=True).start()
