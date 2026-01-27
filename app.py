# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import filedialog, messagebox
import csv
import threading
import pyttsx3

# Keep UI responsive by speaking in a background thread
_speak_lock = threading.Lock()

def _speak_text(text):
    try:
        with _speak_lock:
            engine = pyttsx3.init(driverName="sapi5")
            engine.say(text)
            engine.runAndWait()
    except Exception as e:
        messagebox.showerror("Speech Error", f"Error: {e}")

def speak_word():
    selection = listbox.curselection()
    if not selection:
        messagebox.showinfo("Info", "Please select a word first.")
        return
    word = listbox.get(selection[0])
    threading.Thread(target=_speak_text, args=(word,), daemon=True).start()

def load_words():
    path = filedialog.askopenfilename(
        title="Choose a word list",
        filetypes=[("Text files", "*.txt"), ("CSV files", "*.csv")]
    )
    if not path:
        return

    listbox.delete(0, tk.END)

    if path.endswith(".txt"):
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                word = line.strip()
                if word:
                    listbox.insert(tk.END, word)

    elif path.endswith(".csv"):
        with open(path, "r", encoding="utf-8") as f:
            reader = csv.reader(f)
            for row in reader:
                if row:
                    listbox.insert(tk.END, row[0].strip())

root = tk.Tk()
root.title("Word Speaker")

btn_frame = tk.Frame(root)
btn_frame.pack(pady=8)

btn_load = tk.Button(btn_frame, text="Import", command=load_words)
btn_load.pack(side=tk.LEFT, padx=5)

btn_speak = tk.Button(btn_frame, text="Speak", command=speak_word)
btn_speak.pack(side=tk.LEFT, padx=5)

listbox = tk.Listbox(root, width=40, height=15)
listbox.pack(padx=10, pady=10)

root.mainloop()
