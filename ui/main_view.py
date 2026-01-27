# -*- coding: utf-8 -*-
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from data.store import WordStore
from services.tts import speak_async
from ui.dictation_panel import DictationPanel


class MainView(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.store = WordStore()
        self.dictation_visible = False
        self.history_visible = False

        self.build_ui()
        self.refresh_history()
        self.update_empty_state()

    def build_ui(self):
        header = ttk.Frame(self)
        header.pack(fill=tk.X, pady=(0, 6))

        title = ttk.Label(header, text="Word Speaker", font=("Segoe UI", 14, "bold"))
        title.pack(side=tk.LEFT)

        header_sep = ttk.Separator(self, orient="horizontal")
        header_sep.pack(fill=tk.X, pady=(0, 10))

        main = ttk.Frame(self)
        main.pack()

        left = ttk.Frame(main, style="Card.TFrame")
        left.grid(row=0, column=0, sticky="n")

        mid_sep = ttk.Separator(main, orient="vertical")
        mid_sep.grid(row=0, column=1, sticky="ns", padx=10)

        right = ttk.Frame(main, style="Card.TFrame")
        right.grid(row=0, column=2, sticky="n")

        left_title = ttk.Label(left, text="Word List  [W]", style="Card.TLabel")
        left_title.grid(row=0, column=0, columnspan=2, padx=12, pady=(12, 6), sticky="w")

        btn_load = ttk.Button(
            left, text="Import Words", style="Primary.TButton", command=self.load_words
        )
        btn_load.grid(row=1, column=0, padx=12, pady=6, sticky="w")

        btn_speak = ttk.Button(left, text="Speak Selected", command=self.speak_selected)
        btn_speak.grid(row=1, column=1, padx=12, pady=6, sticky="e")

        self.listbox = tk.Listbox(
            left,
            width=36,
            height=18,
            bg="#ffffff",
            fg="#222222",
            selectbackground="#cce1ff",
            selectforeground="#111111",
            highlightthickness=1,
            highlightbackground="#d9dbe1",
        )
        self.listbox.grid(row=2, column=0, columnspan=2, padx=12, pady=(4, 6))

        self.empty_label = ttk.Label(
            left,
            text="No words yet. Click “Import Words” to get started.",
            style="Card.TLabel",
            foreground="#666",
        )
        self.empty_label.grid(row=3, column=0, columnspan=2, padx=12, pady=(0, 12), sticky="w")

        right_title = ttk.Label(right, text="Action Panel  [*]", style="Card.TLabel")
        right_title.grid(row=0, column=0, padx=12, pady=(12, 6), sticky="w")

        actions = ttk.Frame(right, style="Card.TFrame")
        actions.grid(row=1, column=0, padx=12, pady=(0, 8), sticky="w")

        self.dict_btn = ttk.Button(
            actions,
            text="Dictation  >",
            style="CardButton.TButton",
            command=self.toggle_dictation,
        )
        self.dict_btn.grid(row=0, column=0, padx=(0, 8), pady=4)

        self.hist_btn = ttk.Button(
            actions,
            text="History  >",
            style="CardButton.TButton",
            command=self.toggle_history,
        )
        self.hist_btn.grid(row=0, column=1, padx=(0, 8), pady=4)

        self.dictation_panel = DictationPanel(right, self.store)
        self.dictation_panel.grid(row=2, column=0, padx=12, pady=(0, 12), sticky="w")
        self.dictation_panel.grid_remove()

        self.history_panel = ttk.Frame(right, style="Card.TFrame")
        self.history_panel.grid(row=2, column=0, padx=12, pady=(0, 12), sticky="w")
        self.history_panel.grid_remove()

        hist_title = ttk.Label(self.history_panel, text="History  [H]", style="Card.TLabel")
        hist_title.grid(row=0, column=0, columnspan=2, padx=10, pady=(8, 4), sticky="w")

        self.history_list = tk.Listbox(
            self.history_panel,
            width=42,
            height=8,
            bg="#ffffff",
            fg="#222222",
            selectbackground="#cce1ff",
            selectforeground="#111111",
            highlightthickness=1,
            highlightbackground="#d9dbe1",
        )
        self.history_list.grid(row=1, column=0, columnspan=2, padx=10, pady=(4, 6))
        self.history_list.bind("<Double-1>", self.on_history_open)

        self.history_empty = ttk.Label(
            self.history_panel, text="No history yet.", style="Card.TLabel", foreground="#666"
        )
        self.history_empty.grid(row=2, column=0, columnspan=2, padx=10, pady=(0, 6), sticky="w")

        btn_open = ttk.Button(self.history_panel, text="Import Selected", command=self.on_history_open)
        btn_open.grid(row=3, column=0, padx=10, pady=(0, 10), sticky="w")

        self.history_path = ttk.Label(self.history_panel, text="", style="Card.TLabel")
        self.history_path.grid(row=3, column=1, padx=10, pady=(0, 10), sticky="e")

    def load_words(self):
        path = filedialog.askopenfilename(
            title="Choose a word list",
            filetypes=[("Text files", "*.txt"), ("CSV files", "*.csv")],
        )
        if not path:
            return
        words = self.store.load_from_file(path)
        self.render_words(words)
        self.refresh_history()

    def render_words(self, words):
        self.listbox.delete(0, tk.END)
        for w in words:
            self.listbox.insert(tk.END, w)
        self.update_empty_state()

    def update_empty_state(self):
        if self.store.words:
            self.empty_label.grid_remove()
        else:
            self.empty_label.grid()

    def speak_selected(self):
        selection = self.listbox.curselection()
        if not selection:
            messagebox.showinfo("Info", "Please select a word first.")
            return
        word = self.listbox.get(selection[0])
        speak_async(word)

    def toggle_dictation(self):
        self.dictation_visible = not self.dictation_visible
        if self.dictation_visible:
            self.history_visible = False
            self.history_panel.grid_remove()
            self.hist_btn.config(text="History  >", style="CardButton.TButton")
            self.dictation_panel.grid()
            self.dict_btn.config(text="Dictation  v", style="SelectedCardButton.TButton")
        else:
            self.dictation_panel.grid_remove()
            self.dict_btn.config(text="Dictation  >", style="CardButton.TButton")

    def toggle_history(self):
        self.history_visible = not self.history_visible
        if self.history_visible:
            self.dictation_visible = False
            self.dictation_panel.grid_remove()
            self.dict_btn.config(text="Dictation  >", style="CardButton.TButton")
            self.history_panel.grid()
            self.hist_btn.config(text="History  v", style="SelectedCardButton.TButton")
        else:
            self.history_panel.grid_remove()
            self.hist_btn.config(text="History  >", style="CardButton.TButton")

    def refresh_history(self):
        history = self.store.load_history()
        self.history_list.delete(0, tk.END)
        for item in history:
            label = f"{item.get('name', '')}  ({item.get('time', '')})"
            self.history_list.insert(tk.END, label)
        if history:
            self.history_path.config(text=history[0].get("path", ""))
            self.history_empty.grid_remove()
        else:
            self.history_path.config(text="")
            self.history_empty.grid()

    def on_history_open(self, _event=None):
        sel = self.history_list.curselection()
        if not sel:
            return
        history = self.store.load_history()
        idx = sel[0]
        if idx >= len(history):
            return
        path = history[idx].get("path")
        if not path or not os.path.exists(path):
            messagebox.showinfo("Info", "File not found or moved.")
            return
        words = self.store.load_from_file(path)
        self.render_words(words)
        self.refresh_history()
