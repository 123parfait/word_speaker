# -*- coding: utf-8 -*-
import os
import random
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from data.store import WordStore
from services.tts import speak_async
from services.diff_view import apply_diff


class MainView(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.store = WordStore()

        self.order_random = tk.BooleanVar(value=False)
        self.interval_var = tk.IntVar(value=2)
        self.volume_var = tk.IntVar(value=80)
        self.status_var = tk.StringVar(value="Not started")

        self.play_state = "stopped"  # stopped | playing | paused
        self.queue = []
        self.pos = -1
        self.current_word = None
        self.after_id = None

        self.build_ui()
        self.refresh_history()
        self.update_empty_state()
        self.update_order_button()
        self.update_speed_buttons()
        self.update_play_button()

    def build_ui(self):
        header = ttk.Frame(self)
        header.pack(fill=tk.X, pady=(0, 6))
        title = ttk.Label(header, text="Word Speaker", font=("Segoe UI", 14, "bold"))
        title.pack(side=tk.LEFT)

        ttk.Separator(self, orient="horizontal").pack(fill=tk.X, pady=(0, 10))

        main = ttk.Frame(self)
        main.pack()

        left = ttk.Frame(main, style="Card.TFrame")
        left.grid(row=0, column=0, sticky="n")
        mid = ttk.Frame(main, style="Card.TFrame")
        mid.grid(row=0, column=1, sticky="n", padx=(12, 0))
        right = ttk.Frame(main, style="Card.TFrame")
        right.grid(row=0, column=2, sticky="n", padx=(12, 0))

        # Left: Word list + player bar
        left_title = ttk.Label(left, text="Word List", style="Card.TLabel")
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
        self.empty_label.grid(row=3, column=0, columnspan=2, padx=12, pady=(0, 8), sticky="w")

        player = ttk.Frame(left, style="Card.TFrame")
        player.grid(row=4, column=0, columnspan=2, padx=12, pady=(4, 12), sticky="w")

        self.play_btn = ttk.Button(
            player, text="Start", style="CardButton.TButton", command=self.toggle_play
        )
        self.play_btn.grid(row=0, column=0, padx=(0, 8))

        self.settings_btn = ttk.Button(
            player, text="Settings ▾", command=self.toggle_settings
        )
        self.settings_btn.grid(row=0, column=1, padx=(0, 8))

        self.status_label = ttk.Label(
            left, textvariable=self.status_var, style="Card.TLabel", foreground="#444"
        )
        self.status_label.grid(row=5, column=0, columnspan=2, padx=12, pady=(0, 12), sticky="w")

        self.settings_panel = ttk.Frame(left, style="Card.TFrame")
        self.settings_panel.grid(row=6, column=0, columnspan=2, padx=12, pady=(0, 12), sticky="w")
        self.settings_panel.grid_remove()

        # Settings content
        order_frame = ttk.Frame(self.settings_panel, style="Card.TFrame")
        order_frame.pack(anchor="w", pady=2)
        ttk.Label(order_frame, text="Order:", style="Card.TLabel").pack(side=tk.LEFT)
        self.order_btn = ttk.Button(order_frame, text="", command=self.toggle_order)
        self.order_btn.pack(side=tk.LEFT, padx=6)

        speed_frame = ttk.Frame(self.settings_panel, style="Card.TFrame")
        speed_frame.pack(anchor="w", pady=2)
        ttk.Label(speed_frame, text="Speed:", style="Card.TLabel").pack(side=tk.LEFT)
        self.speed_buttons = []
        for v in [1, 2, 3, 5, 10]:
            btn = ttk.Button(
                speed_frame,
                text=f"{v}x",
                command=lambda val=v: self.set_interval(val),
            )
            btn.pack(side=tk.LEFT, padx=3)
            self.speed_buttons.append((v, btn))

        volume_frame = ttk.Frame(self.settings_panel, style="Card.TFrame")
        volume_frame.pack(anchor="w", pady=2)
        ttk.Label(volume_frame, text="Volume:", style="Card.TLabel").pack(side=tk.LEFT)
        self.volume_scale = ttk.Scale(
            volume_frame, from_=0, to=100, variable=self.volume_var, orient="horizontal", length=140
        )
        self.volume_scale.pack(side=tk.LEFT, padx=6)

        # Middle: History
        hist_title = ttk.Label(mid, text="History", style="Card.TLabel")
        hist_title.grid(row=0, column=0, columnspan=2, padx=12, pady=(12, 6), sticky="w")

        self.history_list = tk.Listbox(
            mid,
            width=34,
            height=20,
            bg="#ffffff",
            fg="#222222",
            selectbackground="#cce1ff",
            selectforeground="#111111",
            highlightthickness=1,
            highlightbackground="#d9dbe1",
        )
        self.history_list.grid(row=1, column=0, columnspan=2, padx=12, pady=(4, 6))
        self.history_list.bind("<Double-1>", self.on_history_open)

        self.history_empty = ttk.Label(
            mid, text="No history yet.", style="Card.TLabel", foreground="#666"
        )
        self.history_empty.grid(row=2, column=0, columnspan=2, padx=12, pady=(0, 6), sticky="w")

        btn_open = ttk.Button(mid, text="Import Selected", command=self.on_history_open)
        btn_open.grid(row=3, column=0, padx=12, pady=(0, 12), sticky="w")

        self.history_path = ttk.Label(mid, text="", style="Card.TLabel")
        self.history_path.grid(row=3, column=1, padx=12, pady=(0, 12), sticky="e")

        # Right: Input Check
        check_title = ttk.Label(right, text="Input Check", style="Card.TLabel")
        check_title.grid(row=0, column=0, columnspan=2, padx=12, pady=(12, 6), sticky="w")

        self.input_entry = ttk.Entry(right, width=36)
        self.input_entry.grid(row=1, column=0, padx=12, pady=(0, 6), sticky="w")

        self.check_btn = ttk.Button(right, text="Check", command=self.check_input)
        self.check_btn.grid(row=1, column=1, padx=6, pady=(0, 6), sticky="e")

        self.result_text = tk.Text(
            right,
            height=12,
            width=38,
            bg="#ffffff",
            fg="#222222",
            highlightthickness=1,
            highlightbackground="#d9dbe1",
        )
        self.result_text.grid(row=2, column=0, columnspan=2, padx=12, pady=(0, 12), sticky="w")
        self.result_text.tag_configure("wrong", foreground="#c62828")
        self.result_text.tag_configure("missing", foreground="#c62828", underline=1)
        self.result_text.tag_configure("extra", foreground="#ef6c00")
        self.result_text.config(state="disabled")

    # Data + history
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

    # Player controls
    def speak_selected(self):
        selection = self.listbox.curselection()
        if not selection:
            messagebox.showinfo("Info", "Please select a word first.")
            return
        word = self.listbox.get(selection[0])
        speak_async(word, self.volume_var.get() / 100.0)

    def toggle_settings(self):
        if self.settings_panel.winfo_ismapped():
            self.settings_panel.grid_remove()
            self.settings_btn.config(text="Settings ▾")
        else:
            self.settings_panel.grid()
            self.settings_btn.config(text="Settings ▴")

    def toggle_order(self):
        self.order_random.set(not self.order_random.get())
        self.update_order_button()
        # reset queue for a clean no-repeat cycle
        if self.play_state == "stopped":
            self.queue = []
            self.pos = -1

    def update_order_button(self):
        if self.order_random.get():
            self.order_btn.config(text="Random (no repeat)")
        else:
            self.order_btn.config(text="In order")

    def set_interval(self, seconds):
        self.interval_var.set(seconds)
        self.update_speed_buttons()

    def update_speed_buttons(self):
        for v, btn in self.speed_buttons:
            if v == self.interval_var.get():
                btn.config(style="SelectedSpeed.TButton")
            else:
                btn.config(style="Speed.TButton")

    def build_queue(self):
        indices = list(range(len(self.store.words)))
        if self.order_random.get():
            random.shuffle(indices)
        return indices

    def toggle_play(self):
        if not self.store.words:
            messagebox.showinfo("Info", "Please import words first.")
            return
        if self.play_state == "playing":
            self.play_state = "paused"
            self.cancel_schedule()
            self.status_var.set("Paused")
            self.update_play_button()
            return

        # start or resume
        if not self.queue:
            self.queue = self.build_queue()
            self.pos = 0
            self.set_current_word()
        elif self.pos < 0:
            self.pos = 0
            self.set_current_word()

        self.play_state = "playing"
        self.update_play_button()
        self.play_current()
        self.schedule_next()

    def play_current(self):
        if not self.current_word:
            return
        speak_async(self.current_word, self.volume_var.get() / 100.0)

    def schedule_next(self):
        self.cancel_schedule()
        if self.play_state != "playing":
            return
        interval = max(1, int(self.interval_var.get()))
        self.after_id = self.after(interval * 1000, self.next_word)

    def next_word(self):
        if self.play_state != "playing":
            return
        if not self.queue:
            self.queue = self.build_queue()
            self.pos = 0
        else:
            self.pos += 1
            if self.pos >= len(self.queue):
                # cycle complete, rebuild
                self.queue = self.build_queue()
                self.pos = 0
        self.set_current_word()
        self.play_current()
        self.schedule_next()

    def set_current_word(self):
        idx = self.queue[self.pos]
        self.current_word = self.store.words[idx]
        self.status_var.set(
            f"Current: {self.pos + 1}/{len(self.queue)}  Word: {self.current_word}"
        )

    def cancel_schedule(self):
        if self.after_id:
            self.after_cancel(self.after_id)
            self.after_id = None

    def update_play_button(self):
        if self.play_state == "playing":
            self.play_btn.config(text="Pause")
        else:
            self.play_btn.config(text="Start")

    # Input check
    def check_input(self):
        if not self.current_word:
            messagebox.showinfo("Info", "No current word yet.")
            return
        user_text = self.input_entry.get()
        apply_diff(self.result_text, self.current_word, user_text)
