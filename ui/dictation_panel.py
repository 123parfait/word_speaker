# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, messagebox
import random

from services.tts import speak_async
from services.diff_view import apply_diff


class DictationPanel(ttk.Frame):
    def __init__(self, master, store):
        super().__init__(master, style="Card.TFrame")
        self.store = store
        self.state = {
            "indices": [],
            "pos": -1,
            "current": None,
            "running": False,
            "after_id": None,
            "paused": False,
        }

        self.mode_var = tk.StringVar(value="order")
        self.style_var = tk.StringVar(value="type")
        self.auto_var = tk.BooleanVar(value=False)
        self.manual_var = tk.BooleanVar(value=False)
        self.interval_var = tk.StringVar(value="2.0")
        self.status_var = tk.StringVar(value="Not started")

        self.build_ui()
        self.on_style_change()

    def build_ui(self):
        title = ttk.Label(self, text="Dictation", style="Card.TLabel")
        title.grid(row=0, column=0, columnspan=2, padx=12, pady=(12, 6), sticky="w")

        mode_frame = ttk.Frame(self, style="Card.TFrame")
        mode_frame.grid(row=1, column=0, sticky="w", padx=12, pady=2)
        ttk.Label(mode_frame, text="Order:", style="Card.TLabel").pack(side=tk.LEFT)
        ttk.Radiobutton(mode_frame, text="In order", variable=self.mode_var, value="order").pack(
            side=tk.LEFT
        )
        ttk.Radiobutton(
            mode_frame, text="Random (no repeat)", variable=self.mode_var, value="random"
        ).pack(side=tk.LEFT)

        style_frame = ttk.Frame(self, style="Card.TFrame")
        style_frame.grid(row=2, column=0, sticky="w", padx=12, pady=2)
        ttk.Label(style_frame, text="Mode:", style="Card.TLabel").pack(side=tk.LEFT)
        ttk.Radiobutton(
            style_frame,
            text="Type & Check",
            variable=self.style_var,
            value="type",
            command=self.on_style_change,
        ).pack(side=tk.LEFT)
        ttk.Radiobutton(
            style_frame,
            text="Paper Only",
            variable=self.style_var,
            value="paper",
            command=self.on_style_change,
        ).pack(side=tk.LEFT)

        interval_frame = ttk.Frame(self, style="Card.TFrame")
        interval_frame.grid(row=3, column=0, sticky="w", padx=12, pady=2)
        ttk.Label(interval_frame, text="Interval (s):", style="Card.TLabel").pack(side=tk.LEFT)
        interval_entry = ttk.Entry(interval_frame, width=6, textvariable=self.interval_var)
        interval_entry.pack(side=tk.LEFT, padx=4)
        ttk.Checkbutton(
            interval_frame,
            text="Auto play",
            variable=self.auto_var,
            command=self.on_auto_toggle,
        ).pack(side=tk.LEFT)
        ttk.Checkbutton(
            interval_frame,
            text="Manual first",
            variable=self.manual_var,
        ).pack(side=tk.LEFT, padx=(6, 0))

        btn_frame = ttk.Frame(self, style="Card.TFrame")
        btn_frame.grid(row=4, column=0, sticky="w", padx=12, pady=6)
        ttk.Button(
            btn_frame, text="Start", command=self.start_dictation, style="Primary.TButton"
        ).grid(row=0, column=0, padx=4, pady=2)
        ttk.Button(btn_frame, text="Play current", command=self.play_current).grid(
            row=0, column=1, padx=4, pady=2
        )
        ttk.Button(btn_frame, text="Next", command=self.next_word).grid(
            row=0, column=2, padx=4, pady=2
        )
        self.pause_btn = ttk.Button(btn_frame, text="Pause", command=self.toggle_pause)
        self.pause_btn.grid(row=0, column=3, padx=4, pady=2)
        ttk.Button(btn_frame, text="Stop", command=self.stop_dictation).grid(
            row=0, column=4, padx=4, pady=2
        )

        status_label = ttk.Label(
            self, textvariable=self.status_var, foreground="#444", style="Card.TLabel"
        )
        status_label.grid(row=5, column=0, sticky="w", padx=12, pady=(2, 6))

        input_frame = ttk.Frame(self, style="Card.TFrame")
        input_frame.grid(row=6, column=0, sticky="ew", padx=12, pady=(4, 6))
        input_frame.grid_columnconfigure(0, weight=1)
        input_title = ttk.Label(input_frame, text="Input Check", style="Card.TLabel")
        input_title.grid(row=0, column=0, columnspan=2, sticky="w", pady=(2, 4))

        self.input_entry = ttk.Entry(input_frame, width=56)
        self.input_entry.grid(row=1, column=0, padx=(0, 6), pady=(0, 6), sticky="ew")

        self.check_btn = ttk.Button(input_frame, text="Check", command=self.check_input)
        self.check_btn.grid(row=1, column=1, pady=(0, 6), sticky="w")

        self.result_text = tk.Text(
            self,
            height=7,
            width=42,
            bg="#ffffff",
            fg="#222222",
            highlightthickness=1,
            highlightbackground="#d9dbe1",
        )
        self.result_text.grid(row=7, column=0, padx=12, pady=(0, 12), sticky="w")
        self.result_text.tag_configure("wrong", foreground="#c62828")
        self.result_text.tag_configure("missing", foreground="#c62828", underline=1)
        self.result_text.tag_configure("extra", foreground="#ef6c00")
        self.result_text.config(state="disabled")

    def build_dictation_queue(self):
        indices = list(range(len(self.store.words)))
        if self.mode_var.get() == "random":
            random.shuffle(indices)
        return indices

    def start_dictation(self):
        if not self.store.words:
            messagebox.showinfo("Info", "Please import words first.")
            return
        self.reset_dictation_state()
        self.state["running"] = True
        self.state["paused"] = False
        self.state["indices"] = self.build_dictation_queue()
        self.state["pos"] = 0
        self.set_current_word()
        self.play_current()
        if self.auto_var.get():
            self.schedule_next()

    def stop_dictation(self):
        self.state["running"] = False
        self.state["paused"] = False
        self.cancel_schedule()
        self.pause_btn.config(text="Pause")

    def next_word(self, auto=False):
        if not self.state["running"]:
            return
        if self.state["paused"] and (auto or not self.manual_var.get()):
            return
        if not self.state["indices"]:
            messagebox.showinfo("Info", "Please start dictation first.")
            return
        self.state["pos"] += 1
        if self.state["pos"] >= len(self.state["indices"]):
            self.state["current"] = None
            self.update_status(done=True)
            self.stop_dictation()
            return
        self.set_current_word()
        self.play_current()
        if self.auto_var.get() and not self.state["paused"]:
            self.schedule_next()

    def set_current_word(self):
        idx = self.state["indices"][self.state["pos"]]
        self.state["current"] = self.store.words[idx]
        self.update_status()

    def play_current(self):
        if not self.state["current"]:
            messagebox.showinfo("Info", "No current word yet.")
            return
        speak_async(self.state["current"])

    def schedule_next(self):
        self.cancel_schedule()
        if not self.state["running"] or self.state["paused"]:
            return
        try:
            interval = float(self.interval_var.get())
        except Exception:
            interval = 2.0
        interval = max(0.5, interval)
        self.state["after_id"] = self.after(
            int(interval * 1000), lambda: self.next_word(auto=True)
        )

    def cancel_schedule(self):
        after_id = self.state.get("after_id")
        if after_id:
            self.after_cancel(after_id)
            self.state["after_id"] = None

    def update_status(self, done=False):
        if done:
            self.status_var.set("Completed")
            return
        if not self.state["indices"]:
            self.status_var.set("Not started")
            return
        self.status_var.set(
            f"Current: {self.state['pos'] + 1}/{len(self.state['indices'])}  Word: {self.state['current']}"
        )

    def on_style_change(self):
        if self.style_var.get() == "paper":
            self.input_entry.config(state="disabled")
            self.check_btn.config(state="disabled")
            self.result_text.config(state="normal")
            self.result_text.delete("1.0", tk.END)
            self.result_text.insert(tk.END, "Paper mode: play only, no checking.\n")
            self.result_text.config(state="disabled")
        else:
            self.input_entry.config(state="normal")
            self.check_btn.config(state="normal")
            self.input_entry.focus_set()
            self.clear_result()

    def on_auto_toggle(self):
        if not self.state["running"]:
            return
        if self.auto_var.get():
            if not self.state["paused"]:
                self.schedule_next()
        else:
            self.cancel_schedule()

    def toggle_pause(self):
        if not self.state["running"]:
            return
        self.state["paused"] = not self.state["paused"]
        if self.state["paused"]:
            self.cancel_schedule()
            self.status_var.set("Paused")
            self.pause_btn.config(text="Resume")
        else:
            self.pause_btn.config(text="Pause")
            self.update_status()
            if self.auto_var.get():
                self.schedule_next()

    def clear_result(self):
        self.result_text.config(state="normal")
        self.result_text.delete("1.0", tk.END)
        self.result_text.config(state="disabled")

    def check_input(self):
        if not self.state["current"]:
            messagebox.showinfo("Info", "Please start dictation first.")
            return
        user_text = self.input_entry.get()
        apply_diff(self.result_text, self.state["current"], user_text)

    def reset_dictation_state(self):
        self.stop_dictation()
        self.state["indices"] = []
        self.state["pos"] = -1
        self.state["current"] = None
        self.state["paused"] = False
        self.update_status()
