# -*- coding: utf-8 -*-
import os
import random
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from data.store import WordStore
from services.tts import speak_async, cancel_all as tts_cancel_all
from services.diff_view import apply_diff


class Tooltip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip = None
        self.widget.bind("<Enter>", self.show)
        self.widget.bind("<Leave>", self.hide)

    def show(self, _event=None):
        if self.tip or not self.text:
            return
        x = self.widget.winfo_rootx() + 10
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.geometry(f"+{x}+{y}")
        label = tk.Label(
            self.tip,
            text=self.text,
            background="#222222",
            foreground="#ffffff",
            padx=6,
            pady=3,
        )
        label.pack()

    def hide(self, _event=None):
        if self.tip:
            self.tip.destroy()
            self.tip = None


class MainView(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.store = WordStore()

        self.order_mode = tk.StringVar(value="order")  # order | random_no_repeat
        self.interval_var = tk.IntVar(value=2)
        self.volume_var = tk.IntVar(value=80)
        self.loop_var = tk.BooleanVar(value=True)
        self.stop_at_end_var = tk.BooleanVar(value=not self.loop_var.get())
        self.status_var = tk.StringVar(value="Not started")

        self.play_state = "stopped"  # stopped | playing | paused
        self.queue = []
        self.pos = -1
        self.current_word = None
        self.after_id = None
        self.play_token = 0

        self.history_visible = True
        self.check_visible = False
        self.wordlist_hidden = False
        self.order_btn = None
        self.order_btn_rand = None
        self.order_tip = None
        self.order_tip_rand = None
        self.loop_btn = None
        self.loop_btn_stop = None
        self.stop_at_end_check = None
        self.speed_buttons = []
        self.volume_scale = None
        self.player_frame = None
        self.play_btn_check = None
        self.settings_btn_check = None
        self.check_btn_toggle_check = None
        self.hist_btn_toggle_check = None
        self.check_controls = None
        self.hide_words_btn = None

        self.build_ui()
        self.refresh_history()
        self.update_empty_state()
        self.update_order_button()
        self.update_loop_button()
        self.update_speed_buttons()
        self.update_play_button()
        self.update_right_visibility()

    def build_ui(self):
        header = ttk.Frame(self)
        header.pack(fill=tk.X, pady=(0, 6))
        title = ttk.Label(header, text="Word Speaker", font=("Segoe UI", 14, "bold"))
        title.pack(side=tk.LEFT)

        ttk.Separator(self, orient="horizontal").pack(fill=tk.X, pady=(0, 10))

        main = ttk.Frame(self)
        main.pack()

        self.left = ttk.Frame(main, style="Card.TFrame")
        self.left.grid(row=0, column=0, sticky="n")
        self.mid_sep = ttk.Separator(main, orient="vertical")
        self.mid_sep.grid(row=0, column=1, sticky="ns", padx=10)
        self.right = ttk.Frame(main, style="Card.TFrame")
        self.right.grid(row=0, column=2, sticky="n")

        # Left: Word list + player bar
        left_title = ttk.Label(self.left, text="Word List", style="Card.TLabel")
        left_title.grid(row=0, column=0, columnspan=2, padx=12, pady=(12, 6), sticky="w")

        btn_load = ttk.Button(self.left, text="⭳", style="Primary.TButton", command=self.load_words)
        btn_load.grid(row=1, column=0, padx=12, pady=6, sticky="w")
        Tooltip(btn_load, "Import Words")

        btn_speak = ttk.Button(self.left, text="🔊", command=self.speak_selected)
        btn_speak.grid(row=1, column=1, padx=12, pady=6, sticky="e")
        Tooltip(btn_speak, "Speak Selected")

        self.listbox = tk.Listbox(
            self.left,
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
        self.listbox.bind("<<ListboxSelect>>", self.on_word_selected)

        self.empty_label = ttk.Label(
            self.left,
            text="No words yet. Click the import button to get started.",
            style="Card.TLabel",
            foreground="#666",
        )
        self.empty_label.grid(row=3, column=0, columnspan=2, padx=12, pady=(0, 8), sticky="w")

        self.player_frame = ttk.Frame(self.left, style="Card.TFrame")
        self.player_frame.grid(row=4, column=0, columnspan=2, padx=12, pady=(4, 6), sticky="ew")
        self.player_frame.grid_columnconfigure(0, weight=1)
        self.player_frame.grid_columnconfigure(1, weight=1)
        self.player_frame.grid_columnconfigure(2, weight=1)
        self.player_frame.grid_columnconfigure(3, weight=1)

        self.play_btn = ttk.Button(
            self.player_frame, text="▶", style="Icon.TButton", width=4, command=self.toggle_play
        )
        self.play_btn.grid(row=0, column=0, padx=4, sticky="ew")
        Tooltip(self.play_btn, "Start / Pause")

        self.settings_btn = ttk.Button(
            self.player_frame, text="⚙", style="Icon.TButton", width=4, command=self.toggle_settings
        )
        self.settings_btn.grid(row=0, column=1, padx=4, sticky="ew")
        Tooltip(self.settings_btn, "Settings")

        self.check_btn_toggle = ttk.Button(
            self.player_frame, text="✅", style="Icon.TButton", width=4, command=self.toggle_check
        )
        self.check_btn_toggle.grid(row=0, column=2, padx=4, sticky="ew")
        Tooltip(self.check_btn_toggle, "Dictation Check")

        self.hist_btn_toggle = ttk.Button(
            self.player_frame, text="🕘", style="Icon.TButton", width=4, command=self.toggle_history
        )
        self.hist_btn_toggle.grid(row=0, column=3, padx=4, sticky="ew")
        Tooltip(self.hist_btn_toggle, "History")

        self.status_label = ttk.Label(
            self.left, textvariable=self.status_var, style="Card.TLabel", foreground="#444"
        )
        self.status_label.grid(row=5, column=0, columnspan=2, padx=12, pady=(0, 12), sticky="w")

        # Settings popup (created on demand)
        self.settings_window = None

        # Right: Dictation Check (toggle) + History (default open)
        self.check_panel = ttk.Frame(self.right, style="Card.TFrame")
        self.check_panel.grid(row=0, column=0, columnspan=2, padx=12, pady=(12, 8), sticky="w")
        self.check_panel.grid_remove()

        self.check_controls = ttk.Frame(self.check_panel, style="Card.TFrame")
        self.check_controls.grid(row=0, column=0, columnspan=2, padx=10, pady=(8, 4), sticky="w")
        self.hide_words_btn = ttk.Button(
            self.check_controls, text="Hide Word List", command=self.toggle_wordlist_visibility
        )
        self.hide_words_btn.pack(side=tk.LEFT, padx=(0, 6))
        Tooltip(self.hide_words_btn, "Hide/Show word list")

        self.play_btn_check = ttk.Button(
            self.check_controls, text="▶", style="Icon.TButton", width=4, command=self.toggle_play
        )
        self.play_btn_check.pack(side=tk.LEFT, padx=3)
        Tooltip(self.play_btn_check, "Start / Pause")

        self.settings_btn_check = ttk.Button(
            self.check_controls, text="⚙", style="Icon.TButton", width=4, command=self.toggle_settings
        )
        self.settings_btn_check.pack(side=tk.LEFT, padx=3)
        Tooltip(self.settings_btn_check, "Settings")

        self.check_btn_toggle_check = ttk.Button(
            self.check_controls, text="✅", style="Icon.TButton", width=4, command=self.toggle_check
        )
        self.check_btn_toggle_check.pack(side=tk.LEFT, padx=3)
        Tooltip(self.check_btn_toggle_check, "Dictation Check")

        self.hist_btn_toggle_check = ttk.Button(
            self.check_controls, text="🕘", style="Icon.TButton", width=4, command=self.toggle_history
        )
        self.hist_btn_toggle_check.pack(side=tk.LEFT, padx=3)
        Tooltip(self.hist_btn_toggle_check, "History")

        check_title = ttk.Label(self.check_panel, text="Dictation Check", style="Card.TLabel")
        check_title.grid(row=1, column=0, columnspan=2, padx=10, pady=(4, 4), sticky="w")

        self.input_entry = ttk.Entry(self.check_panel, width=30)
        self.input_entry.grid(row=2, column=0, padx=10, pady=(0, 6), sticky="w")
        self.input_entry.bind("<Return>", self.on_check_enter)
        self.input_entry.bind("<space>", self.on_input_space)

        self.check_btn = ttk.Button(self.check_panel, text="✔", command=self.check_input)
        self.check_btn.grid(row=2, column=1, padx=6, pady=(0, 6), sticky="e")
        Tooltip(self.check_btn, "Check")

        self.result_text = tk.Text(
            self.check_panel,
            height=8,
            width=32,
            bg="#ffffff",
            fg="#222222",
            highlightthickness=1,
            highlightbackground="#d9dbe1",
        )
        self.result_text.grid(row=3, column=0, columnspan=2, padx=10, pady=(0, 10), sticky="w")
        self.result_text.tag_configure("wrong", foreground="#c62828")
        self.result_text.tag_configure("missing", foreground="#c62828", underline=1)
        self.result_text.tag_configure("extra", foreground="#ef6c00")
        self.result_text.config(state="disabled")

        self.check_shortcuts = ttk.Frame(self.check_panel, style="Card.TFrame")
        self.check_shortcuts.grid(row=3, column=1, padx=(6, 10), pady=(0, 10), sticky="nw")
        ttk.Label(self.check_shortcuts, text="Enter: ✔", style="Card.TLabel").pack(anchor="w")
        ttk.Label(self.check_shortcuts, text="Space: CLR", style="Card.TLabel").pack(anchor="w")

        self.check_sep = ttk.Separator(self.right, orient="horizontal")
        self.check_sep.grid(row=1, column=0, columnspan=2, padx=12, pady=(0, 8), sticky="ew")
        self.check_sep.grid_remove()

        self.history_panel = ttk.Frame(self.right, style="Card.TFrame")
        self.history_panel.grid(row=2, column=0, columnspan=2, padx=12, pady=(0, 12), sticky="w")

        hist_title = ttk.Label(self.history_panel, text="History", style="Card.TLabel")
        hist_title.grid(row=0, column=0, padx=10, pady=(8, 4), sticky="w")

        self.history_list = tk.Listbox(
            self.history_panel,
            width=30,
            height=16,
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

        btn_open = ttk.Button(self.history_panel, text="⭳", command=self.on_history_open)
        btn_open.grid(row=3, column=0, padx=10, pady=(0, 10), sticky="w")
        Tooltip(btn_open, "Import Selected")

        self.history_path = ttk.Label(self.history_panel, text="", style="Card.TLabel")
        self.history_path.grid(row=3, column=1, padx=10, pady=(0, 10), sticky="e")

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
        self.reset_playback_state()

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
        self.reset_playback_state()

    # Player controls
    def speak_selected(self):
        selection = self.listbox.curselection()
        if not selection:
            messagebox.showinfo("Info", "Please select a word first.")
            return
        word = self.listbox.get(selection[0])
        speak_async(word, self.volume_var.get() / 100.0, cancel_before=True)

    def toggle_settings(self):
        if self.settings_window and self.settings_window.winfo_exists():
            self.settings_window.lift()
            return
        self.open_settings_window()

    def toggle_history(self):
        self.history_visible = not self.history_visible
        if self.history_visible:
            self.history_panel.grid()
        else:
            self.history_panel.grid_remove()
        self.update_right_visibility()

    def toggle_check(self):
        self.check_visible = not self.check_visible
        if self.check_visible:
            self.check_panel.grid()
            self.check_sep.grid()
        else:
            self.check_panel.grid_remove()
            self.check_sep.grid_remove()
        self.update_right_visibility()

    def set_order_mode(self, mode):
        self.order_mode.set(mode)
        self.update_order_button()
        if self.order_tip:
            self.order_tip.hide()
        if self.order_tip_rand:
            self.order_tip_rand.hide()
        tts_cancel_all()
        self.play_token += 1
        self.rebuild_queue_on_mode_change()

    def update_order_button(self):
        mode = self.order_mode.get()
        if self.order_btn:
            self.order_btn.config(style="SelectedSpeed.TButton" if mode == "order" else "Speed.TButton")
        if self.order_btn_rand:
            self.order_btn_rand.config(
                style="SelectedSpeed.TButton" if mode == "random_no_repeat" else "Speed.TButton"
            )
        if self.order_tip:
            self.order_tip.text = "In order (from current)"
        if self.order_tip_rand:
            self.order_tip_rand.text = "Random (no repeat)"

    def set_interval(self, seconds):
        self.interval_var.set(seconds)
        self.update_speed_buttons()
        if self.play_state == "playing":
            self.schedule_next()
            self.play_current()

    def update_speed_buttons(self):
        for v, btn in self.speed_buttons:
            if v == self.interval_var.get():
                btn.config(style="SelectedSpeed.TButton")
            else:
                btn.config(style="Speed.TButton")

    def on_volume_change(self, _value=None):
        if self.volume_value:
            self.volume_value.config(text=f"{int(self.volume_var.get())}%")

    def set_loop_mode(self, loop_all):
        self.loop_var.set(bool(loop_all))
        self.update_loop_button()

    def update_loop_button(self):
        self.stop_at_end_var.set(not self.loop_var.get())

    def on_stop_at_end_toggle(self):
        self.loop_var.set(not self.stop_at_end_var.get())
        self.update_loop_button()
        # Rebuild queue so the current play order matches the new stop/loop choice.
        self.rebuild_queue_on_mode_change()

    def open_settings_window(self):
        self.settings_window = tk.Toplevel(self)
        self.settings_window.title("Settings")
        self.settings_window.configure(bg="#f6f7fb")
        self.settings_window.resizable(False, False)

        container = ttk.Frame(self.settings_window, style="Card.TFrame")
        container.pack(padx=10, pady=10)

        left_menu = ttk.Frame(container, style="Card.TFrame", width=120)
        left_menu.grid(row=0, column=0, sticky="n")
        right_panel = ttk.Frame(container, style="Card.TFrame", width=360, height=260)
        right_panel.grid(row=0, column=1, padx=(10, 0), sticky="n")
        right_panel.grid_propagate(False)

        self.settings_sections_visible = {"order": True, "speed": False, "volume": False}
        sections = []

        def rebuild_sections():
            for item in sections:
                item["frame"].pack_forget()
                item["sep"].pack_forget()

            visible_keys = [k for k, v in self.settings_sections_visible.items() if v]
            if not visible_keys:
                right_panel.grid_remove()
                return
            right_panel.grid()

            # pack in fixed order
            for idx, item in enumerate(sections):
                key = item["key"]
                if not self.settings_sections_visible.get(key, False):
                    continue
                item["frame"].pack(fill=tk.X, pady=(0, 6))
                # separator only if there is a visible section after this one
                has_next = False
                for later in sections[idx + 1 :]:
                    if self.settings_sections_visible.get(later["key"], False):
                        has_next = True
                        break
                if has_next:
                    item["sep"].pack(fill=tk.X, pady=(0, 6))

        def toggle_section(key):
            self.settings_sections_visible[key] = not self.settings_sections_visible.get(key, False)
            rebuild_sections()

        # Left menu
        for label, key in [("Order", "order"), ("Speed", "speed"), ("Volume", "volume")]:
            btn = ttk.Button(left_menu, text=label, command=lambda k=key: toggle_section(k))
            btn.pack(fill=tk.X, pady=4)

        # Order section
        order_section = ttk.Frame(right_panel, style="Card.TFrame")
        ttk.Label(order_section, text="Order", style="Card.TLabel").pack(anchor="w")
        ttk.Label(
            order_section,
            text="Choose in-order or random (no repeat).",
            style="Card.TLabel",
            foreground="#666",
        ).pack(anchor="w", pady=(0, 4))
        order_row = ttk.Frame(order_section, style="Card.TFrame")
        order_row.pack(anchor="w")
        self.order_btn = ttk.Button(order_row, text="⇢", command=lambda: self.set_order_mode("order"))
        self.order_btn.pack(side=tk.LEFT, padx=4)
        self.order_tip = Tooltip(self.order_btn, "In order (from current)")
        self.order_btn_rand = ttk.Button(order_row, text="🔀", command=lambda: self.set_order_mode("random_no_repeat"))
        self.order_btn_rand.pack(side=tk.LEFT, padx=4)
        self.order_tip_rand = Tooltip(self.order_btn_rand, "Random (no repeat)")
        self.stop_at_end_check = ttk.Checkbutton(
            order_section,
            text="Stop after list (no repeat list)",
            variable=self.stop_at_end_var,
            command=self.on_stop_at_end_toggle,
        )
        self.stop_at_end_check.pack(anchor="w", pady=(6, 0))
        order_sep = ttk.Separator(right_panel, orient="horizontal")
        sections.append({"key": "order", "frame": order_section, "sep": order_sep})

        # Speed section
        speed_section = ttk.Frame(right_panel, style="Card.TFrame")
        ttk.Label(speed_section, text="Speed", style="Card.TLabel").pack(anchor="w")
        ttk.Label(
            speed_section,
            text="Select interval in seconds or set a custom value.",
            style="Card.TLabel",
            foreground="#666",
        ).pack(anchor="w", pady=(0, 4))
        self.speed_buttons = []
        speed_row = ttk.Frame(speed_section, style="Card.TFrame")
        speed_row.pack(anchor="w")
        for v in [1, 2, 3, 5, 10]:
            btn = ttk.Button(speed_row, text=f"{v}x", command=lambda val=v: self.set_interval(val))
            btn.pack(side=tk.LEFT, padx=3)
            self.speed_buttons.append((v, btn))

        custom_row = ttk.Frame(speed_section, style="Card.TFrame")
        custom_row.pack(anchor="w", pady=(4, 0))
        ttk.Label(custom_row, text="Custom (s):", style="Card.TLabel").pack(side=tk.LEFT)
        self.custom_interval = ttk.Entry(custom_row, width=6)
        self.custom_interval.pack(side=tk.LEFT, padx=4)
        self.custom_interval.bind("<Return>", lambda _e: self.apply_custom_interval())
        ttk.Button(custom_row, text="Apply", command=self.apply_custom_interval).pack(side=tk.LEFT)
        speed_sep = ttk.Separator(right_panel, orient="horizontal")
        sections.append({"key": "speed", "frame": speed_section, "sep": speed_sep})

        # Volume section
        volume_section = ttk.Frame(right_panel, style="Card.TFrame")
        ttk.Label(volume_section, text="Volume", style="Card.TLabel").pack(anchor="w")
        ttk.Label(
            volume_section,
            text="Adjust output volume for playback.",
            style="Card.TLabel",
            foreground="#666",
        ).pack(anchor="w", pady=(0, 4))
        self.volume_scale = tk.Scale(
            volume_section,
            from_=0,
            to=100,
            variable=self.volume_var,
            orient="horizontal",
            length=220,
            showvalue=0,
            troughcolor="#a9c1ff",
            activebackground="#a9c1ff",
            highlightthickness=0,
        )
        self.volume_scale.pack(anchor="w")
        self.volume_value = ttk.Label(volume_section, text=f"{int(self.volume_var.get())}%", style="Card.TLabel")
        self.volume_value.pack(anchor="w", pady=(2, 0))
        self.volume_scale.config(command=self.on_volume_change)
        volume_sep = ttk.Separator(right_panel, orient="horizontal")
        sections.append({"key": "volume", "frame": volume_section, "sep": volume_sep})

        rebuild_sections()

        self.update_order_button()
        self.update_loop_button()
        self.update_speed_buttons()
        self.on_volume_change()

    def apply_custom_interval(self):
        try:
            val = int(self.custom_interval.get())
            if val < 1:
                raise ValueError
        except Exception:
            messagebox.showinfo("Info", "Please enter a valid number (>=1).")
            return
        self.set_interval(val)

    def build_queue(self):
        indices = list(range(len(self.store.words)))
        if self.order_mode.get() == "random_no_repeat":
            random.shuffle(indices)
        return indices

    def rebuild_queue_on_mode_change(self):
        if not self.store.words:
            self.queue = []
            self.pos = -1
            return
        if self.current_word is None:
            self.queue = self.build_queue()
            self.pos = 0
            self.set_current_word()
            return

        current_index = self.store.words.index(self.current_word)
        if self.order_mode.get() == "random_no_repeat":
            # reshuffle all words (no repeat) and restart from the new order
            self.queue = self.build_queue()
            self.pos = 0
        else:
            # ordered from current
            n = len(self.store.words)
            if self.loop_var.get():
                self.queue = list(range(current_index, n)) + list(range(0, current_index))
            else:
                self.queue = list(range(current_index, n))
            self.pos = 0
        self.set_current_word()
        if self.play_state == "playing":
            self.play_current()
            self.schedule_next()

    def toggle_play(self):
        if not self.store.words:
            messagebox.showinfo("Info", "Please import words first.")
            return
        if self.play_state == "playing":
            self.play_state = "paused"
            self.cancel_schedule()
            tts_cancel_all()
            self.play_token += 1
            self.status_var.set("Paused")
            self.update_play_button()
            return

        # start or resume
        if not self.queue or self.pos < 0:
            self.build_queue_from_selection()

        self.play_state = "playing"
        self.play_token += 1
        self.update_play_button()
        self.play_current()
        self.schedule_next()

    def play_current(self):
        if not self.current_word:
            return
        speak_async(self.current_word, self.volume_var.get() / 100.0, cancel_before=True)

    def schedule_next(self):
        self.cancel_schedule()
        if self.play_state != "playing":
            return
        interval = max(1, int(self.interval_var.get()))
        token = self.play_token
        self.after_id = self.after(interval * 1000, lambda: self.next_word(token))

    def next_word(self, token):
        if self.play_state != "playing" or token != self.play_token:
            return
        if not self.queue:
            self.queue = self.build_queue()
            self.pos = 0
        else:
            self.pos += 1
            if self.pos >= len(self.queue):
                if self.loop_var.get():
                    self.queue = self.build_queue()
                    self.pos = 0
                else:
                    self.play_state = "stopped"
                    self.cancel_schedule()
                    self.status_var.set("Completed")
                    self.update_play_button()
                    return
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
            self.play_btn.config(text="⏸")
            if self.play_btn_check:
                self.play_btn_check.config(text="⏸")
        else:
            self.play_btn.config(text="▶")
            if self.play_btn_check:
                self.play_btn_check.config(text="▶")

    def reset_playback_state(self):
        self.cancel_schedule()
        tts_cancel_all()
        self.play_token += 1
        self.play_state = "stopped"
        self.queue = []
        self.pos = -1
        self.current_word = None
        self.status_var.set("Not started")
        self.update_play_button()

    def update_right_visibility(self):
        if self.history_visible or self.check_visible:
            self.right.grid()
            if not self.wordlist_hidden:
                self.mid_sep.grid()
            else:
                self.mid_sep.grid_remove()
        else:
            self.right.grid_remove()
            self.mid_sep.grid_remove()

    def build_queue_from_selection(self):
        # Nothing to build if the word list is empty.
        if not self.store.words:
            self.queue = []
            self.pos = -1
            self.current_word = None
            return
        selection = self.listbox.curselection()
        start_idx = selection[0] if selection else 0

        if self.order_mode.get() == "random_no_repeat":
            indices = list(range(len(self.store.words)))
            if start_idx in indices:
                indices.remove(start_idx)
            random.shuffle(indices)
            self.queue = [start_idx] + indices if self.store.words else []
            self.pos = 0
        else:
            n = len(self.store.words)
            if self.loop_var.get():
                self.queue = list(range(start_idx, n)) + list(range(0, start_idx))
            else:
                self.queue = list(range(start_idx, n))
            self.pos = 0
        self.set_current_word()

    def on_word_selected(self, _event=None):
        # When the user clicks a word, restart playback from that word.
        if not self.store.words:
            return
        if not self.listbox.curselection():
            return
        tts_cancel_all()
        self.play_token += 1
        self.build_queue_from_selection()
        if self.play_state == "playing":
            self.update_play_button()
            self.play_current()
            self.schedule_next()

    def toggle_wordlist_visibility(self):
        # Hide or show the word list during dictation to avoid seeing words.
        self.wordlist_hidden = not self.wordlist_hidden
        if self.wordlist_hidden:
            self.left.grid_remove()
            self.mid_sep.grid_remove()
            self.player_frame.grid_remove()
            self.status_label.grid_remove()
            self.hide_words_btn.config(text="Show Word List")
        else:
            self.left.grid()
            self.player_frame.grid()
            self.status_label.grid()
            self.hide_words_btn.config(text="Hide Word List")
        self.update_right_visibility()

    # Input check
    def on_check_enter(self, _event=None):
        # Allow quick submit with Enter.
        self.check_input()
        return "break"

    def on_input_space(self, _event=None):
        # Clear the input quickly when user presses Space.
        self.input_entry.delete(0, tk.END)
        return "break"

    def check_input(self):
        if not self.current_word:
            messagebox.showinfo("Info", "No current word yet.")
            return
        user_text = self.input_entry.get()
        apply_diff(self.result_text, self.current_word, user_text)
