# -*- coding: utf-8 -*-
import os
import queue
import random
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from data.store import WordStore
from services.tts import (
    speak_async,
    speak_stream_async,
    cancel_all as tts_cancel_all,
    prepare_async as tts_prepare_async,
)
from services.translation import (
    prepare_async as translation_prepare_async,
    translate_words as translate_words_en_zh,
)
from services.ielts_passage import build_ielts_listening_passage
from services.app_config import (
    get_gemini_api_key,
    get_generation_model,
    set_gemini_api_key,
    set_generation_model,
)
from services.gemini_writer import (
    DEFAULT_GEMINI_MODEL,
    choose_preferred_generation_model,
    generate_example_sentence_with_gemini,
    generate_english_passage_with_gemini,
    list_available_gemini_models,
    validate_gemini_api_key,
)
from services.voice_catalog import list_system_voices
from services.voice_manager import (
    SOURCE_KOKORO,
    get_voice_id,
    get_voice_source,
    set_voice_source,
)
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

        self.order_mode = tk.StringVar(value="order")  # order | random_no_repeat | click_to_play
        self.interval_var = tk.DoubleVar(value=2.0)
        self.volume_var = tk.IntVar(value=80)
        self.speech_rate_var = tk.DoubleVar(value=1.0)
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
        self.order_btn_click = None
        self.order_tip = None
        self.order_tip_rand = None
        self.order_tip_click = None
        self.loop_btn = None
        self.loop_btn_stop = None
        self.stop_at_end_check = None
        self.speed_buttons = []
        self.speech_rate_buttons = []
        self.volume_scale = None
        self.player_frame = None
        self.play_btn_check = None
        self.settings_btn_check = None
        self.check_btn_toggle_check = None
        self.hist_btn_toggle_check = None
        self.passage_btn_check = None
        self.check_controls = None
        self.hide_words_btn = None
        self.voice_var = tk.StringVar(value="")
        self.voice_combo = None
        self.voice_map = {}
        self.word_table = None
        self.word_table_scroll = None
        self.word_context_menu = None
        self.word_edit_entry = None
        self.word_edit_row = None
        self.suppress_word_select_action = False
        self.sentence_window = None
        self.manual_words_window = None
        self.manual_words_text = None
        self.sentence_generation_token = 0
        self.sentence_generation_active_token = 0
        self.sentence_event_queue = queue.Queue()
        self.translations = {}
        self.translation_token = 0
        self.passage_window = None
        self.passage_text = None
        self.passage_status_var = tk.StringVar(value="Load words and click Generate.")
        self.gemini_model_var = tk.StringVar(value=get_generation_model() or DEFAULT_GEMINI_MODEL)
        self.gemini_model_combo = None
        self.gemini_model_values = []
        self.gemini_verified = False
        self.gemini_key_window = None
        self.gemini_key_var = tk.StringVar(value=get_gemini_api_key())
        self.gemini_key_status_var = tk.StringVar(value="Paste your Gemini API key, then test it.")
        self.gemini_key_test_btn = None
        self.gemini_validation_token = 0
        self.gemini_validation_active_token = 0
        self.gemini_validation_queue = queue.Queue()
        self.current_passage = ""
        self.current_passage_original = ""
        self.current_passage_words = []
        self.passage_cloze_text = ""
        self.passage_answers = []
        self.passage_is_practice = False
        self.passage_practice_input = None
        self.passage_practice_result = None
        self.passage_generation_token = 0
        self.passage_generation_active_token = 0
        self.passage_event_queue = queue.Queue()

        self.build_ui()
        self.refresh_history()
        self.update_empty_state()
        self.update_order_button()
        self.update_loop_button()
        self.update_speed_buttons()
        self.update_speech_rate_buttons()
        self.update_play_button()
        self.update_right_visibility()
        tts_prepare_async()
        translation_prepare_async()
        self.refresh_gemini_models()
        self.after(150, self.ensure_gemini_api_key)

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

        top_btn_row = ttk.Frame(self.left, style="Card.TFrame")
        top_btn_row.grid(row=1, column=0, columnspan=2, padx=12, pady=6, sticky="ew")
        top_btn_row.grid_columnconfigure(0, weight=1)
        top_btn_row.grid_columnconfigure(1, weight=1)
        top_btn_row.grid_columnconfigure(2, weight=1)

        btn_load = ttk.Button(top_btn_row, text="⭳", style="Primary.TButton", command=self.load_words)
        btn_load.grid(row=0, column=0, padx=(0, 6), sticky="ew")
        Tooltip(btn_load, "Import Words")

        btn_manual = ttk.Button(top_btn_row, text="✍", command=self.open_manual_words_window)
        btn_manual.grid(row=0, column=1, padx=3, sticky="ew")
        Tooltip(btn_manual, "Type/Paste Words")

        btn_speak = ttk.Button(top_btn_row, text="🔊", command=self.speak_selected)
        btn_speak.grid(row=0, column=2, padx=(6, 0), sticky="ew")
        Tooltip(btn_speak, "Speak Selected")

        table_wrap = ttk.Frame(self.left, style="Card.TFrame")
        table_wrap.grid(row=2, column=0, columnspan=2, padx=12, pady=(4, 6), sticky="nsew")
        self.left.grid_rowconfigure(2, weight=1)
        self.left.grid_columnconfigure(0, weight=1)
        self.left.grid_columnconfigure(1, weight=1)

        self.word_table = ttk.Treeview(
            table_wrap,
            columns=("en", "zh"),
            show="headings",
            height=18,
            selectmode="extended",
        )
        self.word_table.heading("en", text="English")
        self.word_table.heading("zh", text="中文")
        self.word_table.column("en", width=170, anchor="w")
        self.word_table.column("zh", width=170, anchor="w")

        self.word_table_scroll = ttk.Scrollbar(table_wrap, orient="vertical", command=self.word_table.yview)
        self.word_table.configure(yscrollcommand=self.word_table_scroll.set)
        self.word_table.grid(row=0, column=0, sticky="nsew")
        self.word_table_scroll.grid(row=0, column=1, sticky="ns")
        table_wrap.grid_rowconfigure(0, weight=1)
        table_wrap.grid_columnconfigure(0, weight=1)
        self.word_table.bind("<<TreeviewSelect>>", self.on_word_selected)
        self.word_table.bind("<Double-1>", self.on_word_double_click)
        self.word_table.bind("<Button-3>", self.on_word_right_click)
        self.word_table.bind("<F2>", self.start_edit_selected_word)
        self.word_table.bind("<Escape>", self.cancel_word_edit)
        self.word_table.bind("<Control-v>", self.on_word_table_paste)
        self.word_table.bind("<Control-V>", self.on_word_table_paste)

        self.word_context_menu = tk.Menu(self, tearoff=0)
        self.word_context_menu.add_command(label="编辑", command=self.start_edit_selected_word)
        self.word_context_menu.add_command(label="造句", command=self.make_sentence_for_selected_word)

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
        self.player_frame.grid_columnconfigure(4, weight=1)

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

        self.passage_btn = ttk.Button(
            self.player_frame, text="📝", style="Icon.TButton", width=4, command=self.open_passage_window
        )
        self.passage_btn.grid(row=0, column=4, padx=4, sticky="ew")
        Tooltip(self.passage_btn, "Generate IELTS Passage")

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

        self.passage_btn_check = ttk.Button(
            self.check_controls, text="📝", style="Icon.TButton", width=4, command=self.open_passage_window
        )
        self.passage_btn_check.pack(side=tk.LEFT, padx=3)
        Tooltip(self.passage_btn_check, "Generate IELTS Passage")

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
        self.cancel_word_edit()
        words = self.store.load_from_file(path)
        self.render_words(words)
        self.refresh_history()
        self.reset_playback_state()
        self.status_var.set(f"Loaded {len(words)} words from file.")

    def _parse_manual_words(self, raw_text):
        text = str(raw_text or "")
        lines = text.splitlines()
        words = []
        for line in lines:
            line = line.strip()
            if not line:
                continue
            parts = re.split(r"[,\t;；，]+", line)
            if len(parts) <= 1:
                words.append(line)
                continue
            for part in parts:
                token = part.strip()
                if token:
                    words.append(token)
        return words

    def _update_word_stats_from_manual_input(self, words):
        try:
            stats = self.store.load_stats()
            for word in words:
                stats[word] = int(stats.get(word, 0)) + 1
            self.store.save_stats(stats)
        except Exception:
            return

    def _apply_manual_words(self, words, mode="replace"):
        normalized = [str(w).strip() for w in words if str(w).strip()]
        if not normalized:
            messagebox.showinfo("Info", "No valid words found.")
            return False
        self.cancel_word_edit()

        if mode == "append":
            merged = list(self.store.words) + normalized
            self.store.set_words(merged)
            self.render_words(merged)
            self._update_word_stats_from_manual_input(normalized)
            status_text = f"Appended {len(normalized)} words."
        else:
            self.store.set_words(normalized)
            self.render_words(normalized)
            self._update_word_stats_from_manual_input(normalized)
            status_text = f"Loaded {len(normalized)} words."
        self.reset_playback_state()
        self.status_var.set(status_text)
        return True

    def _apply_manual_words_from_editor(self, mode):
        if not self.manual_words_text:
            return
        raw = self.manual_words_text.get("1.0", tk.END)
        words = self._parse_manual_words(raw)
        ok = self._apply_manual_words(words, mode=mode)
        if ok and self.manual_words_window and self.manual_words_window.winfo_exists():
            self.manual_words_window.destroy()
            self.manual_words_window = None
            self.manual_words_text = None

    def open_manual_words_window(self):
        if self.manual_words_window and self.manual_words_window.winfo_exists():
            self.manual_words_window.lift()
            return

        self.manual_words_window = tk.Toplevel(self)
        self.manual_words_window.title("Manual Words")
        self.manual_words_window.configure(bg="#f6f7fb")
        self.manual_words_window.resizable(False, False)

        wrap = ttk.Frame(self.manual_words_window, style="Card.TFrame")
        wrap.pack(fill="both", expand=True, padx=10, pady=10)
        ttk.Label(wrap, text="Type or paste words", style="Card.TLabel").pack(anchor="w")
        ttk.Label(
            wrap,
            text="One per line, or split by comma/semicolon/tab.",
            style="Card.TLabel",
            foreground="#666",
        ).pack(anchor="w", pady=(0, 6))

        self.manual_words_text = tk.Text(
            wrap,
            wrap="word",
            height=10,
            width=48,
            bg="#ffffff",
            fg="#222222",
            highlightthickness=1,
            highlightbackground="#d9dbe1",
        )
        self.manual_words_text.pack(fill="x")
        self.manual_words_text.focus_set()

        def _on_close():
            self.manual_words_window.destroy()
            self.manual_words_window = None
            self.manual_words_text = None

        row = ttk.Frame(wrap, style="Card.TFrame")
        row.pack(fill="x", pady=(8, 0))
        ttk.Button(
            row,
            text="Replace List",
            command=lambda: self._apply_manual_words_from_editor("replace"),
        ).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(
            row,
            text="Append",
            command=lambda: self._apply_manual_words_from_editor("append"),
        ).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(
            row,
            text="Close",
            command=_on_close,
        ).pack(side=tk.LEFT)

        self.manual_words_window.protocol("WM_DELETE_WINDOW", _on_close)

    def on_word_table_paste(self, _event=None):
        try:
            raw = self.clipboard_get()
        except Exception:
            return "break"
        words = self._parse_manual_words(raw)
        if not words:
            return "break"
        self._apply_manual_words(words, mode="append")
        return "break"

    def _get_selected_indices(self):
        if not self.word_table:
            return []
        sel = self.word_table.selection()
        if not sel:
            return []
        values = []
        for item in sel:
            try:
                values.append(int(item))
            except Exception:
                continue
        return sorted(set(values))

    def _get_selected_index(self):
        if not self.word_table:
            return None
        focus = str(self.word_table.focus() or "").strip()
        if focus:
            try:
                focus_idx = int(focus)
                if focus in self.word_table.selection():
                    return focus_idx
            except Exception:
                pass
        indices = self._get_selected_indices()
        if not indices:
            return None
        return indices[0]

    def _get_selected_words_for_passage(self):
        selected = []
        for idx in self._get_selected_indices():
            if 0 <= idx < len(self.store.words):
                selected.append(self.store.words[idx])
        if selected:
            return selected
        return list(self.store.words)

    def _save_words_back_to_source(self):
        if not self.store.has_current_source_file():
            return False
        try:
            self.store.save_to_current_file()
            return True
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save source file.\n{e}")
            return False

    def _translate_single_word_async(self, row_idx, word):
        token = self.translation_token

        def _run():
            translated = translate_words_en_zh([word])
            self.after(0, lambda: self._apply_single_translation(token, row_idx, word, translated.get(word) or ""))

        import threading

        threading.Thread(target=_run, daemon=True).start()

    def _apply_single_translation(self, token, row_idx, word, zh_text):
        if token != self.translation_token or not self.word_table:
            return
        if row_idx < 0 or row_idx >= len(self.store.words):
            return
        if self.store.words[row_idx] != word:
            return
        self.translations[word] = zh_text
        iid = str(row_idx)
        if self.word_table.exists(iid):
            self.word_table.item(iid, values=(word, zh_text))

    def start_edit_selected_word(self, _event=None):
        row_idx = self._get_selected_index()
        if row_idx is None or not self.word_table:
            return "break"
        if row_idx >= len(self.store.words):
            return "break"
        self.cancel_word_edit()
        iid = str(row_idx)
        bbox = self.word_table.bbox(iid, "#1")
        if not bbox:
            return "break"

        x, y, width, height = bbox
        current_word = self.store.words[row_idx]
        entry = ttk.Entry(self.word_table)
        entry.insert(0, current_word)
        entry.select_range(0, tk.END)
        entry.focus_set()
        entry.place(x=x, y=y, width=width, height=height)
        entry.bind("<Return>", lambda _e, idx=row_idx: self.finish_edit_word(idx))
        entry.bind("<Escape>", self.cancel_word_edit)
        entry.bind("<FocusOut>", lambda _e, idx=row_idx: self.finish_edit_word(idx))
        self.word_edit_entry = entry
        self.word_edit_row = row_idx
        return "break"

    def cancel_word_edit(self, _event=None):
        if self.word_edit_entry and self.word_edit_entry.winfo_exists():
            self.word_edit_entry.destroy()
        self.word_edit_entry = None
        self.word_edit_row = None
        return "break"

    def finish_edit_word(self, row_idx=None):
        if not self.word_edit_entry:
            return "break"
        idx = self.word_edit_row if row_idx is None else row_idx
        new_word = re.sub(r"\s+", " ", str(self.word_edit_entry.get() or "").strip())
        self.cancel_word_edit()
        if idx is None or idx < 0 or idx >= len(self.store.words):
            return "break"
        if not new_word:
            messagebox.showinfo("Info", "Word cannot be empty.")
            return "break"

        old_word = self.store.words[idx]
        if new_word == old_word:
            return "break"

        self.store.words[idx] = new_word
        iid = str(idx)
        if self.word_table and self.word_table.exists(iid):
            self.word_table.item(iid, values=(new_word, "Translating..."))

        saved = self._save_words_back_to_source()
        self._translate_single_word_async(idx, new_word)
        source_note = " and saved to source file" if saved else ""
        self.status_var.set(f"Updated '{old_word}' to '{new_word}'{source_note}.")
        return "break"

    def render_words(self, words):
        if not self.word_table:
            return
        self.cancel_word_edit()
        self.translation_token += 1
        token = self.translation_token
        self.translations = {}
        self.word_table.delete(*self.word_table.get_children())
        for idx, w in enumerate(words):
            self.word_table.insert("", tk.END, iid=str(idx), values=(w, "Translating..."))
        self.update_empty_state()
        self._start_translation_job(words, token)

    def _start_translation_job(self, words, token):
        if not words:
            return

        def _run():
            translated = translate_words_en_zh(words)
            self.after(0, lambda: self._apply_translations(token, translated))

        import threading

        threading.Thread(target=_run, daemon=True).start()

    def _apply_translations(self, token, translated):
        if token != self.translation_token or not self.word_table:
            return
        self.translations = dict(translated)
        for idx, word in enumerate(self.store.words):
            if not self.word_table.exists(str(idx)):
                continue
            zh = self.translations.get(word) or ""
            self.word_table.item(str(idx), values=(word, zh))

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
        self.cancel_word_edit()
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

    # IELTS passage
    def open_passage_window(self):
        if self.passage_window and self.passage_window.winfo_exists():
            self.passage_window.lift()
            return

        self.passage_window = tk.Toplevel(self)
        self.passage_window.title("IELTS Passage Builder")
        self.passage_window.configure(bg="#f6f7fb")
        self.passage_window.resizable(True, True)
        self.passage_window.minsize(680, 460)

        wrap = ttk.Frame(self.passage_window, style="Card.TFrame")
        wrap.pack(fill="both", expand=True, padx=10, pady=10)

        ttk.Label(wrap, text="IELTS Listening Style Passage", style="Card.TLabel").pack(anchor="w")
        ttk.Label(
            wrap,
            text="Generate with Gemini API and read it with Kokoro. Select words in the main table first if you only want part of the list.",
            style="Card.TLabel",
            foreground="#666",
        ).pack(anchor="w", pady=(0, 8))

        ctrl = ttk.Frame(wrap, style="Card.TFrame")
        ctrl.pack(fill="x", pady=(0, 8))

        btn_generate = ttk.Button(ctrl, text="Generate", command=self.generate_ielts_passage)
        btn_generate.pack(side=tk.LEFT, padx=(0, 6))
        Tooltip(btn_generate, "Build passage from selected words, or the full list if nothing is selected")

        btn_play = ttk.Button(ctrl, text="Read with Kokoro", command=self.play_generated_passage)
        btn_play.pack(side=tk.LEFT, padx=(0, 6))
        Tooltip(btn_play, "Speak current passage")

        btn_stop = ttk.Button(ctrl, text="Stop", command=self.stop_passage_playback)
        btn_stop.pack(side=tk.LEFT, padx=(0, 6))
        Tooltip(btn_stop, "Stop speaking")

        btn_practice = ttk.Button(ctrl, text="Practice", command=self.start_passage_practice)
        btn_practice.pack(side=tk.LEFT, padx=(0, 6))
        Tooltip(btn_practice, "Hide keywords as blanks")

        btn_check_practice = ttk.Button(ctrl, text="Check", command=self.check_passage_practice)
        btn_check_practice.pack(side=tk.LEFT)
        Tooltip(btn_check_practice, "Check your filled answers")

        model_wrap = ttk.Frame(ctrl, style="Card.TFrame")
        model_wrap.pack(side=tk.RIGHT)
        ttk.Label(model_wrap, text="Model:", style="Card.TLabel").pack(side=tk.LEFT, padx=(0, 4))
        self.gemini_model_combo = ttk.Combobox(
            model_wrap,
            textvariable=self.gemini_model_var,
            state="readonly",
            width=20,
        )
        self.gemini_model_combo.pack(side=tk.LEFT)
        self.gemini_model_combo.bind("<<ComboboxSelected>>", self.on_gemini_model_change)

        self.passage_text = tk.Text(
            wrap,
            wrap="word",
            height=18,
            bg="#ffffff",
            fg="#222222",
            highlightthickness=1,
            highlightbackground="#d9dbe1",
        )
        self.passage_text.pack(fill="both", expand=True)

        practice_wrap = ttk.Frame(wrap, style="Card.TFrame")
        practice_wrap.pack(fill="x", pady=(8, 0))
        ttk.Label(
            practice_wrap,
            text="Practice: fill one missing word/phrase per line (in order).",
            style="Card.TLabel",
            foreground="#666",
        ).pack(anchor="w")

        self.passage_practice_input = tk.Text(
            practice_wrap,
            wrap="word",
            height=4,
            bg="#ffffff",
            fg="#222222",
            highlightthickness=1,
            highlightbackground="#d9dbe1",
        )
        self.passage_practice_input.pack(fill="x", pady=(4, 6))

        self.passage_practice_result = tk.Text(
            practice_wrap,
            wrap="word",
            height=5,
            bg="#ffffff",
            fg="#222222",
            highlightthickness=1,
            highlightbackground="#d9dbe1",
        )
        self.passage_practice_result.pack(fill="x")
        self.passage_practice_result.tag_configure("wrong", foreground="#c62828")
        self.passage_practice_result.tag_configure("missing", foreground="#c62828", underline=1)
        self.passage_practice_result.tag_configure("extra", foreground="#ef6c00")
        self.passage_practice_result.config(state="disabled")

        status = ttk.Label(wrap, textvariable=self.passage_status_var, style="Card.TLabel", foreground="#444")
        status.pack(anchor="w", pady=(8, 0))

        if self.current_passage:
            self._set_passage_text(self.current_passage)
        else:
            self._set_passage_text("")
        self.refresh_gemini_models()

        def _on_close():
            self.passage_window.destroy()
            self.passage_window = None
            self.passage_text = None
            self.passage_practice_input = None
            self.passage_practice_result = None

        self.passage_window.protocol("WM_DELETE_WINDOW", _on_close)

    def _set_passage_text(self, text):
        if not self.passage_text:
            return
        self.passage_text.delete("1.0", tk.END)
        self.passage_text.insert("1.0", text or "")

    def _get_passage_text(self):
        if not self.passage_text:
            return self.current_passage
        return self.passage_text.get("1.0", tk.END).strip()

    def _speech_text_from_passage(self, text):
        lines = [line.strip() for line in str(text or "").splitlines()]
        if lines and lines[0].lower().startswith("ielts listening practice -"):
            lines = lines[1:]
        compact = "\n".join(line for line in lines if line)
        return compact.strip()

    def _normalize_answer(self, text):
        return re.sub(r"\s+", " ", str(text or "").strip().casefold())

    def _replace_first_case_insensitive(self, text, target, repl):
        if not target:
            return text, False
        pattern = re.compile(re.escape(str(target)), re.IGNORECASE)
        match = pattern.search(text)
        if not match:
            return text, False
        return text[: match.start()] + repl + text[match.end() :], True

    def _build_cloze_passage(self, passage, keywords, max_blanks=12):
        text = str(passage or "").strip()
        if not text:
            return "", []
        cloze = text
        answers = []
        for word in keywords:
            if len(answers) >= max_blanks:
                break
            key = str(word or "").strip()
            if not key:
                continue
            blank = "____"
            cloze, replaced = self._replace_first_case_insensitive(cloze, key, blank)
            if replaced:
                answers.append(key)
        return cloze, answers

    def _clear_passage_practice_result(self):
        if not self.passage_practice_result:
            return
        self.passage_practice_result.config(state="normal")
        self.passage_practice_result.delete("1.0", tk.END)
        self.passage_practice_result.config(state="disabled")

    def _clear_passage_practice_input(self):
        if self.passage_practice_input:
            self.passage_practice_input.delete("1.0", tk.END)

    def start_passage_practice(self):
        source_text = self.current_passage_original.strip() or self.current_passage.strip()
        if not source_text:
            messagebox.showinfo("Info", "Generate a passage first.")
            return

        keywords = list(self.current_passage_words) if self.current_passage_words else list(self.store.words)
        cloze, answers = self._build_cloze_passage(source_text, keywords, max_blanks=12)
        if not answers:
            messagebox.showinfo("Info", "No suitable keywords found in this passage for practice.")
            return

        self.passage_is_practice = True
        self.passage_cloze_text = cloze
        self.passage_answers = answers
        self._set_passage_text(cloze)
        self._clear_passage_practice_input()
        self._clear_passage_practice_result()
        self.passage_status_var.set(
            f"Practice ready: {len(answers)} blanks. Fill one answer per line, then click Check."
        )

    def check_passage_practice(self):
        if not self.passage_answers:
            messagebox.showinfo("Info", "Click Practice first.")
            return
        if not self.passage_practice_input:
            return

        lines = [
            line.strip()
            for line in self.passage_practice_input.get("1.0", tk.END).splitlines()
            if line.strip()
        ]
        expected = "\n".join(self.passage_answers)
        actual = "\n".join(lines)
        if self.passage_practice_result:
            apply_diff(self.passage_practice_result, expected, actual)

        correct = 0
        for idx, answer in enumerate(self.passage_answers):
            user_value = lines[idx] if idx < len(lines) else ""
            if self._normalize_answer(user_value) == self._normalize_answer(answer):
                correct += 1
        self.passage_status_var.set(
            f"Practice check: {correct}/{len(self.passage_answers)} correct."
        )

    def on_gemini_model_change(self, _event=None):
        set_generation_model(self._get_selected_gemini_model())

    def _get_selected_gemini_model(self):
        value = str(self.gemini_model_var.get() or "").strip()
        return value or DEFAULT_GEMINI_MODEL

    def refresh_gemini_models(self):
        models = list_available_gemini_models()
        if not models:
            models = [DEFAULT_GEMINI_MODEL]
        self.gemini_model_values = list(models)
        if self.gemini_model_combo:
            self.gemini_model_combo["values"] = self.gemini_model_values
        current = self._get_selected_gemini_model()
        if current not in self.gemini_model_values:
            current = choose_preferred_generation_model(self.gemini_model_values, fallback=DEFAULT_GEMINI_MODEL)
        self.gemini_model_var.set(current)
        set_generation_model(current)

    def ensure_gemini_api_key(self):
        self.open_gemini_key_window(force_verify=True)

    def _clear_gemini_validation_queue(self):
        try:
            while True:
                self.gemini_validation_queue.get_nowait()
        except queue.Empty:
            return

    def _emit_gemini_validation_event(self, event_type, token, payload=None):
        try:
            self.gemini_validation_queue.put_nowait((event_type, token, payload))
        except Exception:
            return

    def _poll_gemini_validation_events(self, token):
        if token != self.gemini_validation_active_token:
            return
        done = False
        try:
            while True:
                event_type, event_token, payload = self.gemini_validation_queue.get_nowait()
                if event_token != token:
                    continue
                if event_type == "success":
                    self._finish_gemini_validation_success(payload or {})
                elif event_type == "error":
                    self._finish_gemini_validation_error(str(payload or "Unknown error"))
                elif event_type == "done":
                    done = True
        except queue.Empty:
            pass
        if not done and token == self.gemini_validation_active_token:
            self.after(80, lambda t=token: self._poll_gemini_validation_events(t))

    def open_gemini_key_window(self, force_verify=False):
        self.gemini_verified = self.gemini_verified and not force_verify
        if self.gemini_key_window and self.gemini_key_window.winfo_exists():
            self.gemini_key_window.lift()
            self.gemini_key_window.focus_force()
            return

        self.gemini_key_status_var.set("Paste your Gemini API key, then test it.")
        self.gemini_key_var.set(get_gemini_api_key())
        win = tk.Toplevel(self)
        self.gemini_key_window = win
        win.title("Gemini API Key")
        win.configure(bg="#f6f7fb")
        win.resizable(False, False)
        win.transient(self.winfo_toplevel())

        wrap = ttk.Frame(win, style="Card.TFrame")
        wrap.pack(fill="both", expand=True, padx=12, pady=12)

        ttk.Label(wrap, text="Gemini API setup", style="Card.TLabel").pack(anchor="w")
        ttk.Label(
            wrap,
            text="Paste your own Gemini API key. The app will test it before enabling AI features.",
            style="Card.TLabel",
            foreground="#666",
        ).pack(anchor="w", pady=(0, 8))

        entry = ttk.Entry(wrap, textvariable=self.gemini_key_var, width=54, show="*")
        entry.pack(fill="x")
        entry.focus_set()
        entry.icursor(tk.END)
        entry.bind("<Return>", lambda _event: self.test_and_save_gemini_key())

        ttk.Label(
            wrap,
            text="Model used for article generation and sentence generation:",
            style="Card.TLabel",
            foreground="#666",
        ).pack(anchor="w", pady=(8, 4))
        combo = ttk.Combobox(
            wrap,
            textvariable=self.gemini_model_var,
            values=self.gemini_model_values or list_available_gemini_models(),
            state="readonly",
            width=24,
        )
        combo.pack(anchor="w")
        combo.bind("<<ComboboxSelected>>", self.on_gemini_model_change)

        ttk.Label(wrap, textvariable=self.gemini_key_status_var, style="Card.TLabel", foreground="#444").pack(
            anchor="w", pady=(10, 0)
        )

        btn_row = ttk.Frame(wrap, style="Card.TFrame")
        btn_row.pack(fill="x", pady=(10, 0))
        self.gemini_key_test_btn = ttk.Button(btn_row, text="Test and Save", command=self.test_and_save_gemini_key)
        self.gemini_key_test_btn.pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="Exit", command=self._cancel_gemini_key_setup).pack(side=tk.LEFT)

        win.grab_set()
        win.protocol("WM_DELETE_WINDOW", self._cancel_gemini_key_setup)

    def _cancel_gemini_key_setup(self):
        if self.gemini_key_window and self.gemini_key_window.winfo_exists():
            try:
                self.gemini_key_window.grab_release()
            except Exception:
                pass
            self.gemini_key_window.destroy()
        self.gemini_key_window = None
        self.gemini_key_test_btn = None
        if not self.gemini_verified:
            self.winfo_toplevel().destroy()

    def test_and_save_gemini_key(self):
        api_key = str(self.gemini_key_var.get() or "").strip()
        model_name = self._get_selected_gemini_model()
        if not api_key:
            messagebox.showinfo("Info", "Please paste a Gemini API key first.")
            return

        self.gemini_key_status_var.set(f"Testing Gemini key with {model_name}...")
        if self.gemini_key_test_btn:
            self.gemini_key_test_btn.config(state="disabled")

        self.gemini_validation_token += 1
        token = self.gemini_validation_token
        self.gemini_validation_active_token = token
        self._clear_gemini_validation_queue()

        import threading

        def _run():
            try:
                validate_gemini_api_key(api_key, model=model_name, timeout=25)
                self._emit_gemini_validation_event(
                    "success",
                    token,
                    {"api_key": api_key, "model": model_name},
                )
            except Exception as e:
                self._emit_gemini_validation_event("error", token, str(e))
            self._emit_gemini_validation_event("done", token, None)

        threading.Thread(target=_run, daemon=True).start()
        self.after(80, lambda t=token: self._poll_gemini_validation_events(t))

    def _finish_gemini_validation_success(self, payload):
        api_key = str(payload.get("api_key") or "").strip()
        model_name = str(payload.get("model") or DEFAULT_GEMINI_MODEL).strip()
        set_gemini_api_key(api_key)
        set_generation_model(model_name)
        self.gemini_verified = True
        self.gemini_key_status_var.set("Gemini key is valid.")
        if self.gemini_key_test_btn:
            self.gemini_key_test_btn.config(state="normal")
        if self.gemini_key_window and self.gemini_key_window.winfo_exists():
            try:
                self.gemini_key_window.grab_release()
            except Exception:
                pass
            self.gemini_key_window.destroy()
        self.gemini_key_window = None
        self.gemini_key_test_btn = None
        self.status_var.set("Gemini ready.")

    def _finish_gemini_validation_error(self, message):
        self.gemini_verified = False
        self.gemini_key_status_var.set("Gemini key test failed. Please paste another key.")
        if self.gemini_key_test_btn:
            self.gemini_key_test_btn.config(state="normal")
        messagebox.showerror("Gemini API Key Error", str(message or "Unknown error"))

    def _require_gemini_ready(self):
        if self.gemini_verified and get_gemini_api_key():
            return True
        self.open_gemini_key_window(force_verify=False)
        return False

    def generate_ielts_passage(self):
        if not self.store.words:
            messagebox.showinfo("Info", "Please import words first.")
            return
        if not self._require_gemini_ready():
            return

        words = self._get_selected_words_for_passage()
        model_name = self._get_selected_gemini_model()
        self.passage_generation_token += 1
        token = self.passage_generation_token
        selected_count = len(self._get_selected_indices())
        source_text = f"{len(words)} selected words" if selected_count else f"{len(words)} words"
        self.passage_status_var.set(f"Generating with Gemini ({model_name}) from {source_text}...")
        self.current_passage = ""
        self.current_passage_original = ""
        self.current_passage_words = []
        self.passage_is_practice = False
        self.passage_cloze_text = ""
        self.passage_answers = []
        self._set_passage_text("")
        self._clear_passage_practice_input()
        self._clear_passage_practice_result()
        self.passage_generation_active_token = token
        self._clear_passage_event_queue()

        import threading

        threading.Thread(
            target=lambda: self._run_passage_generation(token, words, model_name),
            daemon=True,
        ).start()
        self.after(80, lambda t=token: self._poll_passage_generation_events(t))

    def _clear_passage_event_queue(self):
        try:
            while True:
                self.passage_event_queue.get_nowait()
        except queue.Empty:
            return

    def _emit_passage_event(self, event_type, token, payload=None):
        try:
            self.passage_event_queue.put_nowait((event_type, token, payload))
        except Exception:
            return

    def _poll_passage_generation_events(self, token):
        if token != self.passage_generation_active_token:
            return

        done = False
        try:
            while True:
                event_type, event_token, payload = self.passage_event_queue.get_nowait()
                if event_token != token:
                    continue
                if event_type == "partial":
                    self._update_partial_passage(token, payload)
                elif event_type == "result":
                    self._apply_generated_passage(token, payload)
                elif event_type == "error":
                    messagebox.showerror("Generate Error", str(payload or "Unknown error"))
                elif event_type == "done":
                    done = True
        except queue.Empty:
            pass

        if not done and token == self.passage_generation_active_token:
            self.after(80, lambda t=token: self._poll_passage_generation_events(t))

    def _update_partial_passage(self, token, text):
        if token != self.passage_generation_token:
            return
        self.current_passage = str(text or "").strip()
        self.current_passage_original = self.current_passage
        if self.current_passage:
            self._set_passage_text(self.current_passage)

    def _run_passage_generation(self, token, words, model_name):
        result = None
        fallback_error = None

        try:
            result = generate_english_passage_with_gemini(
                words,
                api_key=get_gemini_api_key(),
                max_words=24,
                timeout=90,
                model=model_name,
            )
            self._emit_passage_event("partial", token, result.get("passage", ""))
        except Exception as e:
            fallback_error = str(e)
            try:
                result = build_ielts_listening_passage(words, max_words=24)
                result["source"] = "template"
                result["fallback_reason"] = fallback_error
            except Exception as e2:
                template_error = str(e2)
                self._emit_passage_event(
                    "error",
                    token,
                    f"Failed to generate passage.\nGemini: {fallback_error}\nTemplate: {template_error}",
                )
                self._emit_passage_event("done", token, None)
                return

        self._emit_passage_event("result", token, result)
        self._emit_passage_event("done", token, None)

    def _apply_generated_passage(self, token, result):
        if token != self.passage_generation_token:
            return

        self.current_passage = str(result.get("passage", "")).strip()
        self.current_passage_original = self.current_passage
        self.current_passage_words = list(result.get("used_words") or [])
        self.passage_is_practice = False
        self.passage_cloze_text = ""
        self.passage_answers = []
        self._set_passage_text(self.current_passage)

        used_count = len(result.get("used_words") or [])
        source = result.get("source")
        if source == "gemini":
            coverage = int(float(result.get("coverage", 0.0)) * 100)
            model = result.get("model") or DEFAULT_GEMINI_MODEL
            missed_words = list(result.get("missed_words") or [])
            suffix = " (low coverage)" if result.get("low_coverage") else ""
            repaired_suffix = " | repaired" if result.get("repaired") else ""
            missed_suffix = f" | missed: {', '.join(missed_words[:3])}" if missed_words else ""
            self.passage_status_var.set(
                f"Gemini ({model}) | used {used_count} words | coverage {coverage}%{suffix}{repaired_suffix}{missed_suffix}."
            )
            return

        skipped_count = len(result.get("skipped_words") or [])
        scenario = result.get("scenario") or "General"
        reason = result.get("fallback_reason") or "Gemini not available."
        extra = f", skipped {skipped_count}" if skipped_count else ""
        self.passage_status_var.set(
            f"Template fallback ({scenario}) | used {used_count} words{extra}. Reason: {reason}"
        )

    def _pause_word_playback(self):
        self.cancel_schedule()
        self.play_token += 1
        if self.play_state == "playing":
            self.play_state = "paused"
            self.status_var.set("Paused (reading passage)")
            self.update_play_button()
        tts_cancel_all()

    def play_generated_passage(self):
        text = self.current_passage_original.strip() or self.current_passage.strip() or self._get_passage_text()
        if not text:
            messagebox.showinfo("Info", "Generate a passage first.")
            return

        speech_text = self._speech_text_from_passage(text)
        if not speech_text:
            messagebox.showinfo("Info", "Passage is empty.")
            return

        self._pause_word_playback()
        speak_stream_async(
            speech_text,
            self.volume_var.get() / 100.0,
            rate_ratio=self.speech_rate_var.get(),
            cancel_before=False,
            chunk_chars=90,
        )
        self.passage_status_var.set("Reading passage with Kokoro...")

    def stop_passage_playback(self):
        tts_cancel_all()
        self.passage_status_var.set("Stopped.")

    # Player controls
    def speak_selected(self):
        selected_idx = self._get_selected_index()
        if selected_idx is None or selected_idx >= len(self.store.words):
            messagebox.showinfo("Info", "Please select a word first.")
            return
        word = self.store.words[selected_idx]
        speak_async(
            word,
            self.volume_var.get() / 100.0,
            rate_ratio=self.speech_rate_var.get(),
            cancel_before=True,
        )

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
        if self.order_tip_click:
            self.order_tip_click.hide()
        self.cancel_schedule()
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
        if self.order_btn_click:
            self.order_btn_click.config(
                style="SelectedSpeed.TButton" if mode == "click_to_play" else "Speed.TButton"
            )
        if self.order_tip:
            self.order_tip.text = "In order (from current)"
        if self.order_tip_rand:
            self.order_tip_rand.text = "Random (no repeat)"
        if self.order_tip_click:
            self.order_tip_click.text = "Click one word to play one"
        if self.stop_at_end_check:
            # This option only applies to auto-play modes.
            self.stop_at_end_check.config(state=("disabled" if mode == "click_to_play" else "normal"))

    def set_interval(self, seconds):
        self.interval_var.set(float(seconds))
        self.update_speed_buttons()
        if self.play_state == "playing":
            self.schedule_next()
            self.play_current()

    def update_speed_buttons(self):
        for v, btn in self.speed_buttons:
            if abs(float(v) - float(self.interval_var.get())) < 1e-6:
                btn.config(style="SelectedSpeed.TButton")
            else:
                btn.config(style="Speed.TButton")

    def set_speech_rate(self, ratio):
        self.speech_rate_var.set(float(ratio))
        self.update_speech_rate_buttons()
        if self.play_state == "playing":
            self.play_current()

    def update_speech_rate_buttons(self):
        for v, btn in self.speech_rate_buttons:
            if abs(float(v) - float(self.speech_rate_var.get())) < 1e-6:
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

        self.settings_sections_visible = {"source": True, "order": True, "speed": False, "volume": False}
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
        for label, key in [("Source", "source"), ("Order", "order"), ("Speed", "speed"), ("Volume", "volume")]:
            btn = ttk.Button(left_menu, text=label, command=lambda k=key: toggle_section(k))
            btn.pack(fill=tk.X, pady=4)

        # Source section
        source_section = ttk.Frame(right_panel, style="Card.TFrame")
        ttk.Label(source_section, text="Source", style="Card.TLabel").pack(anchor="w")
        ttk.Label(
            source_section,
            text="Choose Kokoro accent (English US / English UK).",
            style="Card.TLabel",
            foreground="#666",
        ).pack(anchor="w", pady=(0, 4))

        self.voice_combo = ttk.Combobox(
            source_section,
            textvariable=self.voice_var,
            state="readonly",
            width=32,
        )
        self.voice_combo.pack(anchor="w")
        self.voice_combo.bind("<<ComboboxSelected>>", self.on_voice_change)
        ttk.Button(source_section, text="Gemini API Key", command=self.open_gemini_key_window).pack(
            anchor="w", pady=(8, 0)
        )

        source_sep = ttk.Separator(right_panel, orient="horizontal")
        sections.append({"key": "source", "frame": source_section, "sep": source_sep})

        # Order section
        order_section = ttk.Frame(right_panel, style="Card.TFrame")
        ttk.Label(order_section, text="Order", style="Card.TLabel").pack(anchor="w")
        ttk.Label(
            order_section,
            text="Choose in-order, random (no repeat), or click-to-play.",
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
        self.order_btn_click = ttk.Button(order_row, text="☝", command=lambda: self.set_order_mode("click_to_play"))
        self.order_btn_click.pack(side=tk.LEFT, padx=4)
        self.order_tip_click = Tooltip(self.order_btn_click, "Click one word to play one")
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
            text="Interval: time between words.",
            style="Card.TLabel",
            foreground="#666",
        ).pack(anchor="w", pady=(0, 4))
        self.speed_buttons = []
        speed_row = ttk.Frame(speed_section, style="Card.TFrame")
        speed_row.pack(anchor="w")
        for v in [1, 2, 3, 5, 10]:
            btn = ttk.Button(speed_row, text=f"{v}s", command=lambda val=v: self.set_interval(val))
            btn.pack(side=tk.LEFT, padx=3)
            self.speed_buttons.append((v, btn))

        custom_row = ttk.Frame(speed_section, style="Card.TFrame")
        custom_row.pack(anchor="w", pady=(4, 0))
        ttk.Label(custom_row, text="Custom (s):", style="Card.TLabel").pack(side=tk.LEFT)
        self.custom_interval = ttk.Entry(custom_row, width=6)
        self.custom_interval.pack(side=tk.LEFT, padx=4)
        self.custom_interval.bind("<Return>", lambda _e: self.apply_custom_interval())
        ttk.Button(custom_row, text="Apply", command=self.apply_custom_interval).pack(side=tk.LEFT)

        ttk.Label(
            speed_section,
            text="Pronunciation: speaking speed for each word.",
            style="Card.TLabel",
            foreground="#666",
        ).pack(anchor="w", pady=(8, 4))
        self.speech_rate_buttons = []
        speech_row = ttk.Frame(speed_section, style="Card.TFrame")
        speech_row.pack(anchor="w")
        for v in [0.6, 0.8, 1.0, 1.2]:
            btn = ttk.Button(speech_row, text=f"{v:.1f}x", command=lambda val=v: self.set_speech_rate(val))
            btn.pack(side=tk.LEFT, padx=3)
            self.speech_rate_buttons.append((v, btn))
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
        self.update_speech_rate_buttons()
        self.on_volume_change()
        self.refresh_voice_list()

    def apply_custom_interval(self):
        try:
            val = float(self.custom_interval.get())
            if val < 0.2:
                raise ValueError
        except Exception:
            messagebox.showinfo("Info", "Please enter a valid number (>=0.2).")
            return
        self.set_interval(val)

    def build_queue(self):
        if self.order_mode.get() == "click_to_play":
            return []
        indices = list(range(len(self.store.words)))
        if self.order_mode.get() == "random_no_repeat":
            random.shuffle(indices)
        return indices

    def rebuild_queue_on_mode_change(self):
        if not self.store.words:
            self.queue = []
            self.pos = -1
            return
        if self.order_mode.get() == "click_to_play":
            self.play_state = "stopped"
            selected_idx = self._get_selected_index()
            if selected_idx is not None:
                self.queue = [selected_idx]
                self.pos = 0
                self.set_current_word()
            else:
                self.queue = []
                self.pos = -1
                self.current_word = None
                self.status_var.set("Click a word to play")
            self.update_play_button()
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
        if self.order_mode.get() == "click_to_play":
            if self._get_selected_index() is None:
                messagebox.showinfo("Info", "Click a word first in Click-to-play mode.")
                return
            self.play_state = "stopped"
            self.cancel_schedule()
            tts_cancel_all()
            self.play_token += 1
            self.build_queue_from_selection()
            self.update_play_button()
            self.play_current()
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
        speak_async(
            self.current_word,
            self.volume_var.get() / 100.0,
            rate_ratio=self.speech_rate_var.get(),
            cancel_before=True,
        )

    def schedule_next(self):
        self.cancel_schedule()
        if self.play_state != "playing":
            return
        interval = max(0.2, float(self.interval_var.get()))
        token = self.play_token
        self.after_id = self.after(int(interval * 1000), lambda: self.next_word(token))

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
        if self.word_table:
            self.word_table.selection_remove(*self.word_table.selection())
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
        selected_idx = self._get_selected_index()
        start_idx = selected_idx if selected_idx is not None else 0

        if self.order_mode.get() == "click_to_play":
            self.queue = [start_idx]
            self.pos = 0
        elif self.order_mode.get() == "random_no_repeat":
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
        if self.suppress_word_select_action:
            self.suppress_word_select_action = False
            return
        # When the user clicks a word, restart playback from that word.
        if not self.store.words:
            return
        if self._get_selected_index() is None:
            return
        tts_cancel_all()
        self.play_token += 1
        self.build_queue_from_selection()
        if self.order_mode.get() == "click_to_play":
            self.play_state = "stopped"
            self.cancel_schedule()
            self.update_play_button()
            self.play_current()
            return
        if self.play_state == "playing":
            self.update_play_button()
            self.play_current()
            self.schedule_next()

    def on_word_double_click(self, _event=None):
        # Double-click enters inline edit mode for the selected word.
        self.start_edit_selected_word()
        return "break"

    def on_word_right_click(self, event):
        if not self.word_table or not self.word_context_menu:
            return
        row_id = self.word_table.identify_row(event.y)
        if not row_id:
            return
        try:
            self.suppress_word_select_action = True
            self.word_table.selection_set(row_id)
            self.word_table.focus(row_id)
        except Exception:
            pass
        try:
            self.word_context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.word_context_menu.grab_release()
        return "break"

    def _fallback_sentence(self, word):
        w = str(word or "").strip()
        if not w:
            return ""
        if " " in w:
            return f'Please use "{w}" in your next speaking practice task.'
        return f"I wrote the word {w} in my notebook for today's review."

    def _clear_sentence_event_queue(self):
        try:
            while True:
                self.sentence_event_queue.get_nowait()
        except queue.Empty:
            return

    def _emit_sentence_event(self, event_type, token, payload=None):
        try:
            self.sentence_event_queue.put_nowait((event_type, token, payload))
        except Exception:
            return

    def _poll_sentence_events(self, token):
        if token != self.sentence_generation_active_token:
            return
        done = False
        try:
            while True:
                event_type, event_token, payload = self.sentence_event_queue.get_nowait()
                if event_token != token:
                    continue
                if event_type == "result":
                    data = payload or {}
                    self._show_sentence_window(
                        data.get("word", ""),
                        data.get("sentence", ""),
                        data.get("source", "Unknown"),
                    )
                elif event_type == "error":
                    messagebox.showerror("Sentence Error", str(payload or "Unknown error"))
                elif event_type == "done":
                    done = True
        except queue.Empty:
            pass

        if not done and token == self.sentence_generation_active_token:
            self.after(80, lambda t=token: self._poll_sentence_events(t))

    def make_sentence_for_selected_word(self):
        selected_idx = self._get_selected_index()
        if selected_idx is None or selected_idx >= len(self.store.words):
            messagebox.showinfo("Info", "Please select a word first.")
            return
        if not self._require_gemini_ready():
            return
        word = self.store.words[selected_idx]
        model_name = self._get_selected_gemini_model()
        self.status_var.set(f"Generating IELTS sentence for '{word}' with {model_name}...")
        self.sentence_generation_token += 1
        token = self.sentence_generation_token
        self.sentence_generation_active_token = token
        self._clear_sentence_event_queue()

        import threading

        def _run():
            try:
                sentence = generate_example_sentence_with_gemini(
                    word=word,
                    api_key=get_gemini_api_key(),
                    model=model_name,
                    timeout=45,
                )
                source = f"Gemini ({model_name}, IELTS)"
            except Exception:
                sentence = self._fallback_sentence(word)
                source = "Fallback"
            self._emit_sentence_event(
                "result",
                token,
                {"word": word, "sentence": sentence, "source": source},
            )
            self._emit_sentence_event("done", token, None)

        threading.Thread(target=_run, daemon=True).start()
        self.after(80, lambda t=token: self._poll_sentence_events(t))

    def _show_sentence_window(self, word, sentence, source):
        self.status_var.set(f"Sentence ready for '{word}'.")
        if self.sentence_window and self.sentence_window.winfo_exists():
            self.sentence_window.destroy()

        self.sentence_window = tk.Toplevel(self)
        self.sentence_window.title(f"Sentence - {word}")
        self.sentence_window.configure(bg="#f6f7fb")
        self.sentence_window.resizable(False, False)

        wrap = ttk.Frame(self.sentence_window, style="Card.TFrame")
        wrap.pack(fill="both", expand=True, padx=10, pady=10)
        ttk.Label(wrap, text=f"Word: {word}", style="Card.TLabel").pack(anchor="w")
        ttk.Label(wrap, text=f"Source: {source}", style="Card.TLabel", foreground="#666").pack(anchor="w", pady=(0, 6))

        text = tk.Text(
            wrap,
            wrap="word",
            height=4,
            width=56,
            bg="#ffffff",
            fg="#222222",
            highlightthickness=1,
            highlightbackground="#d9dbe1",
        )
        text.pack(fill="x")
        text.insert("1.0", sentence or "")
        text.config(state="disabled")

        row = ttk.Frame(wrap, style="Card.TFrame")
        row.pack(fill="x", pady=(8, 0))
        ttk.Button(
            row,
            text="Read Sentence",
            command=lambda s=sentence: speak_async(
                s,
                self.volume_var.get() / 100.0,
                rate_ratio=self.speech_rate_var.get(),
                cancel_before=True,
            ),
        ).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(row, text="Close", command=self.sentence_window.destroy).pack(side=tk.LEFT)

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

    # Voice source
    def refresh_voice_list(self):
        voices = list_system_voices()
        self.voice_map = {}
        options = []

        for v in voices:
            name = v.get("name") or v.get("id") or "Unknown"
            langs = v.get("languages") or []
            lang_parts = []
            for lang in langs:
                if isinstance(lang, bytes):
                    try:
                        lang = lang.decode(errors="ignore")
                    except Exception:
                        lang = ""
                lang = str(lang).strip()
                if lang:
                    lang_parts.append(lang)
            lang_text = ",".join(lang_parts)
            label = f"{name} ({lang_text})" if lang_text else name
            if label not in self.voice_map:
                self.voice_map[label] = (SOURCE_KOKORO, v.get("id"), name)
                options.append(label)

        if self.voice_combo:
            self.voice_combo["values"] = options

        # restore current selection
        current_source = get_voice_source()
        current_id = get_voice_id()
        selected = options[0] if options else ""
        if current_source == SOURCE_KOKORO and current_id:
            for label, data in self.voice_map.items():
                if data[0] == SOURCE_KOKORO and data[1] == current_id:
                    selected = label
                    break
        self.voice_var.set(selected)

    def on_voice_change(self, _event=None):
        label = self.voice_var.get()
        data = self.voice_map.get(label)
        if not data:
            set_voice_source(SOURCE_KOKORO, "bf_emma", "English (UK)")
            return
        source, voice_id, voice_label = data
        set_voice_source(source, voice_id, voice_label)

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
