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
    get_cached_translations,
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
from services.corpus_search import (
    corpus_stats,
    get_nlp_status,
    import_corpus_files,
    list_documents as list_corpus_documents,
    search_corpus,
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
        self.find_btn = None
        self.find_btn_check = None
        self.check_controls = None
        self.hide_words_btn = None
        self.main = None
        self.voice_var = tk.StringVar(value="")
        self.voice_combo = None
        self.voice_map = {}
        self.word_table = None
        self.word_table_scroll = None
        self.word_context_menu = None
        self.word_edit_entry = None
        self.word_edit_row = None
        self.word_edit_column = None
        self.suppress_word_select_action = False
        self.sentence_window = None
        self.manual_words_window = None
        self.manual_words_table = None
        self.manual_words_table_scroll = None
        self.manual_preview_edit_entry = None
        self.manual_preview_edit_item = None
        self.manual_preview_edit_column = None
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
        self.find_window = None
        self.find_search_var = tk.StringVar(value="")
        self.find_limit_var = tk.StringVar(value="20")
        self.find_status_var = tk.StringVar(value="Import docs, then search by word or phrase.")
        self.find_results_table = None
        self.find_preview_text = None
        self.find_docs_list = None
        self.find_doc_items = []
        self.find_result_items = {}
        self.find_task_queue = queue.Queue()
        self.find_task_token = 0
        self.find_active_token = 0

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
        self.refresh_gemini_models()
        self.after(150, self.ensure_gemini_api_key)

    def build_ui(self):
        header = ttk.Frame(self)
        header.pack(fill=tk.X, pady=(0, 6))
        title = ttk.Label(header, text="Word Speaker", font=("Segoe UI", 14, "bold"))
        title.pack(side=tk.LEFT)

        ttk.Separator(self, orient="horizontal").pack(fill=tk.X, pady=(0, 10))

        self.main = ttk.Frame(self)
        self.main.pack(fill="both", expand=True)
        self.main.grid_columnconfigure(0, weight=5)
        self.main.grid_columnconfigure(2, weight=3)
        self.main.grid_rowconfigure(0, weight=1)

        self.left = ttk.Frame(self.main, style="Card.TFrame")
        self.left.grid(row=0, column=0, sticky="nsew")
        self.mid_sep = ttk.Separator(self.main, orient="vertical")
        self.mid_sep.grid(row=0, column=1, sticky="ns", padx=10)
        self.right = ttk.Frame(self.main, style="Card.TFrame")
        self.right.grid(row=0, column=2, sticky="nsew")

        # Left: Word list + player bar
        left_title = ttk.Label(self.left, text="Word List", style="Card.TLabel")
        left_title.grid(row=0, column=0, columnspan=2, padx=12, pady=(12, 6), sticky="w")

        top_btn_row = ttk.Frame(self.left, style="Card.TFrame")
        top_btn_row.grid(row=1, column=0, columnspan=2, padx=12, pady=6, sticky="ew")
        top_btn_row.grid_columnconfigure(0, weight=1)
        top_btn_row.grid_columnconfigure(1, weight=1)
        top_btn_row.grid_columnconfigure(2, weight=1)
        top_btn_row.grid_columnconfigure(3, weight=1)

        btn_load = ttk.Button(top_btn_row, text="⭳", style="Primary.TButton", command=self.load_words)
        btn_load.grid(row=0, column=0, padx=(0, 6), sticky="ew")
        Tooltip(btn_load, "Import Words")

        btn_manual = ttk.Button(top_btn_row, text="✍", command=self.open_manual_words_window)
        btn_manual.grid(row=0, column=1, padx=3, sticky="ew")
        Tooltip(btn_manual, "Type/Paste words or a two-column table")

        btn_speak = ttk.Button(top_btn_row, text="🔊", command=self.speak_selected)
        btn_speak.grid(row=0, column=2, padx=3, sticky="ew")
        Tooltip(btn_speak, "Speak Selected")

        btn_find = ttk.Button(top_btn_row, text="🔎", command=self.open_find_window)
        btn_find.grid(row=0, column=3, padx=(6, 0), sticky="ew")
        Tooltip(btn_find, "Find Corpus Sentences")

        table_wrap = ttk.Frame(self.left, style="Card.TFrame")
        table_wrap.grid(row=2, column=0, columnspan=2, padx=12, pady=(4, 6), sticky="nsew")
        self.left.grid_rowconfigure(2, weight=1)
        self.left.grid_columnconfigure(0, weight=1)
        self.left.grid_columnconfigure(1, weight=1)

        self.word_table = ttk.Treeview(
            table_wrap,
            columns=("en", "note", "zh"),
            show="headings",
            height=18,
            selectmode="extended",
        )
        self.word_table.heading("en", text="English")
        self.word_table.heading("note", text="Notes")
        self.word_table.heading("zh", text="中文")
        self.word_table.column("en", width=260, anchor="w")
        self.word_table.column("note", width=320, anchor="w")
        self.word_table.column("zh", width=240, anchor="w")

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
        self.word_context_menu.add_command(label="Find", command=self.search_selected_word_in_corpus)
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
        self.player_frame.grid_columnconfigure(5, weight=1)

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

        self.find_btn = ttk.Button(
            self.player_frame, text="🔎", style="Icon.TButton", width=4, command=self.open_find_window
        )
        self.find_btn.grid(row=0, column=5, padx=4, sticky="ew")
        Tooltip(self.find_btn, "Find Corpus Sentences")

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

        self.find_btn_check = ttk.Button(
            self.check_controls, text="🔎", style="Icon.TButton", width=4, command=self.open_find_window
        )
        self.find_btn_check.pack(side=tk.LEFT, padx=3)
        Tooltip(self.find_btn_check, "Find Corpus Sentences")

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

    def _parse_manual_rows(self, raw_text):
        text = str(raw_text or "")
        lines = text.splitlines()
        rows = []
        pending_word = None
        pending_note_lines = []
        for line in lines:
            line = line.strip()
            if not line:
                continue
            if "\t" in line:
                if pending_word:
                    rows.append({"word": pending_word, "note": " | ".join(pending_note_lines).strip()})
                    pending_word = None
                    pending_note_lines = []
                parts = [part.strip() for part in line.split("\t")]
                word = str(parts[0] or "").strip()
                note = " | ".join(part for part in parts[1:] if str(part).strip())
                if word:
                    rows.append({"word": word, "note": note})
                continue
            if self._looks_like_word_line(line):
                if pending_word:
                    rows.append({"word": pending_word, "note": " | ".join(pending_note_lines).strip()})
                pending_word = line
                pending_note_lines = []
                continue

            if pending_word:
                pending_note_lines.append(line)
            else:
                parts = re.split(r"[,;；，]+", line)
                if len(parts) > 1 and all(str(part or "").strip() for part in parts):
                    for part in parts:
                        token = str(part or "").strip()
                        if token:
                            rows.append({"word": token, "note": ""})
                else:
                    rows.append({"word": line, "note": ""})

        if pending_word:
            rows.append({"word": pending_word, "note": " | ".join(pending_note_lines).strip()})
        return rows

    def _looks_like_word_line(self, text):
        value = str(text or "").strip()
        if not value:
            return False
        if re.search(r"[\u4e00-\u9fff]", value):
            return False
        letters = len(re.findall(r"[A-Za-z]", value))
        if letters <= 0:
            return False
        cjk_count = len(re.findall(r"[\u4e00-\u9fff]", value))
        return cjk_count == 0 and letters >= max(1, len(value) // 4)

    def _append_manual_preview_rows(self, rows, replace=False):
        if not self.manual_words_table:
            return
        if replace:
            self.manual_words_table.delete(*self.manual_words_table.get_children())
        start = len(self.manual_words_table.get_children())
        for offset, row in enumerate(rows):
            word = str(row.get("word") or "").strip()
            note = str(row.get("note") or "").strip()
            if not word:
                continue
            self.manual_words_table.insert("", tk.END, iid=f"manual_{start + offset}", values=(word, note))

    def _collect_manual_rows_from_table(self):
        rows = []
        if not self.manual_words_table:
            return rows
        for item_id in self.manual_words_table.get_children():
            values = list(self.manual_words_table.item(item_id, "values") or [])
            word = str(values[0] if len(values) > 0 else "").strip()
            note = str(values[1] if len(values) > 1 else "").strip()
            if not word:
                continue
            rows.append({"word": word, "note": note})
        return rows

    def _clear_manual_preview(self):
        self._cancel_manual_preview_edit()
        if self.manual_words_table:
            self.manual_words_table.delete(*self.manual_words_table.get_children())

    def _close_manual_words_window(self):
        self._cancel_manual_preview_edit()
        if self.manual_words_window and self.manual_words_window.winfo_exists():
            self.manual_words_window.destroy()
        self.manual_words_window = None
        self.manual_words_table = None
        self.manual_words_table_scroll = None

    def _cancel_manual_preview_edit(self, _event=None):
        if self.manual_preview_edit_entry and self.manual_preview_edit_entry.winfo_exists():
            self.manual_preview_edit_entry.destroy()
        self.manual_preview_edit_entry = None
        self.manual_preview_edit_item = None
        self.manual_preview_edit_column = None
        return "break"

    def _finish_manual_preview_edit(self, _event=None):
        if not self.manual_preview_edit_entry or not self.manual_words_table:
            return "break"
        item_id = self.manual_preview_edit_item
        column_id = self.manual_preview_edit_column
        new_value = re.sub(r"\s+", " ", str(self.manual_preview_edit_entry.get() or "").strip())
        self._cancel_manual_preview_edit()
        if not item_id or not self.manual_words_table.exists(item_id):
            return "break"
        values = list(self.manual_words_table.item(item_id, "values") or ["", ""])
        while len(values) < 2:
            values.append("")
        if column_id == "#1" and not new_value:
            return "break"
        if column_id == "#1":
            values[0] = new_value
        elif column_id == "#2":
            values[1] = new_value
        self.manual_words_table.item(item_id, values=values)
        return "break"

    def _start_manual_preview_edit(self, event=None, item_id=None, column_id=None):
        if not self.manual_words_table:
            return "break"
        target_item = item_id
        target_column = column_id or "#1"
        if event is not None:
            target_item = self.manual_words_table.identify_row(event.y)
            target_column = self.manual_words_table.identify_column(event.x)
        if not target_item:
            target_item = str(self.manual_words_table.focus() or "").strip()
        if not target_item:
            selection = self.manual_words_table.selection()
            if selection:
                target_item = selection[0]
        if target_column not in ("#1", "#2"):
            target_column = "#1"
        if not target_item or not self.manual_words_table.exists(target_item):
            return "break"
        self._cancel_manual_preview_edit()
        bbox = self.manual_words_table.bbox(target_item, target_column)
        if not bbox:
            return "break"
        x, y, width, height = bbox
        values = list(self.manual_words_table.item(target_item, "values") or ["", ""])
        while len(values) < 2:
            values.append("")
        current_value = values[0] if target_column == "#1" else values[1]
        entry = ttk.Entry(self.manual_words_table)
        entry.insert(0, current_value)
        entry.select_range(0, tk.END)
        entry.focus_set()
        entry.place(x=x, y=y, width=width, height=height)
        entry.bind("<Return>", self._finish_manual_preview_edit)
        entry.bind("<Escape>", self._cancel_manual_preview_edit)
        entry.bind("<FocusOut>", self._finish_manual_preview_edit)
        self.manual_preview_edit_entry = entry
        self.manual_preview_edit_item = target_item
        self.manual_preview_edit_column = target_column
        return "break"

    def _add_manual_preview_row(self):
        if not self.manual_words_table:
            return
        item_id = f"manual_{len(self.manual_words_table.get_children())}"
        self.manual_words_table.insert("", tk.END, iid=item_id, values=("", ""))
        self.manual_words_table.selection_set(item_id)
        self.manual_words_table.focus(item_id)
        self._start_manual_preview_edit(item_id=item_id, column_id="#1")

    def _delete_selected_manual_preview_rows(self):
        if not self.manual_words_table:
            return
        selected = list(self.manual_words_table.selection())
        if not selected:
            return
        self._cancel_manual_preview_edit()
        for item_id in selected:
            if self.manual_words_table.exists(item_id):
                self.manual_words_table.delete(item_id)

    def on_manual_preview_paste(self, _event=None):
        try:
            raw = self.clipboard_get()
        except Exception:
            return "break"
        rows = self._parse_manual_rows(raw)
        if not rows:
            return "break"
        self._append_manual_preview_rows(rows, replace=False)
        return "break"

    def _update_word_stats_from_manual_input(self, words):
        try:
            stats = self.store.load_stats()
            for word in words:
                stats[word] = int(stats.get(word, 0)) + 1
            self.store.save_stats(stats)
        except Exception:
            return

    def _apply_manual_words(self, rows, mode="replace"):
        normalized_words = []
        normalized_notes = []
        for row in rows:
            if isinstance(row, dict):
                word = str(row.get("word") or "").strip()
                note = re.sub(r"\s+", " ", str(row.get("note") or "").strip())
            else:
                word = str(row or "").strip()
                note = ""
            if not word:
                continue
            normalized_words.append(word)
            normalized_notes.append(note)
        if not normalized_words:
            messagebox.showinfo("Info", "No valid words found.")
            return False
        self.cancel_word_edit()

        if mode == "append":
            merged_words = list(self.store.words) + normalized_words
            merged_notes = list(self.store.notes) + normalized_notes
            self.store.set_words(merged_words, merged_notes)
            self.render_words(merged_words)
            self._update_word_stats_from_manual_input(normalized_words)
            status_text = f"Appended {len(normalized_words)} words."
        else:
            self.store.set_words(normalized_words, normalized_notes)
            self.render_words(normalized_words)
            self._update_word_stats_from_manual_input(normalized_words)
            status_text = f"Loaded {len(normalized_words)} words."
        self.reset_playback_state()
        self.status_var.set(status_text)
        return True

    def _apply_manual_words_from_editor(self, mode):
        rows = self._collect_manual_rows_from_table()
        if not rows:
            messagebox.showinfo("Info", "No valid words found.")
            return
        ok = self._apply_manual_words(rows, mode=mode)
        if ok:
            self._close_manual_words_window()

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
        ttk.Label(wrap, text="Paste into preview table", style="Card.TLabel").pack(anchor="w")
        ttk.Label(
            wrap,
            text="Paste a two-column table from Google Docs/Sheets. Column 1 = English, Column 2 = Notes.",
            style="Card.TLabel",
            foreground="#666",
        ).pack(anchor="w", pady=(0, 6))

        table_wrap = ttk.Frame(wrap, style="Card.TFrame")
        table_wrap.pack(fill="both", expand=True)
        self.manual_words_table = ttk.Treeview(
            table_wrap,
            columns=("en", "note"),
            show="headings",
            height=12,
            selectmode="browse",
        )
        self.manual_words_table.heading("en", text="English")
        self.manual_words_table.heading("note", text="Notes")
        self.manual_words_table.column("en", width=260, anchor="w")
        self.manual_words_table.column("note", width=420, anchor="w")
        self.manual_words_table.grid(row=0, column=0, sticky="nsew")
        self.manual_words_table_scroll = ttk.Scrollbar(
            table_wrap, orient="vertical", command=self.manual_words_table.yview
        )
        self.manual_words_table_scroll.grid(row=0, column=1, sticky="ns")
        self.manual_words_table.configure(yscrollcommand=self.manual_words_table_scroll.set)
        table_wrap.grid_rowconfigure(0, weight=1)
        table_wrap.grid_columnconfigure(0, weight=1)
        self.manual_words_table.bind("<Control-v>", self.on_manual_preview_paste)
        self.manual_words_table.bind("<Control-V>", self.on_manual_preview_paste)
        self.manual_words_table.bind("<Double-1>", self._start_manual_preview_edit)
        self.manual_words_table.bind("<F2>", self._start_manual_preview_edit)
        self.manual_words_table.bind("<Delete>", lambda _e: self._delete_selected_manual_preview_rows())
        self.manual_words_table.bind("<Escape>", self._cancel_manual_preview_edit)
        self.manual_words_table.focus_set()

        helper = ttk.Label(
            wrap,
            text="Tip: press Ctrl+V to paste, or add rows and edit cells directly.",
            style="Card.TLabel",
            foreground="#666",
        )
        helper.pack(anchor="w", pady=(6, 0))

        row = ttk.Frame(wrap, style="Card.TFrame")
        row.pack(fill="x", pady=(8, 0))
        ttk.Button(
            row,
            text="Paste Clipboard",
            command=self.on_manual_preview_paste,
        ).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(
            row,
            text="Add Row",
            command=self._add_manual_preview_row,
        ).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(
            row,
            text="Delete Selected",
            command=self._delete_selected_manual_preview_rows,
        ).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(
            row,
            text="Clear",
            command=self._clear_manual_preview,
        ).pack(side=tk.LEFT, padx=(0, 12))
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
            command=self._close_manual_words_window,
        ).pack(side=tk.LEFT)

        self.manual_words_window.protocol("WM_DELETE_WINDOW", self._close_manual_words_window)

    def on_word_table_paste(self, _event=None):
        try:
            raw = self.clipboard_get()
        except Exception:
            return "break"
        rows = self._parse_manual_rows(raw)
        if not rows:
            return "break"
        self._apply_manual_words(rows, mode="append")
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
            note = self.store.notes[row_idx] if row_idx < len(self.store.notes) else ""
            self.word_table.item(iid, values=(word, note, zh_text))

    def start_edit_selected_word(self, _event=None):
        return self.start_edit_word_cell(column_id="#1")

    def start_edit_word_cell(self, event=None, column_id=None):
        if not self.word_table:
            return "break"
        row_idx = self._get_selected_index()
        target_column = column_id or "#1"
        if event is not None:
            row_id = self.word_table.identify_row(event.y)
            hit_column = self.word_table.identify_column(event.x)
            if row_id:
                try:
                    row_idx = int(row_id)
                except Exception:
                    row_idx = self._get_selected_index()
                self.word_table.selection_set(row_id)
                self.word_table.focus(row_id)
            if hit_column in ("#1", "#2"):
                target_column = hit_column
        if row_idx is None or row_idx >= len(self.store.words):
            return "break"
        if target_column not in ("#1", "#2"):
            return "break"
        self.cancel_word_edit()
        iid = str(row_idx)
        bbox = self.word_table.bbox(iid, target_column)
        if not bbox:
            return "break"

        x, y, width, height = bbox
        if target_column == "#2":
            current_value = self.store.notes[row_idx] if row_idx < len(self.store.notes) else ""
        else:
            current_value = self.store.words[row_idx]
        entry = ttk.Entry(self.word_table)
        entry.insert(0, current_value)
        entry.select_range(0, tk.END)
        entry.focus_set()
        entry.place(x=x, y=y, width=width, height=height)
        entry.bind("<Return>", lambda _e, idx=row_idx: self.finish_edit_word(idx))
        entry.bind("<Escape>", self.cancel_word_edit)
        entry.bind("<FocusOut>", lambda _e, idx=row_idx: self.finish_edit_word(idx))
        self.word_edit_entry = entry
        self.word_edit_row = row_idx
        self.word_edit_column = target_column
        return "break"

    def cancel_word_edit(self, _event=None):
        if self.word_edit_entry and self.word_edit_entry.winfo_exists():
            self.word_edit_entry.destroy()
        self.word_edit_entry = None
        self.word_edit_row = None
        self.word_edit_column = None
        return "break"

    def finish_edit_word(self, row_idx=None):
        if not self.word_edit_entry:
            return "break"
        idx = self.word_edit_row if row_idx is None else row_idx
        edit_column = self.word_edit_column or "#1"
        raw_value = str(self.word_edit_entry.get() or "").strip()
        new_word = re.sub(r"\s+", " ", raw_value)
        self.cancel_word_edit()
        if idx is None or idx < 0 or idx >= len(self.store.words):
            return "break"
        if edit_column == "#1" and not new_word:
            messagebox.showinfo("Info", "Word cannot be empty.")
            return "break"
        iid = str(idx)

        if edit_column == "#2":
            old_note = self.store.notes[idx] if idx < len(self.store.notes) else ""
            if idx >= len(self.store.notes):
                self.store.notes.extend([""] * (idx - len(self.store.notes) + 1))
            if new_word == old_note:
                return "break"
            self.store.notes[idx] = new_word
            zh = self.translations.get(self.store.words[idx]) or ""
            if self.word_table and self.word_table.exists(iid):
                self.word_table.item(iid, values=(self.store.words[idx], new_word, zh))
            saved = False
            source_path = str(self.store.get_current_source_path() or "").lower()
            if source_path.endswith(".csv"):
                saved = self._save_words_back_to_source()
            source_note = " and saved to source file" if saved else ""
            self.status_var.set(f"Updated note for '{self.store.words[idx]}'{source_note}.")
            return "break"

        old_word = self.store.words[idx]
        if new_word == old_word:
            return "break"

        old_note = self.store.notes[idx] if idx < len(self.store.notes) else ""
        self.store.words[idx] = new_word
        if self.word_table and self.word_table.exists(iid):
            self.word_table.item(iid, values=(new_word, old_note, "Translating..."))

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
        cached = get_cached_translations(words)
        self.translations = dict(cached)
        self.word_table.delete(*self.word_table.get_children())
        for idx, w in enumerate(words):
            note = self.store.notes[idx] if idx < len(self.store.notes) else ""
            zh = cached.get(w)
            self.word_table.insert("", tk.END, iid=str(idx), values=(w, note, zh if zh is not None else ""))
        self.update_empty_state()
        missing_words = [w for w in words if w not in cached]
        if missing_words:
            self._start_translation_job(missing_words, token)

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
        self.translations.update(translated)
        for idx, word in enumerate(self.store.words):
            if not self.word_table.exists(str(idx)):
                continue
            zh = self.translations.get(word) or ""
            note = self.store.notes[idx] if idx < len(self.store.notes) else ""
            self.word_table.item(str(idx), values=(word, note, zh))

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

    # Corpus find
    def open_find_window(self):
        if self.find_window and self.find_window.winfo_exists():
            self.find_window.lift()
            self._set_find_query_from_selection()
            return

        self.find_window = tk.Toplevel(self)
        self.find_window.title("Find Corpus Sentences")
        self.find_window.configure(bg="#f6f7fb")
        self.find_window.minsize(900, 620)

        wrap = ttk.Frame(self.find_window, style="Card.TFrame")
        wrap.pack(fill="both", expand=True, padx=10, pady=10)
        wrap.grid_columnconfigure(0, weight=3)
        wrap.grid_columnconfigure(1, weight=2)
        wrap.grid_rowconfigure(2, weight=1)

        ttk.Label(wrap, text="Find Corpus Sentences", style="Card.TLabel").grid(
            row=0, column=0, columnspan=2, sticky="w"
        )
        ttk.Label(
            wrap,
            text="Import txt/docx/pdf files, build a local sentence index, then search by word or phrase.",
            style="Card.TLabel",
            foreground="#666",
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 8))

        top = ttk.Frame(wrap, style="Card.TFrame")
        top.grid(row=2, column=0, sticky="nsew", padx=(0, 10))
        top.grid_columnconfigure(0, weight=1)
        top.grid_rowconfigure(2, weight=1)
        top.grid_rowconfigure(4, weight=0)

        search_row = ttk.Frame(top, style="Card.TFrame")
        search_row.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        search_row.grid_columnconfigure(0, weight=1)
        entry = ttk.Entry(search_row, textvariable=self.find_search_var)
        entry.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        entry.bind("<Return>", lambda _e: self.run_find_search())
        ttk.Button(search_row, text="Search", command=self.run_find_search).grid(row=0, column=1, padx=(0, 6))
        ttk.Label(search_row, text="Show", style="Card.TLabel").grid(row=0, column=2, padx=(0, 4))
        limit_combo = ttk.Combobox(
            search_row,
            textvariable=self.find_limit_var,
            state="readonly",
            width=5,
            values=("20", "50", "100"),
        )
        limit_combo.grid(row=0, column=3, padx=(0, 6))
        ttk.Label(search_row, text="results", style="Card.TLabel").grid(row=0, column=4, padx=(0, 6))
        ttk.Button(search_row, text="Use Selected Word", command=self.search_selected_word_in_corpus).grid(
            row=0, column=5, padx=(0, 6)
        )
        ttk.Button(search_row, text="Import Docs", command=self.import_find_documents).grid(row=0, column=6)

        ttk.Label(top, textvariable=self.find_status_var, style="Card.TLabel", foreground="#444").grid(
            row=1, column=0, sticky="w", pady=(0, 8)
        )

        self.find_results_table = ttk.Treeview(
            top,
            columns=("sentence", "source"),
            show="headings",
            height=18,
            selectmode="browse",
        )
        self.find_results_table.heading("sentence", text="Sentence")
        self.find_results_table.heading("source", text="Source")
        self.find_results_table.column("sentence", width=600, anchor="w")
        self.find_results_table.column("source", width=260, anchor="w")
        find_scroll = ttk.Scrollbar(top, orient="vertical", command=self.find_results_table.yview)
        self.find_results_table.configure(yscrollcommand=find_scroll.set)
        self.find_results_table.grid(row=2, column=0, sticky="nsew")
        find_scroll.grid(row=2, column=1, sticky="ns")
        self.find_results_table.bind("<<TreeviewSelect>>", self._on_find_result_select)

        ttk.Label(top, text="Preview", style="Card.TLabel").grid(row=3, column=0, sticky="w", pady=(8, 4))
        self.find_preview_text = tk.Text(
            top,
            wrap="word",
            height=7,
            bg="#ffffff",
            fg="#222222",
            highlightthickness=1,
            highlightbackground="#d9dbe1",
        )
        self.find_preview_text.grid(row=4, column=0, columnspan=2, sticky="ew")
        self.find_preview_text.tag_configure("hit", background="#ffe58f", foreground="#111111")
        self.find_preview_text.configure(state="disabled")

        side = ttk.Frame(wrap, style="Card.TFrame")
        side.grid(row=2, column=1, sticky="nsew")
        side.grid_rowconfigure(1, weight=1)
        side.grid_columnconfigure(0, weight=1)

        side_header = ttk.Frame(side, style="Card.TFrame")
        side_header.grid(row=0, column=0, sticky="ew")
        side_header.grid_columnconfigure(0, weight=1)
        ttk.Label(side_header, text="Indexed Documents", style="Card.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Button(side_header, text="Clear Filter", command=self.clear_find_document_filter).grid(
            row=0, column=1, sticky="e"
        )
        self.find_docs_list = tk.Listbox(
            side,
            bg="#ffffff",
            fg="#222222",
            selectbackground="#cce1ff",
            selectforeground="#111111",
            highlightthickness=1,
            highlightbackground="#d9dbe1",
        )
        self.find_docs_list.grid(row=1, column=0, sticky="nsew", pady=(6, 0))

        self._set_find_query_from_selection()
        self.refresh_find_corpus_summary()

        def _on_close():
            self.find_window.destroy()
            self.find_window = None
            self.find_results_table = None
            self.find_preview_text = None
            self.find_docs_list = None
            self.find_doc_items = []
            self.find_result_items = {}

        self.find_window.protocol("WM_DELETE_WINDOW", _on_close)

    def _set_find_query_from_selection(self):
        selected_idx = self._get_selected_index()
        if selected_idx is None or selected_idx >= len(self.store.words):
            return
        self.find_search_var.set(self.store.words[selected_idx])

    def refresh_find_corpus_summary(self):
        try:
            stats = corpus_stats()
            docs = list_corpus_documents(limit=200)
        except Exception as e:
            if self.find_docs_list:
                self.find_docs_list.delete(0, tk.END)
            self.find_status_var.set(str(e))
            return
        if self.find_docs_list:
            self.find_docs_list.delete(0, tk.END)
            self.find_doc_items = list(docs)
            for item in docs:
                label = (
                    f"{item.get('name')} "
                    f"({int(item.get('chunk_count') or 0)} chunks / {int(item.get('sentence_count') or 0)} sentences)"
                )
                self.find_docs_list.insert(tk.END, label)
        self.find_status_var.set(
            f"Indexed {stats.get('documents', 0)} documents / {stats.get('chunks', 0)} chunks / {stats.get('sentences', 0)} sentences. NLP: {stats.get('nlp_mode')}"
        )

    def _clear_find_task_queue(self):
        try:
            while True:
                self.find_task_queue.get_nowait()
        except queue.Empty:
            return

    def _emit_find_task_event(self, event_type, token, payload=None):
        try:
            self.find_task_queue.put_nowait((event_type, token, payload))
        except Exception:
            return

    def _poll_find_task_events(self, token):
        if token != self.find_active_token:
            return
        done = False
        try:
            while True:
                event_type, event_token, payload = self.find_task_queue.get_nowait()
                if event_token != token:
                    continue
                if event_type == "import_done":
                    self._apply_find_import_result(payload or {})
                elif event_type == "search_done":
                    self._apply_find_search_result(payload or {})
                elif event_type == "error":
                    messagebox.showerror("Find Error", str(payload or "Unknown error"))
                elif event_type == "done":
                    done = True
        except queue.Empty:
            pass
        if not done and token == self.find_active_token:
            self.after(80, lambda t=token: self._poll_find_task_events(t))

    def import_find_documents(self):
        if not self.find_window:
            self.open_find_window()
        try:
            get_nlp_status()
        except Exception as e:
            messagebox.showerror("Find Setup Error", str(e))
            self.find_status_var.set(str(e))
            return
        paths = filedialog.askopenfilenames(
            title="Choose documents",
            filetypes=[("Supported files", "*.txt *.docx *.pdf"), ("Text", "*.txt"), ("Word", "*.docx"), ("PDF", "*.pdf")],
        )
        if not paths:
            return
        self.find_task_token += 1
        token = self.find_task_token
        self.find_active_token = token
        self._clear_find_task_queue()
        self.find_status_var.set("Importing documents and building sentence index...")

        import threading

        def _run():
            try:
                result = import_corpus_files(paths)
                self._emit_find_task_event("import_done", token, result)
            except Exception as e:
                self._emit_find_task_event("error", token, str(e))
            self._emit_find_task_event("done", token, None)

        threading.Thread(target=_run, daemon=True).start()
        self.after(80, lambda t=token: self._poll_find_task_events(t))

    def _apply_find_import_result(self, payload):
        self.refresh_find_corpus_summary()
        files_count = int(payload.get("files") or 0)
        chunk_count = int(payload.get("chunks") or 0)
        sent_count = int(payload.get("sentences") or 0)
        errors = list(payload.get("errors") or [])
        status = f"Imported {files_count} files and indexed {chunk_count} chunks / {sent_count} sentences."
        if errors:
            status += f" Errors: {len(errors)}"
        self.find_status_var.set(status)
        if errors:
            messagebox.showerror("Import Warning", "\n".join(errors[:10]))

    def run_find_search(self):
        query = str(self.find_search_var.get() or "").strip()
        if not query:
            messagebox.showinfo("Info", "Enter a word or phrase first.")
            return
        try:
            limit = int(str(self.find_limit_var.get() or "20").strip())
        except Exception:
            limit = 20
            self.find_limit_var.set("20")
        try:
            get_nlp_status()
        except Exception as e:
            messagebox.showerror("Find Setup Error", str(e))
            self.find_status_var.set(str(e))
            return
        self.find_task_token += 1
        token = self.find_task_token
        self.find_active_token = token
        self._clear_find_task_queue()
        selected_doc = self._get_selected_find_document()
        selected_doc_name = str(selected_doc.get("name") or "").strip() if selected_doc else ""
        if selected_doc_name:
            self.find_status_var.set(
                f"Searching '{query}' in {selected_doc_name} (up to {limit} results)..."
            )
        else:
            self.find_status_var.set(f"Searching corpus for '{query}' (up to {limit} results)...")

        import threading

        def _run():
            try:
                result = search_corpus(
                    query,
                    limit=limit,
                    document_path=selected_doc.get("path") if selected_doc else None,
                )
                self._emit_find_task_event("search_done", token, result)
            except Exception as e:
                self._emit_find_task_event("error", token, str(e))
            self._emit_find_task_event("done", token, None)

        threading.Thread(target=_run, daemon=True).start()
        self.after(80, lambda t=token: self._poll_find_task_events(t))

    def search_selected_word_in_corpus(self):
        if not self.find_window or not self.find_window.winfo_exists():
            self.open_find_window()
        self._set_find_query_from_selection()
        self.run_find_search()

    def _get_selected_find_document(self):
        if not self.find_docs_list or not self.find_doc_items:
            return None
        selection = self.find_docs_list.curselection()
        if not selection:
            return None
        idx = int(selection[0])
        if idx < 0 or idx >= len(self.find_doc_items):
            return None
        return self.find_doc_items[idx]

    def clear_find_document_filter(self):
        if self.find_docs_list:
            self.find_docs_list.selection_clear(0, tk.END)
        self.find_status_var.set("Document filter cleared. Search will use all indexed documents.")

    def _apply_find_search_result(self, payload):
        results = list(payload.get("results") or [])
        query = str(payload.get("query") or "").strip()
        lemmas = list(payload.get("lemmas") or [])
        document_path = str(payload.get("document_path") or "").strip()
        filtered_name = ""
        if document_path and self.find_doc_items:
            for item in self.find_doc_items:
                if str(item.get("path") or "").strip() == document_path:
                    filtered_name = str(item.get("name") or "").strip()
                    break
        self.find_result_items = {}
        if self.find_results_table:
            self.find_results_table.delete(*self.find_results_table.get_children())
            for idx, item in enumerate(results):
                source_bits = [item.get("source_file") or ""]
                for key in ("test_label", "section_label", "part_label", "speaker_label", "question_label"):
                    value = str(item.get(key) or "").strip()
                    if value:
                        source_bits.append(value)
                if item.get("page_num"):
                    source_bits.append(f"p.{item.get('page_num')}")
                source = " · ".join(bit for bit in source_bits if bit)
                item = dict(item)
                item["source_text"] = source
                row_id = f"find_{idx}"
                self.find_result_items[row_id] = item
                self.find_results_table.insert(
                    "",
                    tk.END,
                    iid=row_id,
                    values=(item.get("sentence_text") or "", source),
                )
            children = self.find_results_table.get_children()
            if children:
                first = children[0]
                self.find_results_table.selection_set(first)
                self.find_results_table.focus(first)
                self._show_find_result_preview(first)
            else:
                self._clear_find_preview()
        else:
            self._clear_find_preview()
        if filtered_name:
            self.find_status_var.set(
                f"Found {len(results)} results for '{query}' in {filtered_name}. Lemmas: {', '.join(lemmas) if lemmas else 'n/a'}"
            )
        else:
            self.find_status_var.set(
                f"Found {len(results)} results for '{query}'. Lemmas: {', '.join(lemmas) if lemmas else 'n/a'}"
            )

    def _clear_find_preview(self):
        if not self.find_preview_text:
            return
        self.find_preview_text.configure(state="normal")
        self.find_preview_text.delete("1.0", tk.END)
        self.find_preview_text.configure(state="disabled")

    def _on_find_result_select(self, _event=None):
        if not self.find_results_table:
            return
        selection = self.find_results_table.selection()
        if not selection:
            self._clear_find_preview()
            return
        self._show_find_result_preview(selection[0])

    def _show_find_result_preview(self, row_id):
        item = self.find_result_items.get(row_id)
        if not item or not self.find_preview_text:
            self._clear_find_preview()
            return
        sentence = str(item.get("sentence_text") or "")
        source = str(item.get("source_text") or "")
        ranges = list(item.get("highlight_ranges") or [])
        text = self.find_preview_text
        text.configure(state="normal")
        text.delete("1.0", tk.END)
        text.insert("1.0", sentence)
        for start, end in ranges:
            if start < end:
                text.tag_add("hit", f"1.0+{int(start)}c", f"1.0+{int(end)}c")
        if source:
            text.insert(tk.END, "\n\n")
            source_start = text.index(tk.END)
            text.insert(tk.END, source)
            text.tag_add("source", source_start, tk.END)
        text.tag_configure("source", foreground="#666666")
        text.configure(state="disabled")

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
            if self.wordlist_hidden:
                if self.main:
                    self.main.grid_columnconfigure(0, weight=1)
                    self.main.grid_columnconfigure(2, weight=0)
                self.right.grid_configure(row=0, column=0, columnspan=3, sticky="nsew")
                self.mid_sep.grid_remove()
            else:
                if self.main:
                    self.main.grid_columnconfigure(0, weight=5)
                    self.main.grid_columnconfigure(2, weight=3)
                self.left.grid_configure(columnspan=1)
                self.right.grid_configure(row=0, column=2, columnspan=1, sticky="nsew")
                self.mid_sep.grid()
            self.right.grid()
        else:
            if self.main:
                self.main.grid_columnconfigure(0, weight=1)
                self.main.grid_columnconfigure(2, weight=0)
            self.left.grid_configure(columnspan=3)
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

    def on_word_double_click(self, event=None):
        # Double-click edits the clicked word/note cell.
        self.start_edit_word_cell(event=event)
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
            self.hide_words_btn.config(text="Show Word List")
        else:
            self.left.grid()
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
