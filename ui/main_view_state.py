# -*- coding: utf-8 -*-
from __future__ import annotations

from dataclasses import dataclass, field, fields
from typing import Any
import queue as _queue
import tkinter as tk


@dataclass
class MainViewState:
    order_mode: Any
    interval_var: Any
    volume_var: Any
    dictation_volume_var: Any
    speech_rate_var: Any
    status_var: Any
    play_state: str = "stopped"
    queue: list = field(default_factory=list)
    pos: int = -1
    current_word: str | None = None
    after_id: Any = None
    play_token: int = 0
    playback_schedule_token: int = 0
    history_visible: bool = True
    check_visible: bool = False
    wordlist_hidden: bool = False
    order_btn: Any = None
    order_btn_rand: Any = None
    order_btn_click: Any = None
    order_tip: Any = None
    order_tip_rand: Any = None
    order_tip_click: Any = None
    loop_btn: Any = None
    loop_btn_stop: Any = None
    stop_at_end_check: Any = None
    speed_buttons: list = field(default_factory=list)
    speech_rate_buttons: list = field(default_factory=list)
    volume_scale: Any = None
    player_frame: Any = None
    play_btn_check: Any = None
    dictation_volume_btn: Any = None
    dictation_volume_popup: Any = None
    dictation_volume_scale: Any = None
    dictation_volume_value_label: Any = None
    settings_btn_check: Any = None
    check_btn_toggle_check: Any = None
    hist_btn_toggle_check: Any = None
    passage_btn_check: Any = None
    find_btn: Any = None
    find_btn_check: Any = None
    check_controls: Any = None
    hide_words_btn: Any = None
    main: Any = None
    voice_var: Any = None
    voice_combo: Any = None
    voice_map: dict = field(default_factory=dict)
    tts_status_request: int = 0
    word_table: Any = None
    word_table_scroll: Any = None
    word_context_menu: Any = None
    dictation_context_menu: Any = None
    history_context_menu: Any = None
    word_action_index: Any = None
    word_action_word: str = ""
    word_action_origin: str = "main"
    word_edit_entry: Any = None
    word_edit_row: Any = None
    word_edit_column: Any = None
    suppress_word_select_action: bool = False
    suppress_dictation_select_action: bool = False
    last_word_speak_index: Any = None
    last_word_speak_at: float = 0.0
    last_dictation_preview_key: Any = None
    last_dictation_preview_at: float = 0.0
    sentence_window: Any = None
    synonym_window: Any = None
    manual_words_window: Any = None
    manual_words_table: Any = None
    manual_words_table_scroll: Any = None
    manual_preview_edit_entry: Any = None
    manual_preview_edit_item: Any = None
    manual_preview_edit_column: Any = None
    sentence_generation_token: int = 0
    sentence_generation_active_token: int = 0
    sentence_event_queue: Any = field(default_factory=_queue.Queue)
    synonym_lookup_token: int = 0
    synonym_lookup_active_token: int = 0
    synonym_event_queue: Any = field(default_factory=_queue.Queue)
    translations: dict = field(default_factory=dict)
    word_pos: dict = field(default_factory=dict)
    word_phonetics: dict = field(default_factory=dict)
    word_metadata_event_queue: Any = field(default_factory=_queue.Queue)
    word_metadata_event_after: Any = None
    word_metadata_active_tasks: int = 0
    pending_translation_words: set = field(default_factory=set)
    pending_analysis_words: set = field(default_factory=set)
    pending_phonetic_words: set = field(default_factory=set)
    translation_token: int = 0
    analysis_token: int = 0
    phonetic_token: int = 0
    manual_source_dirty: bool = False
    ui_language_var: Any = None
    passage_window: Any = None
    passage_text: Any = None
    passage_status_var: Any = None
    gemini_model_var: Any = None
    gemini_model_combo: Any = None
    gemini_model_values: list = field(default_factory=list)
    llm_api_provider_var: Any = None
    tts_api_provider_var: Any = None
    gemini_verified: bool = False
    api_key_window: Any = None
    api_key_force_llm: bool = False
    api_key_force_tts: bool = False
    gemini_key_var: Any = None
    gemini_key_status_var: Any = None
    tts_key_var: Any = None
    tts_key_status_var: Any = None
    api_key_test_btn: Any = None
    api_llm_entry: Any = None
    api_tts_entry: Any = None
    gemini_runtime_status_var: Any = None
    gemini_retry_status_var: Any = None
    gemini_status_after: Any = None
    gemini_key_test_btn: Any = None
    tts_key_test_btn: Any = None
    gemini_validation_token: int = 0
    gemini_validation_active_token: int = 0
    gemini_validation_queue: Any = field(default_factory=_queue.Queue)
    current_passage: str = ""
    current_passage_original: str = ""
    current_passage_words: list = field(default_factory=list)
    passage_cloze_text: str = ""
    passage_answers: list = field(default_factory=list)
    passage_is_practice: bool = False
    passage_practice_input: Any = None
    passage_practice_result: Any = None
    passage_generation_token: int = 0
    passage_generation_active_token: int = 0
    passage_event_queue: Any = field(default_factory=_queue.Queue)
    find_window: Any = None
    find_search_var: Any = None
    find_limit_var: Any = None
    find_status_var: Any = None
    find_results_table: Any = None
    find_preview_text: Any = None
    find_docs_list: Any = None
    find_import_btn: Any = None
    find_docs_context_menu: Any = None
    find_doc_items: list = field(default_factory=list)
    find_result_items: dict = field(default_factory=dict)
    find_task_queue: Any = field(default_factory=_queue.Queue)
    find_task_token: int = 0
    find_active_token: int = 0
    audio_precache_token: int = 0
    dictation_mode_var: Any = None
    dictation_feedback_var: Any = None
    dictation_live_feedback_var: Any = None
    dictation_show_answer_var: Any = None
    dictation_show_note_var: Any = None
    dictation_show_phonetic_var: Any = None
    dictation_feedback_seconds_var: Any = None
    dictation_speed_var: Any = None
    dictation_order_var: Any = None
    dictation_setup_status_var: Any = None
    dictation_status_var: Any = None
    dictation_timer_var: Any = None
    dictation_progress_var: Any = None
    dictation_summary_var: Any = None
    dictation_recent_list: Any = None
    dictation_recent_items: list = field(default_factory=list)
    dictation_all_items: list = field(default_factory=list)
    dictation_list_mode_var: Any = None
    dictation_all_tab_var: Any = None
    dictation_recent_tab_var: Any = None
    dictation_setup_frame: Any = None
    dictation_session_frame: Any = None
    dictation_result_frame: Any = None
    dictation_mode_popup: Any = None
    dictation_mode_buttons: list = field(default_factory=list)
    dictation_order_buttons: list = field(default_factory=list)
    dictation_input: Any = None
    dictation_result_label: Any = None
    dictation_progress: Any = None
    dictation_timer_after: Any = None
    dictation_feedback_after: Any = None
    dictation_play_after: Any = None
    dictation_pool: list = field(default_factory=list)
    dictation_index: int = -1
    dictation_current_word: str = ""
    dictation_wrong_items: list = field(default_factory=list)
    dictation_session_attempts: list = field(default_factory=list)
    dictation_correct_count: int = 0
    dictation_answer_revealed: bool = False
    dictation_running: bool = False
    dictation_paused: bool = False
    dictation_seconds_left: int = 0
    dictation_session_source_path: Any = None
    dictation_session_list_mode: str = "all"
    dictation_previous_session_accuracy: float = 0.0
    dictation_answer_review_popup: Any = None
    dictation_answer_review_tree: Any = None
    dictation_answer_review_accuracy_var: Any = None
    dictation_answer_review_last_var: Any = None
    dictation_answer_review_filter_var: Any = None
    dictation_answer_review_show_wrong_only: bool = False
    dictation_result_review_tree: Any = None
    dictation_result_accuracy_var: Any = None
    dictation_result_last_var: Any = None
    dictation_result_filter_var: Any = None
    right_notebook: Any = None
    review_tab: Any = None
    check_tab: Any = None
    check_panel: Any = None
    history_tab: Any = None
    tools_tab: Any = None
    dictation_window: Any = None
    detail_word_var: Any = None
    detail_note_var: Any = None
    detail_translation_var: Any = None
    detail_meta_var: Any = None
    review_source_var: Any = None
    review_stats_var: Any = None
    review_focus_var: Any = None
    tools_hint_var: Any = None
    detail_speak_btn: Any = None
    detail_sentence_btn: Any = None
    detail_find_btn: Any = None
    review_open_source_btn: Any = None
    tools_sentence_btn: Any = None
    tools_passage_btn: Any = None
    tools_find_btn: Any = None
    tools_settings_btn: Any = None
    tools_update_btn: Any = None
    tools_export_cache_btn: Any = None
    tools_import_cache_btn: Any = None
    tools_export_resource_pack_btn: Any = None
    tools_import_resource_pack_btn: Any = None
    save_as_btn: Any = None
    new_list_btn: Any = None
    detail_edit_btn: Any = None
    detail_card: Any = None
    saved_window_geometry: str = ""
    saved_window_minsize: Any = None
    hidden_notebook_tabs: list = field(default_factory=list)

    @classmethod
    def create(
        cls,
        *,
        store,
        ui_language: str,
        generation_model: str,
        llm_api_key: str,
        tts_api_key: str,
        llm_provider_label: str,
        tts_provider_label: str,
        default_gemini_model: str,
    ) -> "MainViewState":
        return cls(
            order_mode=tk.StringVar(value="order"),
            interval_var=tk.DoubleVar(value=2.0),
            volume_var=tk.IntVar(value=80),
            dictation_volume_var=tk.IntVar(value=100),
            speech_rate_var=tk.DoubleVar(value=1.0),
            status_var=tk.StringVar(value="Not started"),
            voice_var=tk.StringVar(value=""),
            ui_language_var=tk.StringVar(value=ui_language),
            passage_status_var=tk.StringVar(value="Load words and click Generate."),
            gemini_model_var=tk.StringVar(value=generation_model or default_gemini_model),
            llm_api_provider_var=tk.StringVar(value=llm_provider_label),
            tts_api_provider_var=tk.StringVar(value=tts_provider_label),
            gemini_key_var=tk.StringVar(value=llm_api_key),
            gemini_key_status_var=tk.StringVar(value="Paste your LLM API key, then test it."),
            tts_key_var=tk.StringVar(value=tts_api_key),
            tts_key_status_var=tk.StringVar(value="Paste your TTS API key, then test it."),
            gemini_runtime_status_var=tk.StringVar(value=""),
            gemini_retry_status_var=tk.StringVar(value=""),
            find_search_var=tk.StringVar(value=""),
            find_limit_var=tk.StringVar(value="20"),
            find_status_var=tk.StringVar(value="Import docs, then search by word or phrase."),
            dictation_mode_var=tk.StringVar(value="online_spelling"),
            dictation_feedback_var=tk.StringVar(value="live"),
            dictation_live_feedback_var=tk.BooleanVar(value=True),
            dictation_show_answer_var=tk.BooleanVar(value=True),
            dictation_show_note_var=tk.BooleanVar(value=True),
            dictation_show_phonetic_var=tk.BooleanVar(value=False),
            dictation_feedback_seconds_var=tk.StringVar(value="2.2"),
            dictation_speed_var=tk.StringVar(value="1.0"),
            dictation_order_var=tk.StringVar(value="order"),
            dictation_setup_status_var=tk.StringVar(value="Recent mistake list"),
            dictation_status_var=tk.StringVar(value="Recent mistake list"),
            dictation_timer_var=tk.StringVar(value="5s"),
            dictation_progress_var=tk.StringVar(value="Spelling (0/0)"),
            dictation_summary_var=tk.StringVar(value=""),
            dictation_list_mode_var=tk.StringVar(value="recent"),
            dictation_all_tab_var=tk.StringVar(value="All (0)"),
            dictation_recent_tab_var=tk.StringVar(value="Recent Wrong (0)"),
            detail_word_var=tk.StringVar(value="No word selected"),
            detail_note_var=tk.StringVar(value="Select a word to see notes and translation."),
            detail_translation_var=tk.StringVar(value=""),
            detail_meta_var=tk.StringVar(value="Import a list to begin."),
            review_source_var=tk.StringVar(value="Source file: none"),
            review_stats_var=tk.StringVar(value="Words: 0"),
            review_focus_var=tk.StringVar(value="Focus: select a word or start playback."),
            tools_hint_var=tk.StringVar(value="Open tools from here instead of hunting across the window."),
            dictation_previous_session_accuracy=store.get_last_dictation_accuracy(),
        )


MAIN_VIEW_STATE_FIELDS = tuple(item.name for item in fields(MainViewState))


def bind_state_properties(cls):
    for name in MAIN_VIEW_STATE_FIELDS:
        if hasattr(cls, name):
            continue

        def _getter(self, attr=name):
            return getattr(self.state, attr)

        def _setter(self, value, attr=name):
            setattr(self.state, attr, value)

        setattr(cls, name, property(_getter, _setter))
