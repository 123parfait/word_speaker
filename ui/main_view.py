# -*- coding: utf-8 -*-
import ctypes
import html
import os
import queue
import random
import re
import time
import tkinter as tk
from html.parser import HTMLParser
from tkinter import ttk, filedialog, messagebox

from data.store import WordStore
from services.tts import (
    speak_async,
    cancel_all as tts_cancel_all,
    clear_word_backend_override as tts_clear_word_backend_override,
    get_recent_wrong_cache_source as tts_get_recent_wrong_cache_source,
    get_word_audio_cache_info as tts_get_word_audio_cache_info,
    precache_word_audio_async,
    prepare_async as tts_prepare_async,
    dedupe_pending_online_queue as tts_dedupe_pending_online_queue,
    set_preferred_pending_source as tts_set_preferred_pending_source,
    set_word_backend_override as tts_set_word_backend_override,
    get_online_tts_queue_status as tts_get_online_tts_queue_status,
    get_runtime_label as tts_get_runtime_label,
    has_cached_word_audio as tts_has_cached_word_audio,
    export_shared_audio_cache_package as tts_export_shared_audio_cache_package,
    import_shared_audio_cache_package as tts_import_shared_audio_cache_package,
    set_error_notifier as tts_set_error_notifier,
)
from services.translation import (
    get_cached_translations,
    prepare_async as translation_prepare_async,
    set_cached_translation,
)
from services.phonetics import get_cached_phonetics, set_cached_phonetic
from services.word_analysis import (
    get_cached_pos,
    set_cached_pos,
)
from services.resource_pack import (
    export_word_resource_pack,
    import_word_resource_pack,
)
from services.bundled_corpus import prepare_async as bundled_corpus_prepare_async
from services.official_library_sync import resolve_official_library_urls, sync_official_library
from services.app_config import (
    get_llm_api_key,
    get_llm_api_provider,
    get_generation_model,
    get_tts_api_key,
    get_tts_api_provider,
    get_ui_language,
    get_shared_cache_manifest_url,
    get_update_manifest_url,
    set_llm_api_key,
    set_llm_api_provider,
    set_generation_model,
    set_shared_cache_manifest_url,
    set_tts_api_key,
    set_tts_api_provider,
    set_ui_language,
    set_update_manifest_url,
)
from services.gemini_writer import (
    DEFAULT_GEMINI_MODEL,
    choose_preferred_generation_model,
    list_available_gemini_models,
)
from services.voice_catalog import kokoro_ready, list_system_voices, piper_ready
from services.voice_manager import (
    SOURCE_GEMINI,
    SOURCE_KOKORO,
    SOURCE_PIPER,
    get_voice_id,
    get_voice_source,
    set_voice_source,
)
from services.update_manager import (
    build_online_manifest,
    build_update_package,
    download_update_package,
    fetch_online_manifest,
    inspect_update_package,
    is_newer_version,
    is_packaged_runtime,
    launch_staged_update,
    load_local_version_info,
    load_version_info,
    stage_update_package,
)
from ui.detail_presenter import build_detail_view_state, build_recent_wrong_detail_view_state
from ui.detail_sidebar import build_detail_card, build_review_tab
from ui.diff_view_adapter import apply_diff
from ui.dictation_controller import DictationController
from ui.dictation_panel import build_dictation_answer_review_popup
from ui.dictation_session_coordinator import (
    advance_word as advance_dictation_word_flow,
    finalize_attempt as finalize_dictation_attempt_flow,
    finish_session as finish_dictation_session_flow,
    normalize_compare_text as normalize_dictation_compare_text,
    on_input_change as on_dictation_input_change_flow,
    pause_session as pause_dictation_session_flow,
    play_current_word as play_dictation_current_word_flow,
    previous_word as previous_dictation_word_flow,
    replay_current_word as replay_dictation_word_flow,
    reset_view as reset_dictation_view_flow,
    restart_timer as restart_dictation_timer_flow,
    start_session as start_online_spelling_session_flow,
    tick_timer as tick_dictation_timer_flow,
    toggle_play_pause as toggle_dictation_play_pause_flow,
)
from ui.dictation_window_coordinator import (
    close_mode_picker as close_dictation_mode_picker_flow,
    close_volume_popup as close_dictation_volume_popup_flow,
    close_window as close_dictation_window_flow,
    confirm_mode_picker as confirm_dictation_mode_picker_flow,
    get_pool as get_dictation_pool_flow,
    get_preview_source_path as get_dictation_preview_source_path_flow,
    get_source_items as get_dictation_source_items_flow,
    on_list_click_play as on_dictation_list_click_play_flow,
    on_review_tree_click as on_dictation_review_tree_click_flow,
    on_list_selected as on_dictation_list_selected_flow,
    open_mode_picker as open_dictation_mode_picker_flow,
    open_window as open_dictation_window_flow,
    refresh_recent_list as refresh_dictation_recent_list_flow,
    seconds_for_speed as dictation_seconds_for_speed_flow,
    set_feedback as set_dictation_feedback_flow,
    set_list_mode as set_dictation_list_mode_flow,
    set_mode as set_dictation_mode_flow,
    set_order as set_dictation_order_flow,
    set_speed as set_dictation_speed_flow,
    show_frame as show_dictation_frame_flow,
    speak_preview as speak_dictation_preview_flow,
    start_from_selected_word as start_dictation_from_selected_word_flow,
    toggle_volume_popup as toggle_dictation_volume_popup_flow,
    on_volume_change as on_dictation_volume_change_flow,
)
from ui.dictation_result_effects import (
    start_result_effect as start_dictation_result_effect_flow,
    stop_result_effect as stop_dictation_result_effect_flow,
)
from ui.async_event_helper import clear_event_queue, drain_event_queue, emit_event
from ui.api_key_async import (
    start_combined_api_validation_task,
    start_gemini_validation_task,
    start_tts_validation_task,
)
from ui.api_key_presenter import (
    build_combined_api_apply_state,
    build_combined_api_local_validation_state,
    build_combined_api_validation_request,
    build_single_api_error_state,
    build_single_llm_success_state,
    build_single_tts_success_state,
)
from ui.find_window_coordinator import (
    apply_import_result as apply_find_import_result_flow,
    apply_search_result as apply_find_search_result_flow,
    clear_document_filter as clear_find_document_filter_flow,
    clear_preview as clear_find_preview_flow,
    clear_task_queue as clear_find_task_queue_flow,
    delete_selected_document as delete_selected_corpus_document_flow,
    emit_task_event as emit_find_task_event_flow,
    get_selected_document as get_selected_find_document_flow,
    handle_task_error as handle_find_task_error_flow,
    import_documents as import_find_documents_flow,
    on_docs_right_click as on_find_docs_right_click_flow,
    on_result_select as on_find_result_select_flow,
    open_window as open_find_window_flow,
    poll_task_events as poll_find_task_events_flow,
    refresh_corpus_summary as refresh_find_corpus_summary_flow,
    run_search as run_find_search_flow,
    search_selected_word as search_selected_word_in_corpus_flow,
    set_query_from_selection as set_find_query_from_selection_flow,
    show_result_preview as show_find_result_preview_flow,
)
from ui.history_presenter import build_history_list_state, build_rename_history_target, get_selected_history_item
from ui.list_presenter import build_word_table_values
from ui.manual_words_presenter import (
    normalize_import_word_text,
    normalize_manual_input_rows,
    parse_manual_rows,
    parse_clipboard_html_rows,
    parse_tabular_text_rows,
    read_clipboard_import_rows,
)
from ui.main_playback_controller import MainPlaybackController
from ui.main_playback_host import (
    next_word as next_main_playback_word_flow,
    play_current as play_current_main_playback_flow,
    rebuild_on_mode_change as rebuild_main_playback_on_mode_change_flow,
    reset_state as reset_main_playback_state_flow,
    schedule_next as schedule_next_main_playback_flow,
    set_current_word as set_current_main_playback_word_flow,
    stop as stop_main_playback_flow,
    sync_state as sync_main_playback_state_flow,
    toggle_play as toggle_main_playback_flow,
    update_play_button as update_main_playback_button_flow,
)
from ui.manual_words_panel import build_manual_words_window
from ui.manual_words_editor import (
    add_manual_preview_row,
    append_manual_preview_rows,
    cancel_manual_preview_edit,
    clear_manual_preview,
    close_manual_words_window,
    collect_manual_rows_from_table,
    delete_selected_manual_preview_rows,
    finish_manual_preview_edit,
    start_manual_preview_edit,
)
from ui.passage_panel import build_passage_window
from ui.passage_presenter import build_passage_practice_check_state, build_passage_practice_state, normalize_answer, speech_text_from_passage
from ui.recent_wrong_controller import RecentWrongController
from ui.settings_presenter import parse_custom_interval
from ui.settings_host_coordinator import (
    clear_validation_queue as clear_gemini_validation_queue_flow,
    close_api_key_window as close_api_key_window_flow,
    close_settings_window as close_settings_window_flow,
    emit_validation_event as emit_gemini_validation_event_flow,
    maybe_close_api_key_window as maybe_close_api_key_window_flow,
    open_api_key_window as open_api_key_window_flow,
    open_settings_window as open_settings_window_flow,
    poll_validation_events as poll_gemini_validation_events_flow,
    refresh_settings_runtime_status as refresh_settings_runtime_status_flow,
    set_api_entry_error as set_api_entry_error_flow,
)
from ui.sidebar_panels import build_history_tab, build_tools_tab
from ui.word_list_panel import build_main_shell, build_word_list_panel
from ui.word_action_coordinator import (
    clear_context as clear_word_action_context_flow,
    dictation_row_to_store_index as dictation_row_to_store_index_flow,
    get_context_audio_source_path as get_context_audio_source_path_flow,
    get_context_or_selected_index as get_context_or_selected_index_flow,
    get_context_word as get_context_word_flow,
    get_word_audio_override_source_path as get_word_audio_override_source_path_flow,
    on_dictation_word_right_click as on_dictation_word_right_click_flow,
    on_word_right_click as on_word_right_click_flow,
    set_context as set_word_action_context_flow,
)
from ui.main_view_state import MainViewState, bind_state_properties
from ui.word_list_controller import WordListController
from ui.word_metadata_coordinator import (
    apply_phonetics as apply_phonetics_flow,
    apply_pos_analysis as apply_pos_analysis_flow,
    apply_single_phonetic as apply_single_phonetic_flow,
    apply_single_translation as apply_single_translation_flow,
    apply_translations as apply_translations_flow,
    ensure_word_metadata as ensure_word_metadata_flow,
    render_words as render_words_flow,
    start_analysis_job as start_analysis_job_flow,
    start_phonetic_job as start_phonetic_job_flow,
    start_single_phonetic as start_single_phonetic_flow,
    start_single_translation as start_single_translation_flow,
    start_translation_job as start_translation_job_flow,
)
from ui.word_tools_async import start_passage_generation_task
from ui.tool_host_coordinator import (
    apply_generated_passage as apply_generated_passage_flow,
    clear_passage_event_queue as clear_passage_event_queue_flow,
    clear_sentence_event_queue as clear_sentence_event_queue_flow,
    clear_synonym_event_queue as clear_synonym_event_queue_flow,
    emit_passage_event as emit_passage_event_flow,
    emit_sentence_event as emit_sentence_event_flow,
    emit_synonym_event as emit_synonym_event_flow,
    generate_passage as generate_ielts_passage_flow,
    lookup_synonyms_for_selected_word as lookup_synonyms_for_selected_word_flow,
    make_sentence_for_selected_word as make_sentence_for_selected_word_flow,
    open_passage_window as open_passage_window_flow,
    pause_word_playback_for_passage as pause_word_playback_for_passage_flow,
    play_generated_passage as play_generated_passage_flow,
    poll_passage_generation_events as poll_passage_generation_events_flow,
    poll_sentence_events as poll_sentence_events_flow,
    poll_synonym_events as poll_synonym_events_flow,
    show_sentence_window as show_sentence_window_flow,
    show_synonym_window as show_synonym_window_flow,
    stop_passage_playback as stop_passage_playback_flow,
    update_partial_passage as update_partial_passage_flow,
)
from ui.tts_status_bridge import watch_tts_backend as watch_tts_backend_flow


class _ClipboardTableHTMLParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.rows = []
        self._current_row = None
        self._current_cell = None
        self._table_depth = 0

    def handle_starttag(self, tag, attrs):
        tag_name = str(tag or "").strip().lower()
        if tag_name == "table":
            self._table_depth += 1
        if self._table_depth <= 0:
            return
        if tag_name == "tr":
            self._current_row = []
        elif tag_name in {"td", "th"}:
            self._current_cell = []
        elif tag_name == "br":
            if self._current_cell is not None:
                self._current_cell.append("\n")
        elif tag_name in {"p", "div", "li"}:
            if self._current_cell is not None and self._current_cell:
                self._current_cell.append("\n")

    def handle_endtag(self, tag):
        tag_name = str(tag or "").strip().lower()
        if self._table_depth <= 0 and tag_name != "table":
            return
        if tag_name in {"td", "th"} and self._current_row is not None and self._current_cell is not None:
            self._current_row.append(self._normalize_cell("".join(self._current_cell)))
            self._current_cell = None
        elif tag_name == "tr" and self._current_row is not None:
            if any(cell for cell in self._current_row):
                self.rows.append(list(self._current_row))
            self._current_row = None
        elif tag_name == "table" and self._table_depth > 0:
            self._table_depth -= 1

    def handle_data(self, data):
        if self._current_cell is None:
            return
        value = html.unescape(str(data or ""))
        if value:
            self._current_cell.append(value)

    @staticmethod
    def _normalize_cell(text):
        lines = []
        for raw_line in str(text or "").replace("\r", "\n").split("\n"):
            line = re.sub(r"[ \t]+", " ", raw_line).strip()
            if line:
                lines.append(line)
        return "\n".join(lines).strip()


UI_TEXTS = {
    "zh": {
        "word_list": "单词表",
        "word_list_desc": "导入词表、直接编辑，然后从当前选中单词开始学习。",
        "import": "导入",
        "paste_type": "粘贴 / 输入",
        "save_as": "另存为",
        "new_list": "新建词表",
        "word": "单词",
        "notes": "备注",
        "edit_note": "编辑备注",
        "edit_word": "编辑单词",
        "add_word": "添加单词",
        "add_word_title": "添加单词",
        "add_word_prompt": "输入要添加的单词：",
        "edit_word_title": "编辑单词",
        "edit_word_prompt": "修改这个单词：",
        "edit_note_title": "编辑备注",
        "edit_note_prompt": "修改这个单词的备注：",
        "error_type": "错因",
        "no_words": "还没有单词，先点导入开始。",
        "play": "▶ 播放",
        "settings": "设置",
        "dictation": "听写",
        "current_word": "当前单词",
        "speak_word": "朗读单词",
        "generate_sentence": "生成例句",
        "find_in_corpus": "语料检索",
        "edit_pos_translation": "编辑词性 / 中文",
        "synonyms": "近义词",
        "lookup_synonyms": "查询近义词",
        "synonyms_title": "近义词",
        "synonyms_error": "近义词错误",
        "synonyms_ready": "{word} 的近义词已准备好。",
        "no_synonyms_found": "没有找到合适的近义词。",
        "synonyms_source": "来源：{source}",
        "synonyms_source_gemini": "Gemini",
        "synonyms_source_local": "本地回退（spaCy + WordNet）",
        "synonyms_focus": "匹配词：{word}",
        "inspect_audio_cache": "查询音频缓存",
        "replace_audio_with_piper": "替换语音为 Piper",
        "clear_word_audio_override": "恢复默认语音",
        "word_audio_override_missing_source": "当前词表还没有源文件路径，先保存或从文件导入后再设置单词专属语音。",
        "word_audio_override_piper_unavailable": "本地 Piper 还没准备好。请先在 data/models/piper 下放入模型。",
        "word_audio_replaced_piper": "已将 {word} 在当前词表中的语音固定为 Piper。",
        "word_audio_override_cleared": "已恢复 {word} 在当前词表中的默认语音。",
        "delete_word": "删除单词",
        "delete_word_confirm": "要从当前词表里删除这个单词吗？\n\n{word}",
        "word_deleted": "已删除单词：{word}",
        "review": "复习",
        "history": "历史",
        "tools": "工具",
        "open_history": "打开历史",
        "delete_history": "从历史中删除",
        "rename_history_file": "重命名文件",
        "rename_history_prompt": "输入新的文件名：",
        "rename_history_invalid": "请输入有效的文件名，不要包含路径。",
        "rename_history_exists": "目标文件已存在，请换一个名字。",
        "rename_history_missing": "源文件不存在，无法重命名。",
        "rename_history_done": "已重命名为 {name}，并迁移 {count} 个缓存项，更新 {queued} 个等待任务。",
        "rename_history_failed": "重命名失败：{error}",
        "open_tools": "打开工具",
        "study_focus": "学习焦点",
        "no_history": "还没有历史记录。",
        "learning_tools": "学习工具",
        "find_corpus_sentences": "语料句子检索",
        "generate_ielts_passage": "生成 IELTS 篇章",
        "voice_model_settings": "音源 / 模型设置",
        "shared_cache_tools": "共享音频缓存",
        "shared_cache_tip": "把已生成的单词音频打包分享，或导入别人准备好的缓存包以节省 TTS 额度。",
        "export_shared_cache": "导出共享缓存",
        "import_shared_cache": "导入共享缓存",
        "sync_shared_cache": "更新词库",
        "build_shared_cache_manifest": "生成共享库清单",
        "release_checklist": "发布清单",
        "resource_pack_tools": "词表资源包",
        "resource_pack_tip": "导出当前词表、备注和人工校对过的词性 / 中文；导入时只会写入这些词条，不会碰整个 data 目录。",
        "export_resource_pack": "导出词表资源包",
        "import_resource_pack": "导入词表资源包",
        "resource_pack_type": "Word Speaker 词表资源包",
        "resource_pack_export_title": "导出词表资源包",
        "resource_pack_import_title": "导入词表资源包",
        "resource_pack_export_empty": "当前没有可导出的词表内容。",
        "resource_pack_export_done": "已导出 {count} 个词条到：\n{path}",
        "resource_pack_import_done": "已导入 {count} 个词条。\n备注 {notes} 条，中文 {translations} 条，词性 {pos} 条，音标 {phonetics} 条。",
        "resource_pack_export_failed": "导出词表资源包失败：{error}",
        "resource_pack_import_failed": "导入词表资源包失败：{error}",
        "shared_cache_package_type": "Word Speaker 缓存包",
        "shared_cache_export_title": "导出共享音频缓存",
        "shared_cache_import_title": "导入共享音频缓存",
        "shared_cache_export_empty": "当前还没有可导出的共享单词音频缓存。",
        "shared_cache_export_done": "已导出 {count} 个缓存音频到：\n{path}",
        "shared_cache_import_done": "导入完成。\n新增 {imported} 个，替换 {replaced} 个，跳过相同 {same} 个，跳过较旧 {older} 个。\n全局数据：中文 {translations} 条，词性 {pos} 条，音标 {phonetics} 条。",
        "shared_cache_import_errors": "导入过程中有 {count} 个问题：\n{detail}",
        "shared_cache_import_failed": "导入缓存包失败：{error}",
        "shared_cache_export_failed": "导出缓存包失败：{error}",
        "shared_cache_sync_title": "同步官方共享音频库",
        "shared_cache_sync_url_prompt": "输入官方词库更新清单地址（manifest.json）：",
        "shared_cache_sync_missing_url": "没有填写官方词库更新地址。",
        "shared_cache_sync_no_default_url": "当前没有配置官方词库更新地址。请先在 version.json 里配置 GitHub Release 地址。",
        "shared_cache_sync_checking": "正在检查官方词库更新……",
        "shared_cache_sync_downloading": "正在下载官方词库资源……",
        "shared_cache_sync_confirm": "发现官方词库版本 {version}。\n\n是否现在一键更新？\n\n会尝试同时：\n1. 同步共享词音\n2. 导入官方词表资源包\n3. 更新 bundled corpus 语料",
        "shared_cache_sync_done": "官方词库更新完成。\n版本 {version}\n共享词音：新增 {imported}，替换 {replaced}，跳过 {same}/{older}\n全局数据：中文 {translations} 条，词性 {pos} 条，音标 {phonetics} 条\n词表资源包：导入 {pack_count} 个词条\n语料库：导入 {corpus_imported} 个文件，共 {corpus_available} 个",
        "shared_cache_sync_status": "官方词库已更新：词音 +{imported}，全局数据 {translations}/{pos}/{phonetics}，词条 {pack_count}，语料 {corpus_imported}/{corpus_available}。",
        "shared_cache_sync_failed": "官方词库更新失败：{error}",
        "shared_cache_manifest_build_title": "生成共享音频库清单",
        "shared_cache_manifest_version_prompt": "输入这次共享音频库版本号：",
        "shared_cache_manifest_url_prompt": "输入这个共享缓存包的在线下载地址：",
        "shared_cache_manifest_notes_prompt": "可选：输入这次共享音频库说明：",
        "shared_cache_manifest_done": "已生成共享缓存包和在线清单：\n缓存包：{package_path}\n清单：{manifest_path}\n条目数：{count}",
        "release_checklist_title": "发布清单",
        "release_checklist_copy": "复制清单",
        "release_checklist_copied": "发布清单已复制到剪贴板。",
        "update_app": "更新程序",
        "update_title": "程序更新",
        "update_desc": "可以联网检查新版本，也可以直接导入本地更新包。更新时会尽量保留本机缓存和配置。",
        "update_current_version": "当前版本：{version}",
        "update_online": "在线更新",
        "update_offline": "导入更新包",
        "build_update_package": "生成更新包",
        "update_package_type": "Word Speaker 更新包",
        "update_online_url_prompt": "输入在线更新清单地址（manifest.json）：",
        "update_online_missing_url": "没有填写在线更新地址。",
        "update_online_no_default_url": "当前没有配置在线更新地址。请先在 version.json 里配置 GitHub Release manifest 地址。",
        "update_source_dir": "选择已打包的程序目录",
        "update_build_done": "已生成更新包：\n{path}\n\n版本：{version}\n文件数：{count}",
        "update_manifest_done": "已生成在线更新清单：\n{path}",
        "update_manifest_url_prompt": "输入这个更新包的在线下载地址：",
        "update_manifest_notes_prompt": "可选：输入这次更新说明：",
        "update_manifest_create": "是否顺手生成一个在线 manifest.json？",
        "update_not_packaged": "当前是源码运行环境。自更新仅在打包后的程序里可用。",
        "update_checking": "正在检查更新……",
        "update_downloading": "正在下载更新包……",
        "update_no_update": "当前已经是最新版本。",
        "update_online_available": "发现新版本 {version}。\n\n是否下载并安装？",
        "update_offline_confirm": "准备安装更新包版本 {version}。\n\n是否现在更新？",
        "update_reinstall_confirm": "更新包版本是 {version}，与当前版本相同或更旧。\n\n仍然要继续安装吗？",
        "update_started": "更新程序已经启动。当前窗口会关闭，随后会出现更新进度窗口，并在完成后自动打开新版本。",
        "update_failed": "启动更新失败：{error}",
        "tools_tip": "提示：先选中单词，再生成例句或做定向语料检索。",
        "settings_title": "设置",
        "ui_language": "界面语言",
        "ui_language_restart": "界面语言已保存。重开程序后会完整应用。",
        "language_zh": "中文",
        "language_en": "English",
        "source": "音源",
        "order": "顺序",
        "speed": "速度",
        "volume": "音量",
        "gemini_api_key": "大模型 API",
        "llm_api": "大模型 API",
        "tts_api": "TTS API",
        "api_setup": "API 设置",
        "api_provider": "提供方",
        "provider_gemini": "Gemini",
        "provider_elevenlabs": "ElevenLabs",
        "llm_api_setup": "大模型 API 设置",
        "llm_key_desc": "用于文章生成、例句生成等 AI 功能。当前实现支持 Gemini。",
        "tts_api_setup": "TTS API 设置",
        "tts_key_desc": "用于在线语音生成。当前实现支持 Gemini TTS 和 ElevenLabs。ElevenLabs 默认使用偏英式的标准发音。",
        "api_setup_desc": "在同一个窗口里配置大模型 API 和在线 TTS API。两项会分别测试并保存。",
        "tts_api_key_error": "TTS API Key 错误",
        "paste_tts_key_first": "请先粘贴 TTS API Key。",
        "tts_status_normal": "{provider} 正常",
        "tts_status_limited": "{provider} 限流中",
        "tts_status_error": "{provider} 错误",
        "tts_status_idle": "{provider} 空闲",
        "tts_status_retry_at": "下次请求时间：{time}",
        "tts_status_retry_at_in": "下次请求时间：{time}（约 {seconds} 秒后）",
        "tts_status_retry_none": "下次请求时间：-",
        "tts_status_queue": "等待队列：{count}",
        "tts_status_queue_processing": "队列处理中：{count}",
        "tts_status_queue_waiting": "等待重试：{count}",
        "order_desc": "选择顺序、随机不重复，或点击播放。",
        "stop_after_list": "播完整个列表后停止",
        "speed_desc": "间隔：单词与单词之间的时间。",
        "custom_seconds": "自定义（秒）：",
        "apply": "应用",
        "pronunciation_speed": "发音：每个单词的朗读速度。",
        "volume_desc": "调节播放输出音量。",
        "settings_toggle_source": "音源",
        "settings_toggle_order": "顺序",
        "settings_toggle_speed": "速度",
        "settings_toggle_volume": "音量",
        "start_from_word": "从某词开始",
        "start_learning": "开始学习",
        "back": "返回",
        "cancel": "取消",
        "dictation_volume": "听写音量",
        "dictation_volume_button": "🔊",
        "dictation_volume_level": "音量：{value}%",
        "dictation_volume_tip": "可单独放大听写音频，最高 600%。",
        "start_here": "从这里开始",
        "answer": "答案",
        "study_mode": "学习模式",
        "playback_speed": "播放倍速",
        "dictation_order": "出题顺序",
        "dictation_order_sequential": "顺序",
        "dictation_order_random": "乱序不重复",
        "feedback": "反馈",
        "dictation_feedback_display": "答错后显示",
        "dictation_feedback_show_answer": "显示答案",
        "dictation_feedback_show_note": "显示备注",
        "dictation_feedback_show_phonetic": "显示音标",
        "dictation_feedback_duration": "反馈时间",
        "seconds_short": "s",
        "next_word": "下个单词",
        "pause": "暂停",
        "replay": "重播",
        "current_session_accuracy": "本次正确率",
        "view_wrong_words": "查看错词",
        "back_to_list": "返回列表",
        "from_selected_word": "从指定单词开始听写",
        "part_of_speech": "词性",
        "chinese_translation": "中文翻译",
        "save": "保存",
        "clear": "清空",
        "close": "关闭",
        "paste_preview": "粘贴到预览表格",
        "paste_preview_desc": "可从 Google Docs / Sheets 粘贴两列表格。第 1 列为英文，第 2 列为备注。",
        "english": "英文",
        "paste_clipboard": "粘贴剪贴板",
        "add_row": "添加行",
        "delete_selected": "删除选中",
        "replace_list": "替换列表",
        "append": "追加",
        "find_desc": "导入 txt/docx/pdf，建立本地句子索引，再按单词或短语搜索。",
        "search": "搜索",
        "show": "显示",
        "results": "条结果",
        "use_selected_word": "用当前选中词",
        "import_docs": "导入文档",
        "sentence": "句子",
        "preview": "预览",
        "indexed_documents": "已索引文档",
        "clear_filter": "清除筛选",
        "passage_title": "IELTS 听力风格篇章",
        "passage_desc": "用 Gemini 生成篇章，并用当前在线 TTS 朗读。如果只想用部分单词，请先在主表里选中。",
        "generate": "生成",
        "read_with_gemini": "用在线 TTS 朗读",
        "stop": "停止",
        "practice": "练习",
        "check": "检查",
        "model": "模型：",
        "practice_tip": "练习：每行按顺序填写一个缺失单词或词组。",
        "gemini_api_setup": "大模型 API 设置",
        "gemini_key_desc": "粘贴你的大模型 API Key。程序会先测试，再启用 AI 功能。当前实现支持 Gemini。",
        "gemini_model_desc": "用于文章生成和例句生成的模型：",
        "test_and_save": "测试并保存",
        "exit": "退出",
        "read_sentence": "朗读句子",
        "show_word_list": "显示词表",
        "hide_word_list": "隐藏词表",
        "mode_settings": "学习模式",
        "mode_picker_title": "学习模式",
        "mode_quiz": "测一测",
        "mode_word_mode": "背词模式",
        "mode_answer_review": "听写对答案",
        "mode_online_spelling": "在线拼写",
        "mode_not_ready": "这个模式下一步再做，当前先完成在线拼写。",
        "confirm": "确定",
        "dictation_settings": "模式设置",
        "previous_word": "上一个",
        "play_pause": "播放 / 暂停",
        "no_live_feedback": "拼写中不反馈",
        "live_feedback": "拼写中反馈",
        "adaptive_speed": "自适应",
        "info": "提示",
        "error": "错误",
        "save_error": "保存失败",
        "wrong_words": "错词",
        "answer_review": "对答案",
        "answer_review_title": "对答案",
        "answer_review_so_far": "到目前为止正确率",
        "last_session_accuracy": "上次听写正确率",
        "session_attempts": "本轮作答",
        "your_answer": "你的输入",
        "wrong_times": "错过次数",
        "show_wrong_only": "查看错词",
        "show_all_answers": "查看全部",
        "no_answers_yet": "本轮还没有作答记录。",
        "blank_answer": "(空白)",
        "add_wrong_word": "手动添加错词",
        "wrong_word_added": "已把 {word} 加入错词表。",
        "delete_history_confirm": "要从 app 历史里删除这个文件记录，并清理对应缓存吗？\n\n{name}\n\n不会删除电脑上的原文件。",
        "history_deleted": "已从 app 历史中删除 {name}，并清理 {count} 个缓存文件。电脑上的原文件没有删除。",
        "delete_corpus_doc": "从语料库列表移除",
        "delete_corpus_doc_confirm": "要从 app 的语料库索引里移除这个文档吗？\n\n{name}\n\n不会删除电脑上的原文件。",
        "corpus_doc_deleted": "已从 app 语料库索引中移除 {name}。电脑上的原文件没有删除。",
        "enter_wrong_word": "输入要加入错词表的单词或词组：",
        "find_error": "检索错误",
        "find_setup_error": "检索初始化错误",
        "import_warning": "导入警告",
        "gemini_api_key_error": "大模型 API Key 错误",
        "audio_cache_info_title": "音频缓存信息",
        "generate_error": "生成错误",
        "sentence_error": "例句错误",
        "no_words_to_save": "没有可保存的单词。",
        "save_failed": "保存词表失败。\n{error}",
        "import_words_first": "请先导入单词。",
        "no_words_available": "当前没有可用单词。",
        "select_word_first": "请先选中一个单词。",
        "no_words_for_dictation": "当前没有可用于听写的单词。",
        "no_wrong_words_session": "这次没有错词。",
        "no_valid_words": "没有找到有效单词。",
        "word_cannot_be_empty": "单词不能为空。",
        "file_not_found_moved": "文件不存在或已被移动。",
        "enter_word_or_phrase": "请先输入要搜索的单词或短语。",
        "generate_passage_first": "请先生成篇章。",
        "no_keywords_for_practice": "这篇文章里没有适合做练习的关键词。",
        "click_practice_first": "请先点击 Practice。",
        "paste_gemini_key_first": "请先粘贴大模型 API Key。",
        "passage_empty": "篇章为空。",
        "valid_number_needed": "请输入有效数字（>= 0.2）。",
        "click_word_first_mode": "点击播放模式下，请先点一个单词。",
        "ui_created_blank_list": "已创建空白词表。可以用“粘贴 / 输入”或添加行开始。",
        "ui_saved_word_list": "词表已保存到 {name}。",
        "ui_precache_start": "正在后台为 {count} 个单词准备音频缓存……",
        "ui_precache_progress": "正在准备音频缓存…… {done}/{total}",
        "ui_precache_done": "音频预缓存完成：{detail}。",
        "audio_cache_missing": "这个单词当前还没有缓存音频。",
        "audio_cache_backend": "当前缓存来源：{backend}",
        "audio_cache_target": "目标音源：{backend}",
        "audio_cache_pending": "Gemini 替换状态：等待后台替换",
        "audio_cache_ready": "Gemini 替换状态：无需替换",
        "audio_cache_path": "缓存文件：{path}",
        "audio_cache_shared_reused": "当前文件缓存：来自全局共享复用",
        "audio_cache_meta_path": "元数据文件：{path}",
        "audio_cache_shared_path": "全局共享缓存：{path}",
        "audio_cache_shared_missing": "全局共享缓存：当前没有",
        "dictation_recent_hint": "优先复习近期错词。没有错词时会回退到当前词表。",
        "dictation_all_hint": "全部词表。可以从当前列表任意单词开始。",
        "dictation_all_start": "从全部词表中指定起点",
        "dictation_recent_start": "从近期错词中指定起点",
        "dictation_playing": "听写：正在播放“{word}”。",
        "dictation_paused": "已暂停。点击播放继续。",
        "dictation_keep_spelling": "继续拼写……",
        "dictation_wrong_live": "拼写中反馈：错误",
        "dictation_correct": "正确",
        "dictation_wrong_plain": "错误。",
        "dictation_wrong_answer": "错误。答案：{word}",
        "dictation_wrong_answer_line": "答案：{word}",
        "dictation_feedback_note": "备注：{note}",
        "dictation_feedback_phonetic": "音标：{phonetic}",
        "dictation_listen_type": "请听音并拼写单词。",
        "dictation_session_complete": "本轮结束。",
        "dictation_recent_title": "近期错词列表",
        "dictation_empty_recent": "近期还没有错词记录。",
        "dictation_empty_list": "当前词表为空。",
        "dictation_scope_all": "当前按钮将作用于“全部”词表。",
        "dictation_scope_recent": "当前按钮将作用于“近期错词”列表。",
    },
    "en": {
        "word_list": "Word List",
        "word_list_desc": "Import a list, edit inline, then study from the selected word.",
        "import": "Import",
        "paste_type": "Paste / Type",
        "save_as": "Save As",
        "new_list": "New List",
        "word": "Word",
        "notes": "Notes",
        "edit_note": "Edit Note",
        "edit_word": "Edit Word",
        "add_word": "Add Word",
        "add_word_title": "Add Word",
        "add_word_prompt": "Enter the word to add:",
        "edit_word_title": "Edit Word",
        "edit_word_prompt": "Change this word:",
        "edit_note_title": "Edit Note",
        "edit_note_prompt": "Change the note for this word:",
        "error_type": "Error Type",
        "no_words": "No words yet. Click the import button to get started.",
        "play": "▶ Play",
        "settings": "Settings",
        "dictation": "Dictation",
        "current_word": "Current Word",
        "speak_word": "Speak Word",
        "generate_sentence": "Generate Sentence",
        "find_in_corpus": "Find In Corpus",
        "edit_pos_translation": "Edit POS / 中文",
        "synonyms": "Synonyms",
        "lookup_synonyms": "Lookup Synonyms",
        "synonyms_title": "Synonyms",
        "synonyms_error": "Synonyms Error",
        "synonyms_ready": "Synonyms are ready for {word}.",
        "no_synonyms_found": "No suitable synonyms were found.",
        "synonyms_source": "Source: {source}",
        "synonyms_source_gemini": "Gemini",
        "synonyms_source_local": "Local fallback (spaCy + WordNet)",
        "synonyms_focus": "Matched token: {word}",
        "inspect_audio_cache": "Inspect Audio Cache",
        "replace_audio_with_piper": "Replace Audio With Piper",
        "clear_word_audio_override": "Restore Default Audio",
        "word_audio_override_missing_source": "This word list does not have a source file path yet. Save it or import it from a file first.",
        "word_audio_override_piper_unavailable": "Piper is not ready yet. Add a Piper model under data/models/piper first.",
        "word_audio_replaced_piper": "Pinned {word} to Piper for this word list.",
        "word_audio_override_cleared": "Restored the default audio backend for {word} in this word list.",
        "delete_word": "Delete Word",
        "delete_word_confirm": "Delete this word from the current list?\n\n{word}",
        "word_deleted": "Deleted word: {word}",
        "review": "Review",
        "history": "History",
        "tools": "Tools",
        "open_history": "Open History",
        "delete_history": "Remove From History",
        "rename_history_file": "Rename File",
        "rename_history_prompt": "Enter a new file name:",
        "rename_history_invalid": "Enter a valid file name without any path separators.",
        "rename_history_exists": "A file with that name already exists.",
        "rename_history_missing": "The source file no longer exists.",
        "rename_history_done": "Renamed to {name}. Migrated {count} cache items and updated {queued} pending tasks.",
        "rename_history_failed": "Rename failed: {error}",
        "open_tools": "Open Tools",
        "study_focus": "Study Focus",
        "no_history": "No history yet.",
        "learning_tools": "Learning Tools",
        "find_corpus_sentences": "Find Corpus Sentences",
        "generate_ielts_passage": "Generate IELTS Passage",
        "voice_model_settings": "Voice / Model Settings",
        "shared_cache_tools": "Shared Audio Cache",
        "shared_cache_tip": "Export reusable word-audio cache packs for sharing, or import one to save TTS quota.",
        "export_shared_cache": "Export Shared Cache",
        "import_shared_cache": "Import Shared Cache",
        "sync_shared_cache": "Update Word Library",
        "build_shared_cache_manifest": "Build Cache Manifest",
        "release_checklist": "Release Checklist",
        "resource_pack_tools": "Word Resource Packs",
        "resource_pack_tip": "Export the current list, notes, and manually corrected POS / translations. Import only touches those entries and never ships the whole data folder.",
        "export_resource_pack": "Export Resource Pack",
        "import_resource_pack": "Import Resource Pack",
        "resource_pack_type": "Word Speaker Resource Pack",
        "resource_pack_export_title": "Export Word Resource Pack",
        "resource_pack_import_title": "Import Word Resource Pack",
        "resource_pack_export_empty": "There is no word list content to export right now.",
        "resource_pack_export_done": "Exported {count} entries to:\n{path}",
        "resource_pack_import_done": "Imported {count} entries.\nNotes: {notes}, translations: {translations}, POS: {pos}, phonetics: {phonetics}.",
        "resource_pack_export_failed": "Failed to export the word resource pack: {error}",
        "resource_pack_import_failed": "Failed to import the word resource pack: {error}",
        "shared_cache_package_type": "Word Speaker Cache Pack",
        "shared_cache_export_title": "Export Shared Audio Cache",
        "shared_cache_import_title": "Import Shared Audio Cache",
        "shared_cache_export_empty": "There is no shared word-audio cache to export yet.",
        "shared_cache_export_done": "Exported {count} cached audio items to:\n{path}",
        "shared_cache_import_done": "Import complete.\nAdded {imported}, replaced {replaced}, skipped same {same}, skipped older {older}.\nGlobal data: translations {translations}, POS {pos}, phonetics {phonetics}.",
        "shared_cache_import_errors": "The import reported {count} issues:\n{detail}",
        "shared_cache_import_failed": "Failed to import the cache package: {error}",
        "shared_cache_export_failed": "Failed to export the cache package: {error}",
        "shared_cache_sync_title": "Sync Official Shared Audio Cache",
        "shared_cache_sync_url_prompt": "Enter the official word-library manifest URL (manifest.json):",
        "shared_cache_sync_missing_url": "No official word-library URL was provided.",
        "shared_cache_sync_no_default_url": "No official word-library URL is configured. Add the GitHub Release URL to version.json first.",
        "shared_cache_sync_checking": "Checking the official word-library update...",
        "shared_cache_sync_downloading": "Downloading official word-library resources...",
        "shared_cache_sync_confirm": "Official word-library version {version} is available.\n\nUpdate everything now?\n\nThis will try to:\n1. Sync shared audio\n2. Import the official word resource pack\n3. Refresh bundled corpus files",
        "shared_cache_sync_done": "Official word-library update complete.\nVersion {version}\nShared audio: +{imported}, replaced {replaced}, skipped same {same}, skipped older {older}\nGlobal data: translations {translations}, POS {pos}, phonetics {phonetics}\nWord pack: imported {pack_count} entries\nCorpus: imported {corpus_imported} files, {corpus_available} available",
        "shared_cache_sync_status": "Word library updated: audio +{imported}, global data {translations}/{pos}/{phonetics}, entries {pack_count}, corpus {corpus_imported}/{corpus_available}.",
        "shared_cache_sync_failed": "Failed to update the official word library: {error}",
        "shared_cache_manifest_build_title": "Build Shared Audio Cache Manifest",
        "shared_cache_manifest_version_prompt": "Enter the shared-audio release version:",
        "shared_cache_manifest_url_prompt": "Enter the online download URL for this shared-cache package:",
        "shared_cache_manifest_notes_prompt": "Optional: enter notes for this shared-audio release:",
        "shared_cache_manifest_done": "Built the shared-cache package and manifest:\nPackage: {package_path}\nManifest: {manifest_path}\nEntries: {count}",
        "release_checklist_title": "Release Checklist",
        "release_checklist_copy": "Copy Checklist",
        "release_checklist_copied": "The release checklist was copied to the clipboard.",
        "update_app": "Update App",
        "update_title": "App Update",
        "update_desc": "Check for a newer version online or import a local update package. The updater will try to preserve local cache and settings.",
        "update_current_version": "Current version: {version}",
        "update_online": "Online Update",
        "update_offline": "Import Update Package",
        "build_update_package": "Build Update Package",
        "update_package_type": "Word Speaker Update Package",
        "update_online_url_prompt": "Enter the online update manifest URL (manifest.json):",
        "update_online_missing_url": "No online update URL was provided.",
        "update_online_no_default_url": "No online update manifest URL is configured. Add the GitHub Release manifest URL to version.json first.",
        "update_source_dir": "Choose the packaged app folder",
        "update_build_done": "Built update package:\n{path}\n\nVersion: {version}\nFiles: {count}",
        "update_manifest_done": "Built online update manifest:\n{path}",
        "update_manifest_url_prompt": "Enter the online download URL for this update package:",
        "update_manifest_notes_prompt": "Optional: enter release notes for this update:",
        "update_manifest_create": "Also generate an online manifest.json now?",
        "update_not_packaged": "This is a source-code runtime. Self-update is only available in the packaged app.",
        "update_checking": "Checking for updates...",
        "update_downloading": "Downloading the update package...",
        "update_no_update": "This version is already up to date.",
        "update_online_available": "Version {version} is available.\n\nDownload and install it now?",
        "update_offline_confirm": "Ready to install update package version {version}.\n\nUpdate now?",
        "update_reinstall_confirm": "The package version is {version}, which is the same or older than the current version.\n\nInstall it anyway?",
        "update_started": "The updater has started. This window will close, an update progress window will appear, and the new version will launch when the update finishes.",
        "update_failed": "Failed to start the update: {error}",
        "tools_tip": "Tip: select a word first for sentence generation and targeted corpus search.",
        "settings_title": "Settings",
        "ui_language": "UI Language",
        "ui_language_restart": "UI language saved. Restart the app to fully apply it.",
        "language_zh": "Chinese",
        "language_en": "English",
        "source": "Source",
        "order": "Order",
        "speed": "Speed",
        "volume": "Volume",
        "gemini_api_key": "LLM API",
        "llm_api": "LLM API",
        "tts_api": "TTS API",
        "api_setup": "API Setup",
        "api_provider": "Provider",
        "provider_gemini": "Gemini",
        "provider_elevenlabs": "ElevenLabs",
        "llm_api_setup": "LLM API setup",
        "llm_key_desc": "Used for passage generation, sentence generation, and other AI features. Gemini is currently implemented.",
        "tts_api_setup": "TTS API setup",
        "tts_key_desc": "Used for online speech synthesis. Gemini TTS and ElevenLabs are currently implemented. ElevenLabs defaults to a British-style standard voice.",
        "api_setup_desc": "Configure the LLM API and online TTS API in one window. Each section is tested and saved independently.",
        "tts_api_key_error": "TTS API Key Error",
        "paste_tts_key_first": "Please paste a TTS API key first.",
        "tts_status_normal": "{provider} OK",
        "tts_status_limited": "{provider} Rate Limited",
        "tts_status_error": "{provider} Error",
        "tts_status_idle": "{provider} Idle",
        "tts_status_retry_at": "Next request: {time}",
        "tts_status_retry_at_in": "Next request: {time} (in about {seconds}s)",
        "tts_status_retry_none": "Next request: -",
        "tts_status_queue": "Queue: {count}",
        "tts_status_queue_processing": "Processing queue: {count}",
        "tts_status_queue_waiting": "Waiting to retry: {count}",
        "order_desc": "Choose in-order, random (no repeat), or click-to-play.",
        "stop_after_list": "Stop after list (no repeat list)",
        "speed_desc": "Interval: time between words.",
        "custom_seconds": "Custom (s):",
        "apply": "Apply",
        "pronunciation_speed": "Pronunciation: speaking speed for each word.",
        "volume_desc": "Adjust output volume for playback.",
        "settings_toggle_source": "Source",
        "settings_toggle_order": "Order",
        "settings_toggle_speed": "Speed",
        "settings_toggle_volume": "Volume",
        "start_from_word": "Start From Word",
        "start_learning": "Start Learning",
        "back": "Back",
        "cancel": "Cancel",
        "dictation_volume": "Dictation Volume",
        "dictation_volume_button": "🔊",
        "dictation_volume_level": "Volume: {value}%",
        "dictation_volume_tip": "Boost dictation playback only, up to 600%.",
        "start_here": "Start Here",
        "answer": "Answer",
        "study_mode": "Study Mode",
        "playback_speed": "Playback Speed",
        "dictation_order": "Question Order",
        "dictation_order_sequential": "Sequential",
        "dictation_order_random": "Random (No Repeat)",
        "feedback": "Feedback",
        "dictation_feedback_display": "Show after a wrong answer",
        "dictation_feedback_show_answer": "Show Answer",
        "dictation_feedback_show_note": "Show Note",
        "dictation_feedback_show_phonetic": "Show Phonetic",
        "dictation_feedback_duration": "Feedback Time",
        "seconds_short": "s",
        "next_word": "Next Word",
        "pause": "Pause",
        "replay": "Replay",
        "current_session_accuracy": "Current session accuracy",
        "view_wrong_words": "View Wrong Words",
        "back_to_list": "Back To List",
        "from_selected_word": "Start Dictation From Selected Word",
        "part_of_speech": "Part of speech",
        "chinese_translation": "Chinese translation",
        "save": "Save",
        "clear": "Clear",
        "close": "Close",
        "paste_preview": "Paste into preview table",
        "paste_preview_desc": "Paste a two-column table from Google Docs/Sheets. Column 1 = English, Column 2 = Notes.",
        "english": "English",
        "paste_clipboard": "Paste Clipboard",
        "add_row": "Add Row",
        "delete_selected": "Delete Selected",
        "replace_list": "Replace List",
        "append": "Append",
        "find_desc": "Import txt/docx/pdf files, build a local sentence index, then search by word or phrase.",
        "search": "Search",
        "show": "Show",
        "results": "results",
        "use_selected_word": "Use Selected Word",
        "import_docs": "Import Docs",
        "sentence": "Sentence",
        "preview": "Preview",
        "indexed_documents": "Indexed Documents",
        "clear_filter": "Clear Filter",
        "passage_title": "IELTS Listening Style Passage",
        "passage_desc": "Generate with the Gemini LLM and read it with the current online TTS. Select words in the main table first if you only want part of the list.",
        "generate": "Generate",
        "read_with_gemini": "Read with Online TTS",
        "stop": "Stop",
        "practice": "Practice",
        "check": "Check",
        "model": "Model:",
        "practice_tip": "Practice: fill one missing word/phrase per line (in order).",
        "gemini_api_setup": "LLM API setup",
        "gemini_key_desc": "Paste your LLM API key. The app will test it before enabling AI features. Gemini is currently implemented.",
        "gemini_model_desc": "Model used for article generation and sentence generation:",
        "test_and_save": "Test and Save",
        "exit": "Exit",
        "read_sentence": "Read Sentence",
        "show_word_list": "Show Word List",
        "hide_word_list": "Hide Word List",
        "mode_settings": "Study Mode",
        "mode_picker_title": "Study Mode",
        "mode_quiz": "Quiz",
        "mode_word_mode": "Word Mode",
        "mode_answer_review": "Answer Review",
        "mode_online_spelling": "Online Spelling",
        "mode_not_ready": "This mode is not implemented yet. Online Spelling is the current working mode.",
        "confirm": "Confirm",
        "dictation_settings": "Mode Settings",
        "previous_word": "Previous",
        "play_pause": "Play / Pause",
        "no_live_feedback": "No live feedback",
        "live_feedback": "Live feedback",
        "adaptive_speed": "Adaptive",
        "info": "Info",
        "error": "Error",
        "save_error": "Save Error",
        "wrong_words": "Wrong Words",
        "answer_review": "Answer Review",
        "answer_review_title": "Answer Review",
        "answer_review_so_far": "Accuracy so far",
        "last_session_accuracy": "Last session accuracy",
        "session_attempts": "Session Attempts",
        "your_answer": "Your Answer",
        "wrong_times": "Wrong Count",
        "show_wrong_only": "View Wrong Words",
        "show_all_answers": "View All",
        "no_answers_yet": "No answers yet in this session.",
        "blank_answer": "(blank)",
        "add_wrong_word": "Add Wrong Word",
        "wrong_word_added": "Added {word} to the wrong-word list.",
        "delete_history_confirm": "Remove this file from app history and clear its related cache?\n\n{name}\n\nThe original file on your computer will not be deleted.",
        "history_deleted": "Removed {name} from app history and cleared {count} cache files. The original file was kept on disk.",
        "delete_corpus_doc": "Remove From Corpus List",
        "delete_corpus_doc_confirm": "Remove this document from the app corpus index?\n\n{name}\n\nThe original file on your computer will not be deleted.",
        "corpus_doc_deleted": "Removed {name} from the app corpus index. The original file was kept on disk.",
        "enter_wrong_word": "Enter the word or phrase to add to the wrong-word list:",
        "find_error": "Find Error",
        "find_setup_error": "Find Setup Error",
        "import_warning": "Import Warning",
        "gemini_api_key_error": "LLM API Key Error",
        "audio_cache_info_title": "Audio Cache Info",
        "generate_error": "Generate Error",
        "sentence_error": "Sentence Error",
        "no_words_to_save": "No words to save.",
        "save_failed": "Failed to save word list.\n{error}",
        "import_words_first": "Please import words first.",
        "no_words_available": "No words available.",
        "select_word_first": "Please select a word first.",
        "no_words_for_dictation": "No words available for dictation.",
        "no_wrong_words_session": "No wrong words in this session.",
        "no_valid_words": "No valid words found.",
        "word_cannot_be_empty": "Word cannot be empty.",
        "file_not_found_moved": "File not found or moved.",
        "enter_word_or_phrase": "Enter a word or phrase first.",
        "generate_passage_first": "Generate a passage first.",
        "no_keywords_for_practice": "No suitable keywords found in this passage for practice.",
        "click_practice_first": "Click Practice first.",
        "paste_gemini_key_first": "Please paste an LLM API key first.",
        "passage_empty": "Passage is empty.",
        "valid_number_needed": "Please enter a valid number (>=0.2).",
        "click_word_first_mode": "Click a word first in Click-to-play mode.",
        "ui_created_blank_list": "Created a new blank list. Use Paste / Type or Add Row to start.",
        "ui_saved_word_list": "Saved word list to {name}.",
        "ui_precache_start": "Preparing audio cache for {count} words in the background...",
        "ui_precache_progress": "Preparing audio cache... {done}/{total}",
        "ui_precache_done": "Audio pre-cache complete: {detail}.",
        "audio_cache_missing": "There is no cached audio for this word yet.",
        "audio_cache_backend": "Current cached backend: {backend}",
        "audio_cache_target": "Target backend: {backend}",
        "audio_cache_pending": "Gemini replacement: queued for background replacement",
        "audio_cache_ready": "Gemini replacement: not needed",
        "audio_cache_path": "Cache file: {path}",
        "audio_cache_shared_reused": "Current file cache: reused from shared cache",
        "audio_cache_meta_path": "Metadata file: {path}",
        "audio_cache_shared_path": "Shared cache: {path}",
        "audio_cache_shared_missing": "Shared cache: none yet",
        "dictation_recent_hint": "Recent wrong words are reviewed first. If there are none, the current list is used.",
        "dictation_all_hint": "Full word list. You can start from any word in the current list.",
        "dictation_all_start": "Choose a start point from the full list",
        "dictation_recent_start": "Choose a start point from recent wrong words",
        "dictation_playing": "Dictation: playing '{word}'.",
        "dictation_paused": "Paused. Press Play to continue.",
        "dictation_keep_spelling": "Keep spelling...",
        "dictation_wrong_live": "Typing feedback: wrong",
        "dictation_correct": "Correct",
        "dictation_wrong_plain": "Wrong.",
        "dictation_wrong_answer": "Wrong. Answer: {word}",
        "dictation_wrong_answer_line": "Answer: {word}",
        "dictation_feedback_note": "Note: {note}",
        "dictation_feedback_phonetic": "Phonetic: {phonetic}",
        "dictation_listen_type": "Listen and type the word.",
        "dictation_session_complete": "Session complete.",
        "dictation_recent_title": "Recent mistake list",
        "dictation_empty_recent": "No recent wrong words yet.",
        "dictation_empty_list": "Current list is empty.",
        "dictation_scope_all": "These buttons currently apply to the full word list.",
        "dictation_scope_recent": "These buttons currently apply to the recent wrong-word list.",
    },
}


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
        self.dictation_controller = DictationController(self.store)
        self.recent_wrong_controller = RecentWrongController(self.store)
        self.word_list_controller = WordListController(self.store)
        self.main_playback_controller = MainPlaybackController()
        self.state = MainViewState.create(
            store=self.store,
            ui_language=get_ui_language(),
            generation_model=get_generation_model(),
            llm_api_key=get_llm_api_key(),
            tts_api_key=get_tts_api_key(),
            llm_provider_label="Gemini",
            tts_provider_label="ElevenLabs" if str(get_tts_api_provider() or "").strip().lower() == "elevenlabs" else "Gemini",
            default_gemini_model=DEFAULT_GEMINI_MODEL,
        )

        self.build_ui()
        try:
            self.winfo_toplevel().protocol("WM_DELETE_WINDOW", self.on_main_window_close)
            self.winfo_toplevel().bind("<space>", self.on_main_window_space_toggle, add="+")
        except Exception:
            pass
        self.refresh_history()
        self.update_empty_state()
        self.update_speed_buttons()
        self.update_speech_rate_buttons()
        self.update_play_button()
        self.update_right_visibility()
        tts_set_error_notifier(lambda message: messagebox.showerror("Speech Error", f"Error: {message}"))
        tts_prepare_async()
        translation_prepare_async()
        bundled_corpus_prepare_async()
        self.refresh_gemini_models()
        self.after(150, self.ensure_api_credentials)

    def tr(self, key):
        language = "en" if self.ui_language_var.get() == "en" else "zh"
        return UI_TEXTS.get(language, UI_TEXTS["zh"]).get(key, key)

    def trf(self, key, **kwargs):
        try:
            return self.tr(key).format(**kwargs)
        except Exception:
            return self.tr(key)

    def _tts_provider_options(self):
        return {
            self.tr("provider_gemini"): "gemini",
            self.tr("provider_elevenlabs"): "elevenlabs",
        }

    def _tts_provider_label(self, provider=None):
        provider_key = str(provider or get_tts_api_provider() or "gemini").strip().lower()
        return self.tr("provider_elevenlabs") if provider_key == "elevenlabs" else self.tr("provider_gemini")

    def _tts_provider_value(self):
        value = str(self.tts_api_provider_var.get() or "").strip()
        return self._tts_provider_options().get(value, "elevenlabs" if value.lower() == "elevenlabs" else "gemini")

    def _sync_provider_vars(self):
        self.llm_api_provider_var.set(self.tr("provider_gemini"))
        self.tts_api_provider_var.set(self._tts_provider_label(get_tts_api_provider()))

    def show_info(self, key, **kwargs):
        messagebox.showinfo(self.tr("info"), self.trf(key, **kwargs))

    def show_error(self, title_key, message_key=None, **kwargs):
        text = self.trf(message_key or title_key, **kwargs)
        messagebox.showerror(self.tr(title_key), text)

    def _space_shortcut_available(self, widget):
        if widget is None:
            return False
        blocked_classes = {
            "Entry",
            "TEntry",
            "Text",
            "Spinbox",
            "TCombobox",
            "Combobox",
            "Listbox",
            "Button",
            "TButton",
        }
        try:
            widget_class = str(widget.winfo_class() or "")
        except Exception:
            widget_class = ""
        if widget_class in blocked_classes:
            return False
        if self.word_edit_entry and widget == self.word_edit_entry:
            return False
        if self.dictation_window and self.dictation_window.winfo_exists():
            try:
                if widget.winfo_toplevel() == self.dictation_window:
                    return False
            except Exception:
                pass
        return True

    def on_main_window_space_toggle(self, event=None):
        widget = getattr(event, "widget", None)
        if not self._space_shortcut_available(widget):
            return None
        self.toggle_play()
        return "break"

    def on_ui_language_change(self, _event=None):
        new_language = "en" if self.ui_language_var.get() == "en" else "zh"
        set_ui_language(new_language)
        messagebox.showinfo(self.tr("ui_language"), self.tr("ui_language_restart"))

    def build_ui(self):
        build_main_shell(self)
        build_word_list_panel(self, Tooltip)

        # Settings popup (created on demand)
        self.settings_window = None

        self.right.grid_rowconfigure(1, weight=1)
        self.right.grid_columnconfigure(0, weight=1)

        build_detail_card(self)

        self.right_notebook = ttk.Notebook(self.right)
        self.right_notebook.grid(row=1, column=0, padx=12, pady=(0, 12), sticky="nsew")

        self.review_tab = ttk.Frame(self.right_notebook, style="Card.TFrame")
        self.history_tab = ttk.Frame(self.right_notebook, style="Card.TFrame")
        self.tools_tab = ttk.Frame(self.right_notebook, style="Card.TFrame")
        self.right_notebook.add(self.review_tab, text=self.tr("review"))
        self.right_notebook.add(self.history_tab, text=self.tr("history"))
        self.right_notebook.add(self.tools_tab, text=self.tr("tools"))

        build_review_tab(self)

        build_history_tab(self, Tooltip)
        build_tools_tab(self)

        self._refresh_selection_details()

    def _select_sidebar_tab(self, name):
        if not self.right_notebook:
            return
        mapping = {
            "review": self.review_tab,
            "history": self.history_tab,
            "tools": self.tools_tab,
        }
        target = mapping.get(str(name or "").strip().lower())
        if target is not None:
            self.right_notebook.select(target)

    def open_dictation_window(self):
        open_dictation_window_flow(self)

    def close_dictation_window(self):
        close_dictation_window_flow(self)

    def _has_unsaved_manual_words(self):
        return bool(self.manual_source_dirty and self.store.words and not self.store.has_current_source_file())

    def _discard_temporary_session_artifacts(self):
        if not self.store.is_temp_source_active():
            return
        self.word_list_controller.discard_temporary_session()
        self.manual_source_dirty = False

    def _mark_manual_words_dirty(self):
        self.manual_source_dirty = True
        self._refresh_selection_details()

    def _prompt_save_unsaved_manual_words(self, title="Save word list?"):
        if not self._has_unsaved_manual_words():
            return True
        answer = messagebox.askyesnocancel(
            title,
            "This pasted word list has not been saved yet.\nDo you want to save it first?"
            if self.ui_language_var.get() == "en"
            else "这份粘贴的词表还没有保存。\n要先保存吗？",
        )
        if answer is None:
            return False
        if answer:
            return self.save_words_as()
        self._discard_temporary_session_artifacts()
        return True

    def new_blank_list(self):
        if not self._prompt_save_unsaved_manual_words(title="Create new list?"):
            return
        self.cancel_word_edit()
        self.word_list_controller.create_blank_list()
        self.translations = {}
        self.word_pos = {}
        self.word_phonetics = {}
        self.manual_source_dirty = True
        self.render_words([])
        self.reset_playback_state()
        self.status_var.set(self.tr("ui_created_blank_list"))
        self._refresh_selection_details()

    def save_words_as(self):
        if not self.store.words:
            self.show_info("no_words_to_save")
            return False
        current_source = self.store.get_current_source_path()
        was_manual_session = not current_source
        words_snapshot = [str(word) for word in self.store.words]
        initial_name = (
            "word_list.csv"
            if self.store.is_temp_source_active()
            else (os.path.basename(current_source) if current_source else "word_list.csv")
        )
        default_ext = ".csv"
        if str(initial_name).lower().endswith(".txt"):
            default_ext = ".txt"
        path = filedialog.asksaveasfilename(
            title="Save word list as",
            defaultextension=default_ext,
            initialfile=initial_name,
            filetypes=[("CSV files", "*.csv"), ("Text files", "*.txt")],
        )
        if not path:
            return False
        try:
            self.word_list_controller.save_words_as(
                path,
                words_snapshot=words_snapshot,
                was_manual_session=was_manual_session,
            )
            self.refresh_history()
            self.manual_source_dirty = False
            self.status_var.set(self.trf("ui_saved_word_list", name=os.path.basename(path)))
            self._refresh_selection_details()
            return True
        except Exception as e:
            self.show_error("save_error", "save_failed", error=e)
            return False

    def on_main_window_close(self):
        if not self._prompt_save_unsaved_manual_words():
            return
        tts_set_preferred_pending_source(None)
        self._discard_temporary_session_artifacts()
        try:
            self.winfo_toplevel().destroy()
        except Exception:
            pass

    def _refresh_selection_details(self):
        total_words = len(self.store.words)
        selected_idx = self._get_selected_index()
        if selected_idx is None and self.current_word in self.store.words:
            selected_idx = self.store.words.index(self.current_word)
        context_word = self._get_context_word()
        context_idx = self._get_context_or_selected_index()
        if (
            self.word_action_origin == "dictation"
            and self.dictation_list_mode_var.get() == "recent"
            and context_word
            and (context_idx is None or context_idx < 0 or context_idx >= total_words)
        ):
            stats = self.store.get_dictation_word_stats(context_word)
            note = str(stats.get("note") or "").strip()
            zh = str(self.translations.get(context_word) or get_cached_translations([context_word]).get(context_word) or "").strip()
            pos_label = str(self.word_pos.get(context_word) or get_cached_pos([context_word]).get(context_word) or "").strip()
            phonetic = str(
                self.word_phonetics.get(context_word) or get_cached_phonetics([context_word]).get(context_word) or ""
            ).strip()
            state = build_recent_wrong_detail_view_state(
                context_word=context_word,
                wrong_count=int(stats.get("wrong_count", 0) or 0),
                note=note,
                translation=zh,
                pos_label=pos_label,
                phonetic=phonetic,
                current_source_path=self.store.get_display_source_path(),
                has_current_source_file=self.store.has_current_source_file(),
                has_unsaved_manual_words=self._has_unsaved_manual_words(),
            )
            self.detail_word_var.set(state.detail_word)
            self.detail_translation_var.set(state.detail_translation)
            self.detail_note_var.set(state.detail_note)
            self.detail_meta_var.set(state.detail_meta)
            self.review_focus_var.set(state.review_focus)
            self.review_source_var.set(state.review_source)
            if self.review_open_source_btn:
                self.review_open_source_btn.state(["!disabled"] if state.review_open_source_enabled else ["disabled"])
            return

        selected_word = None
        selected_note = ""
        selected_translation = ""
        selected_pos = ""
        selected_phonetic = ""
        if selected_idx is not None and selected_idx < total_words:
            word = self.store.words[selected_idx]
            selected_word = word
            selected_note = self.store.notes[selected_idx] if selected_idx < len(self.store.notes) else ""
            selected_translation = self.translations.get(word) or ""
            selected_pos = str(self.word_pos.get(word) or "").strip()
            selected_phonetic = str(self.word_phonetics.get(word) or "").strip()
        state = build_detail_view_state(
            total_words=total_words,
            selected_idx=selected_idx,
            current_word=self.current_word,
            selected_word=selected_word,
            selected_note=selected_note,
            selected_translation=selected_translation,
            selected_pos=selected_pos,
            selected_phonetic=selected_phonetic,
            current_source_path=self.store.get_display_source_path(),
            has_current_source_file=self.store.has_current_source_file(),
            has_unsaved_manual_words=self._has_unsaved_manual_words(),
            order_mode=self.order_mode.get(),
            play_state=self.play_state,
        )
        self.detail_word_var.set(state.detail_word)
        self.detail_translation_var.set(state.detail_translation)
        self.detail_note_var.set(state.detail_note)
        self.detail_meta_var.set(state.detail_meta)
        self.review_focus_var.set(state.review_focus)
        self.review_source_var.set(state.review_source)
        if state.review_stats is not None:
            self.review_stats_var.set(state.review_stats)
        if self.review_open_source_btn:
            self.review_open_source_btn.state(["!disabled"] if state.review_open_source_enabled else ["disabled"])
        selection_state = "normal" if state.has_selection else "disabled"
        for btn in (
            self.tools_sentence_btn,
            self.tools_find_btn,
        ):
            if btn:
                btn.config(state=selection_state)
        if self.tools_passage_btn:
            self.tools_passage_btn.config(state=("normal" if state.has_words else "disabled"))
        if self.save_as_btn:
            self.save_as_btn.config(state=("normal" if state.has_words else "disabled"))
        if self.new_list_btn:
            self.new_list_btn.config(state="normal")

    def _build_word_table_values(self, idx, word, note=None):
        note_value = self.store.notes[idx] if note is None and idx < len(self.store.notes) else (note or "")
        return build_word_table_values(
            idx,
            word,
            note=note_value,
            word_pos=self.word_pos,
            translations=self.translations,
            phonetics=self.word_phonetics,
        )

    def _start_analysis_job(self, words, token):
        start_analysis_job_flow(self, words, token)

    def _start_phonetic_job(self, words, token):
        start_phonetic_job_flow(self, words, token)

    def _apply_pos_analysis(self, token, requested_words, analyzed):
        apply_pos_analysis_flow(self, token, requested_words, analyzed)

    def _apply_phonetics(self, token, requested_words, phonetics):
        apply_phonetics_flow(self, token, requested_words, phonetics)

    def _start_audio_precache_job(self, words):
        if not words:
            return
        self.audio_precache_token += 1
        token = self.audio_precache_token
        source_path = self.store.get_current_source_path()
        tts_set_preferred_pending_source(source_path)
        total_words = len({str(word or "").strip().casefold() for word in words if str(word or "").strip()})
        self.status_var.set(self.trf("ui_precache_start", count=total_words))

        def _on_progress(done_count, total_count, _current_text):
            if token != self.audio_precache_token:
                return
            if done_count < total_count and done_count % 8 != 0:
                return
            self.after(
                0,
                lambda d=done_count, t=total_count: self._update_audio_precache_progress(token, d, t),
            )

        def _on_done(success_count, skipped_count, pending_count, error_count):
            self.after(
                0,
                lambda s=success_count, sk=skipped_count, p=pending_count, e=error_count: self._finish_audio_precache(
                    token, s, sk, p, e
                ),
            )

        precache_word_audio_async(
            words,
            source_path=source_path,
            rate_ratio=self.speech_rate_var.get(),
            on_progress=_on_progress,
            on_done=_on_done,
        )

    def _update_audio_precache_progress(self, token, done_count, total_count):
        if token != self.audio_precache_token:
            return
        self.status_var.set(self.trf("ui_precache_progress", done=done_count, total=total_count))

    def _finish_audio_precache(self, token, success_count, skipped_count, pending_count, error_count):
        if token != self.audio_precache_token:
            return
        parts = []
        if success_count:
            parts.append(f"{success_count} generated")
        if skipped_count:
            parts.append(f"{skipped_count} cached")
        if pending_count:
            parts.append(f"{pending_count} queued for Gemini")
        if error_count:
            parts.append(f"{error_count} failed")
        detail = ", ".join(parts) if parts else "nothing to do"
        self.status_var.set(self.trf("ui_precache_done", detail=detail))

    def refresh_dictation_recent_list(self):
        refresh_dictation_recent_list_flow(self)

    def _get_dictation_source_items(self):
        return get_dictation_source_items_flow(self)

    def set_dictation_list_mode(self, mode, refresh=True):
        set_dictation_list_mode_flow(self, mode, refresh=refresh)

    def _show_dictation_frame(self, target):
        show_dictation_frame_flow(self, target)

    def open_dictation_mode_picker(self, auto_start=True):
        open_dictation_mode_picker_flow(self, auto_start=auto_start)

    def close_dictation_mode_picker(self):
        close_dictation_mode_picker_flow(self)

    def set_dictation_mode(self, mode):
        set_dictation_mode_flow(self, mode)

    def confirm_dictation_mode_picker(self, auto_start=True):
        confirm_dictation_mode_picker_flow(self, auto_start=auto_start)

    def start_dictation_from_selected_word(self):
        start_dictation_from_selected_word_flow(self)

    def _get_dictation_pool(self):
        return get_dictation_pool_flow(self)

    def _get_recent_wrong_cache_source_path(self):
        return tts_get_recent_wrong_cache_source()

    def _get_dictation_preview_source_path(self):
        return get_dictation_preview_source_path_flow(self)

    def on_dictation_list_selected(self, _event=None):
        on_dictation_list_selected_flow(self, _event=_event)

    def on_dictation_list_click_play(self, event=None):
        return on_dictation_list_click_play_flow(self, event=event)

    def on_dictation_review_tree_click(self, event=None):
        return on_dictation_review_tree_click_flow(self, event=event)

    def _speak_dictation_preview(self, word=None, store_index=None):
        speak_dictation_preview_flow(self, word=word, store_index=store_index)

    def set_dictation_speed(self, value):
        set_dictation_speed_flow(self, value)

    def set_dictation_order(self, value):
        set_dictation_order_flow(self, value)

    def set_dictation_feedback(self, value):
        set_dictation_feedback_flow(self, value)

    def update_dictation_feedback_layout(self):
        from ui.dictation_window_coordinator import update_feedback_layout as update_dictation_feedback_layout_flow

        update_dictation_feedback_layout_flow(self)

    def _dictation_seconds_for_speed(self):
        return dictation_seconds_for_speed_flow(self)

    def _cancel_dictation_feedback_reset(self):
        if self.dictation_feedback_after:
            self.after_cancel(self.dictation_feedback_after)
            self.dictation_feedback_after = None

    def _cancel_dictation_play_start(self):
        if self.dictation_play_after:
            self.after_cancel(self.dictation_play_after)
            self.dictation_play_after = None

    def _cancel_dictation_timer(self):
        if self.dictation_timer_after:
            self.after_cancel(self.dictation_timer_after)
            self.dictation_timer_after = None

    def _set_dictation_input_color(self, mode="neutral"):
        if not self.dictation_input:
            return
        if mode == "correct":
            self.dictation_input.config(bg="#e8fff0", highlightbackground="#22a06b", highlightcolor="#22a06b")
        elif mode == "wrong":
            self.dictation_input.config(bg="#fff0f0", highlightbackground="#ef4444", highlightcolor="#ef4444")
        else:
            self.dictation_input.config(bg="#f6f6f8", highlightbackground="#d9dbe1", highlightcolor="#8a8f98")

    def _dictation_playback_volume_ratio(self):
        return max(0.0, float(self.dictation_volume_var.get()) / 100.0)

    def start_online_spelling_session(self, start_index=0):
        start_online_spelling_session_flow(self, start_index=start_index)

    def play_dictation_current_word(self):
        play_dictation_current_word_flow(self)

    def replay_dictation_word(self):
        replay_dictation_word_flow(self)

    def toggle_dictation_play_pause(self):
        toggle_dictation_play_pause_flow(self)

    def pause_dictation_session(self):
        pause_dictation_session_flow(self)

    def update_dictation_play_button(self):
        if not self.play_btn_check:
            return
        if self.dictation_running and not self.dictation_paused:
            self.play_btn_check.config(text=f"⏸ {self.tr('pause')}")
        else:
            self.play_btn_check.config(text=self.tr("play"))

    def previous_dictation_word(self):
        previous_dictation_word_flow(self)

    def _focus_dictation_input(self):
        if not self.dictation_input:
            return
        try:
            self.dictation_input.focus_force()
            self.dictation_input.icursor(tk.END)
        except Exception:
            try:
                self.dictation_input.focus_set()
                self.dictation_input.icursor(tk.END)
            except Exception:
                pass

    def _restart_dictation_timer(self):
        restart_dictation_timer_flow(self)

    def _tick_dictation_timer(self):
        tick_dictation_timer_flow(self)

    def on_dictation_input_change(self, _event=None):
        on_dictation_input_change_flow(self)

    def on_dictation_enter(self, _event=None):
        if not self.dictation_running or not self.dictation_current_word:
            return "break"
        if self.dictation_feedback_after or self.dictation_answer_revealed:
            return "break"
        self.finalize_dictation_attempt(trigger="manual")
        return "break"

    def _normalize_dictation_compare_text(self, text):
        return normalize_dictation_compare_text(text)

    def finalize_dictation_attempt(self, trigger="manual"):
        finalize_dictation_attempt_flow(self, trigger=trigger)

    def _go_to_next_dictation_word(self):
        self.dictation_feedback_after = None
        self.advance_dictation_word()

    def advance_dictation_word(self, initial=False):
        advance_dictation_word_flow(self, initial=initial)

    def finish_dictation_session(self):
        finish_dictation_session_flow(self)

    def reset_dictation_view(self):
        reset_dictation_view_flow(self)

    def _dictation_accuracy_so_far(self):
        return self.dictation_controller.accuracy_so_far(self.dictation_session_attempts)

    def _dictation_review_rows(self):
        return self.dictation_controller.build_review_rows(
            self.dictation_session_attempts,
            translations=self.translations,
            word_pos=self.word_pos,
            blank_answer_label=self.tr("blank_answer"),
            wrong_only=self.dictation_answer_review_show_wrong_only,
        )

    def close_dictation_answer_review_popup(self):
        if self.dictation_answer_review_popup and self.dictation_answer_review_popup.winfo_exists():
            self.dictation_answer_review_popup.destroy()
        self.dictation_answer_review_popup = None
        self.dictation_answer_review_tree = None

    def _render_dictation_answer_review_tree(self, tree):
        if not tree:
            return
        rows = self._dictation_review_rows()
        tree.delete(*tree.get_children())
        if not rows:
            tree.insert("", tk.END, iid="empty", values=("", self.tr("no_answers_yet"), ""))
            return
        for idx, row in enumerate(rows, start=1):
            tag_name = "correct" if row.get("correct") else "wrong"
            tree.insert(
                "",
                tk.END,
                iid=f"attempt_{idx}",
                values=(
                    f"{idx}. {row['word']}\n{row['subtitle']}" if row.get("subtitle") else f"{idx}. {row['word']}",
                    row["input"],
                    row["wrong_count"],
                ),
                tags=(tag_name,),
            )

    def _render_dictation_answer_review_views(self):
        accuracy = self._dictation_accuracy_so_far()
        previous_accuracy = self.dictation_previous_session_accuracy
        filter_text = self.tr("show_all_answers") if self.dictation_answer_review_show_wrong_only else self.tr("show_wrong_only")

        if self.dictation_answer_review_accuracy_var is not None:
            self.dictation_answer_review_accuracy_var.set(f"{accuracy:.2f}%")
        if self.dictation_result_accuracy_var is not None:
            self.dictation_result_accuracy_var.set(f"{accuracy:.2f}%")

        last_text = "-" if previous_accuracy is None else f"{float(previous_accuracy):.2f}%"
        if self.dictation_answer_review_last_var is not None:
            self.dictation_answer_review_last_var.set(last_text)
        if self.dictation_result_last_var is not None:
            self.dictation_result_last_var.set(last_text)

        if self.dictation_answer_review_filter_var is not None:
            self.dictation_answer_review_filter_var.set(filter_text)
        if self.dictation_result_filter_var is not None:
            self.dictation_result_filter_var.set(filter_text)

        self._render_dictation_answer_review_tree(self.dictation_answer_review_tree)
        self._render_dictation_answer_review_tree(self.dictation_result_review_tree)

        if self.dictation_answer_review_popup and self.dictation_answer_review_popup.winfo_exists():
            self.dictation_answer_review_popup.update_idletasks()
        if self.dictation_result_frame and self.dictation_result_frame.winfo_exists():
            self.dictation_result_frame.update_idletasks()

    def _refresh_dictation_answer_review_popup(self):
        self._render_dictation_answer_review_views()

    def _toggle_dictation_answer_review_filter(self):
        self.dictation_answer_review_show_wrong_only = not self.dictation_answer_review_show_wrong_only
        self._render_dictation_answer_review_views()

    def start_dictation_result_effect(self, accuracy):
        start_dictation_result_effect_flow(self, accuracy)

    def stop_dictation_result_effect(self):
        stop_dictation_result_effect_flow(self)

    def _return_from_dictation_answer_review(self):
        self.close_dictation_answer_review_popup()
        self.reset_dictation_view()

    def open_dictation_answer_review_popup(self):
        if self.dictation_answer_review_popup and self.dictation_answer_review_popup.winfo_exists():
            self._render_dictation_answer_review_views()
            self.dictation_answer_review_popup.deiconify()
            self.dictation_answer_review_popup.lift()
            return
        build_dictation_answer_review_popup(self)
        self._render_dictation_answer_review_views()

    def show_dictation_wrong_words(self):
        if not self.dictation_wrong_items:
            self.show_info("no_wrong_words_session")
            return
        lines = []
        for item in self.dictation_wrong_items:
            lines.append(f"{item['word']}    <-    {item.get('input') or '(blank)'}")
        messagebox.showinfo(self.tr("wrong_words"), "\n".join(lines[:40]))

    def inspect_selected_word_audio_cache(self):
        word = self._get_context_word()
        if not word:
            self.show_info("select_word_first")
            return
        source_path = self._get_context_audio_source_path()
        info = tts_get_word_audio_cache_info(word, source_path=source_path)
        if not info.get("exists") and not info.get("pending_gemini_replacement"):
            messagebox.showinfo(
                self.tr("audio_cache_info_title"),
                self.tr("audio_cache_missing"),
            )
            return

        backend_label = info.get("backend_label") or (info.get("backend") or "Unknown")
        desired_label = info.get("desired_backend_label") or backend_label
        lines = [
            f"Word: {word}",
            self.trf("audio_cache_backend", backend=backend_label),
            self.trf("audio_cache_target", backend=desired_label),
            self.tr("audio_cache_pending") if info.get("pending_gemini_replacement") else self.tr("audio_cache_ready"),
        ]
        if info.get("exists") or info.get("cache_path"):
            lines.append(self.trf("audio_cache_path", path=info.get("cache_path") or ""))
        if info.get("uses_shared_cache"):
            lines.append(self.tr("audio_cache_shared_reused"))
        shared_path = str(info.get("shared_cache_path") or "").strip()
        if info.get("shared_exists") and shared_path:
            lines.append(self.trf("audio_cache_shared_path", path=shared_path))
        else:
            lines.append(self.tr("audio_cache_shared_missing"))
        meta_path = str(info.get("meta_path") or "").strip()
        if meta_path:
            lines.append(self.trf("audio_cache_meta_path", path=meta_path))
        messagebox.showinfo(self.tr("audio_cache_info_title"), "\n".join(lines))

    def replace_selected_word_audio_with_piper(self):
        word = self._get_context_word()
        if not word:
            self.show_info("select_word_first")
            return
        source_path = self._get_word_audio_override_source_path()
        if not source_path:
            self.show_info("word_audio_override_missing_source")
            return
        if not piper_ready():
            self.show_info("word_audio_override_piper_unavailable")
            return
        if not tts_set_word_backend_override(word, source_path=source_path, backend=SOURCE_PIPER):
            self.show_info("word_audio_override_missing_source")
            return
        self.status_var.set(self.trf("word_audio_replaced_piper", word=word))

    def clear_selected_word_audio_override(self):
        word = self._get_context_word()
        if not word:
            self.show_info("select_word_first")
            return
        source_path = self._get_word_audio_override_source_path()
        if not source_path:
            self.show_info("word_audio_override_missing_source")
            return
        tts_clear_word_backend_override(word, source_path=source_path)
        self.status_var.set(self.trf("word_audio_override_cleared", word=word))

    def _format_size_label(self, size_bytes):
        try:
            value = float(size_bytes or 0)
        except Exception:
            value = 0.0
        units = ["B", "KB", "MB", "GB"]
        unit = units[0]
        for unit in units:
            if value < 1024 or unit == units[-1]:
                break
            value /= 1024.0
        if unit == "B":
            return f"{int(value)} {unit}"
        return f"{value:.1f} {unit}"

    def _collect_word_resource_entries(self):
        entries = []
        cached_translations = get_cached_translations(self.store.words)
        cached_pos = get_cached_pos(self.store.words)
        cached_phonetics = get_cached_phonetics(self.store.words)
        for idx, word in enumerate(self.store.words):
            token = str(word or "").strip()
            if not token:
                continue
            note = str(self.store.notes[idx] if idx < len(self.store.notes) else "").strip()
            translation = str(self.translations.get(token) or cached_translations.get(token) or "").strip()
            pos_label = str(self.word_pos.get(token) or cached_pos.get(token) or "").strip()
            phonetic = str(self.word_phonetics.get(token) or cached_phonetics.get(token) or "").strip()
            entries.append(
                {
                    "word": token,
                    "note": note,
                    "translation": translation,
                    "pos": pos_label,
                    "phonetic": phonetic,
                }
            )
        return entries

    def _load_word_resource_pack_entries(self, entries):
        rows = []
        translation_count = 0
        pos_count = 0
        phonetic_count = 0
        for item in entries or []:
            if not isinstance(item, dict):
                continue
            word = str(item.get("word") or "").strip()
            if not word:
                continue
            note = str(item.get("note") or "").strip()
            translation = str(item.get("translation") or "").strip()
            pos_label = str(item.get("pos") or "").strip()
            phonetic = str(item.get("phonetic") or "").strip()
            rows.append({"word": word, "note": note})
            if translation:
                set_cached_translation(word, translation)
                translation_count += 1
            if pos_label:
                set_cached_pos(word, pos_label)
                pos_count += 1
            if phonetic:
                set_cached_phonetic(word, phonetic)
                phonetic_count += 1
        if not rows:
            return None
        self.cancel_word_edit()
        words = [str(row.get("word") or "").strip() for row in rows]
        notes = [str(row.get("note") or "").strip() for row in rows]
        self.store.set_words(words, notes, preserve_source=False)
        self.manual_source_dirty = False
        self.render_words(words)
        self.reset_playback_state()
        return {
            "count": len(rows),
            "note_count": sum(1 for row in rows if str(row.get("note") or "").strip()),
            "translation_count": translation_count,
            "pos_count": pos_count,
            "phonetic_count": phonetic_count,
        }

    def export_word_resource_pack_tool(self):
        if not self.store.words:
            messagebox.showinfo(
                self.tr("resource_pack_export_title"),
                self.tr("resource_pack_export_empty"),
            )
            return
        path = filedialog.asksaveasfilename(
            title=self.tr("resource_pack_export_title"),
            defaultextension=".wspack",
            filetypes=[(self.tr("resource_pack_type"), "*.wspack")],
            initialfile="wordspeaker_word_resource_pack.wspack",
        )
        if not path:
            return
        metadata = {}
        current_source = str(self.store.get_current_source_path() or "").strip()
        if current_source:
            metadata["source_name"] = os.path.basename(current_source)
        try:
            result = export_word_resource_pack(path, self._collect_word_resource_entries(), metadata=metadata)
        except Exception as exc:
            messagebox.showerror(
                self.tr("resource_pack_export_title"),
                self.trf("resource_pack_export_failed", error=str(exc)),
            )
            return
        self.status_var.set(f"Word resource pack exported: {int(result.get('entry_count') or 0)} entries.")
        messagebox.showinfo(
            self.tr("resource_pack_export_title"),
            self.trf("resource_pack_export_done", count=int(result.get("entry_count") or 0), path=path),
        )

    def import_word_resource_pack_tool(self, package_path=None):
        if package_path is None:
            path = filedialog.askopenfilename(
                title=self.tr("resource_pack_import_title"),
                filetypes=[(self.tr("resource_pack_type"), "*.wspack")],
            )
        else:
            path = package_path
        if not path:
            return False
        if not self._prompt_save_unsaved_manual_words(title=self.tr("resource_pack_import_title")):
            return False
        try:
            result = import_word_resource_pack(path)
        except Exception as exc:
            messagebox.showerror(
                self.tr("resource_pack_import_title"),
                self.trf("resource_pack_import_failed", error=str(exc)),
            )
            return False
        load_result = self._load_word_resource_pack_entries(result.get("entries") or [])
        if not load_result:
            messagebox.showinfo(
                self.tr("resource_pack_import_title"),
                self.tr("no_valid_words"),
            )
            return False
        self.status_var.set(f"Word resource pack imported: {int(load_result.get('count') or 0)} entries.")
        messagebox.showinfo(
            self.tr("resource_pack_import_title"),
            self.trf(
                "resource_pack_import_done",
                count=int(load_result.get("count") or 0),
                notes=int(load_result.get("note_count") or 0),
                translations=int(load_result.get("translation_count") or 0),
                pos=int(load_result.get("pos_count") or 0),
                phonetics=int(load_result.get("phonetic_count") or 0),
            ),
        )
        return True

    def export_shared_cache_package(self):
        path = filedialog.asksaveasfilename(
            title=self.tr("shared_cache_export_title"),
            defaultextension=".zip",
            filetypes=[(self.tr("shared_cache_package_type"), "*.zip")],
            initialfile="wordspeaker_shared_audio_cache.zip",
        )
        if not path:
            return
        try:
            result = tts_export_shared_audio_cache_package(path)
        except Exception as exc:
            messagebox.showerror(
                self.tr("shared_cache_export_title"),
                self.trf("shared_cache_export_failed", error=str(exc)),
            )
            return
        if not result.get("ok"):
            messagebox.showinfo(
                self.tr("shared_cache_export_title"),
                self.tr("shared_cache_export_empty"),
            )
            return
        count = int(result.get("entries") or 0)
        size_label = self._format_size_label(result.get("bytes") or 0)
        self.status_var.set(f"Shared audio cache exported: {count} items, {size_label}.")
        messagebox.showinfo(
            self.tr("shared_cache_export_title"),
            self.trf("shared_cache_export_done", count=count, path=path),
        )

    def import_shared_cache_package(self):
        path = filedialog.askopenfilename(
            title=self.tr("shared_cache_import_title"),
            filetypes=[(self.tr("shared_cache_package_type"), "*.zip")],
        )
        if not path:
            return
        try:
            result = tts_import_shared_audio_cache_package(path)
        except Exception as exc:
            messagebox.showerror(
                self.tr("shared_cache_import_title"),
                self.trf("shared_cache_import_failed", error=str(exc)),
            )
            return
        imported = int(result.get("imported") or 0)
        replaced = int(result.get("replaced") or 0)
        skipped_same = int(result.get("skipped_same") or 0)
        skipped_older = int(result.get("skipped_older") or 0)
        metadata_translations = int(result.get("metadata_translations") or 0)
        metadata_pos = int(result.get("metadata_pos") or 0)
        metadata_phonetics = int(result.get("metadata_phonetics") or 0)
        if self.store.words:
            self.render_words(list(self.store.words))
        self.status_var.set(
            "Shared audio cache imported: "
            f"+{imported}, replaced {replaced}, skipped {skipped_same + skipped_older}, "
            f"metadata T/P/Pn {metadata_translations}/{metadata_pos}/{metadata_phonetics}."
        )
        messagebox.showinfo(
            self.tr("shared_cache_import_title"),
            self.trf(
                "shared_cache_import_done",
                imported=imported,
                replaced=replaced,
                same=skipped_same,
                older=skipped_older,
                translations=metadata_translations,
                pos=metadata_pos,
                phonetics=metadata_phonetics,
            ),
        )
        errors = list(result.get("errors") or [])
        if errors:
            detail = "\n".join(errors[:6])
            if len(errors) > 6:
                detail = f"{detail}\n..."
            messagebox.showwarning(
                self.tr("shared_cache_import_title"),
                self.trf("shared_cache_import_errors", count=len(errors), detail=detail),
            )

    def sync_shared_cache_package_online(self):
        current_info = load_local_version_info()
        manifest_url = (
            get_shared_cache_manifest_url()
            or str(current_info.get("shared_cache_manifest_url") or "").strip()
        )
        if not manifest_url:
            messagebox.showinfo(self.tr("shared_cache_sync_title"), self.tr("shared_cache_sync_no_default_url"))
            return
        set_shared_cache_manifest_url(manifest_url)
        self.status_var.set(self.tr("shared_cache_sync_checking"))
        try:
            manifest = fetch_online_manifest(manifest_url)
        except Exception as exc:
            messagebox.showerror(
                self.tr("shared_cache_sync_title"),
                self.trf("shared_cache_sync_failed", error=str(exc)),
            )
            return
        remote_version = str(manifest.get("version") or "unknown").strip() or "unknown"
        confirm_text = self.trf("shared_cache_sync_confirm", version=remote_version)
        notes = str(manifest.get("notes") or "").strip()
        if notes:
            confirm_text = f"{confirm_text}\n\n{notes}"
        if not messagebox.askyesno(self.tr("shared_cache_sync_title"), confirm_text):
            return
        if not self._prompt_save_unsaved_manual_words(title=self.tr("resource_pack_import_title")):
            return
        self.status_var.set(self.tr("shared_cache_sync_downloading"))
        try:
            urls = resolve_official_library_urls(
                current_info=current_info,
                release_base_url=self._release_asset_base_url(),
            )
            sync_result = sync_official_library(
                shared_cache_package_url=manifest.get("package_url"),
                resource_pack_url=urls.get("resource_pack_url"),
                bundled_corpus_url=urls.get("bundled_corpus_url"),
                load_word_resource_entries=self._load_word_resource_pack_entries,
            )
        except Exception as exc:
            messagebox.showerror(
                self.tr("shared_cache_sync_title"),
                self.trf("shared_cache_sync_failed", error=str(exc)),
            )
            return
        result = dict(sync_result.get("shared_cache_result") or {})
        load_result = dict(sync_result.get("word_pack_result") or {})
        corpus_result = dict(sync_result.get("corpus_result") or {})
        imported = int(result.get("imported") or 0)
        replaced = int(result.get("replaced") or 0)
        skipped_same = int(result.get("skipped_same") or 0)
        skipped_older = int(result.get("skipped_older") or 0)
        metadata_translations = int(result.get("metadata_translations") or 0)
        metadata_pos = int(result.get("metadata_pos") or 0)
        metadata_phonetics = int(result.get("metadata_phonetics") or 0)
        pack_count = int((load_result or {}).get("count") or 0)
        corpus_imported = int((corpus_result or {}).get("imported") or (corpus_result or {}).get("files") or 0)
        corpus_available = int((corpus_result or {}).get("available") or 0)
        self.status_var.set(
            self.trf(
                "shared_cache_sync_status",
                imported=imported,
                replaced=replaced,
                skipped=skipped_same + skipped_older,
                translations=metadata_translations,
                pos=metadata_pos,
                phonetics=metadata_phonetics,
                pack_count=pack_count,
                corpus_imported=corpus_imported,
                corpus_available=corpus_available,
            )
        )
        messagebox.showinfo(
            self.tr("shared_cache_sync_title"),
            self.trf(
                "shared_cache_sync_done",
                version=remote_version,
                imported=imported,
                replaced=replaced,
                same=skipped_same,
                older=skipped_older,
                translations=metadata_translations,
                pos=metadata_pos,
                phonetics=metadata_phonetics,
                pack_count=pack_count,
                corpus_imported=corpus_imported,
                corpus_available=corpus_available,
            ),
        )
        errors = list(result.get("errors") or [])
        if errors:
            detail = "\n".join(errors[:6])
            if len(errors) > 6:
                detail = f"{detail}\n..."
            messagebox.showwarning(
                self.tr("shared_cache_sync_title"),
                self.trf("shared_cache_import_errors", count=len(errors), detail=detail),
            )

    def build_shared_cache_manifest_tool(self):
        package_path = filedialog.asksaveasfilename(
            title=self.tr("shared_cache_export_title"),
            defaultextension=".zip",
            filetypes=[(self.tr("shared_cache_package_type"), "*.zip")],
            initialfile="wordspeaker_shared_audio_cache.zip",
        )
        if not package_path:
            return
        try:
            export_result = tts_export_shared_audio_cache_package(package_path)
        except Exception as exc:
            messagebox.showerror(
                self.tr("shared_cache_manifest_build_title"),
                self.trf("shared_cache_export_failed", error=str(exc)),
            )
            return
        if not export_result.get("ok"):
            messagebox.showinfo(
                self.tr("shared_cache_manifest_build_title"),
                self.tr("shared_cache_export_empty"),
            )
            return
        default_version = time.strftime("%Y.%m.%d.%H%M")
        version_text = self._prompt_text_input(
            self.tr("shared_cache_manifest_build_title"),
            self.tr("shared_cache_manifest_version_prompt"),
            initial_value=default_version,
        )
        if version_text is None:
            return
        version_text = str(version_text or "").strip() or default_version
        package_url = self._prompt_text_input(
            self.tr("shared_cache_manifest_build_title"),
            self.tr("shared_cache_manifest_url_prompt"),
            initial_value=(
                f"{self._release_asset_base_url()}/wordspeaker_shared_audio_cache.zip"
                if self._release_asset_base_url()
                else ""
            ),
        )
        if package_url is None:
            return
        package_url = str(package_url or "").strip()
        if not package_url:
            messagebox.showinfo(
                self.tr("shared_cache_manifest_build_title"),
                self.tr("shared_cache_sync_missing_url"),
            )
            return
        notes = self._prompt_text_input(
            self.tr("shared_cache_manifest_build_title"),
            self.tr("shared_cache_manifest_notes_prompt"),
        )
        manifest_path = filedialog.asksaveasfilename(
            title=self.tr("shared_cache_manifest_build_title"),
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
            initialfile="shared_audio_manifest.json",
        )
        if not manifest_path:
            return
        try:
            manifest_info = build_online_manifest(
                version_text,
                package_url,
                manifest_path,
                notes=notes or "",
            )
        except Exception as exc:
            messagebox.showerror(
                self.tr("shared_cache_manifest_build_title"),
                self.trf("shared_cache_sync_failed", error=str(exc)),
            )
            return
        count = int(export_result.get("entries") or 0)
        self.status_var.set(f"Shared audio cache manifest built: {count} items.")
        messagebox.showinfo(
            self.tr("shared_cache_manifest_build_title"),
            self.trf(
                "shared_cache_manifest_done",
                package_path=package_path,
                manifest_path=manifest_info.get("output_path") or manifest_path,
                count=count,
            ),
        )

    def _build_release_checklist_text(self):
        current_info = load_local_version_info()
        version_text = str(current_info.get("version") or "0.0.0").strip() or "0.0.0"
        portable_name = f"WordSpeaker-{version_text}-portable.zip"
        update_name = f"WordSpeaker-update-{version_text}.zip"
        lines = []
        if self.ui_language_var.get() == "en":
            lines.extend(
                [
                    f"Version: {version_text}",
                    "",
                    "Release assets",
                    "- dist/WordSpeaker/",
                    f"- {portable_name}",
                    f"- {update_name}",
                    "- manifest.json",
                    "- wordspeaker_shared_audio_cache.zip",
                    "- shared_audio_manifest.json",
                    "- optional: *.wspack",
                    "",
                    "Release order",
                    "1. Update version.json.",
                    "2. Rebuild the packaged app folder: dist/WordSpeaker/.",
                    f"3. Zip the whole dist/WordSpeaker/ folder as {portable_name}.",
                    f"4. Create {update_name} from the packaged app folder.",
                    "5. Generate manifest.json for the GitHub Release asset URL.",
                    "6. Export the official shared audio cache zip.",
                    "7. Create shared_audio_manifest.json for the shared-cache release URL.",
                    "8. Upload the portable zip, update zip, update manifest, shared-cache zip, and shared-audio manifest.",
                    "",
                    "Notes",
                    "- Update-package creation expects the packaged app folder, not the source repo.",
                    "- Update App is for program updates; Update Word Library is for shared audio only.",
                    "- Shared-cache sync only merges missing/newer global audio and keeps user cache.",
                ]
            )
        else:
            lines.extend(
                [
                    f"当前版本：{version_text}",
                    "",
                    "本次发布物",
                    "- dist/WordSpeaker/",
                    f"- {portable_name}",
                    f"- {update_name}",
                    "- manifest.json",
                    "- wordspeaker_shared_audio_cache.zip",
                    "- shared_audio_manifest.json",
                    "- 可选：*.wspack",
                    "",
                    "发布顺序",
                    "1. 先更新 version.json。",
                    "2. 重新打包完整程序目录：dist/WordSpeaker/。",
                    f"3. 把整个 dist/WordSpeaker/ 压成 {portable_name}。",
                    f"4. 从已打包目录制作 {update_name}。",
                    "5. 给更新包生成 manifest.json，默认指向 GitHub Release 资源地址。",
                    "6. 导出官方共享音频缓存 zip。",
                    "7. 为共享词音包生成 shared_audio_manifest.json。",
                    "8. 上传 portable zip、update zip、更新 manifest、共享词音 zip、共享词音 manifest。",
                    "",
                    "补充说明",
                    "- 更新包制作要基于已打包好的 dist/WordSpeaker/，不是源码目录。",
                    "- “更新程序”是更新程序本体，“更新词库”只同步共享词音。",
                    "- 共享词音同步只会补缺或更新较新的 global 音频，不会清空用户自己的缓存。",
                ]
            )
        return "\n".join(lines)

    def _release_asset_base_url(self):
        current_info = load_local_version_info()
        for key in ("channel_url", "shared_cache_manifest_url"):
            raw = str(current_info.get(key) or "").strip()
            if raw and "/" in raw:
                return raw.rsplit("/", 1)[0]
        return ""

    def open_release_checklist(self):
        win = tk.Toplevel(self)
        win.title(self.tr("release_checklist_title"))
        win.configure(bg="#f6f7fb")
        win.resizable(True, True)
        win.minsize(680, 520)

        wrap = ttk.Frame(win, style="Card.TFrame")
        wrap.pack(fill="both", expand=True, padx=10, pady=10)
        wrap.grid_columnconfigure(0, weight=1)
        wrap.grid_rowconfigure(1, weight=1)

        ttk.Label(wrap, text=self.tr("release_checklist_title"), style="Card.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, 8)
        )

        text_wrap = ttk.Frame(wrap, style="Card.TFrame")
        text_wrap.grid(row=1, column=0, sticky="nsew")
        text_wrap.grid_columnconfigure(0, weight=1)
        text_wrap.grid_rowconfigure(0, weight=1)

        text_widget = tk.Text(
            text_wrap,
            wrap="word",
            font=("Consolas", 11),
            relief="solid",
            bd=1,
            bg="#fbfcfe",
            fg="#1f2937",
        )
        text_widget.grid(row=0, column=0, sticky="nsew")
        text_scroll = ttk.Scrollbar(text_wrap, orient="vertical", command=text_widget.yview)
        text_scroll.grid(row=0, column=1, sticky="ns")
        text_widget.configure(yscrollcommand=text_scroll.set)
        checklist_text = self._build_release_checklist_text()
        text_widget.insert("1.0", checklist_text)
        text_widget.config(state="disabled")

        def _copy():
            try:
                self.clipboard_clear()
                self.clipboard_append(checklist_text)
            except Exception:
                return
            self.status_var.set(self.tr("release_checklist_copied"))

        btn_row = ttk.Frame(wrap, style="Card.TFrame")
        btn_row.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(btn_row, text=self.tr("release_checklist_copy"), command=_copy).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text=self.tr("close"), command=win.destroy).pack(side=tk.LEFT)

        win.transient(self.winfo_toplevel())
        win.grab_set()
        win.focus_set()

    def _launch_staged_update(self, staged_info):
        if not staged_info:
            return
        if not is_packaged_runtime():
            messagebox.showinfo(self.tr("update_title"), self.tr("update_not_packaged"))
            return
        if not self._prompt_save_unsaved_manual_words(title=self.tr("update_title")):
            return
        try:
            launch_staged_update(
                staged_info.get("source_dir"),
                entry_exe=staged_info.get("entry_exe"),
            )
        except Exception as exc:
            messagebox.showerror(self.tr("update_title"), self.trf("update_failed", error=str(exc)))
            return
        messagebox.showinfo(self.tr("update_title"), self.tr("update_started"))
        self.after(120, self.on_main_window_close)

    def _install_update_package(self, package_path, *, confirm=True):
        try:
            package_info = inspect_update_package(package_path)
        except Exception as exc:
            messagebox.showerror(self.tr("update_title"), self.trf("update_failed", error=str(exc)))
            return
        current_version = str(load_local_version_info().get("version") or "0.0.0").strip() or "0.0.0"
        target_version = str(package_info.get("version") or "0.0.0").strip() or "0.0.0"
        if confirm:
            if is_newer_version(target_version, current_version):
                should_install = messagebox.askyesno(
                    self.tr("update_title"),
                    self.trf("update_offline_confirm", version=target_version),
                )
            else:
                should_install = messagebox.askyesno(
                    self.tr("update_title"),
                    self.trf("update_reinstall_confirm", version=target_version),
                )
            if not should_install:
                return
        try:
            staged_info = stage_update_package(package_path)
        except Exception as exc:
            messagebox.showerror(self.tr("update_title"), self.trf("update_failed", error=str(exc)))
            return
        self._launch_staged_update(staged_info)

    def start_offline_update(self):
        if not is_packaged_runtime():
            messagebox.showinfo(self.tr("update_title"), self.tr("update_not_packaged"))
            return
        path = filedialog.askopenfilename(
            title=self.tr("update_title"),
            filetypes=[(self.tr("update_package_type"), "*.zip")],
        )
        if not path:
            return
        self._install_update_package(path, confirm=True)

    def start_online_update(self):
        if not is_packaged_runtime():
            messagebox.showinfo(self.tr("update_title"), self.tr("update_not_packaged"))
            return
        current_info = load_local_version_info()
        manifest_url = get_update_manifest_url() or str(current_info.get("channel_url") or "").strip()
        if not manifest_url:
            messagebox.showinfo(self.tr("update_title"), self.tr("update_online_no_default_url"))
            return
        set_update_manifest_url(manifest_url)
        self.status_var.set(self.tr("update_checking"))
        try:
            manifest = fetch_online_manifest(manifest_url)
        except Exception as exc:
            messagebox.showerror(self.tr("update_title"), self.trf("update_failed", error=str(exc)))
            return
        current_version = str(current_info.get("version") or "0.0.0").strip() or "0.0.0"
        target_version = str(manifest.get("version") or "0.0.0").strip() or "0.0.0"
        if not is_newer_version(target_version, current_version):
            self.status_var.set(self.tr("update_no_update"))
            messagebox.showinfo(self.tr("update_title"), self.tr("update_no_update"))
            return
        if not messagebox.askyesno(
            self.tr("update_title"),
            self.trf("update_online_available", version=target_version),
        ):
            return
        self.status_var.set(self.tr("update_downloading"))
        try:
            package_path = download_update_package(manifest.get("package_url"))
        except Exception as exc:
            messagebox.showerror(self.tr("update_title"), self.trf("update_failed", error=str(exc)))
            return
        self._install_update_package(package_path, confirm=False)

    def open_update_dialog(self):
        current_info = load_local_version_info()
        version_text = str(current_info.get("version") or "0.0.0").strip() or "0.0.0"
        win = tk.Toplevel(self)
        win.title(self.tr("update_title"))
        win.configure(bg="#f6f7fb")
        win.resizable(False, False)

        wrap = ttk.Frame(win, style="Card.TFrame")
        wrap.pack(fill="both", expand=True, padx=10, pady=10)
        ttk.Label(wrap, text=self.tr("update_title"), style="Card.TLabel").pack(anchor="w")
        ttk.Label(
            wrap,
            text=self.trf("update_current_version", version=version_text),
            style="Card.TLabel",
            foreground="#4b5563",
        ).pack(anchor="w", pady=(4, 0))
        ttk.Label(
            wrap,
            text=self.tr("update_desc"),
            style="Card.TLabel",
            foreground="#667085",
            wraplength=420,
            justify="left",
        ).pack(anchor="w", pady=(8, 12))

        btn_row = ttk.Frame(wrap, style="Card.TFrame")
        btn_row.pack(fill="x")

        def _run_online():
            win.destroy()
            self.start_online_update()

        def _run_offline():
            win.destroy()
            self.start_offline_update()

        ttk.Button(btn_row, text=self.tr("update_online"), command=_run_online).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text=self.tr("update_offline"), command=_run_offline).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text=self.tr("close"), command=win.destroy).pack(side=tk.LEFT)
        win.transient(self)
        win.grab_set()
        win.focus_set()

    def build_update_package_tool(self):
        source_dir = filedialog.askdirectory(title=self.tr("update_source_dir"))
        if not source_dir:
            return
        try:
            info = load_version_info(source_dir)
        except Exception as exc:
            messagebox.showerror(self.tr("update_title"), self.trf("update_failed", error=str(exc)))
            return
        version_text = str(info.get("version") or "0.0.0").strip() or "0.0.0"
        default_name = f"WordSpeaker-update-{version_text}.zip"
        output_path = filedialog.asksaveasfilename(
            title=self.tr("build_update_package"),
            defaultextension=".zip",
            filetypes=[(self.tr("update_package_type"), "*.zip")],
            initialfile=default_name,
        )
        if not output_path:
            return
        try:
            result = build_update_package(source_dir, output_path)
        except Exception as exc:
            messagebox.showerror(self.tr("update_title"), self.trf("update_failed", error=str(exc)))
            return
        self.status_var.set(
            f"Update package built: {result.get('version')} ({int(result.get('files') or 0)} files)."
        )
        messagebox.showinfo(
            self.tr("update_title"),
            self.trf(
                "update_build_done",
                path=result.get("output_path") or output_path,
                version=result.get("version") or version_text,
                count=int(result.get("files") or 0),
            ),
        )
        if not messagebox.askyesno(self.tr("update_title"), self.tr("update_manifest_create")):
            return
        package_url = self._prompt_text_input(
            self.tr("update_title"),
            self.tr("update_manifest_url_prompt"),
            initial_value=(
                f"{self._release_asset_base_url()}/{default_name}"
                if self._release_asset_base_url()
                else ""
            ),
        )
        if package_url is None:
            return
        package_url = str(package_url or "").strip()
        if not package_url:
            messagebox.showinfo(self.tr("update_title"), self.tr("update_online_missing_url"))
            return
        notes = self._prompt_text_input(
            self.tr("update_title"),
            self.tr("update_manifest_notes_prompt"),
        )
        manifest_path = filedialog.asksaveasfilename(
            title=self.tr("update_title"),
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
            initialfile="manifest.json",
        )
        if not manifest_path:
            return
        try:
            manifest_info = build_online_manifest(
                result.get("version") or version_text,
                package_url,
                manifest_path,
                notes=notes or "",
            )
        except Exception as exc:
            messagebox.showerror(self.tr("update_title"), self.trf("update_failed", error=str(exc)))
            return
        messagebox.showinfo(
            self.tr("update_title"),
            self.trf("update_manifest_done", path=manifest_info.get("output_path") or manifest_path),
        )

    def edit_selected_word_meta(self):
        selected_idx = self._get_context_or_selected_index()
        word = self._get_context_word()
        if not word:
            messagebox.showinfo("Info", "Please select a word first.")
            return
        current_pos = str(self.word_pos.get(word) or "").strip()
        current_zh = str(self.translations.get(word) or "").strip()

        win = tk.Toplevel(self)
        win.title(f"Edit POS / Translation - {word}")
        win.configure(bg="#f6f7fb")
        win.resizable(False, False)

        wrap = ttk.Frame(win, style="Card.TFrame")
        wrap.pack(fill="both", expand=True, padx=10, pady=10)
        ttk.Label(wrap, text=f"Word: {word}", style="Card.TLabel").pack(anchor="w")
        ttk.Label(
            wrap,
            text="Change the displayed part of speech and Chinese translation for this word."
            if self.ui_language_var.get() == "en"
            else "修改这个单词显示出来的词性和中文翻译。",
            style="Card.TLabel",
            foreground="#666",
            wraplength=420,
            justify="left",
        ).pack(anchor="w", pady=(0, 8))

        ttk.Label(wrap, text=self.tr("part_of_speech"), style="Card.TLabel").pack(anchor="w")
        pos_var = tk.StringVar(value=current_pos)
        pos_entry = ttk.Entry(wrap, textvariable=pos_var, width=30)
        pos_entry.pack(anchor="w", fill="x", pady=(2, 8))

        ttk.Label(wrap, text=self.tr("chinese_translation"), style="Card.TLabel").pack(anchor="w")
        zh_var = tk.StringVar(value=current_zh)
        zh_entry = ttk.Entry(wrap, textvariable=zh_var, width=40)
        zh_entry.pack(anchor="w", fill="x", pady=(2, 10))

        row = ttk.Frame(wrap, style="Card.TFrame")
        row.pack(fill="x")

        def _save():
            new_pos = str(pos_var.get() or "").strip()
            new_zh = str(zh_var.get() or "").strip()
            set_cached_pos(word, new_pos)
            set_cached_translation(word, new_zh)
            self.word_pos[word] = new_pos
            self.translations[word] = new_zh
            iid = str(selected_idx)
            if selected_idx is not None and self.word_table and self.word_table.exists(iid):
                note = self.store.notes[selected_idx] if selected_idx < len(self.store.notes) else ""
                tag = "even" if selected_idx % 2 == 0 else "odd"
                self.word_table.item(iid, values=self._build_word_table_values(selected_idx, word, note), tags=(tag,))
            self.refresh_dictation_recent_list()
            self.status_var.set(f"Updated POS / translation for '{word}'.")
            self._refresh_selection_details()
            win.destroy()

        ttk.Button(row, text=self.tr("save"), command=_save).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(row, text=self.tr("clear"), command=lambda: (pos_var.set(""), zh_var.set(""))).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(row, text=self.tr("cancel"), command=win.destroy).pack(side=tk.LEFT)

        pos_entry.focus_set()
        win.transient(self.winfo_toplevel())
        win.grab_set()

    # Data + history
    def load_words(self):
        path = filedialog.askopenfilename(
            title="Choose a word list",
            filetypes=[
                (self.tr("resource_pack_type"), "*.wspack"),
                ("Text files", "*.txt"),
                ("CSV files", "*.csv"),
            ],
        )
        if not path:
            return
        if not self._prompt_save_unsaved_manual_words(title="Open word list?"):
            return
        if str(path).lower().endswith(".wspack"):
            self.import_word_resource_pack_tool(package_path=path)
            return
        self.cancel_word_edit()
        words = self.word_list_controller.load_words(path)
        self.manual_source_dirty = False
        self.render_words(words)
        self.refresh_history()
        self.reset_playback_state()
        self.status_var.set(f"Loaded {len(words)} words from file.")

    def _parse_manual_rows(self, raw_text):
        return parse_manual_rows(raw_text)

    def _get_windows_clipboard_html(self):
        if os.name != "nt":
            return ""
        try:
            import win32clipboard

            fmt = win32clipboard.RegisterClipboardFormat("HTML Format")
            win32clipboard.OpenClipboard()
            try:
                if not win32clipboard.IsClipboardFormatAvailable(fmt):
                    return ""
                raw = win32clipboard.GetClipboardData(fmt)
            finally:
                win32clipboard.CloseClipboard()
            if isinstance(raw, bytes):
                raw = raw.rstrip(b"\x00")
                for encoding in ("utf-8", "utf-16le", "latin-1"):
                    try:
                        return raw.decode(encoding, errors="ignore")
                    except Exception:
                        continue
            return str(raw or "")
        except Exception:
            pass

        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32
        fmt = user32.RegisterClipboardFormatW("HTML Format")
        if not fmt:
            return ""
        if not user32.OpenClipboard(None):
            return ""
        try:
            handle = user32.GetClipboardData(fmt)
            if not handle:
                return ""
            locked = kernel32.GlobalLock(handle)
            if not locked:
                return ""
            try:
                size = int(kernel32.GlobalSize(handle) or 0)
                if size <= 0:
                    return ""
                raw = ctypes.string_at(locked, size).rstrip(b"\x00")
            finally:
                kernel32.GlobalUnlock(handle)
        finally:
            user32.CloseClipboard()
        if not raw:
            return ""
        for encoding in ("utf-8", "utf-16le", "latin-1"):
            try:
                return raw.decode(encoding, errors="ignore")
            except Exception:
                continue
        return ""

    def _extract_clipboard_html_fragment(self, raw_html):
        from ui.manual_words_presenter import extract_clipboard_html_fragment

        return extract_clipboard_html_fragment(raw_html)

    def _parse_clipboard_html_rows(self, raw_html):
        return parse_clipboard_html_rows(raw_html, table_parser_cls=_ClipboardTableHTMLParser)

    def _parse_tabular_text_rows(self, raw_text):
        return parse_tabular_text_rows(raw_text)

    def _read_clipboard_import_rows(self):
        try:
            raw = self.clipboard_get()
        except Exception:
            raw = ""
        return read_clipboard_import_rows(
            html_text=self._get_windows_clipboard_html(),
            raw_text=raw,
            table_parser_cls=_ClipboardTableHTMLParser,
        )

    def _normalize_import_word_text(self, text):
        return normalize_import_word_text(text)

    def _looks_like_word_line(self, text, next_line=""):
        from ui.manual_words_presenter import looks_like_word_line

        return looks_like_word_line(text, next_line=next_line)

    def _looks_like_contextual_phrase_word_line(self, text, next_line=""):
        from ui.manual_words_presenter import looks_like_contextual_phrase_word_line

        return looks_like_contextual_phrase_word_line(text, next_line=next_line)

    def _append_manual_preview_rows(self, rows, replace=False):
        append_manual_preview_rows(self.manual_words_table, rows, replace=replace)

    def _collect_manual_rows_from_table(self):
        return collect_manual_rows_from_table(self.manual_words_table)

    def _clear_manual_preview(self):
        clear_manual_preview(self.manual_words_table, self._cancel_manual_preview_edit)

    def _close_manual_words_window(self):
        close_manual_words_window(self.manual_words_window, self._cancel_manual_preview_edit)
        self.manual_words_window = None
        self.manual_words_table = None
        self.manual_words_table_scroll = None

    def _cancel_manual_preview_edit(self, _event=None):
        cancel_manual_preview_edit(self.manual_preview_edit_entry)
        self.manual_preview_edit_entry = None
        self.manual_preview_edit_item = None
        self.manual_preview_edit_column = None
        return "break"

    def _finish_manual_preview_edit(self, _event=None):
        return finish_manual_preview_edit(
            entry=self.manual_preview_edit_entry,
            table=self.manual_words_table,
            item_id=self.manual_preview_edit_item,
            column_id=self.manual_preview_edit_column,
            cancel_edit=self._cancel_manual_preview_edit,
        )

    def _start_manual_preview_edit(self, event=None, item_id=None, column_id=None):
        result = start_manual_preview_edit(
            table=self.manual_words_table,
            event=event,
            item_id=item_id,
            column_id=column_id,
            cancel_edit=self._cancel_manual_preview_edit,
            on_finish=self._finish_manual_preview_edit,
        )
        if not result:
            return "break"
        self.manual_preview_edit_entry = result["entry"]
        self.manual_preview_edit_item = result["item_id"]
        self.manual_preview_edit_column = result["column_id"]
        return "break"

    def _add_manual_preview_row(self):
        item_id = add_manual_preview_row(self.manual_words_table)
        if item_id:
            self._start_manual_preview_edit(item_id=item_id, column_id="#1")

    def _delete_selected_manual_preview_rows(self):
        delete_selected_manual_preview_rows(self.manual_words_table, self._cancel_manual_preview_edit)

    def on_manual_preview_paste(self, _event=None):
        rows = self._read_clipboard_import_rows()
        if not rows:
            return "break"
        self._append_manual_preview_rows(rows, replace=False)
        return "break"

    def _apply_manual_words(self, rows, mode="replace"):
        normalized_words, normalized_notes = normalize_manual_input_rows(rows)
        if not normalized_words:
            messagebox.showinfo("Info", "No valid words found.")
            return False
        self.cancel_word_edit()
        result = self.word_list_controller.apply_manual_words(
            normalized_words,
            normalized_notes,
            mode=mode,
            ui_language=self.ui_language_var.get(),
        )
        self.render_words(result.words)
        if result.source_bound:
            if result.saved_to_source:
                self.manual_source_dirty = False
        else:
            self._mark_manual_words_dirty()
        self.reset_playback_state()
        self.status_var.set(result.status_text)
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
        build_manual_words_window(self)

    def on_word_table_paste(self, _event=None):
        rows = self._read_clipboard_import_rows()
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

    def _set_word_action_context(self, index, origin="main", word=None):
        return set_word_action_context_flow(self, index, origin=origin, word=word)

    def _clear_word_action_context(self):
        clear_word_action_context_flow(self)

    def _get_context_or_selected_index(self):
        return get_context_or_selected_index_flow(self)

    def _get_context_word(self):
        return get_context_word_flow(self)

    def _get_context_audio_source_path(self):
        return get_context_audio_source_path_flow(self)

    def _get_word_audio_override_source_path(self):
        return get_word_audio_override_source_path_flow(self)

    def _dictation_row_to_store_index(self, tree, row_id=None):
        return dictation_row_to_store_index_flow(self, tree, row_id=row_id)

    def _get_selected_words_for_passage(self):
        selected = []
        for idx in self._get_selected_indices():
            if 0 <= idx < len(self.store.words):
                selected.append(self.store.words[idx])
        if selected:
            return selected
        return list(self.store.words)

    def _save_words_back_to_source(self):
        try:
            return self.word_list_controller.save_back_to_source()
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save source file.\n{e}")
            return False

    def _translate_single_word_async(self, row_idx, word):
        start_single_translation_flow(self, row_idx, word)

    def _phonetic_single_word_async(self, row_idx, word):
        start_single_phonetic_flow(self, row_idx, word)

    def _apply_single_translation(self, token, row_idx, word, zh_text):
        apply_single_translation_flow(self, token, row_idx, word, zh_text)

    def _apply_single_phonetic(self, token, row_idx, word, phonetic_text):
        apply_single_phonetic_flow(self, token, row_idx, word, phonetic_text)

    def start_edit_selected_word(self, _event=None):
        return self.start_edit_word_cell(column_id="#2")

    def start_edit_selected_note(self, _event=None):
        return self.start_edit_word_cell(column_id="#3")

    def prompt_edit_context_word(self, _event=None):
        selected_idx = self._get_context_or_selected_index()
        current_word = self._get_context_word()
        if not current_word:
            self.show_info("select_word_first")
            return
        new_value = self._prompt_text_input(
            self.tr("edit_word_title"),
            self.tr("edit_word_prompt"),
            initial_value=current_word,
        )
        if new_value is None:
            return
        if selected_idx is None or selected_idx >= len(self.store.words):
            return self._apply_recent_wrong_word_edit(current_word, new_value)
        return self._apply_word_edit_value(selected_idx, new_value)

    def prompt_edit_context_note(self, _event=None):
        selected_idx = self._get_context_or_selected_index()
        current_word = self._get_context_word()
        if not current_word:
            self.show_info("select_word_first")
            return
        if selected_idx is not None and selected_idx < len(self.store.words):
            current_note = self.store.notes[selected_idx] if selected_idx < len(self.store.notes) else ""
        else:
            current_note = str(self.store.get_dictation_word_stats(current_word).get("note") or "")
        new_value = self._prompt_text_input(
            self.tr("edit_note_title"),
            self.tr("edit_note_prompt"),
            initial_value=current_note,
        )
        if new_value is None:
            return
        if selected_idx is None or selected_idx >= len(self.store.words):
            return self._apply_recent_wrong_note_edit(current_word, new_value)
        return self._apply_note_edit_value(selected_idx, new_value)

    def prompt_add_word(self, _event=None):
        new_value = self._prompt_text_input(
            self.tr("add_word_title"),
            self.tr("add_word_prompt"),
        )
        if new_value is None:
            return
        normalized_word = re.sub(r"\s+", " ", str(new_value or "").strip())
        if not normalized_word:
            messagebox.showinfo("Info", "Word cannot be empty.")
            return
        try:
            result = self.word_list_controller.add_word(normalized_word, "")
        except ValueError:
            messagebox.showinfo("Info", "Word cannot be empty.")
            return
        self.cancel_word_edit()
        self.render_words(self.store.words)
        if not result.saved_to_source and not self.store.has_current_source_file():
            self._mark_manual_words_dirty()
        self.refresh_dictation_recent_list()
        self._set_word_action_context(result.index, origin="main")
        if self.word_table:
            iid = str(result.index)
            try:
                self.suppress_word_select_action = True
                self.word_table.selection_set(iid)
                self.word_table.focus(iid)
                self.word_table.see(iid)
            except Exception:
                pass
        self.status_var.set(f"Added word '{result.word}'.")
        self._ensure_word_metadata(result.word)
        self._refresh_selection_details()

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
                self.suppress_word_select_action = True
                self.word_table.selection_set(row_id)
                self.word_table.focus(row_id)
            if hit_column in ("#2", "#3"):
                target_column = hit_column
        if row_idx is None or row_idx >= len(self.store.words):
            return "break"
        if target_column not in ("#2", "#3"):
            return "break"
        self.cancel_word_edit()
        iid = str(row_idx)
        bbox = self.word_table.bbox(iid, target_column)
        if not bbox:
            return "break"

        x, y, width, height = bbox
        if target_column == "#3":
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

    def _apply_note_edit_value(self, idx, raw_value):
        if idx is None or idx < 0 or idx >= len(self.store.words):
            return "break"
        result = self.word_list_controller.update_note(idx, raw_value)
        if not result.changed:
            return "break"
        iid = str(idx)
        if self.word_table and self.word_table.exists(iid):
            tag = "even" if idx % 2 == 0 else "odd"
            self.word_table.item(
                iid,
                values=self._build_word_table_values(idx, self.store.words[idx], result.note),
                tags=(tag,),
            )
        if not result.saved_to_source and not self.store.has_current_source_file():
            self._mark_manual_words_dirty()
        self.refresh_dictation_recent_list()
        source_note = " and saved to source file" if result.saved_to_source else ""
        self.status_var.set(f"Updated note for '{result.word}'{source_note}.")
        self._refresh_selection_details()
        return "break"

    def _apply_recent_wrong_note_edit(self, word, raw_value):
        token = self._normalize_import_word_text(word or "")
        if not token:
            return "break"
        result = self.recent_wrong_controller.update_note(token, raw_value)
        if not result.changed:
            return "break"
        self.refresh_dictation_recent_list()
        self.status_var.set(f"Updated note for '{result.word}'.")
        self._refresh_selection_details()
        return "break"

    def _apply_word_edit_value(self, idx, raw_value):
        if idx is None or idx < 0 or idx >= len(self.store.words):
            return "break"
        normalized_value = re.sub(r"\s+", " ", str(raw_value or "").strip())
        if not normalized_value:
            messagebox.showinfo("Info", "Word cannot be empty.")
            return "break"
        try:
            result = self.word_list_controller.update_word(
                idx,
                normalized_value,
                translations=self.translations,
                word_pos=self.word_pos,
            )
        except ValueError:
            messagebox.showinfo("Info", "Word cannot be empty.")
            return "break"
        if not result.changed:
            return "break"
        iid = str(idx)
        if self.word_table and self.word_table.exists(iid):
            tag = "even" if idx % 2 == 0 else "odd"
            self.word_table.item(
                iid,
                values=(f"{idx + 1}.", f"{result.new_word}\nTranslating...", result.note),
                tags=(tag,),
            )
        if not result.saved_to_source and not self.store.has_current_source_file():
            self._mark_manual_words_dirty()
        self._translate_single_word_async(idx, result.new_word)
        self.phonetic_token += 1
        self._phonetic_single_word_async(idx, result.new_word)
        self.analysis_token += 1
        self._start_analysis_job([result.new_word], self.analysis_token)
        self.refresh_dictation_recent_list()
        source_note = " and saved to source file" if result.saved_to_source else ""
        self.status_var.set(f"Updated '{result.old_word}' to '{result.new_word}'{source_note}.")
        self._refresh_selection_details()
        return "break"

    def _apply_recent_wrong_word_edit(self, old_word, raw_value):
        source_word = self._normalize_import_word_text(old_word or "")
        new_word = self._normalize_import_word_text(raw_value or "")
        if not source_word:
            return "break"
        if not new_word:
            messagebox.showinfo("Info", "Word cannot be empty.")
            return "break"
        try:
            result = self.recent_wrong_controller.update_word(
                source_word,
                new_word,
                translations=self.translations,
                word_pos=self.word_pos,
            )
        except ValueError:
            messagebox.showinfo("Info", "Word cannot be empty.")
            return "break"
        if not result.changed:
            return "break"
        self.word_action_word = result.new_word
        self.refresh_dictation_recent_list()
        self.status_var.set(f"Updated '{result.old_word}' to '{result.new_word}'.")
        self._refresh_selection_details()
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
        if edit_column == "#3":
            return self._apply_note_edit_value(idx, new_word)
        return self._apply_word_edit_value(idx, new_word)

    def delete_selected_word(self):
        selected_idx = self._get_context_or_selected_index()
        context_word = self._get_context_word()
        if (selected_idx is None or selected_idx >= len(self.store.words)) and self.word_action_origin == "dictation":
            if not context_word:
                self.show_info("select_word_first")
                return
            if not messagebox.askyesno(
                self.tr("delete_word"),
                self.trf("delete_word_confirm", word=context_word),
            ):
                return
            self.recent_wrong_controller.clear_wrong_word(
                context_word,
                recent_wrong_source_path=self._get_recent_wrong_cache_source_path(),
            )
            self._clear_word_action_context()
            self.refresh_dictation_recent_list()
            self.status_var.set(self.trf("word_deleted", word=context_word))
            self._refresh_selection_details()
            return
        if selected_idx is None or selected_idx >= len(self.store.words):
            self.show_info("select_word_first")
            return
        word = self.store.words[selected_idx]
        if not messagebox.askyesno(
            self.tr("delete_word"),
            self.trf("delete_word_confirm", word=word),
        ):
            return

        self.cancel_word_edit()
        result = self.word_list_controller.delete_word(
            selected_idx,
            translations=self.translations,
            word_pos=self.word_pos,
        )
        if not result.saved_to_source and not self.store.has_current_source_file():
            self._mark_manual_words_dirty()

        self._clear_word_action_context()
        self.render_words(list(self.store.words))
        self.refresh_dictation_recent_list()
        if self.store.words and self.word_table:
            next_idx = min(selected_idx, len(self.store.words) - 1)
            if self.word_table.exists(str(next_idx)):
                self.suppress_word_select_action = True
                self.word_table.selection_set(str(next_idx))
                self.word_table.focus(str(next_idx))
        self.status_var.set(self.trf("word_deleted", word=result.word))

    def render_words(self, words):
        render_words_flow(self, words)

    def _start_translation_job(self, words, token):
        start_translation_job_flow(self, words, token)

    def _apply_translations(self, token, requested_words, translated):
        apply_translations_flow(self, token, requested_words, translated)

    def _ensure_word_metadata(self, word):
        ensure_word_metadata_flow(self, word)

    def update_empty_state(self):
        if self.store.words:
            self.empty_label.grid_remove()
        else:
            self.empty_label.grid()
        self._refresh_selection_details()

    def refresh_history(self):
        history = self.store.load_history()
        state = build_history_list_state(history)
        self.history_list.delete(0, tk.END)
        for label in state.labels:
            self.history_list.insert(tk.END, label)
        if not state.is_empty:
            self.history_path.config(text=state.current_path)
            self.history_empty.grid_remove()
        else:
            self.history_path.config(text="")
            self.history_empty.grid()
        self.refresh_dictation_recent_list()

    def on_history_open(self, _event=None):
        sel = self.history_list.curselection()
        self.cancel_word_edit()
        if not self._prompt_save_unsaved_manual_words(title="Open history file?"):
            return
        history = self.store.load_history()
        selected = get_selected_history_item(history, sel)
        if selected is None:
            return
        path = selected.path
        if not path or not os.path.exists(path):
            messagebox.showinfo("Info", "File not found or moved.")
            return
        words = self.word_list_controller.open_history_path(path)
        self.manual_source_dirty = False
        self.render_words(words)
        self.refresh_history()
        self.reset_playback_state()

    def on_history_right_click(self, event):
        if not self.history_list or not self.history_context_menu:
            return
        index = self.history_list.nearest(event.y)
        if index < 0:
            return
        try:
            self.history_list.selection_clear(0, tk.END)
            self.history_list.selection_set(index)
            self.history_list.activate(index)
        except Exception:
            pass
        try:
            self.history_context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.history_context_menu.grab_release()
        return "break"

    def delete_selected_history_item(self):
        history = self.store.load_history()
        selected = get_selected_history_item(history, self.history_list.curselection())
        if selected is None:
            return
        path = selected.path
        name = selected.name
        if not messagebox.askyesno(self.tr("history"), self.trf("delete_history_confirm", name=name)):
            return
        result = self.word_list_controller.delete_history_item(path)
        if result.detached_current_source:
            self.manual_source_dirty = bool(self.store.words)
            self._refresh_selection_details()
        self.refresh_history()
        self.show_info("history_deleted", name=result.name, count=result.removed_count)

    def rename_selected_history_item(self):
        history = self.store.load_history()
        selected = get_selected_history_item(history, self.history_list.curselection())
        if selected is None:
            return
        old_path = os.path.abspath(selected.path)
        if not old_path or not os.path.exists(old_path):
            self.show_info("rename_history_missing")
            return
        old_name = os.path.basename(old_path)
        new_name = self._prompt_text_input(
            self.tr("rename_history_file"),
            self.tr("rename_history_prompt"),
            initial_value=old_name,
        )
        if new_name is None:
            return
        try:
            rename_target = build_rename_history_target(old_path, new_name)
        except ValueError as exc:
            if str(exc) == "invalid_name":
                self.show_info("rename_history_invalid")
            return
        if not rename_target.new_name:
            return
        if not rename_target.changed:
            return
        if os.path.exists(rename_target.new_path):
            self.show_info("rename_history_exists")
            return
        try:
            result = self.word_list_controller.rename_history_item(rename_target.old_path, rename_target.new_path)
            self.refresh_history()
            self._refresh_selection_details()
            self.show_info(
                "rename_history_done",
                name=os.path.basename(result.path),
                count=result.migrated,
                queued=result.queued,
            )
        except Exception as e:
            self.show_error("history", "rename_history_failed", error=e)

    def add_manual_wrong_word(self):
        word = self._prompt_text_input(self.tr("add_wrong_word"), self.tr("enter_wrong_word"))
        if word is None:
            return
        token = str(word or "").strip()
        if not token:
            return
        result = self.recent_wrong_controller.add_manual_wrong_word(
            token,
            recent_wrong_source_path=self._get_recent_wrong_cache_source_path(),
        )
        if result.added_to_word_list:
            self.manual_source_dirty = True
            self.render_words(self.store.words)
        self.refresh_dictation_recent_list()
        self._refresh_selection_details()
        self.show_info("wrong_word_added", word=result.word)

    def _prompt_text_input(self, title, prompt, initial_value=""):
        win = tk.Toplevel(self)
        win.title(str(title or "Input"))
        win.configure(bg="#f6f7fb")
        win.resizable(False, False)
        result = {"value": None}

        wrap = ttk.Frame(win, style="Card.TFrame")
        wrap.pack(fill="both", expand=True, padx=10, pady=10)
        ttk.Label(wrap, text=str(prompt or ""), style="Card.TLabel", wraplength=360, justify="left").pack(anchor="w")
        value_var = tk.StringVar(value=str(initial_value or ""))
        entry = ttk.Entry(wrap, textvariable=value_var, width=36)
        entry.pack(fill="x", pady=(8, 10))

        def _submit():
            result["value"] = value_var.get()
            win.destroy()

        def _cancel():
            win.destroy()

        btn_row = ttk.Frame(wrap, style="Card.TFrame")
        btn_row.pack(fill="x")
        ttk.Button(btn_row, text=self.tr("confirm"), command=_submit).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text=self.tr("cancel"), command=_cancel).pack(side=tk.LEFT)
        entry.bind("<Return>", lambda _e: _submit())
        entry.focus_set()
        win.transient(self)
        win.grab_set()
        self.wait_window(win)
        return result["value"]

    # Corpus find
    def open_find_window(self):
        open_find_window_flow(self)

    def _set_find_query_from_selection(self):
        set_find_query_from_selection_flow(self)

    def refresh_find_corpus_summary(self):
        refresh_find_corpus_summary_flow(self)

    def on_find_docs_right_click(self, event):
        return on_find_docs_right_click_flow(self, event)

    def delete_selected_corpus_document(self):
        delete_selected_corpus_document_flow(self)

    def _clear_find_task_queue(self):
        clear_find_task_queue_flow(self)

    def _emit_find_task_event(self, event_type, token, payload=None):
        emit_find_task_event_flow(self, event_type, token, payload)

    def _poll_find_task_events(self, token):
        poll_find_task_events_flow(self, token)

    def _handle_find_task_error(self, message):
        handle_find_task_error_flow(self, message)

    def import_find_documents(self):
        import_find_documents_flow(self)

    def _apply_find_import_result(self, payload):
        apply_find_import_result_flow(self, payload)

    def run_find_search(self):
        run_find_search_flow(self)

    def search_selected_word_in_corpus(self):
        search_selected_word_in_corpus_flow(self)

    def _get_selected_find_document(self):
        return get_selected_find_document_flow(self)

    def clear_find_document_filter(self):
        clear_find_document_filter_flow(self)

    def _apply_find_search_result(self, payload):
        apply_find_search_result_flow(self, payload)

    def _clear_find_preview(self):
        clear_find_preview_flow(self)

    def _on_find_result_select(self, _event=None):
        on_find_result_select_flow(self, _event=_event)

    def _show_find_result_preview(self, row_id):
        show_find_result_preview_flow(self, row_id)

    # IELTS passage
    def open_passage_window(self):
        open_passage_window_flow(self, Tooltip)

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
        return speech_text_from_passage(text)

    def _normalize_answer(self, text):
        return normalize_answer(text)

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
        state = build_passage_practice_state(
            current_passage_original=self.current_passage_original,
            current_passage=self.current_passage,
            current_passage_words=self.current_passage_words,
            store_words=self.store.words,
        )
        if state.get("error") == "missing_passage":
            messagebox.showinfo("Info", "Generate a passage first.")
            return
        if state.get("error") == "no_keywords":
            messagebox.showinfo("Info", "No suitable keywords found in this passage for practice.")
            return

        self.passage_is_practice = True
        self.passage_cloze_text = state["cloze"]
        self.passage_answers = state["answers"]
        self._set_passage_text(state["cloze"])
        self._clear_passage_practice_input()
        self._clear_passage_practice_result()
        self.passage_status_var.set(state["status_text"])

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
        state = build_passage_practice_check_state(answers=self.passage_answers, user_lines=lines)
        if self.passage_practice_result:
            apply_diff(self.passage_practice_result, state["expected_text"], state["actual_text"])
        self.passage_status_var.set(state["status_text"])

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

    def ensure_api_credentials(self):
        need_llm = not str(get_llm_api_key() or "").strip()
        need_tts = not str(get_tts_api_key() or "").strip()
        if not need_llm and not need_tts:
            return
        self.open_api_key_window(force_llm=need_llm, force_tts=need_tts)

    def ensure_gemini_api_key(self):
        self.open_api_key_window(force_llm=True, force_tts=False, initial_section="llm")

    def ensure_tts_api_key(self):
        if str(get_tts_api_key() or "").strip():
            return
        self.open_api_key_window(force_llm=False, force_tts=True, initial_section="tts")

    def list_gemini_models(self):
        return list_available_gemini_models()

    def _on_llm_provider_selected(self):
        set_llm_api_provider("gemini")
        self.llm_api_provider_var.set(self.tr("provider_gemini"))

    def _clear_gemini_validation_queue(self):
        clear_gemini_validation_queue_flow(self)

    def _emit_gemini_validation_event(self, event_type, token, payload=None):
        emit_gemini_validation_event_flow(self, event_type, token, payload)

    def _poll_gemini_validation_events(self, token):
        poll_gemini_validation_events_flow(self, token)

    def open_api_key_window(self, force_llm=False, force_tts=False, initial_section="llm"):
        open_api_key_window_flow(self, force_llm=force_llm, force_tts=force_tts, initial_section=initial_section)

    def open_gemini_key_window(self, force_verify=False):
        self.open_api_key_window(force_llm=force_verify, force_tts=False, initial_section="llm")

    def _close_api_key_window(self):
        close_api_key_window_flow(self)

    def _set_api_entry_error(self, field, has_error):
        set_api_entry_error_flow(self, field, has_error)

    def test_and_save_gemini_key(self):
        api_key = str(self.gemini_key_var.get() or "").strip()
        model_name = self._get_selected_gemini_model()
        if not api_key:
            messagebox.showinfo(self.tr("info"), self.tr("paste_gemini_key_first"))
            return

        self.gemini_key_status_var.set(f"Testing LLM key with {model_name}...")
        if self.gemini_key_test_btn:
            self.gemini_key_test_btn.config(state="disabled")

        self.gemini_validation_token += 1
        token = self.gemini_validation_token
        self.gemini_validation_active_token = token
        self._clear_gemini_validation_queue()
        start_gemini_validation_task(
            token=token,
            api_key=api_key,
            model_name=model_name,
            emit_event=self._emit_gemini_validation_event,
        )
        self.after(80, lambda t=token: self._poll_gemini_validation_events(t))

    def test_and_save_api_keys(self):
        request = build_combined_api_validation_request(
            llm_key=self.gemini_key_var.get(),
            tts_key=self.tts_key_var.get(),
            tts_provider=self._tts_provider_value(),
            model_name=self._get_selected_gemini_model(),
            force_llm=self.api_key_force_llm,
            force_tts=self.api_key_force_tts,
        )

        self._set_api_entry_error("llm", False)
        self._set_api_entry_error("tts", False)
        local_state = build_combined_api_local_validation_state(request)
        self.gemini_key_status_var.set(local_state["llm_status"])
        self.tts_key_status_var.set(local_state["tts_status"])
        if local_state["llm_error"]:
            self._set_api_entry_error("llm", True)
        if local_state["tts_error"]:
            self._set_api_entry_error("tts", True)
        if local_state["has_local_error"]:
            return
        if not local_state["has_any_request"]:
            messagebox.showinfo(self.tr("info"), "Please enter at least one API key.")
            return

        if request["llm_required"]:
            self.gemini_key_status_var.set(f"Testing LLM key with {request['model_name']}...")
        if request["tts_required"]:
            self.tts_key_status_var.set(
                f"Testing TTS API key with {self._tts_provider_label(request['tts_provider'])}..."
            )
        if self.api_key_test_btn:
            self.api_key_test_btn.config(state="disabled")

        self.gemini_validation_token += 1
        token = self.gemini_validation_token
        self.gemini_validation_active_token = token
        self._clear_gemini_validation_queue()
        start_combined_api_validation_task(
            token=token,
            llm_required=request["llm_required"],
            tts_required=request["tts_required"],
            llm_key=request["llm_key"],
            tts_key=request["tts_key"],
            model_name=request["model_name"],
            tts_provider=request["tts_provider"],
            emit_event=self._emit_gemini_validation_event,
        )
        self.after(80, lambda t=token: self._poll_gemini_validation_events(t))

    def _finish_gemini_validation_success(self, payload):
        state = build_single_llm_success_state(payload, default_model=DEFAULT_GEMINI_MODEL)
        set_llm_api_key(state["api_key"])
        set_llm_api_provider("gemini")
        set_generation_model(state["model_name"])
        self.gemini_verified = True
        self.gemini_key_status_var.set(state["status_text"])
        if self.gemini_key_test_btn:
            self.gemini_key_test_btn.config(state="normal")
        self.status_var.set(state["main_status"])
        self._maybe_close_api_key_window()

    def _finish_combined_api_validation(self, payload):
        state = build_combined_api_apply_state(payload)

        if state["llm_required"] and state["llm_ok"]:
            set_llm_api_key(state["llm_api_key"])
            set_llm_api_provider("gemini")
            set_generation_model(state["llm_model"] or DEFAULT_GEMINI_MODEL)
            self.gemini_verified = True
            self.gemini_key_status_var.set("LLM API key is valid.")
            self._set_api_entry_error("llm", False)
        elif state["llm_required"]:
            self.gemini_verified = False
            self.gemini_key_status_var.set(state["llm_error_message"])
            self._set_api_entry_error("llm", True)

        if state["tts_required"] and state["tts_ok"]:
            set_tts_api_key(state["tts_api_key"])
            set_tts_api_provider(state["tts_provider"])
            self.tts_api_provider_var.set(self._tts_provider_label(state["tts_provider"]))
            self.tts_key_status_var.set("TTS API key is valid.")
            self._set_api_entry_error("tts", False)
            self.refresh_voice_list()
        elif state["tts_required"]:
            self.tts_key_status_var.set(state["tts_error_message"])
            self._set_api_entry_error("tts", True)

        if self.api_key_test_btn:
            self.api_key_test_btn.config(state="normal")

        if state["all_ok"]:
            self.status_var.set("API ready.")
            self._maybe_close_api_key_window()

    def _finish_gemini_validation_error(self, message):
        state = build_single_api_error_state(kind="llm")
        self.gemini_verified = False
        self.gemini_key_status_var.set(state["status_text"])
        if self.gemini_key_test_btn:
            self.gemini_key_test_btn.config(state="normal")
        messagebox.showerror(self.tr("gemini_api_key_error"), str(message or "Unknown error"))

    def _require_gemini_ready(self):
        if self.gemini_verified and get_llm_api_key():
            return True
        self.open_gemini_key_window(force_verify=False)
        return False

    def open_tts_key_window(self):
        self.open_api_key_window(force_llm=False, force_tts=False, initial_section="tts")

    def _on_tts_provider_selected(self, _event=None):
        provider = self._tts_provider_value()
        set_tts_api_provider(provider)
        tts_dedupe_pending_online_queue(provider)
        self.tts_api_provider_var.set(self._tts_provider_label(provider))
        self.refresh_voice_list()

    def test_and_save_tts_key(self):
        api_key = str(self.tts_key_var.get() or "").strip()
        provider = self._tts_provider_value()
        if not api_key:
            messagebox.showinfo(self.tr("info"), self.tr("paste_tts_key_first"))
            return
        self.tts_key_status_var.set("Testing TTS API key...")
        if self.tts_key_test_btn:
            self.tts_key_test_btn.config(state="disabled")

        self.gemini_validation_token += 1
        token = self.gemini_validation_token
        self.gemini_validation_active_token = token
        self._clear_gemini_validation_queue()
        start_tts_validation_task(
            token=token,
            api_key=api_key,
            provider=provider,
            emit_event=self._emit_gemini_validation_event,
        )
        self.after(80, lambda t=token: self._poll_gemini_validation_events(t))

    def _finish_tts_validation_success(self, payload):
        state = build_single_tts_success_state(payload)
        set_tts_api_key(state["api_key"])
        set_tts_api_provider(state["provider"])
        self.tts_key_status_var.set(state["status_text"])
        if self.tts_key_test_btn:
            self.tts_key_test_btn.config(state="normal")
        self.tts_api_provider_var.set(self._tts_provider_label(state["provider"]))
        self.refresh_voice_list()
        self.status_var.set(state["main_status"])
        self._maybe_close_api_key_window()

    def _finish_tts_validation_error(self, message):
        state = build_single_api_error_state(kind="tts")
        self.tts_key_status_var.set(state["status_text"])
        if self.tts_key_test_btn:
            self.tts_key_test_btn.config(state="normal")
        messagebox.showerror(self.tr("tts_api_key_error"), str(message or "Unknown error"))

    def _maybe_close_api_key_window(self):
        maybe_close_api_key_window_flow(self)

    def generate_ielts_passage(self):
        generate_ielts_passage_flow(self)

    def _clear_passage_event_queue(self):
        clear_passage_event_queue_flow(self)

    def _emit_passage_event(self, event_type, token, payload=None):
        emit_passage_event_flow(self, event_type, token, payload)

    def _poll_passage_generation_events(self, token):
        poll_passage_generation_events_flow(self, token)

    def _update_partial_passage(self, token, text):
        update_partial_passage_flow(self, token, text)

    def _run_passage_generation(self, token, words, model_name):
        start_passage_generation_task(
            token=token,
            words=words,
            model_name=model_name,
            emit_event=self._emit_passage_event,
        )

    def _apply_generated_passage(self, token, result):
        apply_generated_passage_flow(self, token, result)

    def _pause_word_playback(self):
        pause_word_playback_for_passage_flow(self)

    def play_generated_passage(self):
        play_generated_passage_flow(self)

    def stop_passage_playback(self):
        stop_passage_playback_flow(self)

    # Player controls
    def speak_selected(self):
        selected_idx = self._get_selected_index()
        if selected_idx is None or selected_idx >= len(self.store.words):
            messagebox.showinfo("Info", "Please select a word first.")
            return
        word = self.store.words[selected_idx]
        runtime = tts_get_runtime_label()
        source_path = self.store.get_current_source_path()
        cached = get_voice_source() == SOURCE_GEMINI and tts_has_cached_word_audio(word, source_path=source_path)
        token = speak_async(
            word,
            self.volume_var.get() / 100.0,
            rate_ratio=self.speech_rate_var.get(),
            cancel_before=True,
            source_path=source_path,
        )
        if cached:
            self.status_var.set(f"Playing cached audio for '{word}'.")
        else:
            self.status_var.set(f"Generating '{word}' with {runtime}...")
        self._watch_tts_backend(token, target="status", text_label=word)

    def toggle_settings(self):
        if self.settings_window and self.settings_window.winfo_exists():
            self.settings_window.lift()
            return
        self.open_settings_window()

    def _close_settings_window(self):
        close_settings_window_flow(self)

    def _refresh_settings_gemini_status(self):
        refresh_settings_runtime_status_flow(self)

    def toggle_history(self):
        self.history_visible = True
        self.check_visible = False
        self._select_sidebar_tab("history")
        self._refresh_selection_details()

    def toggle_check(self):
        self.check_visible = True
        self.history_visible = True
        self.open_dictation_window()
        self._refresh_selection_details()

    def set_order_mode(self, mode):
        self.order_mode.set("order")

    def update_order_button(self):
        return

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
        return

    def on_dictation_volume_change(self, _value=None):
        on_dictation_volume_change_flow(self, _value)

    def close_dictation_volume_popup(self):
        close_dictation_volume_popup_flow(self)

    def toggle_dictation_volume_popup(self):
        toggle_dictation_volume_popup_flow(self)

    def set_loop_mode(self, loop_all):
        return

    def update_loop_button(self):
        return

    def on_stop_at_end_toggle(self):
        return

    def open_settings_window(self):
        open_settings_window_flow(self)

    def apply_custom_interval(self):
        try:
            val = parse_custom_interval(self.custom_interval.get(), minimum=0.2)
        except Exception:
            self.show_info("valid_number_needed")
            return
        self.set_interval(val)

    def build_queue(self):
        return self.main_playback_controller.build_queue(self.store.words)

    def _sync_main_playback_state(self):
        sync_main_playback_state_flow(self)

    def rebuild_queue_on_mode_change(self):
        rebuild_main_playback_on_mode_change_flow(self)

    def toggle_play(self):
        toggle_main_playback_flow(self)

    def play_current(self):
        play_current_main_playback_flow(self)

    def schedule_next(self, playback_token):
        schedule_next_main_playback_flow(self, playback_token)

    def next_word(self, token):
        next_main_playback_word_flow(self, token)

    def set_current_word(self):
        set_current_main_playback_word_flow(self)

    def cancel_schedule(self):
        self.playback_schedule_token += 1
        if self.after_id:
            self.after_cancel(self.after_id)
            self.after_id = None

    def update_play_button(self):
        update_main_playback_button_flow(self)

    def reset_playback_state(self):
        reset_main_playback_state_flow(self)

    def update_right_visibility(self):
        if self.wordlist_hidden:
            if self.main:
                self.main.grid_columnconfigure(0, weight=1)
                self.main.grid_columnconfigure(2, weight=0)
            if self.detail_card:
                self.detail_card.grid_remove()
            self.right.grid_configure(row=0, column=0, columnspan=3, sticky="nsew")
            self.mid_sep.grid_remove()
        else:
            if self.main:
                self.main.grid_columnconfigure(0, weight=5)
                self.main.grid_columnconfigure(2, weight=4)
            if self.detail_card:
                self.detail_card.grid()
            self.left.grid_configure(columnspan=1)
            self.right.grid_configure(row=0, column=2, columnspan=1, sticky="nsew")
            self.mid_sep.grid()
        self.right.grid()

    def build_queue_from_selection(self):
        self.main_playback_controller.set_current_by_selection(
            words=self.store.words,
            selected_idx=self._get_selected_index(),
        )
        self._sync_main_playback_state()
        if not self.store.words:
            return
        self.set_current_word()

    def on_word_selected(self, _event=None):
        if self.suppress_word_select_action:
            self.suppress_word_select_action = False
            return
        self._clear_word_action_context()
        if not self.store.words:
            return
        selected_index = self._get_selected_index()
        if selected_index is None:
            self._refresh_selection_details()
            return
        if 0 <= selected_index < len(self.store.words):
            self._ensure_word_metadata(self.store.words[selected_index])
        self._refresh_selection_details()
        if self._dictation_window_active():
            return
        self._speak_selected_word_if_needed(selected_index)

    def on_word_double_click(self, event=None):
        self._clear_word_action_context()
        if not self.store.words:
            return "break"
        row_id = self.word_table.identify_row(event.y) if event is not None and self.word_table else ""
        if row_id:
            try:
                self.suppress_word_select_action = True
                self.word_table.selection_set(row_id)
                self.word_table.focus(row_id)
            except Exception:
                pass
        selected_index = self._get_selected_index()
        if selected_index is not None:
            self._speak_selected_word_if_needed(selected_index, force=True)
        return "break"

    def _speak_selected_word_if_needed(self, selected_index, force=False):
        if selected_index is None:
            return
        if not force and self._dictation_window_active():
            return
        now = time.time()
        if not force and self.last_word_speak_index == selected_index and (now - self.last_word_speak_at) < 0.35:
            return
        self.last_word_speak_index = selected_index
        self.last_word_speak_at = now
        self.speak_selected()

    def _dictation_window_active(self):
        return bool(self.dictation_window and self.dictation_window.winfo_exists())

    def _stop_main_word_playback(self):
        stop_main_playback_flow(self)

    def on_word_right_click(self, event):
        return on_word_right_click_flow(self, event)

    def on_dictation_word_right_click(self, event):
        return on_dictation_word_right_click_flow(self, event)

    def _fallback_sentence(self, word):
        w = str(word or "").strip()
        if not w:
            return ""
        if " " in w:
            return f'Please use "{w}" in your next speaking practice task.'
        return f"I wrote the word {w} in my notebook for today's review."

    def _clear_sentence_event_queue(self):
        clear_sentence_event_queue_flow(self)

    def _emit_sentence_event(self, event_type, token, payload=None):
        emit_sentence_event_flow(self, event_type, token, payload)

    def _clear_synonym_event_queue(self):
        clear_synonym_event_queue_flow(self)

    def _emit_synonym_event(self, event_type, token, payload=None):
        emit_synonym_event_flow(self, event_type, token, payload)

    def _poll_synonym_events(self, token):
        poll_synonym_events_flow(self, token)

    def _poll_sentence_events(self, token):
        poll_sentence_events_flow(self, token)

    def make_sentence_for_selected_word(self):
        make_sentence_for_selected_word_flow(self)

    def lookup_synonyms_for_selected_word(self):
        lookup_synonyms_for_selected_word_flow(self)

    def _show_sentence_window(self, word, sentence, source):
        show_sentence_window_flow(self, word, sentence, source)

    def _show_synonym_window(self, word, focus, synonyms, source=None):
        show_synonym_window_flow(self, word, focus, synonyms, source=source)

    def toggle_wordlist_visibility(self):
        # Hide or show the word list during dictation to avoid seeing words.
        top = self.winfo_toplevel()
        self.wordlist_hidden = not self.wordlist_hidden
        if self.wordlist_hidden:
            self.hidden_notebook_tabs = []
            try:
                self.saved_window_geometry = top.geometry()
                minsize_values = top.tk.call("wm", "minsize", top._w)
                if len(minsize_values) >= 2:
                    self.saved_window_minsize = (int(minsize_values[0]), int(minsize_values[1]))
            except Exception:
                self.saved_window_geometry = ""
                self.saved_window_minsize = None
            self.left.grid_remove()
            self.hide_words_btn.config(text=self.tr("show_word_list"))
            if self.right_notebook:
                for tab in (self.review_tab, self.history_tab, self.tools_tab):
                    try:
                        self.right_notebook.tab(tab, state="hidden")
                        self.hidden_notebook_tabs.append(tab)
                    except Exception:
                        pass
            try:
                top.minsize(560, 430)
                top.geometry("620x500")
            except Exception:
                pass
        else:
            self.left.grid()
            self.hide_words_btn.config(text=self.tr("hide_word_list"))
            if self.right_notebook:
                for tab in self.hidden_notebook_tabs:
                    try:
                        self.right_notebook.tab(tab, state="normal")
                    except Exception:
                        pass
                self.hidden_notebook_tabs = []
            try:
                if self.saved_window_minsize:
                    top.minsize(self.saved_window_minsize[0], self.saved_window_minsize[1])
                else:
                    top.minsize(1320, 760)
                if self.saved_window_geometry:
                    top.geometry(self.saved_window_geometry)
                else:
                    top.geometry("1480x860")
            except Exception:
                pass
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
                self.voice_map[label] = (v.get("source") or SOURCE_GEMINI, v.get("id"), name)
                options.append(label)

        if self.voice_combo:
            self.voice_combo["values"] = options

        # restore current selection
        current_source = get_voice_source()
        current_id = get_voice_id()
        selected = options[0] if options else ""
        if current_source and current_id:
            for label, data in self.voice_map.items():
                if data[0] == current_source and data[1] == current_id:
                    selected = label
                    break
        self.voice_var.set(selected)

    def on_voice_change(self, _event=None):
        label = self.voice_var.get()
        data = self.voice_map.get(label)
        if not data:
            set_voice_source(SOURCE_GEMINI, "gemini-kore", None)
            return
        source, voice_id, voice_label = data
        if source == "kokoro" and not kokoro_ready():
            set_voice_source(SOURCE_GEMINI, "gemini-kore", None)
            self.refresh_voice_list()
            messagebox.showinfo(
                "Kokoro Not Ready",
                "Kokoro is listed for convenience, but it is not installed yet.\n\n"
                "To enable it, add the Kokoro model files under:\n"
                "data/models/kokoro/\n\n"
                "The source has been switched back to the online TTS provider for now.",
            )
            return
        if source == "piper" and not piper_ready():
            if kokoro_ready():
                set_voice_source(SOURCE_KOKORO, "bf_emma", "Kokoro English (UK)")
                self.refresh_voice_list()
                messagebox.showinfo(
                    "Piper Not Ready",
                    "Piper is listed for convenience, but it is not configured yet.\n\n"
                    "To enable it, add at least one Piper .onnx model and matching .onnx.json config under:\n"
                    "data/models/piper/\n\n"
                    "The source has been switched to Kokoro for now.",
                )
            else:
                set_voice_source(SOURCE_GEMINI, "gemini-kore", None)
                self.refresh_voice_list()
                messagebox.showinfo(
                    "Piper Not Ready",
                    "Piper is listed for convenience, but it is not configured yet.\n\n"
                    "To enable it, add at least one Piper .onnx model and matching .onnx.json config under:\n"
                    "data/models/piper/\n\n"
                    "The source has been switched back to the online TTS provider for now.",
                )
            return
        set_voice_source(source, voice_id, voice_label)

    def _watch_tts_backend(self, token, target, text_label):
        watch_tts_backend_flow(self, token, target, text_label)


bind_state_properties(MainView)

