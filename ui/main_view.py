# -*- coding: utf-8 -*-
import ctypes
import html
import os
import queue
import random
import re
import shutil
import time
import unicodedata
import tkinter as tk
from html.parser import HTMLParser
from tkinter import ttk, filedialog, messagebox

from data.store import WordStore
from services.tts import (
    speak_async,
    speak_stream_async,
    cancel_all as tts_cancel_all,
    clear_word_backend_override as tts_clear_word_backend_override,
    cleanup_manual_session_cache as tts_cleanup_manual_session_cache,
    get_recent_wrong_cache_source as tts_get_recent_wrong_cache_source,
    get_word_audio_cache_info as tts_get_word_audio_cache_info,
    precache_word_audio_async,
    prepare_async as tts_prepare_async,
    dedupe_pending_online_queue as tts_dedupe_pending_online_queue,
    set_preferred_pending_source as tts_set_preferred_pending_source,
    set_word_backend_override as tts_set_word_backend_override,
    get_backend_status as tts_get_backend_status,
    get_online_tts_queue_status as tts_get_online_tts_queue_status,
    get_runtime_label as tts_get_runtime_label,
    has_cached_word_audio as tts_has_cached_word_audio,
    export_shared_audio_cache_package as tts_export_shared_audio_cache_package,
    import_shared_audio_cache_package as tts_import_shared_audio_cache_package,
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
from services.bundled_corpus import import_bundled_corpus_package, prepare_async as bundled_corpus_prepare_async
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
    validate_gemini_api_key,
)
from services.tts import validate_tts_api_key
from services.voice_catalog import kokoro_ready, list_system_voices, piper_ready
from services.voice_manager import (
    SOURCE_GEMINI,
    SOURCE_KOKORO,
    SOURCE_PIPER,
    get_voice_id,
    get_voice_source,
    set_voice_source,
)
from services.corpus_search import (
    corpus_stats,
    get_nlp_status,
    list_documents as list_corpus_documents,
    remove_document as remove_corpus_document,
)
from services.diff_view import apply_diff
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
from ui.dictation_controller import DictationController
from ui.dictation_panel import build_dictation_answer_review_popup, build_dictation_panel
from ui.async_event_helper import clear_event_queue, drain_event_queue, emit_event
from ui.find_async import start_find_import_task, start_find_search_task
from ui.find_controller import (
    build_find_clear_filter_status,
    build_find_import_start_state,
    build_find_search_start_state,
)
from ui.find_panel import build_find_window
from ui.find_presenter import (
    build_find_corpus_summary_state,
    build_find_import_completion_message,
    build_find_import_status,
    build_find_preview_state,
    build_find_search_result_state,
    build_find_search_status,
    get_selected_find_document,
)
from ui.history_presenter import build_history_list_state, build_rename_history_target, get_selected_history_item
from ui.list_presenter import build_dictation_list_state, build_word_table_values
from ui.manual_words_presenter import (
    normalize_import_word_text,
    normalize_manual_input_rows,
    parse_manual_rows,
    parse_clipboard_html_rows,
    parse_tabular_text_rows,
    read_clipboard_import_rows,
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
from ui.passage_presenter import (
    build_generated_passage_state,
    build_partial_passage_state,
    build_passage_audio_status,
    build_passage_practice_check_state,
    build_passage_practice_state,
    normalize_answer,
    speech_text_from_passage,
)
from ui.recent_wrong_controller import RecentWrongController
from ui.sidebar_panels import build_history_tab, build_tools_tab
from ui.word_list_panel import build_main_shell, build_word_list_panel
from ui.word_list_controller import WordListController
from ui.word_metadata_async import (
    start_analysis_task,
    start_phonetic_task,
    start_single_phonetic_task,
    start_single_translation_task,
    start_translation_task,
)
from ui.word_metadata_presenter import (
    build_render_words_state,
    can_apply_batch_metadata,
    can_apply_single_translation,
    normalize_requested_words,
)
from ui.word_table_helper import refresh_word_table_rows
from ui.word_tools_async import (
    start_passage_generation_task,
    start_sentence_generation_task,
    start_synonym_lookup_task,
)
from ui.word_tools_panel import build_sentence_window, build_synonym_window
from ui.word_tools_presenter import build_sentence_view_state, build_synonym_view_state


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
        "update_started": "更新程序已经启动。当前窗口会关闭，并在更新完成后自动打开新版本。",
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
        "feedback": "反馈",
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
        "dictation_wrong_answer": "错误。答案：{word}",
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
        "update_started": "The updater has started. This window will close and the new version will launch when the update finishes.",
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
        "feedback": "Feedback",
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
        "dictation_wrong_answer": "Wrong. Answer: {word}",
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

        self.order_mode = tk.StringVar(value="order")
        self.interval_var = tk.DoubleVar(value=2.0)
        self.volume_var = tk.IntVar(value=80)
        self.dictation_volume_var = tk.IntVar(value=100)
        self.speech_rate_var = tk.DoubleVar(value=1.0)
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
        self.dictation_volume_btn = None
        self.dictation_volume_popup = None
        self.dictation_volume_scale = None
        self.dictation_volume_value_label = None
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
        self.tts_status_request = 0
        self.word_table = None
        self.word_table_scroll = None
        self.word_context_menu = None
        self.dictation_context_menu = None
        self.history_context_menu = None
        self.word_action_index = None
        self.word_action_word = ""
        self.word_action_origin = "main"
        self.word_edit_entry = None
        self.word_edit_row = None
        self.word_edit_column = None
        self.suppress_word_select_action = False
        self.suppress_dictation_select_action = False
        self.last_word_speak_index = None
        self.last_word_speak_at = 0.0
        self.last_dictation_preview_key = None
        self.last_dictation_preview_at = 0.0
        self.sentence_window = None
        self.synonym_window = None
        self.manual_words_window = None
        self.manual_words_table = None
        self.manual_words_table_scroll = None
        self.manual_preview_edit_entry = None
        self.manual_preview_edit_item = None
        self.manual_preview_edit_column = None
        self.sentence_generation_token = 0
        self.sentence_generation_active_token = 0
        self.sentence_event_queue = queue.Queue()
        self.synonym_lookup_token = 0
        self.synonym_lookup_active_token = 0
        self.synonym_event_queue = queue.Queue()
        self.translations = {}
        self.word_pos = {}
        self.word_phonetics = {}
        self.pending_translation_words = set()
        self.pending_analysis_words = set()
        self.pending_phonetic_words = set()
        self.translation_token = 0
        self.analysis_token = 0
        self.phonetic_token = 0
        self.manual_source_dirty = False
        self.ui_language_var = tk.StringVar(value=get_ui_language())
        self.passage_window = None
        self.passage_text = None
        self.passage_status_var = tk.StringVar(value="Load words and click Generate.")
        self.gemini_model_var = tk.StringVar(value=get_generation_model() or DEFAULT_GEMINI_MODEL)
        self.gemini_model_combo = None
        self.gemini_model_values = []
        self.llm_api_provider_var = tk.StringVar(value=self.tr("provider_gemini"))
        self.tts_api_provider_var = tk.StringVar(value=self._tts_provider_label(get_tts_api_provider()))
        self.gemini_verified = False
        self.api_key_window = None
        self.api_key_force_llm = False
        self.api_key_force_tts = False
        self.gemini_key_var = tk.StringVar(value=get_llm_api_key())
        self.gemini_key_status_var = tk.StringVar(value="Paste your LLM API key, then test it.")
        self.tts_key_var = tk.StringVar(value=get_tts_api_key())
        self.tts_key_status_var = tk.StringVar(value="Paste your TTS API key, then test it.")
        self.api_key_test_btn = None
        self.api_llm_entry = None
        self.api_tts_entry = None
        self.gemini_runtime_status_var = tk.StringVar(value="")
        self.gemini_retry_status_var = tk.StringVar(value="")
        self.gemini_status_after = None
        self.gemini_key_test_btn = None
        self.tts_key_test_btn = None
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
        self.find_import_btn = None
        self.find_docs_context_menu = None
        self.find_doc_items = []
        self.find_result_items = {}
        self.find_task_queue = queue.Queue()
        self.find_task_token = 0
        self.find_active_token = 0
        self.audio_precache_token = 0
        self.dictation_mode_var = tk.StringVar(value="online_spelling")
        self.dictation_feedback_var = tk.StringVar(value="live")
        self.dictation_speed_var = tk.StringVar(value="1.0")
        self.dictation_status_var = tk.StringVar(value="Recent mistake list")
        self.dictation_timer_var = tk.StringVar(value="5s")
        self.dictation_progress_var = tk.StringVar(value="Spelling (0/0)")
        self.dictation_summary_var = tk.StringVar(value="")
        self.dictation_recent_list = None
        self.dictation_recent_items = []
        self.dictation_all_items = []
        self.dictation_list_mode_var = tk.StringVar(value="recent")
        self.dictation_all_tab_var = tk.StringVar(value="All (0)")
        self.dictation_recent_tab_var = tk.StringVar(value="Recent Wrong (0)")
        self.dictation_setup_frame = None
        self.dictation_session_frame = None
        self.dictation_result_frame = None
        self.dictation_mode_popup = None
        self.dictation_mode_buttons = []
        self.dictation_input = None
        self.dictation_result_label = None
        self.dictation_progress = None
        self.dictation_timer_after = None
        self.dictation_feedback_after = None
        self.dictation_play_after = None
        self.dictation_pool = []
        self.dictation_index = -1
        self.dictation_current_word = ""
        self.dictation_wrong_items = []
        self.dictation_session_attempts = []
        self.dictation_correct_count = 0
        self.dictation_answer_revealed = False
        self.dictation_running = False
        self.dictation_paused = False
        self.dictation_seconds_left = 0
        self.dictation_session_source_path = None
        self.dictation_session_list_mode = "all"
        self.dictation_previous_session_accuracy = self.store.get_last_dictation_accuracy()
        self.dictation_answer_review_popup = None
        self.dictation_answer_review_tree = None
        self.dictation_answer_review_accuracy_var = None
        self.dictation_answer_review_last_var = None
        self.dictation_answer_review_filter_var = None
        self.dictation_answer_review_show_wrong_only = False
        self.dictation_result_review_tree = None
        self.dictation_result_accuracy_var = None
        self.dictation_result_last_var = None
        self.dictation_result_filter_var = None
        self.right_notebook = None
        self.review_tab = None
        self.check_tab = None
        self.check_panel = None
        self.history_tab = None
        self.tools_tab = None
        self.dictation_window = None
        self.detail_word_var = tk.StringVar(value="No word selected")
        self.detail_note_var = tk.StringVar(value="Select a word to see notes and translation.")
        self.detail_translation_var = tk.StringVar(value="")
        self.detail_meta_var = tk.StringVar(value="Import a list to begin.")
        self.review_source_var = tk.StringVar(value="Source file: none")
        self.review_stats_var = tk.StringVar(value="Words: 0")
        self.review_focus_var = tk.StringVar(value="Focus: select a word or start playback.")
        self.tools_hint_var = tk.StringVar(value="Open tools from here instead of hunting across the window.")
        self.detail_speak_btn = None
        self.detail_sentence_btn = None
        self.detail_find_btn = None
        self.review_open_source_btn = None
        self.tools_sentence_btn = None
        self.tools_passage_btn = None
        self.tools_find_btn = None
        self.tools_settings_btn = None
        self.tools_update_btn = None
        self.tools_export_cache_btn = None
        self.tools_import_cache_btn = None
        self.tools_export_resource_pack_btn = None
        self.tools_import_resource_pack_btn = None
        self.save_as_btn = None
        self.new_list_btn = None
        self.detail_edit_btn = None
        self.detail_card = None
        self.saved_window_geometry = ""
        self.saved_window_minsize = None
        self.hidden_notebook_tabs = []

        self.build_ui()
        try:
            self.winfo_toplevel().protocol("WM_DELETE_WINDOW", self.on_main_window_close)
        except Exception:
            pass
        self.refresh_history()
        self.update_empty_state()
        self.update_speed_buttons()
        self.update_speech_rate_buttons()
        self.update_play_button()
        self.update_right_visibility()
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
        if self.dictation_window and self.dictation_window.winfo_exists():
            self.dictation_window.deiconify()
            self.dictation_window.lift()
            self.refresh_dictation_recent_list()
            return

        self._stop_main_word_playback()
        self.dictation_window = tk.Toplevel(self)
        self.dictation_window.title("Dictation")
        self.dictation_window.configure(bg="#f6f7fb")
        self.dictation_window.minsize(640, 620)
        self.dictation_window.geometry("720x700")
        self.dictation_window.protocol("WM_DELETE_WINDOW", self.close_dictation_window)

        self.check_panel = ttk.Frame(self.dictation_window, style="Card.TFrame")
        self.check_panel.pack(fill="both", expand=True, padx=12, pady=12)
        self.check_panel.grid_columnconfigure(0, weight=1)
        self.check_panel.grid_rowconfigure(0, weight=1)
        build_dictation_panel(self, self.check_panel)
        self.refresh_dictation_recent_list()
        self._show_dictation_frame(self.dictation_setup_frame)

    def close_dictation_window(self):
        self.close_dictation_mode_picker()
        self.close_dictation_answer_review_popup()
        self.close_dictation_volume_popup()
        self._cancel_dictation_timer()
        self._cancel_dictation_feedback_reset()
        if self.dictation_window and self.dictation_window.winfo_exists():
            self.dictation_window.destroy()
        self.dictation_window = None
        self.check_panel = None
        self.dictation_setup_frame = None
        self.dictation_session_frame = None
        self.dictation_result_frame = None
        self.dictation_recent_list = None
        self.dictation_mode_hint_label = None
        self.dictation_input = None
        self.dictation_result_label = None
        self.dictation_progress = None
        self.dictation_answer_review_tree = None
        self.play_btn_check = None
        self.dictation_volume_btn = None
        self.dictation_speed_buttons = []
        self.dictation_feedback_buttons = []
        self.dictation_play_after = None

    def _has_unsaved_manual_words(self):
        return bool(self.manual_source_dirty and self.store.words and not self.store.has_current_source_file())

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
        initial_name = os.path.basename(current_source) if current_source else "word_list.csv"
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
        if not self.store.has_current_source_file():
            tts_cleanup_manual_session_cache()
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
                current_source_path=self.store.get_current_source_path(),
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
            current_source_path=self.store.get_current_source_path(),
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
        )

    def _start_analysis_job(self, words, token):
        requested_words = normalize_requested_words(words)
        if not requested_words:
            return
        self.pending_analysis_words.update(requested_words)
        start_analysis_task(
            requested_words=requested_words,
            token=token,
            after=self.after,
            on_complete=self._apply_pos_analysis,
        )

    def _start_phonetic_job(self, words, token):
        requested_words = normalize_requested_words(words)
        if not requested_words:
            return
        self.pending_phonetic_words.update(requested_words)
        start_phonetic_task(
            requested_words=requested_words,
            token=token,
            after=self.after,
            on_complete=self._apply_phonetics,
        )

    def _apply_pos_analysis(self, token, requested_words, analyzed):
        for word in requested_words or []:
            self.pending_analysis_words.discard(word)
        if not can_apply_batch_metadata(
            token=token,
            active_token=self.analysis_token,
            has_word_table=bool(self.word_table),
        ):
            return
        self.word_pos.update(analyzed)
        refresh_word_table_rows(
            table=self.word_table,
            words=self.store.words,
            notes=self.store.notes,
            build_values=self._build_word_table_values,
        )
        self._refresh_selection_details()

    def _apply_phonetics(self, token, requested_words, phonetics):
        for word in requested_words or []:
            self.pending_phonetic_words.discard(word)
        if not can_apply_batch_metadata(
            token=token,
            active_token=self.phonetic_token,
            has_word_table=bool(self.word_table),
        ):
            return
        self.word_phonetics.update(phonetics)
        self._refresh_selection_details()

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
        if not self.dictation_recent_list:
            return
        state = build_dictation_list_state(
            words=self.store.words,
            notes=self.store.notes,
            recent_items=self.store.recent_wrong_words(limit=100),
            mode=self.dictation_list_mode_var.get(),
            word_pos=self.word_pos,
            translations=self.translations,
            tr=self.tr,
        )
        self.dictation_all_items = state.all_items
        self.dictation_recent_items = state.recent_items
        self.dictation_all_tab_var.set(state.all_tab_label)
        self.dictation_recent_tab_var.set(state.recent_tab_label)
        self.dictation_recent_list.delete(*self.dictation_recent_list.get_children())
        if state.rows:
            for row_id, values, tag in state.rows:
                self.dictation_recent_list.insert("", tk.END, iid=row_id, values=values, tags=(tag,))
            self.suppress_dictation_select_action = True
            self.dictation_recent_list.selection_set("0")
            self.dictation_recent_list.focus("0")
        else:
            self.dictation_recent_list.insert("", tk.END, iid=state.empty_row[0], values=state.empty_row[1])
        self.set_dictation_list_mode(self.dictation_list_mode_var.get(), refresh=False)

    def _get_dictation_source_items(self):
        if self.dictation_list_mode_var.get() == "recent":
            return list(self.dictation_recent_items)
        return list(self.dictation_all_items)

    def set_dictation_list_mode(self, mode, refresh=True):
        target = "recent" if str(mode or "").strip().lower() == "recent" else "all"
        self.dictation_list_mode_var.set(target)
        if hasattr(self, "dictation_all_tab_btn") and self.dictation_all_tab_btn:
            self.dictation_all_tab_btn.config(
                style="SelectedSpeed.TButton" if target == "all" else "Speed.TButton"
            )
        if hasattr(self, "dictation_recent_tab_btn") and self.dictation_recent_tab_btn:
            self.dictation_recent_tab_btn.config(
                style="SelectedSpeed.TButton" if target == "recent" else "Speed.TButton"
            )
        if self.dictation_recent_list:
            self.dictation_recent_list.heading("meta", text=(self.tr("error_type") if target == "recent" else self.tr("notes")))
        if target == "recent":
            self.dictation_status_var.set(self.tr("dictation_recent_hint"))
        else:
            self.dictation_status_var.set(self.tr("dictation_all_hint"))
        if refresh:
            self.refresh_dictation_recent_list()

    def _show_dictation_frame(self, target):
        for frame in (
            self.dictation_setup_frame,
            self.dictation_session_frame,
            self.dictation_result_frame,
        ):
            if not frame:
                continue
            if frame is target:
                frame.grid()
            else:
                frame.grid_remove()

    def open_dictation_mode_picker(self, auto_start=True):
        has_recent_wrong = bool(self.store.recent_wrong_words(limit=1))
        if not self.store.words and not has_recent_wrong:
            self.show_info("import_words_first")
            return
        if self.dictation_mode_popup and self.dictation_mode_popup.winfo_exists():
            self.dictation_mode_popup.lift()
            return

        parent = self.dictation_window if self.dictation_window and self.dictation_window.winfo_exists() else self
        win = tk.Toplevel(parent)
        self.dictation_mode_popup = win
        win.title(self.tr("mode_picker_title"))
        win.configure(bg="#f6f7fb")
        win.resizable(False, False)
        win.transient(parent)
        win.grab_set()

        wrap = ttk.Frame(win, style="Card.TFrame")
        wrap.pack(fill="both", expand=True, padx=12, pady=12)
        wrap.grid_columnconfigure(0, weight=1)

        ttk.Label(wrap, text=self.tr("mode_picker_title"), style="Card.TLabel").grid(row=0, column=0, sticky="w")

        mode_row = ttk.Frame(wrap, style="Card.TFrame")
        mode_row.grid(row=1, column=0, sticky="ew", pady=(10, 12))
        self.dictation_mode_buttons = []
        for idx, (value, label_key) in enumerate(
            (
                ("quiz", "mode_quiz"),
                ("word_mode", "mode_word_mode"),
                ("answer_review", "mode_answer_review"),
                ("online_spelling", "mode_online_spelling"),
            )
        ):
            btn = ttk.Button(mode_row, text=self.tr(label_key), command=lambda v=value: self.set_dictation_mode(v))
            btn.grid(row=0, column=idx, padx=(0 if idx == 0 else 6, 0), sticky="ew")
            mode_row.grid_columnconfigure(idx, weight=1)
            self.dictation_mode_buttons.append((value, btn))

        options_card = ttk.Frame(wrap, style="Card.TFrame")
        options_card.grid(row=2, column=0, sticky="ew")
        options_card.grid_columnconfigure(0, weight=1)

        ttk.Label(options_card, text=self.tr("playback_speed"), style="Card.TLabel").grid(row=0, column=0, sticky="w")
        speed_row = ttk.Frame(options_card, style="Card.TFrame")
        speed_row.grid(row=1, column=0, sticky="w", pady=(6, 10))
        self.dictation_speed_buttons = []
        for idx, value in enumerate(("1.0", "1.2", "1.4", "1.6", "adaptive")):
            text = self.tr("adaptive_speed") if value == "adaptive" else f"x{value}"
            btn = ttk.Button(speed_row, text=text, command=lambda v=value: self.set_dictation_speed(v))
            btn.grid(row=0, column=idx, padx=(0 if idx == 0 else 6, 0))
            self.dictation_speed_buttons.append((value, btn))

        ttk.Label(options_card, text=self.tr("feedback"), style="Card.TLabel").grid(row=2, column=0, sticky="w")
        feedback_row = ttk.Frame(options_card, style="Card.TFrame")
        feedback_row.grid(row=3, column=0, sticky="w", pady=(6, 0))
        self.dictation_feedback_buttons = []
        for idx, (value, text_key) in enumerate((("none", "no_live_feedback"), ("live", "live_feedback"))):
            btn = ttk.Button(feedback_row, text=self.tr(text_key), command=lambda v=value: self.set_dictation_feedback(v))
            btn.grid(row=0, column=idx, padx=(0 if idx == 0 else 6, 0))
            self.dictation_feedback_buttons.append((value, btn))

        self.dictation_mode_hint_label = ttk.Label(
            wrap,
            text="",
            style="Card.TLabel",
            foreground="#667085",
            wraplength=420,
            justify="left",
        )
        self.dictation_mode_hint_label.grid(row=3, column=0, sticky="w", pady=(10, 0))

        btn_row = ttk.Frame(wrap, style="Card.TFrame")
        btn_row.grid(row=4, column=0, sticky="ew", pady=(14, 0))
        btn_row.grid_columnconfigure(0, weight=1)
        btn_row.grid_columnconfigure(1, weight=1)
        ttk.Button(btn_row, text=self.tr("cancel"), command=self.close_dictation_mode_picker).grid(
            row=0, column=0, padx=(0, 6), sticky="ew"
        )
        ttk.Button(
            btn_row,
            text=self.tr("confirm"),
            style="Primary.TButton",
            command=lambda a=auto_start: self.confirm_dictation_mode_picker(auto_start=a),
        ).grid(row=0, column=1, padx=(6, 0), sticky="ew")

        self.set_dictation_mode(self.dictation_mode_var.get())
        self.set_dictation_speed(self.dictation_speed_var.get())
        self.set_dictation_feedback(self.dictation_feedback_var.get())

    def close_dictation_mode_picker(self):
        if self.dictation_mode_popup and self.dictation_mode_popup.winfo_exists():
            try:
                self.dictation_mode_popup.grab_release()
            except Exception:
                pass
            self.dictation_mode_popup.destroy()
        self.dictation_mode_popup = None

    def set_dictation_mode(self, mode):
        selected = str(mode or "online_spelling").strip().lower()
        self.dictation_mode_var.set(selected)
        for value, btn in getattr(self, "dictation_mode_buttons", []):
            btn.config(style="SelectedSpeed.TButton" if value == selected else "Speed.TButton")
        hint = ""
        if selected != "online_spelling":
            hint = self.tr("mode_not_ready")
        if getattr(self, "dictation_mode_hint_label", None):
            self.dictation_mode_hint_label.config(text=hint)

    def confirm_dictation_mode_picker(self, auto_start=True):
        selected = self.dictation_mode_var.get()
        if selected != "online_spelling":
            self.set_dictation_mode("online_spelling")
        self.close_dictation_mode_picker()
        if auto_start:
            self._show_dictation_frame(self.dictation_session_frame)
            self.start_online_spelling_session()

    def start_dictation_from_selected_word(self):
        if not self.dictation_recent_list:
            return
        selection = self.dictation_recent_list.selection()
        if not selection or selection[0] == "empty":
            self.show_info("select_word_first")
            return
        try:
            start_index = int(selection[0])
        except Exception:
            start_index = 0
        self._show_dictation_frame(self.dictation_session_frame)
        self.set_dictation_speed(self.dictation_speed_var.get())
        self.set_dictation_feedback(self.dictation_feedback_var.get())
        self.start_online_spelling_session(start_index=start_index)

    def _get_dictation_pool(self):
        items = self._get_dictation_source_items()
        if items:
            words = [str(item.get("word") or "").strip() for item in items if str(item.get("word") or "").strip()]
            if words:
                return words
        return list(self.store.words)

    def _get_recent_wrong_cache_source_path(self):
        return tts_get_recent_wrong_cache_source()

    def _get_dictation_preview_source_path(self):
        if self.dictation_list_mode_var.get() == "recent":
            return self._get_recent_wrong_cache_source_path()
        return self.store.get_current_source_path()

    def on_dictation_list_selected(self, _event=None):
        if self.suppress_dictation_select_action:
            self.suppress_dictation_select_action = False
            return
        if not self.dictation_recent_list:
            return
        selection = self.dictation_recent_list.selection()
        if not selection or selection[0] == "empty":
            return
        store_index = self._dictation_row_to_store_index(self.dictation_recent_list, row_id=selection[0])
        selected_word = ""
        try:
            view_index = int(selection[0])
            items = self._get_dictation_source_items()
            if 0 <= view_index < len(items):
                selected_word = str(items[view_index].get("word") or "").strip()
        except Exception:
            selected_word = ""
        if store_index is not None and store_index < len(self.store.words):
            self._set_word_action_context(store_index, origin="dictation", word=selected_word)
        else:
            self._set_word_action_context(None, origin="dictation", word=selected_word)
        self._refresh_selection_details()
        self._speak_dictation_preview(store_index=store_index)

    def on_dictation_list_click_play(self, event=None):
        tree = self.dictation_recent_list
        if not tree:
            return "break"
        row_id = str(tree.identify_row(event.y) or "").strip() if event is not None else ""
        if not row_id or row_id == "empty":
            return "break"
        store_index = self._dictation_row_to_store_index(tree, row_id=row_id)
        if store_index is not None and 0 <= store_index < len(self.store.words):
            word = self.store.words[store_index]
            self._set_word_action_context(store_index, origin="dictation")
        else:
            try:
                view_index = int(row_id)
            except Exception:
                return "break"
            items = self._get_dictation_source_items()
            if view_index < 0 or view_index >= len(items):
                return "break"
            word = str(items[view_index].get("word") or "").strip()
            if not word:
                return "break"
        self._speak_dictation_preview(word=word, store_index=store_index)
        return "break"

    def _speak_dictation_preview(self, word=None, store_index=None):
        preview_word = str(word or "").strip()
        if not preview_word and store_index is not None and 0 <= store_index < len(self.store.words):
            preview_word = str(self.store.words[store_index] or "").strip()
        if not preview_word:
            return
        source_path = self._get_dictation_preview_source_path()
        preview_key = (str(source_path or "").strip(), str(store_index if store_index is not None else preview_word).strip().lower())
        now = time.time()
        if self.last_dictation_preview_key == preview_key and (now - self.last_dictation_preview_at) < 0.35:
            return
        self.last_dictation_preview_key = preview_key
        self.last_dictation_preview_at = now
        runtime = tts_get_runtime_label()
        cached = get_voice_source() == SOURCE_GEMINI and tts_has_cached_word_audio(
            preview_word,
            source_path=source_path,
        )
        token = speak_async(
            preview_word,
            self._dictation_playback_volume_ratio(),
            rate_ratio=self.speech_rate_var.get(),
            cancel_before=True,
            source_path=source_path,
        )
        if cached:
            self.status_var.set(f"Playing cached audio for '{preview_word}'.")
        else:
            self.status_var.set(f"Generating '{preview_word}' with {runtime}...")
        self._watch_tts_backend(token, target="status", text_label=preview_word)

    def set_dictation_speed(self, value):
        self.dictation_speed_var.set(str(value))
        for current, btn in getattr(self, "dictation_speed_buttons", []):
            btn.config(style="SelectedSpeed.TButton" if current == self.dictation_speed_var.get() else "Speed.TButton")

    def set_dictation_feedback(self, value):
        self.dictation_feedback_var.set(str(value))
        for current, btn in getattr(self, "dictation_feedback_buttons", []):
            btn.config(style="SelectedSpeed.TButton" if current == self.dictation_feedback_var.get() else "Speed.TButton")

    def _dictation_seconds_for_speed(self):
        mapping = {"1.0": 5, "1.2": 4, "1.4": 3, "1.6": 2, "adaptive": 0}
        return int(mapping.get(self.dictation_speed_var.get(), 5))

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
        self._stop_main_word_playback()
        pool = self._get_dictation_pool()
        if not pool:
            self.show_info("no_words_for_dictation")
            self.reset_dictation_view()
            return
        state = self.dictation_controller.build_session_state(
            pool=pool,
            list_mode=self.dictation_list_mode_var.get(),
            session_source_path=self._get_dictation_preview_source_path(),
            start_index=start_index,
        )
        self.dictation_pool = list(state.pool)
        self.dictation_session_list_mode = state.session_list_mode
        self.dictation_session_source_path = state.session_source_path
        self.dictation_previous_session_accuracy = state.previous_accuracy
        self.dictation_index = state.index
        self.dictation_wrong_items = list(state.wrong_items)
        self.dictation_session_attempts = list(state.attempts)
        self.dictation_correct_count = state.correct_count
        self.dictation_current_word = state.current_word
        self.dictation_answer_revealed = state.answer_revealed
        self.dictation_running = state.running
        self.dictation_paused = state.paused
        self.dictation_summary_var.set(state.summary_text)
        self.update_dictation_play_button()
        self._show_dictation_frame(self.dictation_session_frame)
        self.advance_dictation_word(initial=True)

    def play_dictation_current_word(self):
        if not self.dictation_running or not self.dictation_current_word:
            return
        self._cancel_dictation_play_start()
        self.dictation_paused = False
        self.update_dictation_play_button()
        self.status_var.set(self.trf("dictation_playing", word=self.dictation_current_word))
        delay_ms = 120

        def _start_playback():
            self.dictation_play_after = None
            if not self.dictation_running or self.dictation_paused or not self.dictation_current_word:
                return
            speak_async(
                self.dictation_current_word,
                self._dictation_playback_volume_ratio(),
                rate_ratio=1.0 if self.dictation_speed_var.get() == "adaptive" else float(self.dictation_speed_var.get()),
                cancel_before=True,
                source_path=self.dictation_session_source_path or self._get_dictation_preview_source_path(),
            )
            self._restart_dictation_timer()
            self._focus_dictation_input()

        self.dictation_play_after = self.after(delay_ms, _start_playback)

    def replay_dictation_word(self):
        if not self.dictation_current_word:
            return
        self.play_dictation_current_word()

    def toggle_dictation_play_pause(self):
        if not self.dictation_running or not self.dictation_current_word:
            return
        if self.dictation_paused:
            self.play_dictation_current_word()
        else:
            self.pause_dictation_session()

    def pause_dictation_session(self):
        self.dictation_paused = True
        self._cancel_dictation_play_start()
        tts_cancel_all()
        self._cancel_dictation_timer()
        self.update_dictation_play_button()
        self.dictation_status_var.set(self.tr("dictation_paused"))
        self._focus_dictation_input()

    def update_dictation_play_button(self):
        if not self.play_btn_check:
            return
        if self.dictation_running and not self.dictation_paused:
            self.play_btn_check.config(text=f"⏸ {self.tr('pause')}")
        else:
            self.play_btn_check.config(text=self.tr("play"))

    def previous_dictation_word(self):
        if not self.dictation_running or not self.dictation_pool:
            return
        self._cancel_dictation_play_start()
        self._cancel_dictation_feedback_reset()
        self._cancel_dictation_timer()
        tts_cancel_all()
        self.dictation_paused = False

        if self.dictation_answer_revealed:
            target_position = self.dictation_index
        else:
            target_position = self.dictation_index - 1
        if target_position < 0:
            target_position = 0

        invalidated_attempt = None
        for idx, item in enumerate(self.dictation_session_attempts):
            if int(item.get("position", -1)) == int(target_position):
                invalidated_attempt = self.dictation_session_attempts.pop(idx)
                break
        if invalidated_attempt:
            self.dictation_controller.revert_attempt(
                invalidated_attempt,
                recent_wrong_source_path=self._get_recent_wrong_cache_source_path(),
            )
            if invalidated_attempt.get("correct"):
                self.dictation_correct_count = max(0, self.dictation_correct_count - 1)
            self.dictation_wrong_items = [
                item
                for item in self.dictation_wrong_items
                if int(item.get("position", -1)) != int(target_position)
            ]
            self.refresh_dictation_recent_list()
            self._refresh_dictation_answer_review_popup()

        self.dictation_index = max(-1, target_position - 1)
        self.advance_dictation_word(initial=True)

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
        self._cancel_dictation_timer()
        self.dictation_seconds_left = self._dictation_seconds_for_speed()
        if self.dictation_seconds_left <= 0:
            self.dictation_timer_var.set("")
            return
        self.dictation_timer_var.set(f"{self.dictation_seconds_left}s")
        self.dictation_timer_after = self.after(1000, self._tick_dictation_timer)

    def _tick_dictation_timer(self):
        if self.dictation_paused or not self.dictation_running:
            return
        self.dictation_seconds_left -= 1
        if self.dictation_seconds_left <= 0:
            self.dictation_timer_var.set("0s")
            self.finalize_dictation_attempt(trigger="timeout")
            return
        self.dictation_timer_var.set(f"{self.dictation_seconds_left}s")
        self.dictation_timer_after = self.after(1000, self._tick_dictation_timer)

    def on_dictation_input_change(self, _event=None):
        if not self.dictation_running or not self.dictation_current_word or not self.dictation_input:
            return
        value = str(self.dictation_input.get() or "").strip()
        target = str(self.dictation_current_word or "").strip()
        if self.dictation_feedback_var.get() != "live":
            self._set_dictation_input_color("neutral")
            return
        if not value:
            self._set_dictation_input_color("neutral")
            return
        value_key = self._normalize_dictation_compare_text(value)
        target_key = self._normalize_dictation_compare_text(target)
        if value_key == target_key:
            self._set_dictation_input_color("correct")
            self.finalize_dictation_attempt(trigger="input")
            return
        if target_key.startswith(value_key):
            self._set_dictation_input_color("neutral")
            self.dictation_status_var.set(self.tr("dictation_keep_spelling"))
            return
        self._set_dictation_input_color("wrong")
        self.dictation_status_var.set(self.tr("dictation_wrong_live"))

    def on_dictation_enter(self, _event=None):
        if not self.dictation_running or not self.dictation_current_word:
            return "break"
        if self.dictation_feedback_after or self.dictation_answer_revealed:
            return "break"
        self.finalize_dictation_attempt(trigger="manual")
        return "break"

    def _normalize_dictation_compare_text(self, text):
        raw = unicodedata.normalize("NFKC", str(text or "").casefold())
        raw = raw.replace("£", " ").replace("$", " ").replace("€", " ").replace("¥", " ")
        raw = raw.replace("'", "").replace('"', "")
        raw = raw.replace("-", " ").replace("/", " ").replace("\\", " ")
        raw = re.sub(r"[.,:;!?()\[\]{}]", " ", raw)
        raw = re.sub(r"\s+", " ", raw).strip()
        compact = re.sub(r"[^0-9a-z ]+", "", raw)
        return re.sub(r"\s+", " ", compact).strip()

    def finalize_dictation_attempt(self, trigger="manual"):
        if not self.dictation_running or self.dictation_answer_revealed or not self.dictation_current_word:
            return
        self._cancel_dictation_timer()
        self.dictation_answer_revealed = True
        user_text = str(self.dictation_input.get() or "").strip() if self.dictation_input else ""
        target = str(self.dictation_current_word or "").strip()
        is_correct = self._normalize_dictation_compare_text(user_text) == self._normalize_dictation_compare_text(target)
        if is_correct:
            self.dictation_correct_count += 1
            self._set_dictation_input_color("correct")
            self.dictation_status_var.set(self.tr("dictation_correct"))
        else:
            if self.dictation_input and self.dictation_input.winfo_exists():
                try:
                    self.dictation_input.delete(0, tk.END)
                    self.dictation_input.insert(0, target)
                    self.dictation_input.select_range(0, tk.END)
                    self.dictation_input.icursor(tk.END)
                except Exception:
                    pass
            self._set_dictation_input_color("wrong")
            self.dictation_status_var.set(self.trf("dictation_wrong_answer", word=target))
        if self.dictation_session_frame and self.dictation_session_frame.winfo_exists():
            try:
                self.dictation_session_frame.update_idletasks()
            except Exception:
                pass
        result = self.dictation_controller.record_attempt(
            target=target,
            user_text=user_text,
            is_correct=is_correct,
            position=self.dictation_index,
            list_mode=self.dictation_session_list_mode,
            recent_wrong_source_path=self._get_recent_wrong_cache_source_path(),
            session_source_path=self.dictation_session_source_path,
        )
        attempt_entry = dict(result.attempt_entry)
        replaced = False
        for idx, item in enumerate(self.dictation_session_attempts):
            if int(item.get("position", -1)) == int(self.dictation_index):
                self.dictation_session_attempts[idx] = attempt_entry
                replaced = True
                break
        if not replaced:
            self.dictation_session_attempts.append(attempt_entry)
        if result.cleared_recent_wrong:
            self.refresh_dictation_recent_list()
        if result.appended_wrong_item:
            self.dictation_wrong_items.append(dict(result.appended_wrong_item))
        self._refresh_dictation_answer_review_popup()
        if not is_correct:
            delay = 2200
        elif trigger == "input":
            delay = 1150
        else:
            delay = 1450
        self._cancel_dictation_feedback_reset()
        self.dictation_feedback_after = self.after(delay, self._go_to_next_dictation_word)

    def _go_to_next_dictation_word(self):
        self.dictation_feedback_after = None
        self.advance_dictation_word()

    def advance_dictation_word(self, initial=False):
        if not self.dictation_running and not initial:
            return
        if not initial and self.dictation_current_word and not self.dictation_answer_revealed:
            self.finalize_dictation_attempt(trigger="manual")
            return

        self._cancel_dictation_play_start()
        self._cancel_dictation_feedback_reset()
        self._cancel_dictation_timer()
        self.dictation_index += 1
        total = len(self.dictation_pool)
        if self.dictation_index >= total:
            self.finish_dictation_session()
            return

        self.dictation_current_word = self.dictation_pool[self.dictation_index]
        self.dictation_answer_revealed = False
        progress_text = f"Spelling ({self.dictation_index + 1}/{total})"
        self.dictation_progress_var.set(progress_text)
        self.dictation_progress["value"] = ((self.dictation_index + 1) / max(1, total)) * 100.0
        self.dictation_status_var.set(self.tr("dictation_listen_type"))
        seconds_for_speed = self._dictation_seconds_for_speed()
        self.dictation_timer_var.set(f"{seconds_for_speed}s" if seconds_for_speed else "")
        if self.dictation_input:
            self.dictation_input.delete(0, tk.END)
            self.dictation_input.focus_set()
        self._set_dictation_input_color("neutral")
        self.update_dictation_play_button()
        self.play_dictation_current_word()

    def finish_dictation_session(self):
        self.dictation_running = False
        self.dictation_paused = False
        self._cancel_dictation_play_start()
        self._cancel_dictation_timer()
        self._cancel_dictation_feedback_reset()
        self.update_dictation_play_button()
        summary = self.dictation_controller.finish_session(
            correct_count=self.dictation_correct_count,
            total=len(self.dictation_pool),
        )
        self.dictation_summary_var.set(f"{summary.accuracy:.2f}%")
        self.dictation_status_var.set(self.tr("dictation_session_complete"))
        self._render_dictation_answer_review_views()
        self.refresh_dictation_recent_list()
        self._show_dictation_frame(self.dictation_result_frame)

    def reset_dictation_view(self):
        state = self.dictation_controller.build_reset_state()
        self.dictation_running = state.running
        self.dictation_paused = state.paused
        self.dictation_pool = list(state.pool)
        self.dictation_index = state.index
        self.dictation_current_word = state.current_word
        self.dictation_session_source_path = state.session_source_path
        self.dictation_session_list_mode = state.session_list_mode
        self.dictation_session_attempts = list(state.attempts)
        self.dictation_wrong_items = list(state.wrong_items)
        self.dictation_correct_count = state.correct_count
        self.dictation_answer_revealed = state.answer_revealed
        self._cancel_dictation_play_start()
        self._cancel_dictation_timer()
        self._cancel_dictation_feedback_reset()
        self.update_dictation_play_button()
        self.dictation_progress_var.set(state.progress_text)
        self.dictation_timer_var.set(state.timer_text)
        self.dictation_status_var.set(self.tr("dictation_recent_title"))
        self.dictation_progress["value"] = 0
        if self.dictation_input:
            self.dictation_input.delete(0, tk.END)
        self._set_dictation_input_color("neutral")
        if self.dictation_result_accuracy_var is not None:
            self.dictation_result_accuracy_var.set("0.00%")
        if self.dictation_result_last_var is not None:
            self.dictation_result_last_var.set("-")
        if self.dictation_result_filter_var is not None:
            self.dictation_result_filter_var.set(self.tr("show_wrong_only"))
        self._render_dictation_answer_review_tree(self.dictation_result_review_tree)
        self.close_dictation_answer_review_popup()
        self._show_dictation_frame(self.dictation_setup_frame)
        self.refresh_dictation_recent_list()

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
        self.status_var.set(self.tr("shared_cache_sync_downloading"))
        package_path = ""
        resource_pack_path = ""
        bundled_corpus_path = ""
        try:
            package_path = download_update_package(manifest.get("package_url"))
            result = tts_import_shared_audio_cache_package(package_path)

            release_base = self._release_asset_base_url()
            resource_pack_url = (
                str(current_info.get("word_resource_pack_url") or "").strip()
                or (f"{release_base}/wordspeaker_word_resource_pack.wspack" if release_base else "")
            )
            bundled_corpus_url = (
                str(current_info.get("bundled_corpus_package_url") or "").strip()
                or (f"{release_base}/wordspeaker_bundled_corpus.zip" if release_base else "")
            )
            if not resource_pack_url:
                raise RuntimeError("Official word resource pack URL is missing.")
            if not bundled_corpus_url:
                raise RuntimeError("Official bundled corpus package URL is missing.")

            if not self._prompt_save_unsaved_manual_words(title=self.tr("resource_pack_import_title")):
                return
            resource_pack_path = download_update_package(resource_pack_url)
            resource_pack_result = import_word_resource_pack(resource_pack_path)
            load_result = self._load_word_resource_pack_entries(resource_pack_result.get("entries") or [])
            if not load_result:
                raise RuntimeError("Official word resource pack contained no valid entries.")

            bundled_corpus_path = download_update_package(bundled_corpus_url)
            corpus_result = import_bundled_corpus_package(bundled_corpus_path)
        except Exception as exc:
            messagebox.showerror(
                self.tr("shared_cache_sync_title"),
                self.trf("shared_cache_sync_failed", error=str(exc)),
            )
            return
        finally:
            for path in (package_path, resource_pack_path, bundled_corpus_path):
                if not path:
                    continue
                try:
                    shutil.rmtree(os.path.dirname(path), ignore_errors=True)
                except Exception:
                    pass
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
        try:
            idx = int(index)
        except Exception:
            idx = None
        token = self._normalize_import_word_text(word or "")
        if idx is None or idx < 0 or idx >= len(self.store.words):
            self.word_action_index = None
            self.word_action_word = token
            self.word_action_origin = "main"
            if token:
                self.word_action_origin = str(origin or "main")
            return None
        self.word_action_index = idx
        self.word_action_word = str(self.store.words[idx] or "").strip()
        self.word_action_origin = str(origin or "main")
        return idx

    def _clear_word_action_context(self):
        self.word_action_index = None
        self.word_action_word = ""
        self.word_action_origin = "main"

    def _get_context_or_selected_index(self):
        if self.word_action_index is not None and 0 <= self.word_action_index < len(self.store.words):
            return self.word_action_index
        return self._get_selected_index()

    def _get_context_word(self):
        if self.word_action_index is not None and 0 <= self.word_action_index < len(self.store.words):
            return str(self.store.words[self.word_action_index] or "").strip()
        token = str(self.word_action_word or "").strip()
        if token:
            return token
        selected_idx = self._get_selected_index()
        if selected_idx is None or selected_idx >= len(self.store.words):
            return ""
        return str(self.store.words[selected_idx] or "").strip()

    def _get_context_audio_source_path(self):
        if self.word_action_origin == "dictation" and self.dictation_list_mode_var.get() == "recent":
            return self._get_recent_wrong_cache_source_path()
        return self.store.get_current_source_path()

    def _get_word_audio_override_source_path(self):
        return self.store.get_current_source_path() or self._get_context_audio_source_path()

    def _dictation_row_to_store_index(self, tree, row_id=None):
        if not tree:
            return None
        item_id = str(row_id or tree.focus() or "").strip()
        if not item_id or item_id == "empty":
            selection = tree.selection()
            if selection:
                item_id = str(selection[0] or "").strip()
        if not item_id or item_id == "empty":
            return None
        try:
            view_index = int(item_id)
        except Exception:
            return None
        items = self._get_dictation_source_items()
        if view_index < 0 or view_index >= len(items):
            return None
        word = str(items[view_index].get("word") or "").strip()
        if not word:
            return None
        try:
            return self.store.words.index(word)
        except ValueError:
            return None

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
        token = self.translation_token
        start_single_translation_task(
            word=word,
            row_idx=row_idx,
            token=token,
            after=self.after,
            on_complete=self._apply_single_translation,
        )

    def _phonetic_single_word_async(self, row_idx, word):
        token = self.phonetic_token
        start_single_phonetic_task(
            word=word,
            row_idx=row_idx,
            token=token,
            after=self.after,
            on_complete=self._apply_single_phonetic,
        )

    def _apply_single_translation(self, token, row_idx, word, zh_text):
        if not can_apply_single_translation(
            token=token,
            active_token=self.translation_token,
            row_idx=row_idx,
            word=word,
            current_words=self.store.words,
            has_word_table=bool(self.word_table),
        ):
            return
        self.translations[word] = zh_text
        iid = str(row_idx)
        if self.word_table.exists(iid):
            note = self.store.notes[row_idx] if row_idx < len(self.store.notes) else ""
            tag = "even" if row_idx % 2 == 0 else "odd"
            self.word_table.item(iid, values=self._build_word_table_values(row_idx, word, note), tags=(tag,))
        self._refresh_selection_details()

    def _apply_single_phonetic(self, token, row_idx, word, phonetic_text):
        if not can_apply_single_translation(
            token=token,
            active_token=self.phonetic_token,
            row_idx=row_idx,
            word=word,
            current_words=self.store.words,
            has_word_table=bool(self.word_table),
        ):
            return
        self.word_phonetics[word] = phonetic_text
        self._refresh_selection_details()

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
                self.word_table.selection_set(str(next_idx))
                self.word_table.focus(str(next_idx))
        self.status_var.set(self.trf("word_deleted", word=result.word))

    def render_words(self, words):
        if not self.word_table:
            return
        self.cancel_word_edit()
        self.translation_token += 1
        self.analysis_token += 1
        self.phonetic_token += 1
        token = self.translation_token
        analysis_token = self.analysis_token
        phonetic_token = self.phonetic_token
        cached = get_cached_translations(words)
        cached_pos = get_cached_pos(words)
        cached_phonetics = get_cached_phonetics(words)
        state = build_render_words_state(
            words=words,
            cached_translations=cached,
            cached_pos=cached_pos,
            cached_phonetics=cached_phonetics,
        )
        self.translations = dict(state["translations"])
        self.word_pos = dict(state["word_pos"])
        self.word_phonetics = dict(state["word_phonetics"])
        self.pending_translation_words.clear()
        self.pending_analysis_words.clear()
        self.pending_phonetic_words.clear()
        self.word_table.delete(*self.word_table.get_children())
        for idx, w in enumerate(words):
            note = self.store.notes[idx] if idx < len(self.store.notes) else ""
            tag = "even" if idx % 2 == 0 else "odd"
            self.word_table.insert(
                "",
                tk.END,
                iid=str(idx),
                values=self._build_word_table_values(idx, w, note),
                tags=(tag,),
            )
        self.update_empty_state()
        self._refresh_selection_details()
        self.refresh_dictation_recent_list()
        missing_words = state["missing_translations"]
        missing_pos = state["missing_pos"]
        missing_phonetics = state["missing_phonetics"]
        if missing_words:
            self._start_translation_job(missing_words, token)
        if missing_pos:
            self._start_analysis_job(missing_pos, analysis_token)
        if missing_phonetics:
            self._start_phonetic_job(missing_phonetics, phonetic_token)
        if words:
            self._start_audio_precache_job(words)

    def _start_translation_job(self, words, token):
        requested_words = normalize_requested_words(words)
        if not requested_words:
            return
        self.pending_translation_words.update(requested_words)
        start_translation_task(
            requested_words=requested_words,
            token=token,
            after=self.after,
            on_complete=self._apply_translations,
        )

    def _apply_translations(self, token, requested_words, translated):
        for word in requested_words or []:
            self.pending_translation_words.discard(word)
        if not can_apply_batch_metadata(
            token=token,
            active_token=self.translation_token,
            has_word_table=bool(self.word_table),
        ):
            return
        self.translations.update(translated)
        refresh_word_table_rows(
            table=self.word_table,
            words=self.store.words,
            notes=self.store.notes,
            build_values=self._build_word_table_values,
        )
        self._refresh_selection_details()

    def _ensure_word_metadata(self, word):
        target = str(word or "").strip()
        if not target:
            return
        if not str(self.word_pos.get(target) or "").strip() and target not in self.pending_analysis_words:
            self._start_analysis_job([target], self.analysis_token)
        if not str(self.translations.get(target) or "").strip() and target not in self.pending_translation_words:
            self._start_translation_job([target], self.translation_token)
        if not str(self.word_phonetics.get(target) or "").strip() and target not in self.pending_phonetic_words:
            self._start_phonetic_job([target], self.phonetic_token)

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
        if self.find_window and self.find_window.winfo_exists():
            self.find_window.lift()
            self._set_find_query_from_selection()
            return
        build_find_window(self)

        self._set_find_query_from_selection()
        self.refresh_find_corpus_summary()

    def _set_find_query_from_selection(self):
        word = self._get_context_word()
        if not word:
            return
        self.find_search_var.set(word)

    def refresh_find_corpus_summary(self):
        try:
            stats = corpus_stats()
            docs = list_corpus_documents(limit=200)
        except Exception as e:
            if self.find_docs_list:
                self.find_docs_list.delete(0, tk.END)
            self.find_status_var.set(str(e))
            return
        state = build_find_corpus_summary_state(stats, docs)
        if self.find_docs_list:
            self.find_docs_list.delete(0, tk.END)
            self.find_doc_items = list(docs)
            for label in state.doc_labels:
                self.find_docs_list.insert(tk.END, label)
        self.find_status_var.set(state.status_text)

    def on_find_docs_right_click(self, event):
        if not self.find_docs_list or not self.find_docs_context_menu:
            return
        index = self.find_docs_list.nearest(event.y)
        if index < 0 or index >= len(self.find_doc_items):
            return
        try:
            self.find_docs_list.selection_clear(0, tk.END)
            self.find_docs_list.selection_set(index)
            self.find_docs_list.activate(index)
        except Exception:
            pass
        try:
            self.find_docs_context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.find_docs_context_menu.grab_release()
        return "break"

    def delete_selected_corpus_document(self):
        if not self.find_docs_list or not self.find_doc_items:
            return
        selection = self.find_docs_list.curselection()
        if not selection:
            return
        idx = int(selection[0])
        if idx < 0 or idx >= len(self.find_doc_items):
            return
        item = self.find_doc_items[idx]
        path = str(item.get("path") or "").strip()
        name = str(item.get("name") or os.path.basename(path) or path)
        if not path:
            return
        if not messagebox.askyesno(self.tr("find_corpus_sentences"), self.trf("delete_corpus_doc_confirm", name=name)):
            return
        removed = remove_corpus_document(path)
        self.clear_find_document_filter()
        self.refresh_find_corpus_summary()
        self.run_find_search()
        if removed:
            messagebox.showinfo(self.tr("find_corpus_sentences"), self.trf("corpus_doc_deleted", name=name))

    def _clear_find_task_queue(self):
        clear_event_queue(self.find_task_queue)

    def _emit_find_task_event(self, event_type, token, payload=None):
        emit_event(self.find_task_queue, event_type, token, payload)

    def _poll_find_task_events(self, token):
        done = drain_event_queue(
            target_queue=self.find_task_queue,
            token=token,
            active_token=self.find_active_token,
            handlers={
                "import_done": lambda payload: self._apply_find_import_result(payload or {}),
                "search_done": lambda payload: self._apply_find_search_result(payload or {}),
                "error": lambda payload: self._handle_find_task_error(str(payload or "Unknown error")),
            },
        )
        if not done and token == self.find_active_token:
            self.after(80, lambda t=token: self._poll_find_task_events(t))

    def _handle_find_task_error(self, message):
        if self.find_import_btn:
            self.find_import_btn.state(["!disabled"])
        self.find_status_var.set(message)
        messagebox.showerror("Find Error", message)

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
        state = build_find_import_start_state(paths)
        if not state:
            return
        self.find_task_token += 1
        token = self.find_task_token
        self.find_active_token = token
        self._clear_find_task_queue()
        if self.find_import_btn:
            self.find_import_btn.state(["disabled"])
        self.find_status_var.set(state.status_text)
        start_find_import_task(paths=state.paths, token=token, emit_event=self._emit_find_task_event)
        self.after(80, lambda t=token: self._poll_find_task_events(t))

    def _apply_find_import_result(self, payload):
        self.refresh_find_corpus_summary()
        status, errors = build_find_import_status(payload)
        self.find_status_var.set(status)
        if self.find_import_btn:
            self.find_import_btn.state(["!disabled"])
        messagebox.showinfo("导入完成", build_find_import_completion_message(payload))
        if errors:
            messagebox.showerror("Import Warning", "\n".join(errors[:10]))

    def run_find_search(self):
        selected_doc = self._get_selected_find_document()
        try:
            state = build_find_search_start_state(
                query_text=self.find_search_var.get(),
                limit_text=self.find_limit_var.get(),
                selected_doc=selected_doc,
                status_builder=build_find_search_status,
            )
        except ValueError:
            messagebox.showinfo("Info", "Enter a word or phrase first.")
            return
        self.find_limit_var.set(state.limit_text)
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
        self.find_status_var.set(state.status_text)
        start_find_search_task(
            query=state.query,
            limit=state.limit,
            document_path=state.document_path,
            token=token,
            emit_event=self._emit_find_task_event,
        )
        self.after(80, lambda t=token: self._poll_find_task_events(t))

    def search_selected_word_in_corpus(self):
        if not self.find_window or not self.find_window.winfo_exists():
            self.open_find_window()
        self._set_find_query_from_selection()
        self.run_find_search()

    def _get_selected_find_document(self):
        if not self.find_docs_list:
            return None
        return get_selected_find_document(self.find_doc_items, self.find_docs_list.curselection())

    def clear_find_document_filter(self):
        if self.find_docs_list:
            self.find_docs_list.selection_clear(0, tk.END)
        self.find_status_var.set(build_find_clear_filter_status())

    def _apply_find_search_result(self, payload):
        state = build_find_search_result_state(payload=payload, doc_items=self.find_doc_items)
        self.find_result_items = state.result_items
        if self.find_results_table:
            self.find_results_table.delete(*self.find_results_table.get_children())
            for row_id, values in state.result_rows:
                self.find_results_table.insert("", tk.END, iid=row_id, values=values)
            if state.first_row_id:
                self.find_results_table.selection_set(state.first_row_id)
                self.find_results_table.focus(state.first_row_id)
                self._show_find_result_preview(state.first_row_id)
            else:
                self._clear_find_preview()
        else:
            self._clear_find_preview()
        self.find_status_var.set(state.status_text)

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
        state = build_find_preview_state(self.find_result_items.get(row_id))
        if not state or not self.find_preview_text:
            self._clear_find_preview()
            return
        text = self.find_preview_text
        text.configure(state="normal")
        text.delete("1.0", tk.END)
        text.insert("1.0", state.sentence)
        for start, end in state.highlight_ranges:
            if start < end:
                text.tag_add("hit", f"1.0+{int(start)}c", f"1.0+{int(end)}c")
        if state.source:
            text.insert(tk.END, "\n\n")
            source_start = text.index(tk.END)
            text.insert(tk.END, state.source)
            text.tag_add("source", source_start, tk.END)
        text.tag_configure("source", foreground="#666666")
        text.configure(state="disabled")

    # IELTS passage
    def open_passage_window(self):
        if self.passage_window and self.passage_window.winfo_exists():
            self.passage_window.lift()
            return
        build_passage_window(self, Tooltip)

        if self.current_passage:
            self._set_passage_text(self.current_passage)
        else:
            self._set_passage_text("")
        self.refresh_gemini_models()

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
                elif event_type == "success_tts":
                    self._finish_tts_validation_success(payload or {})
                elif event_type == "success_api_setup":
                    self._finish_combined_api_validation(payload or {})
                elif event_type == "error":
                    self._finish_gemini_validation_error(str(payload or "Unknown error"))
                elif event_type == "error_tts":
                    self._finish_tts_validation_error(str(payload or "Unknown error"))
                elif event_type == "done":
                    done = True
        except queue.Empty:
            pass
        if not done and token == self.gemini_validation_active_token:
            self.after(80, lambda t=token: self._poll_gemini_validation_events(t))

    def open_api_key_window(self, force_llm=False, force_tts=False, initial_section="llm"):
        self.gemini_verified = self.gemini_verified and not force_llm
        self.api_key_force_llm = self.api_key_force_llm or force_llm
        self.api_key_force_tts = self.api_key_force_tts or force_tts
        if self.api_key_window and self.api_key_window.winfo_exists():
            self.api_key_window.lift()
            self.api_key_window.focus_force()
            return

        self.gemini_key_status_var.set("Paste your LLM API key, then test it.")
        self.tts_key_status_var.set("Paste your TTS API key, then test it.")
        self.gemini_key_var.set(get_llm_api_key())
        self.tts_key_var.set(get_tts_api_key())
        self.llm_api_provider_var.set(self.tr("provider_gemini"))
        self.tts_api_provider_var.set(self._tts_provider_label(get_tts_api_provider()))

        win = tk.Toplevel(self)
        self.api_key_window = win
        win.title(self.tr("api_setup"))
        win.configure(bg="#f6f7fb")
        win.resizable(False, False)
        win.transient(self.winfo_toplevel())

        wrap = ttk.Frame(win, style="Card.TFrame")
        wrap.pack(fill="both", expand=True, padx=12, pady=12)

        ttk.Label(wrap, text=self.tr("api_setup"), style="Card.TLabel").pack(anchor="w")
        ttk.Label(
            wrap,
            text=self.tr("api_setup_desc"),
            style="Card.TLabel",
            foreground="#666",
        ).pack(anchor="w", pady=(0, 10))

        llm_section = ttk.Frame(wrap, style="Card.TFrame")
        llm_section.pack(fill="x")
        ttk.Label(llm_section, text=self.tr("llm_api_setup"), style="Card.TLabel").pack(anchor="w")
        ttk.Label(
            llm_section,
            text=self.tr("llm_key_desc"),
            style="Card.TLabel",
            foreground="#666",
        ).pack(anchor="w", pady=(0, 8))
        llm_provider_row = ttk.Frame(llm_section, style="Card.TFrame")
        llm_provider_row.pack(anchor="w", pady=(0, 8), fill="x")
        ttk.Label(llm_provider_row, text=f"{self.tr('api_provider')}:", style="Card.TLabel").pack(side=tk.LEFT)
        llm_provider_combo = ttk.Combobox(
            llm_provider_row,
            textvariable=self.llm_api_provider_var,
            values=[self.tr("provider_gemini")],
            state="readonly",
            width=18,
        )
        llm_provider_combo.pack(side=tk.LEFT, padx=(6, 0))
        llm_provider_combo.bind("<<ComboboxSelected>>", lambda _e: set_llm_api_provider("gemini"))

        llm_entry = tk.Entry(
            llm_section,
            textvariable=self.gemini_key_var,
            width=54,
            show="*",
            relief="solid",
            bd=1,
            highlightthickness=1,
            highlightbackground="#cbd5e1",
            highlightcolor="#2563eb",
            bg="white",
        )
        llm_entry.pack(fill="x")
        llm_entry.icursor(tk.END)
        llm_entry.bind("<Return>", lambda _event: self.test_and_save_api_keys())
        llm_entry.bind("<KeyRelease>", lambda _event: self._set_api_entry_error("llm", False))

        ttk.Label(
            llm_section,
            text=self.tr("gemini_model_desc"),
            style="Card.TLabel",
            foreground="#666",
        ).pack(anchor="w", pady=(8, 4))
        combo = ttk.Combobox(
            llm_section,
            textvariable=self.gemini_model_var,
            values=self.gemini_model_values or list_available_gemini_models(),
            state="readonly",
            width=24,
        )
        combo.pack(anchor="w")
        combo.bind("<<ComboboxSelected>>", self.on_gemini_model_change)
        ttk.Label(llm_section, textvariable=self.gemini_key_status_var, style="Card.TLabel", foreground="#444").pack(
            anchor="w", pady=(10, 0)
        )
        ttk.Separator(wrap, orient="horizontal").pack(fill="x", pady=12)

        tts_section = ttk.Frame(wrap, style="Card.TFrame")
        tts_section.pack(fill="x")
        ttk.Label(tts_section, text=self.tr("tts_api_setup"), style="Card.TLabel").pack(anchor="w")
        ttk.Label(
            tts_section,
            text=self.tr("tts_key_desc"),
            style="Card.TLabel",
            foreground="#666",
        ).pack(anchor="w", pady=(0, 8))
        tts_provider_row = ttk.Frame(tts_section, style="Card.TFrame")
        tts_provider_row.pack(anchor="w", pady=(0, 8), fill="x")
        ttk.Label(tts_provider_row, text=f"{self.tr('api_provider')}:", style="Card.TLabel").pack(side=tk.LEFT)
        tts_provider_combo = ttk.Combobox(
            tts_provider_row,
            textvariable=self.tts_api_provider_var,
            values=list(self._tts_provider_options().keys()),
            state="readonly",
            width=18,
        )
        tts_provider_combo.pack(side=tk.LEFT, padx=(6, 0))
        tts_provider_combo.bind("<<ComboboxSelected>>", self._on_tts_provider_selected)

        tts_entry = tk.Entry(
            tts_section,
            textvariable=self.tts_key_var,
            width=54,
            show="*",
            relief="solid",
            bd=1,
            highlightthickness=1,
            highlightbackground="#cbd5e1",
            highlightcolor="#2563eb",
            bg="white",
        )
        tts_entry.pack(fill="x")
        tts_entry.icursor(tk.END)
        tts_entry.bind("<Return>", lambda _event: self.test_and_save_api_keys())
        tts_entry.bind("<KeyRelease>", lambda _event: self._set_api_entry_error("tts", False))

        ttk.Label(tts_section, textvariable=self.tts_key_status_var, style="Card.TLabel", foreground="#444").pack(
            anchor="w", pady=(10, 0)
        )
        footer = ttk.Frame(wrap, style="Card.TFrame")
        footer.pack(fill="x", pady=(12, 0))
        self.api_key_test_btn = ttk.Button(footer, text=self.tr("test_and_save"), command=self.test_and_save_api_keys)
        self.api_key_test_btn.pack(side=tk.LEFT)
        ttk.Button(footer, text=self.tr("close"), command=self._close_api_key_window).pack(side=tk.RIGHT)

        self.api_llm_entry = llm_entry
        self.api_tts_entry = tts_entry
        self._set_api_entry_error("llm", False)
        self._set_api_entry_error("tts", False)

        if initial_section == "tts":
            tts_entry.focus_set()
        else:
            llm_entry.focus_set()

        win.grab_set()
        win.protocol("WM_DELETE_WINDOW", self._close_api_key_window)

    def open_gemini_key_window(self, force_verify=False):
        self.open_api_key_window(force_llm=force_verify, force_tts=False, initial_section="llm")

    def _close_api_key_window(self):
        if self.api_key_window and self.api_key_window.winfo_exists():
            try:
                self.api_key_window.grab_release()
            except Exception:
                pass
            self.api_key_window.destroy()
        self.api_key_window = None
        self.api_key_test_btn = None
        self.api_llm_entry = None
        self.api_tts_entry = None
        self.gemini_key_test_btn = None
        self.tts_key_test_btn = None
        llm_missing = self.api_key_force_llm and not str(get_llm_api_key() or "").strip()
        tts_missing = self.api_key_force_tts and not str(get_tts_api_key() or "").strip()
        self.api_key_force_llm = False
        self.api_key_force_tts = False
        if llm_missing or tts_missing:
            self.winfo_toplevel().destroy()

    def _set_api_entry_error(self, field, has_error):
        widget = self.api_llm_entry if field == "llm" else self.api_tts_entry
        if not widget or not widget.winfo_exists():
            return
        if has_error:
            widget.configure(bg="#fff1f2", highlightbackground="#ef4444", highlightcolor="#ef4444")
        else:
            widget.configure(bg="white", highlightbackground="#cbd5e1", highlightcolor="#2563eb")

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

    def test_and_save_api_keys(self):
        llm_key = str(self.gemini_key_var.get() or "").strip()
        tts_key = str(self.tts_key_var.get() or "").strip()
        tts_provider = self._tts_provider_value()
        model_name = self._get_selected_gemini_model()
        llm_required = self.api_key_force_llm or bool(llm_key)
        tts_required = self.api_key_force_tts or bool(tts_key)

        self._set_api_entry_error("llm", False)
        self._set_api_entry_error("tts", False)

        has_local_error = False
        if llm_required and not llm_key:
            self.gemini_key_status_var.set("Please enter an LLM API key.")
            self._set_api_entry_error("llm", True)
            has_local_error = True
        else:
            self.gemini_key_status_var.set("Paste your LLM API key, then test it.")

        if tts_required and not tts_key:
            self.tts_key_status_var.set("Please enter a TTS API key.")
            self._set_api_entry_error("tts", True)
            has_local_error = True
        else:
            self.tts_key_status_var.set("Paste your TTS API key, then test it.")

        if has_local_error:
            return
        if not llm_required and not tts_required:
            messagebox.showinfo(self.tr("info"), "Please enter at least one API key.")
            return

        if llm_required:
            self.gemini_key_status_var.set(f"Testing LLM key with {model_name}...")
        if tts_required:
            self.tts_key_status_var.set(f"Testing TTS API key with {self._tts_provider_label(tts_provider)}...")
        if self.api_key_test_btn:
            self.api_key_test_btn.config(state="disabled")

        self.gemini_validation_token += 1
        token = self.gemini_validation_token
        self.gemini_validation_active_token = token
        self._clear_gemini_validation_queue()

        import threading

        def _run():
            result = {
                "llm_required": llm_required,
                "tts_required": tts_required,
                "llm_ok": False,
                "tts_ok": False,
                "llm_error": "",
                "tts_error": "",
                "llm_api_key": llm_key,
                "tts_api_key": tts_key,
                "llm_model": model_name,
                "tts_provider": tts_provider,
            }
            try:
                if llm_required:
                    try:
                        validate_gemini_api_key(llm_key, model=model_name, timeout=25)
                        result["llm_ok"] = True
                    except Exception as e:
                        result["llm_error"] = str(e)
                if tts_required:
                    try:
                        validate_tts_api_key(tts_key, tts_provider, timeout=30)
                        result["tts_ok"] = True
                    except Exception as e:
                        result["tts_error"] = str(e)
                self._emit_gemini_validation_event("success_api_setup", token, result)
            finally:
                self._emit_gemini_validation_event("done", token, None)

        threading.Thread(target=_run, daemon=True).start()
        self.after(80, lambda t=token: self._poll_gemini_validation_events(t))

    def _finish_gemini_validation_success(self, payload):
        api_key = str(payload.get("api_key") or "").strip()
        model_name = str(payload.get("model") or DEFAULT_GEMINI_MODEL).strip()
        set_llm_api_key(api_key)
        set_llm_api_provider("gemini")
        set_generation_model(model_name)
        self.gemini_verified = True
        self.gemini_key_status_var.set("LLM API key is valid.")
        if self.gemini_key_test_btn:
            self.gemini_key_test_btn.config(state="normal")
        self.status_var.set("LLM API ready.")
        self._maybe_close_api_key_window()

    def _finish_combined_api_validation(self, payload):
        llm_required = bool(payload.get("llm_required"))
        tts_required = bool(payload.get("tts_required"))
        llm_ok = bool(payload.get("llm_ok"))
        tts_ok = bool(payload.get("tts_ok"))
        llm_error = str(payload.get("llm_error") or "").strip()
        tts_error = str(payload.get("tts_error") or "").strip()

        if llm_required and llm_ok:
            set_llm_api_key(str(payload.get("llm_api_key") or "").strip())
            set_llm_api_provider("gemini")
            set_generation_model(str(payload.get("llm_model") or DEFAULT_GEMINI_MODEL).strip())
            self.gemini_verified = True
            self.gemini_key_status_var.set("LLM API key is valid.")
            self._set_api_entry_error("llm", False)
        elif llm_required:
            self.gemini_verified = False
            message = llm_error or "LLM API key test failed. Please paste another key."
            self.gemini_key_status_var.set(message)
            self._set_api_entry_error("llm", True)

        if tts_required and tts_ok:
            provider = str(payload.get("tts_provider") or "gemini").strip().lower()
            set_tts_api_key(str(payload.get("tts_api_key") or "").strip())
            set_tts_api_provider(provider)
            self.tts_api_provider_var.set(self._tts_provider_label(provider))
            self.tts_key_status_var.set("TTS API key is valid.")
            self._set_api_entry_error("tts", False)
            self.refresh_voice_list()
        elif tts_required:
            message = tts_error or "TTS API key test failed. Please paste another key."
            self.tts_key_status_var.set(message)
            self._set_api_entry_error("tts", True)

        if self.api_key_test_btn:
            self.api_key_test_btn.config(state="normal")

        if (not llm_required or llm_ok) and (not tts_required or tts_ok):
            self.status_var.set("API ready.")
            self._maybe_close_api_key_window()

    def _finish_gemini_validation_error(self, message):
        self.gemini_verified = False
        self.gemini_key_status_var.set("LLM API key test failed. Please paste another key.")
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

        import threading

        def _run():
            try:
                validate_tts_api_key(api_key, provider, timeout=30)
                self._emit_gemini_validation_event(
                    "success_tts",
                    token,
                    {"api_key": api_key, "provider": provider},
                )
            except Exception as e:
                self._emit_gemini_validation_event("error_tts", token, str(e))
            self._emit_gemini_validation_event("done", token, None)

        threading.Thread(target=_run, daemon=True).start()
        self.after(80, lambda t=token: self._poll_gemini_validation_events(t))

    def _finish_tts_validation_success(self, payload):
        api_key = str(payload.get("api_key") or "").strip()
        provider = str(payload.get("provider") or "gemini").strip().lower()
        set_tts_api_key(api_key)
        set_tts_api_provider(provider)
        self.tts_key_status_var.set("TTS API key is valid.")
        if self.tts_key_test_btn:
            self.tts_key_test_btn.config(state="normal")
        self.tts_api_provider_var.set(self._tts_provider_label(provider))
        self.refresh_voice_list()
        self.status_var.set("TTS API ready.")
        self._maybe_close_api_key_window()

    def _finish_tts_validation_error(self, message):
        self.tts_key_status_var.set("TTS API key test failed. Please paste another key.")
        if self.tts_key_test_btn:
            self.tts_key_test_btn.config(state="normal")
        messagebox.showerror(self.tr("tts_api_key_error"), str(message or "Unknown error"))

    def _maybe_close_api_key_window(self):
        llm_ready = bool(str(get_llm_api_key() or "").strip())
        tts_ready = bool(str(get_tts_api_key() or "").strip())
        if self.api_key_force_llm and not llm_ready:
            return
        if self.api_key_force_tts and not tts_ready:
            return
        if self.api_key_window and self.api_key_window.winfo_exists():
            self._close_api_key_window()

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
        clear_event_queue(self.passage_event_queue)

    def _emit_passage_event(self, event_type, token, payload=None):
        emit_event(self.passage_event_queue, event_type, token, payload)

    def _poll_passage_generation_events(self, token):
        done = drain_event_queue(
            target_queue=self.passage_event_queue,
            token=token,
            active_token=self.passage_generation_active_token,
            handlers={
                "partial": lambda payload: self._update_partial_passage(token, payload),
                "result": lambda payload: self._apply_generated_passage(token, payload),
                "error": lambda payload: messagebox.showerror("Generate Error", str(payload or "Unknown error")),
            },
        )
        if not done and token == self.passage_generation_active_token:
            self.after(80, lambda t=token: self._poll_passage_generation_events(t))

    def _update_partial_passage(self, token, text):
        if token != self.passage_generation_token:
            return
        state = build_partial_passage_state(text)
        self.current_passage = state["passage"]
        self.current_passage_original = self.current_passage
        if state["has_passage"]:
            self._set_passage_text(self.current_passage)

    def _run_passage_generation(self, token, words, model_name):
        start_passage_generation_task(
            token=token,
            words=words,
            model_name=model_name,
            emit_event=self._emit_passage_event,
        )

    def _apply_generated_passage(self, token, result):
        if token != self.passage_generation_token:
            return
        state = build_generated_passage_state(result, default_model=DEFAULT_GEMINI_MODEL)
        self.current_passage = state["passage"]
        self.current_passage_original = self.current_passage
        self.current_passage_words = list(state["used_words"])
        self.passage_is_practice = False
        self.passage_cloze_text = ""
        self.passage_answers = []
        self._set_passage_text(self.current_passage)
        self.passage_status_var.set(state["status_text"])

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
        runtime = tts_get_runtime_label()
        token = speak_stream_async(
            speech_text,
            self.volume_var.get() / 100.0,
            rate_ratio=self.speech_rate_var.get(),
            cancel_before=False,
            chunk_chars=90,
        )
        self.passage_status_var.set(build_passage_audio_status(runtime))
        self._watch_tts_backend(token, target="passage", text_label="passage")

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
        if self.gemini_status_after:
            try:
                self.after_cancel(self.gemini_status_after)
            except Exception:
                pass
        self.gemini_status_after = None
        if self.settings_window and self.settings_window.winfo_exists():
            self.settings_window.destroy()
        self.settings_window = None

    def _refresh_settings_gemini_status(self):
        status = tts_get_online_tts_queue_status()
        provider_label = self._tts_provider_label(get_tts_api_provider())
        state = str(status.get("state") or "idle").strip().lower()
        if state == "rate_limited":
            status_text = self.trf("tts_status_limited", provider=provider_label)
        elif state == "ok":
            status_text = self.trf("tts_status_normal", provider=provider_label)
        elif state == "error":
            status_text = self.trf("tts_status_error", provider=provider_label)
        else:
            status_text = self.trf("tts_status_idle", provider=provider_label)
        queue_count = int(status.get("queue_count") or 0)
        if state == "rate_limited":
            queue_text = self.trf("tts_status_queue_waiting", count=queue_count)
        elif queue_count > 0:
            queue_text = self.trf("tts_status_queue_processing", count=queue_count)
        else:
            queue_text = self.trf("tts_status_queue", count=queue_count)
        self.gemini_runtime_status_var.set(f"{status_text} | {queue_text}")

        next_retry_at = float(status.get("next_retry_at") or 0.0)
        if next_retry_at > 0:
            now = time.time()
            remaining_seconds = max(0, int(round(next_retry_at - now)))
            retry_text = time.strftime("%H:%M:%S", time.localtime(next_retry_at))
            self.gemini_retry_status_var.set(
                self.trf("tts_status_retry_at_in", time=retry_text, seconds=remaining_seconds)
            )
        else:
            self.gemini_retry_status_var.set(self.tr("tts_status_retry_none"))

        if self.settings_window and self.settings_window.winfo_exists():
            self.gemini_status_after = self.after(1000, self._refresh_settings_gemini_status)
        else:
            self.gemini_status_after = None

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
        value = int(self.dictation_volume_var.get())
        if self.dictation_volume_value_label and self.dictation_volume_value_label.winfo_exists():
            self.dictation_volume_value_label.config(text=self.trf("dictation_volume_level", value=value))

    def close_dictation_volume_popup(self):
        if self.dictation_volume_popup and self.dictation_volume_popup.winfo_exists():
            self.dictation_volume_popup.destroy()
        self.dictation_volume_popup = None
        self.dictation_volume_scale = None
        self.dictation_volume_value_label = None

    def toggle_dictation_volume_popup(self):
        if not self._dictation_window_active() or not self.dictation_volume_btn:
            return
        if self.dictation_volume_popup and self.dictation_volume_popup.winfo_exists():
            self.close_dictation_volume_popup()
            return

        popup = tk.Toplevel(self.dictation_window)
        popup.title(self.tr("dictation_volume"))
        popup.configure(bg="#f6f7fb")
        popup.resizable(False, False)
        popup.transient(self.dictation_window)
        popup.protocol("WM_DELETE_WINDOW", self.close_dictation_volume_popup)
        self.dictation_volume_popup = popup

        wrap = ttk.Frame(popup, style="Card.TFrame", padding=12)
        wrap.pack(fill="both", expand=True)
        ttk.Label(wrap, text=self.tr("dictation_volume"), style="Card.TLabel").pack(anchor="w")
        ttk.Label(
            wrap,
            text=self.tr("dictation_volume_tip"),
            style="Card.TLabel",
            foreground="#667085",
            wraplength=260,
            justify="left",
        ).pack(anchor="w", pady=(4, 10))
        self.dictation_volume_scale = tk.Scale(
            wrap,
            from_=0,
            to=600,
            orient=tk.HORIZONTAL,
            length=260,
            resolution=10,
            showvalue=False,
            variable=self.dictation_volume_var,
            highlightthickness=0,
            command=self.on_dictation_volume_change,
        )
        self.dictation_volume_scale.pack(anchor="w")
        self.dictation_volume_value_label = ttk.Label(
            wrap,
            text=self.trf("dictation_volume_level", value=int(self.dictation_volume_var.get())),
            style="Card.TLabel",
        )
        self.dictation_volume_value_label.pack(anchor="w", pady=(6, 0))
        self.on_dictation_volume_change()

        try:
            self.dictation_window.update_idletasks()
            popup.update_idletasks()
            x = self.dictation_volume_btn.winfo_rootx() - 210
            y = self.dictation_volume_btn.winfo_rooty() + self.dictation_volume_btn.winfo_height() + 6
            popup.geometry(f"+{max(0, x)}+{max(0, y)}")
        except Exception:
            pass

    def set_loop_mode(self, loop_all):
        return

    def update_loop_button(self):
        return

    def on_stop_at_end_toggle(self):
        return

    def open_settings_window(self):
        self._sync_provider_vars()
        self.settings_window = tk.Toplevel(self)
        self.settings_window.title(self.tr("settings_title"))
        self.settings_window.configure(bg="#f6f7fb")
        self.settings_window.resizable(False, False)
        self.settings_window.protocol("WM_DELETE_WINDOW", self._close_settings_window)

        container = ttk.Frame(self.settings_window, style="Card.TFrame")
        container.pack(padx=10, pady=10)

        left_menu = ttk.Frame(container, style="Card.TFrame", width=120)
        left_menu.grid(row=0, column=0, sticky="n")
        right_panel = ttk.Frame(container, style="Card.TFrame", width=360, height=260)
        right_panel.grid(row=0, column=1, padx=(10, 0), sticky="n")
        right_panel.grid_propagate(False)

        self.settings_sections_visible = {
            "source": True,
            "speed": True,
            "language": False,
        }
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
        for label, key in [
            (self.tr("settings_toggle_source"), "source"),
            (self.tr("settings_toggle_speed"), "speed"),
            (self.tr("ui_language"), "language"),
        ]:
            btn = ttk.Button(left_menu, text=label, command=lambda k=key: toggle_section(k))
            btn.pack(fill=tk.X, pady=4)

        # Source section
        source_section = ttk.Frame(right_panel, style="Card.TFrame")
        ttk.Label(source_section, text=self.tr("source"), style="Card.TLabel").pack(anchor="w")

        self.voice_combo = ttk.Combobox(
            source_section,
            textvariable=self.voice_var,
            state="readonly",
            width=32,
        )
        self.voice_combo.pack(anchor="w")
        self.voice_combo.bind("<<ComboboxSelected>>", self.on_voice_change)
        llm_row = ttk.Frame(source_section, style="Card.TFrame")
        llm_row.pack(anchor="w", pady=(8, 0), fill="x")
        ttk.Label(llm_row, text=f"{self.tr('llm_api')}:", style="Card.TLabel").pack(side=tk.LEFT)
        ttk.Combobox(
            llm_row,
            textvariable=self.llm_api_provider_var,
            values=[self.tr("provider_gemini")],
            state="readonly",
            width=12,
        ).pack(side=tk.LEFT, padx=(6, 6))
        ttk.Button(llm_row, text=self.tr("llm_api"), command=self.open_gemini_key_window).pack(side=tk.LEFT)

        tts_row = ttk.Frame(source_section, style="Card.TFrame")
        tts_row.pack(anchor="w", pady=(8, 0), fill="x")
        ttk.Label(tts_row, text=f"{self.tr('tts_api')}:", style="Card.TLabel").pack(side=tk.LEFT)
        tts_provider_combo = ttk.Combobox(
            tts_row,
            textvariable=self.tts_api_provider_var,
            values=list(self._tts_provider_options().keys()),
            state="readonly",
            width=12,
        )
        tts_provider_combo.pack(side=tk.LEFT, padx=(6, 6))
        tts_provider_combo.bind("<<ComboboxSelected>>", self._on_tts_provider_selected)
        ttk.Button(tts_row, text=self.tr("tts_api"), command=self.open_tts_key_window).pack(side=tk.LEFT)
        ttk.Label(
            source_section,
            textvariable=self.gemini_runtime_status_var,
            style="Card.TLabel",
            foreground="#4b5563",
        ).pack(anchor="w", pady=(8, 0))
        ttk.Label(
            source_section,
            textvariable=self.gemini_retry_status_var,
            style="Card.TLabel",
            foreground="#667085",
        ).pack(anchor="w", pady=(2, 0))

        source_sep = ttk.Separator(right_panel, orient="horizontal")
        sections.append({"key": "source", "frame": source_section, "sep": source_sep})

        # Speed section
        speed_section = ttk.Frame(right_panel, style="Card.TFrame")
        ttk.Label(speed_section, text=self.tr("speed"), style="Card.TLabel").pack(anchor="w")
        ttk.Label(
            speed_section,
            text=self.tr("speed_desc"),
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
        ttk.Label(custom_row, text=self.tr("custom_seconds"), style="Card.TLabel").pack(side=tk.LEFT)
        self.custom_interval = ttk.Entry(custom_row, width=6)
        self.custom_interval.pack(side=tk.LEFT, padx=4)
        self.custom_interval.bind("<Return>", lambda _e: self.apply_custom_interval())
        ttk.Button(custom_row, text=self.tr("apply"), command=self.apply_custom_interval).pack(side=tk.LEFT)

        ttk.Label(
            speed_section,
            text=self.tr("pronunciation_speed"),
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

        language_section = ttk.Frame(right_panel, style="Card.TFrame")
        ttk.Label(language_section, text=self.tr("ui_language"), style="Card.TLabel").pack(anchor="w")
        language_combo = ttk.Combobox(
            language_section,
            textvariable=self.ui_language_var,
            state="readonly",
            width=18,
            values=("zh", "en"),
        )
        language_combo.pack(anchor="w", pady=(4, 0))
        language_combo.bind("<<ComboboxSelected>>", self.on_ui_language_change)
        ttk.Label(
            language_section,
            text=f"zh = {self.tr('language_zh')}   |   en = {self.tr('language_en')}",
            style="Card.TLabel",
            foreground="#666",
        ).pack(anchor="w", pady=(6, 0))
        language_sep = ttk.Separator(right_panel, orient="horizontal")
        sections.append({"key": "language", "frame": language_section, "sep": language_sep})

        rebuild_sections()

        self.update_speed_buttons()
        self.update_speech_rate_buttons()
        self.refresh_voice_list()
        self._refresh_settings_gemini_status()

    def apply_custom_interval(self):
        try:
            val = float(self.custom_interval.get())
            if val < 0.2:
                raise ValueError
        except Exception:
            self.show_info("valid_number_needed")
            return
        self.set_interval(val)

    def build_queue(self):
        return list(range(len(self.store.words)))

    def rebuild_queue_on_mode_change(self):
        if not self.store.words:
            self.queue = []
            self.pos = -1
            return
        if self.current_word is None:
            self.queue = self.build_queue()
            selected_idx = self._get_selected_index()
            self.pos = selected_idx if selected_idx is not None else 0
            self.set_current_word()
            return
        self.queue = self.build_queue()
        try:
            self.pos = self.store.words.index(self.current_word)
        except Exception:
            selected_idx = self._get_selected_index()
            self.pos = selected_idx if selected_idx is not None else 0
        self.set_current_word()
        if self.play_state == "playing":
            self.play_current()
            self.schedule_next()

    def toggle_play(self):
        if not self.store.words:
            self.show_info("import_words_first")
            return
        if self.play_state == "playing":
            self.play_state = "paused"
            self.cancel_schedule()
            tts_cancel_all()
            self.play_token += 1
            self.status_var.set("已暂停顺序播放。")
            self.update_play_button()
            return

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
        runtime = tts_get_runtime_label()
        source_path = self.store.get_current_source_path()
        cached = get_voice_source() == SOURCE_GEMINI and tts_has_cached_word_audio(
            self.current_word,
            source_path=source_path,
        )
        token = speak_async(
            self.current_word,
            self.volume_var.get() / 100.0,
            rate_ratio=self.speech_rate_var.get(),
            cancel_before=True,
            source_path=source_path,
        )
        if cached:
            self.status_var.set(f"Playing cached audio for '{self.current_word}'.")
        else:
            self.status_var.set(f"Generating '{self.current_word}' with {runtime}...")
        self._watch_tts_backend(token, target="status", text_label=self.current_word)

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
                self.queue = self.build_queue()
                self.pos = 0
        self.set_current_word()
        self.play_current()
        self.schedule_next()

    def set_current_word(self):
        idx = self.queue[self.pos]
        self.current_word = self.store.words[idx]
        if self.word_table and self.word_table.exists(str(idx)):
            try:
                self.suppress_word_select_action = True
                row_id = str(idx)
                self.word_table.selection_set(row_id)
                self.word_table.focus(row_id)
                self.word_table.see(row_id)
            except Exception:
                pass
        self.status_var.set(f"顺序播放：{idx + 1}/{len(self.store.words)}  {self.current_word}")
        self._refresh_selection_details()

    def cancel_schedule(self):
        if self.after_id:
            self.after_cancel(self.after_id)
            self.after_id = None

    def update_play_button(self):
        if self.play_state == "playing":
            self.play_btn.config(text=("⏸ Pause" if self.ui_language_var.get() == "en" else "⏸ 暂停"))
        else:
            self.play_btn.config(text=self.tr("play"))
        self._refresh_selection_details()

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
        self.status_var.set("未开始顺序播放。")
        self.update_play_button()
        self._select_sidebar_tab("review")
        self._refresh_selection_details()

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
        if not self.store.words:
            self.queue = []
            self.pos = -1
            self.current_word = None
            return
        self.queue = self.build_queue()
        selected_idx = self._get_selected_index()
        self.pos = selected_idx if selected_idx is not None else 0
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
        self.cancel_schedule()
        self.play_token += 1
        self.play_state = "stopped"
        self.update_play_button()
        tts_cancel_all()

    def on_word_right_click(self, event):
        if not self.word_table or not self.word_context_menu:
            return
        row_id = self.word_table.identify_row(event.y)
        if not row_id:
            return
        try:
            row_idx = int(row_id)
        except Exception:
            return
        try:
            self.suppress_word_select_action = True
            self.word_table.selection_set(row_id)
            self.word_table.focus(row_id)
        except Exception:
            pass
        self._set_word_action_context(row_idx, origin="main")
        try:
            self.word_context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.word_context_menu.grab_release()
        return "break"

    def on_dictation_word_right_click(self, event):
        tree = event.widget if isinstance(event.widget, ttk.Treeview) else None
        if not tree or not self.dictation_context_menu:
            return
        row_id = str(tree.identify_row(event.y) or "").strip()
        if not row_id or row_id == "empty":
            return
        try:
            self.suppress_dictation_select_action = True
            tree.selection_set(row_id)
            tree.focus(row_id)
        except Exception:
            pass
        selected_idx = self._dictation_row_to_store_index(tree, row_id=row_id)
        selected_word = ""
        try:
            view_index = int(row_id)
            items = self._get_dictation_source_items()
            if 0 <= view_index < len(items):
                selected_word = str(items[view_index].get("word") or "").strip()
        except Exception:
            selected_word = ""
        if selected_idx is not None:
            self._set_word_action_context(selected_idx, origin="dictation", word=selected_word)
        elif selected_word:
            self._set_word_action_context(None, origin="dictation", word=selected_word)
        else:
            return "break"
        try:
            self.dictation_context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.dictation_context_menu.grab_release()
        return "break"

    def _fallback_sentence(self, word):
        w = str(word or "").strip()
        if not w:
            return ""
        if " " in w:
            return f'Please use "{w}" in your next speaking practice task.'
        return f"I wrote the word {w} in my notebook for today's review."

    def _clear_sentence_event_queue(self):
        clear_event_queue(self.sentence_event_queue)

    def _emit_sentence_event(self, event_type, token, payload=None):
        emit_event(self.sentence_event_queue, event_type, token, payload)

    def _clear_synonym_event_queue(self):
        clear_event_queue(self.synonym_event_queue)

    def _emit_synonym_event(self, event_type, token, payload=None):
        emit_event(self.synonym_event_queue, event_type, token, payload)

    def _poll_synonym_events(self, token):
        done = drain_event_queue(
            target_queue=self.synonym_event_queue,
            token=token,
            active_token=self.synonym_lookup_active_token,
            handlers={
                "result": lambda payload: self._show_synonym_window(
                    (payload or {}).get("word", ""),
                    (payload or {}).get("focus", ""),
                    (payload or {}).get("synonyms") or [],
                    (payload or {}).get("source", ""),
                ),
                "error": lambda payload: messagebox.showerror(self.tr("synonyms_error"), str(payload or "Unknown error")),
            },
        )
        if not done and token == self.synonym_lookup_active_token:
            self.after(80, lambda t=token: self._poll_synonym_events(t))

    def _poll_sentence_events(self, token):
        done = drain_event_queue(
            target_queue=self.sentence_event_queue,
            token=token,
            active_token=self.sentence_generation_active_token,
            handlers={
                "result": lambda payload: self._show_sentence_window(
                    (payload or {}).get("word", ""),
                    (payload or {}).get("sentence", ""),
                    (payload or {}).get("source", "Unknown"),
                ),
                "error": lambda payload: messagebox.showerror("Sentence Error", str(payload or "Unknown error")),
            },
        )
        if not done and token == self.sentence_generation_active_token:
            self.after(80, lambda t=token: self._poll_sentence_events(t))

    def make_sentence_for_selected_word(self):
        word = self._get_context_word()
        if not word:
            self.show_info("select_word_first")
            return
        if not self._require_gemini_ready():
            return
        model_name = self._get_selected_gemini_model()
        self.status_var.set(f"Generating IELTS sentence for '{word}' with {model_name}...")
        self.sentence_generation_token += 1
        token = self.sentence_generation_token
        self.sentence_generation_active_token = token
        self._clear_sentence_event_queue()
        start_sentence_generation_task(
            token=token,
            word=word,
            model_name=model_name,
            fallback_sentence=self._fallback_sentence(word),
            emit_event=self._emit_sentence_event,
        )
        self.after(80, lambda t=token: self._poll_sentence_events(t))

    def lookup_synonyms_for_selected_word(self):
        word = self._get_context_word()
        if not word:
            self.show_info("select_word_first")
            return
        self.status_var.set(f"Looking up synonyms for '{word}'...")
        self.synonym_lookup_token += 1
        token = self.synonym_lookup_token
        self.synonym_lookup_active_token = token
        self._clear_synonym_event_queue()
        start_synonym_lookup_task(
            token=token,
            word=word,
            emit_event=self._emit_synonym_event,
        )
        self.after(80, lambda t=token: self._poll_synonym_events(t))

    def _show_sentence_window(self, word, sentence, source):
        state = build_sentence_view_state(word, sentence, source)
        self.status_var.set(state["status_text"])
        build_sentence_window(
            self,
            state=state,
            on_read=lambda s=sentence: speak_async(
                s,
                self.volume_var.get() / 100.0,
                rate_ratio=self.speech_rate_var.get(),
                cancel_before=True,
                source_path=self.store.get_current_source_path(),
            ),
        )

    def _show_synonym_window(self, word, focus, synonyms, source=None):
        state = build_synonym_view_state(
            tr=self.tr,
            trf=self.trf,
            word=word,
            focus=focus,
            synonyms=synonyms,
            source=source,
        )
        self.status_var.set(state["status_text"])
        build_synonym_window(self, state=state)

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
        self.tts_status_request += 1
        request_id = self.tts_status_request

        def _poll(attempt=0):
            if request_id != self.tts_status_request:
                return
            status = tts_get_backend_status(token)
            if status:
                label = status.get("label") or "TTS"
                fallback = bool(status.get("fallback"))
                from_cache = bool(status.get("from_cache"))
                if target == "passage":
                    if fallback:
                        self.passage_status_var.set(f"Playing passage via {label} after Gemini fallback.")
                    elif from_cache:
                        self.passage_status_var.set(f"Playing passage via cached {label} audio.")
                    else:
                        self.passage_status_var.set(f"Playing passage via {label}.")
                else:
                    if fallback:
                        self.status_var.set(f"Playing '{text_label}' via {label} after Gemini fallback.")
                    elif from_cache:
                        self.status_var.set(f"Playing cached audio for '{text_label}' via {label}.")
                    else:
                        self.status_var.set(f"Playing '{text_label}' via {label}.")
                return
            if attempt < 120:
                self.after(250, lambda: _poll(attempt + 1))

        self.after(250, _poll)

