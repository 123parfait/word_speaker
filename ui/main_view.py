# -*- coding: utf-8 -*-
import os
import queue
import random
import re
import time
import unicodedata
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from data.store import WordStore
from services.tts import (
    speak_async,
    speak_stream_async,
    cancel_all as tts_cancel_all,
    cleanup_cache_for_source_path as tts_cleanup_cache_for_source_path,
    cleanup_word_audio_cache as tts_cleanup_word_audio_cache,
    cleanup_manual_session_cache as tts_cleanup_manual_session_cache,
    get_recent_wrong_cache_source as tts_get_recent_wrong_cache_source,
    get_word_audio_cache_info as tts_get_word_audio_cache_info,
    promote_word_audio_to_recent_wrong as tts_promote_word_audio_to_recent_wrong,
    precache_word_audio_async,
    prepare_async as tts_prepare_async,
    queue_word_audio_generation as tts_queue_word_audio_generation,
    rename_cache_source_path as tts_rename_cache_source_path,
    rebind_manual_session_cache_to_source as tts_rebind_manual_session_cache_to_source,
    set_preferred_pending_source as tts_set_preferred_pending_source,
    get_backend_status as tts_get_backend_status,
    get_gemini_queue_status as tts_get_gemini_queue_status,
    get_runtime_label as tts_get_runtime_label,
    has_cached_word_audio as tts_has_cached_word_audio,
)
from services.translation import (
    get_cached_translations,
    set_cached_translation,
    translate_words as translate_words_en_zh,
)
from services.word_analysis import (
    analyze_words,
    get_cached_pos,
    set_cached_pos,
)
from services.synonyms import get_synonyms as get_local_synonyms
from services.ielts_passage import build_ielts_listening_passage
from services.app_config import (
    get_llm_api_key,
    get_llm_api_provider,
    get_generation_model,
    get_tts_api_key,
    get_tts_api_provider,
    get_ui_language,
    set_llm_api_key,
    set_llm_api_provider,
    set_generation_model,
    set_tts_api_key,
    set_tts_api_provider,
    set_ui_language,
)
from services.gemini_writer import (
    DEFAULT_GEMINI_MODEL,
    choose_preferred_generation_model,
    generate_example_sentence_with_gemini,
    generate_english_passage_with_gemini,
    list_available_gemini_models,
    validate_gemini_api_key,
)
from services.tts import validate_tts_api_key
from services.voice_catalog import kokoro_ready, list_system_voices, piper_ready
from services.voice_manager import (
    SOURCE_GEMINI,
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
    remove_document as remove_corpus_document,
    search_corpus,
)
from services.diff_view import apply_diff


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
        "synonyms_source": "来源：spaCy + WordNet（本地）",
        "synonyms_focus": "匹配词：{word}",
        "inspect_audio_cache": "查询音频缓存",
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
        "api_provider": "提供方",
        "provider_gemini": "Gemini",
        "provider_elevenlabs": "ElevenLabs",
        "llm_api_setup": "大模型 API 设置",
        "llm_key_desc": "用于文章生成、例句生成等 AI 功能。当前实现支持 Gemini。",
        "tts_api_setup": "TTS API 设置",
        "tts_key_desc": "用于在线语音生成。当前实现支持 Gemini TTS 和 ElevenLabs。ElevenLabs 默认使用偏英式的标准发音。",
        "tts_api_key_error": "TTS API Key 错误",
        "paste_tts_key_first": "请先粘贴 TTS API Key。",
        "tts_status_normal": "{provider} 正常",
        "tts_status_limited": "{provider} 限流中",
        "tts_status_error": "{provider} 错误",
        "tts_status_idle": "{provider} 空闲",
        "tts_status_retry_at": "下次请求时间：{time}",
        "tts_status_retry_none": "下次请求时间：-",
        "tts_status_queue": "等待队列：{count}",
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
        "synonyms_source": "Source: spaCy + WordNet (Local)",
        "synonyms_focus": "Matched token: {word}",
        "inspect_audio_cache": "Inspect Audio Cache",
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
        "api_provider": "Provider",
        "provider_gemini": "Gemini",
        "provider_elevenlabs": "ElevenLabs",
        "llm_api_setup": "LLM API setup",
        "llm_key_desc": "Used for passage generation, sentence generation, and other AI features. Gemini is currently implemented.",
        "tts_api_setup": "TTS API setup",
        "tts_key_desc": "Used for online speech synthesis. Gemini TTS and ElevenLabs are currently implemented. ElevenLabs defaults to a British-style standard voice.",
        "tts_api_key_error": "TTS API Key Error",
        "paste_tts_key_first": "Please paste a TTS API key first.",
        "tts_status_normal": "{provider} OK",
        "tts_status_limited": "{provider} Rate Limited",
        "tts_status_error": "{provider} Error",
        "tts_status_idle": "{provider} Idle",
        "tts_status_retry_at": "Next request: {time}",
        "tts_status_retry_none": "Next request: -",
        "tts_status_queue": "Queue: {count}",
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
        self.tts_status_request = 0
        self.word_table = None
        self.word_table_scroll = None
        self.word_context_menu = None
        self.history_context_menu = None
        self.word_action_index = None
        self.word_action_origin = "main"
        self.word_edit_entry = None
        self.word_edit_row = None
        self.word_edit_column = None
        self.suppress_word_select_action = False
        self.suppress_dictation_select_action = False
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
        self.translation_token = 0
        self.analysis_token = 0
        self.manual_source_dirty = False
        self.passage_window = None
        self.passage_text = None
        self.passage_status_var = tk.StringVar(value="Load words and click Generate.")
        self.gemini_model_var = tk.StringVar(value=get_generation_model() or DEFAULT_GEMINI_MODEL)
        self.gemini_model_combo = None
        self.gemini_model_values = []
        self.llm_api_provider_var = tk.StringVar(value=self.tr("provider_gemini"))
        self.tts_api_provider_var = tk.StringVar(value=self._tts_provider_label(get_tts_api_provider()))
        self.gemini_verified = False
        self.gemini_key_window = None
        self.gemini_key_var = tk.StringVar(value=get_llm_api_key())
        self.gemini_key_status_var = tk.StringVar(value="Paste your LLM API key, then test it.")
        self.tts_key_window = None
        self.tts_key_var = tk.StringVar(value=get_tts_api_key())
        self.tts_key_status_var = tk.StringVar(value="Paste your TTS API key, then test it.")
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
        self.find_docs_context_menu = None
        self.find_doc_items = []
        self.find_result_items = {}
        self.find_task_queue = queue.Queue()
        self.find_task_token = 0
        self.find_active_token = 0
        self.audio_precache_token = 0
        self.ui_language_var = tk.StringVar(value=get_ui_language())
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
        self.dictation_pick_frame = None
        self.dictation_session_frame = None
        self.dictation_result_frame = None
        self.dictation_start_word_list = None
        self.dictation_mode_popup = None
        self.dictation_mode_buttons = []
        self.dictation_input = None
        self.dictation_result_label = None
        self.dictation_progress = None
        self.dictation_timer_after = None
        self.dictation_feedback_after = None
        self.dictation_pool = []
        self.dictation_index = -1
        self.dictation_current_word = ""
        self.dictation_wrong_items = []
        self.dictation_correct_count = 0
        self.dictation_answer_revealed = False
        self.dictation_running = False
        self.dictation_paused = False
        self.dictation_seconds_left = 0
        self.dictation_session_source_path = None
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
        self.update_order_button()
        self.update_loop_button()
        self.update_speed_buttons()
        self.update_speech_rate_buttons()
        self.update_play_button()
        self.update_right_visibility()
        tts_prepare_async()
        self.refresh_gemini_models()
        self.after(150, self.ensure_gemini_api_key)

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
        header = ttk.Frame(self)
        header.pack(fill=tk.X, pady=(0, 6))
        title = ttk.Label(header, text="Word Speaker", font=("Segoe UI", 14, "bold"))
        title.pack(side=tk.LEFT)

        ttk.Separator(self, orient="horizontal").pack(fill=tk.X, pady=(0, 10))

        self.main = ttk.Frame(self)
        self.main.pack(fill="both", expand=True)
        self.main.grid_columnconfigure(0, weight=5)
        self.main.grid_columnconfigure(2, weight=4)
        self.main.grid_rowconfigure(0, weight=1)

        self.left = ttk.Frame(self.main, style="Card.TFrame")
        self.left.grid(row=0, column=0, sticky="nsew")
        self.mid_sep = ttk.Separator(self.main, orient="vertical")
        self.mid_sep.grid(row=0, column=1, sticky="ns", padx=10)
        self.right = ttk.Frame(self.main, style="Card.TFrame")
        self.right.grid(row=0, column=2, sticky="nsew")

        # Left: Word list + player bar
        left_title = ttk.Label(self.left, text=self.tr("word_list"), style="Card.TLabel")
        left_title.grid(row=0, column=0, columnspan=2, padx=12, pady=(12, 2), sticky="w")
        ttk.Label(
            self.left,
            text=self.tr("word_list_desc"),
            style="Card.TLabel",
            foreground="#667085",
        ).grid(row=1, column=0, columnspan=2, padx=12, pady=(0, 6), sticky="w")

        top_btn_row = ttk.Frame(self.left, style="Card.TFrame")
        top_btn_row.grid(row=2, column=0, columnspan=2, padx=12, pady=6, sticky="ew")
        top_btn_row.grid_columnconfigure(0, weight=1)
        top_btn_row.grid_columnconfigure(1, weight=1)
        top_btn_row.grid_columnconfigure(2, weight=1)
        top_btn_row.grid_columnconfigure(3, weight=1)

        btn_load = ttk.Button(top_btn_row, text=self.tr("import"), style="Primary.TButton", command=self.load_words)
        btn_load.grid(row=0, column=0, padx=(0, 6), sticky="ew")
        Tooltip(btn_load, "Import Words")

        btn_manual = ttk.Button(top_btn_row, text=self.tr("paste_type"), command=self.open_manual_words_window)
        btn_manual.grid(row=0, column=1, padx=3, sticky="ew")
        Tooltip(btn_manual, "Type/Paste words or a two-column table")

        self.save_as_btn = ttk.Button(top_btn_row, text=self.tr("save_as"), command=self.save_words_as)
        self.save_as_btn.grid(row=0, column=2, padx=3, sticky="ew")
        Tooltip(self.save_as_btn, "Save the current list to a txt or csv file")

        self.new_list_btn = ttk.Button(top_btn_row, text=self.tr("new_list"), command=self.new_blank_list)
        self.new_list_btn.grid(row=0, column=3, padx=(6, 0), sticky="ew")
        Tooltip(self.new_list_btn, "Create a new empty list")

        table_wrap = ttk.Frame(self.left, style="Card.TFrame")
        table_wrap.grid(row=3, column=0, columnspan=2, padx=12, pady=(4, 6), sticky="nsew")
        self.left.grid_rowconfigure(3, weight=1)
        self.left.grid_columnconfigure(0, weight=1)
        self.left.grid_columnconfigure(1, weight=1)
        self.word_table = ttk.Treeview(
            table_wrap,
            columns=("idx", "word", "note"),
            show="headings",
            height=18,
            selectmode="extended",
            style="WordList.Treeview",
        )
        self.word_table.heading("idx", text="#")
        self.word_table.heading("word", text=self.tr("word"))
        self.word_table.heading("note", text=self.tr("notes"))
        self.word_table.column("idx", width=70, anchor="center", stretch=False)
        self.word_table.column("word", width=500, anchor="w")
        self.word_table.column("note", width=240, anchor="w")

        self.word_table_scroll = ttk.Scrollbar(table_wrap, orient="vertical", command=self.word_table.yview)
        self.word_table.configure(yscrollcommand=self.word_table_scroll.set)
        self.word_table.grid(row=0, column=0, sticky="nsew")
        self.word_table_scroll.grid(row=0, column=1, sticky="ns")
        self.word_table.tag_configure("even", background="#ffffff")
        self.word_table.tag_configure("odd", background="#fbfcfe")
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
        self.word_context_menu.add_command(label="编辑词性/翻译", command=self.edit_selected_word_meta)
        self.word_context_menu.add_command(label="Find", command=self.search_selected_word_in_corpus)
        self.word_context_menu.add_command(label="造句", command=self.make_sentence_for_selected_word)
        self.word_context_menu.add_command(label=self.tr("lookup_synonyms"), command=self.lookup_synonyms_for_selected_word)
        self.word_context_menu.add_command(label=self.tr("inspect_audio_cache"), command=self.inspect_selected_word_audio_cache)
        self.word_context_menu.add_separator()
        self.word_context_menu.add_command(label=self.tr("delete_word"), command=self.delete_selected_word)

        self.empty_label = ttk.Label(
            self.left,
            text=self.tr("no_words"),
            style="Card.TLabel",
            foreground="#666",
        )
        self.empty_label.grid(row=4, column=0, columnspan=2, padx=12, pady=(0, 8), sticky="w")

        self.player_frame = ttk.Frame(self.left, style="Card.TFrame")
        self.player_frame.grid(row=5, column=0, columnspan=2, padx=12, pady=(4, 6), sticky="ew")
        self.player_frame.grid_columnconfigure(0, weight=1)
        self.player_frame.grid_columnconfigure(1, weight=1)
        self.player_frame.grid_columnconfigure(2, weight=1)

        self.play_btn = ttk.Button(
            self.player_frame, text=self.tr("play"), style="Icon.TButton", command=self.toggle_play
        )
        self.play_btn.grid(row=0, column=0, padx=4, sticky="ew")
        Tooltip(self.play_btn, "Start / Pause")

        self.settings_btn = ttk.Button(
            self.player_frame, text=self.tr("settings"), style="Icon.TButton", command=self.toggle_settings
        )
        self.settings_btn.grid(row=0, column=1, padx=4, sticky="ew")
        Tooltip(self.settings_btn, "Settings")

        self.dictation_btn = ttk.Button(
            self.player_frame, text=self.tr("dictation"), style="Icon.TButton", command=self.open_dictation_window
        )
        self.dictation_btn.grid(row=0, column=2, padx=4, sticky="ew")
        Tooltip(self.dictation_btn, "Open Dictation Window")

        self.status_label = ttk.Label(
            self.left, textvariable=self.status_var, style="Card.TLabel", foreground="#444"
        )
        self.status_label.grid(row=6, column=0, columnspan=2, padx=12, pady=(0, 12), sticky="w")

        # Settings popup (created on demand)
        self.settings_window = None

        self.right.grid_rowconfigure(1, weight=1)
        self.right.grid_columnconfigure(0, weight=1)

        self.detail_card = ttk.Frame(self.right, style="Card.TFrame")
        self.detail_card.grid(row=0, column=0, padx=12, pady=(12, 6), sticky="ew")
        self.detail_card.grid_columnconfigure(0, weight=1)

        ttk.Label(self.detail_card, text=self.tr("current_word"), style="Card.TLabel").grid(
            row=0, column=0, padx=12, pady=(12, 4), sticky="w"
        )
        ttk.Label(
            self.detail_card,
            textvariable=self.detail_word_var,
            style="Card.TLabel",
            font=("Segoe UI", 16, "bold"),
        ).grid(row=1, column=0, padx=12, sticky="w")
        ttk.Label(
            self.detail_card,
            textvariable=self.detail_translation_var,
            style="Card.TLabel",
            foreground="#1e3a8a",
        ).grid(row=2, column=0, padx=12, pady=(2, 0), sticky="w")
        ttk.Label(
            self.detail_card,
            textvariable=self.detail_note_var,
            style="Card.TLabel",
            foreground="#4b5563",
            wraplength=420,
            justify="left",
        ).grid(row=3, column=0, padx=12, pady=(4, 0), sticky="w")
        ttk.Label(
            self.detail_card,
            textvariable=self.detail_meta_var,
            style="Card.TLabel",
            foreground="#667085",
            wraplength=420,
            justify="left",
        ).grid(row=4, column=0, padx=12, pady=(6, 12), sticky="w")

        self.right_notebook = ttk.Notebook(self.right)
        self.right_notebook.grid(row=1, column=0, padx=12, pady=(0, 12), sticky="nsew")

        self.review_tab = ttk.Frame(self.right_notebook, style="Card.TFrame")
        self.history_tab = ttk.Frame(self.right_notebook, style="Card.TFrame")
        self.tools_tab = ttk.Frame(self.right_notebook, style="Card.TFrame")
        self.right_notebook.add(self.review_tab, text=self.tr("review"))
        self.right_notebook.add(self.history_tab, text=self.tr("history"))
        self.right_notebook.add(self.tools_tab, text=self.tr("tools"))

        self.review_tab.grid_columnconfigure(0, weight=1)
        review_card = ttk.Frame(self.review_tab, style="Card.TFrame")
        review_card.pack(fill="both", expand=True, padx=10, pady=10)
        review_card.grid_columnconfigure(0, weight=1)
        ttk.Label(review_card, text=self.tr("study_focus"), style="Card.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            review_card,
            textvariable=self.review_focus_var,
            style="Card.TLabel",
            foreground="#4b5563",
            wraplength=420,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(4, 10))
        ttk.Label(
            review_card,
            textvariable=self.review_source_var,
            style="Card.TLabel",
            foreground="#667085",
            wraplength=420,
            justify="left",
        ).grid(row=2, column=0, sticky="w", pady=(0, 4))
        ttk.Label(
            review_card,
            textvariable=self.review_stats_var,
            style="Card.TLabel",
            foreground="#667085",
            wraplength=420,
            justify="left",
        ).grid(row=3, column=0, sticky="w", pady=(0, 12))
        review_actions = ttk.Frame(review_card, style="Card.TFrame")
        review_actions.grid(row=4, column=0, sticky="ew")
        review_actions.grid_columnconfigure(0, weight=1)
        review_actions.grid_columnconfigure(1, weight=1)
        self.review_open_source_btn = ttk.Button(
            review_actions,
            text=self.tr("open_history"),
            command=self.toggle_history,
        )
        self.review_open_source_btn.grid(row=0, column=0, padx=(0, 6), sticky="ew")
        ttk.Button(
            review_actions,
            text=self.tr("open_tools"),
            command=lambda: self._select_sidebar_tab("tools"),
        ).grid(row=0, column=1, padx=(6, 0), sticky="ew")

        # History tab
        self.history_panel = self.history_tab

        hist_title = ttk.Label(self.history_panel, text=self.tr("history"), style="Card.TLabel")
        hist_title.grid(row=0, column=0, padx=10, pady=(10, 4), sticky="w")

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
        self.history_list.bind("<Button-3>", self.on_history_right_click)

        self.history_context_menu = tk.Menu(self, tearoff=0)
        self.history_context_menu.add_command(label=self.tr("rename_history_file"), command=self.rename_selected_history_item)
        self.history_context_menu.add_command(label=self.tr("delete_history"), command=self.delete_selected_history_item)

        self.history_empty = ttk.Label(
            self.history_panel, text=self.tr("no_history"), style="Card.TLabel", foreground="#666"
        )
        self.history_empty.grid(row=2, column=0, columnspan=2, padx=10, pady=(0, 6), sticky="w")

        btn_open = ttk.Button(self.history_panel, text="⭳", command=self.on_history_open)
        btn_open.grid(row=3, column=0, padx=10, pady=(0, 10), sticky="w")
        Tooltip(btn_open, "Import Selected")

        self.history_path = ttk.Label(self.history_panel, text="", style="Card.TLabel")
        self.history_path.grid(row=3, column=1, padx=10, pady=(0, 10), sticky="e")

        # Tools tab
        self.tools_tab.grid_columnconfigure(0, weight=1)
        tools_wrap = ttk.Frame(self.tools_tab, style="Card.TFrame")
        tools_wrap.pack(fill="both", expand=True, padx=10, pady=10)
        tools_wrap.grid_columnconfigure(0, weight=1)
        tools_wrap.grid_columnconfigure(1, weight=1)
        ttk.Label(tools_wrap, text=self.tr("learning_tools"), style="Card.TLabel").grid(
            row=0, column=0, columnspan=2, sticky="w"
        )
        ttk.Label(
            tools_wrap,
            textvariable=self.tools_hint_var,
            style="Card.TLabel",
            foreground="#667085",
            wraplength=420,
            justify="left",
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 10))
        self.tools_sentence_btn = ttk.Button(
            tools_wrap,
            text=self.tr("generate_sentence"),
            command=self.make_sentence_for_selected_word,
        )
        self.tools_sentence_btn.grid(row=2, column=0, padx=(0, 6), pady=4, sticky="ew")
        self.tools_find_btn = ttk.Button(
            tools_wrap,
            text=self.tr("find_corpus_sentences"),
            command=self.open_find_window,
        )
        self.tools_find_btn.grid(row=2, column=1, padx=(6, 0), pady=4, sticky="ew")
        self.tools_passage_btn = ttk.Button(
            tools_wrap,
            text=self.tr("generate_ielts_passage"),
            command=self.open_passage_window,
        )
        self.tools_passage_btn.grid(row=3, column=0, padx=(0, 6), pady=4, sticky="ew")
        self.tools_settings_btn = ttk.Button(
            tools_wrap,
            text=self.tr("voice_model_settings"),
            command=self.toggle_settings,
        )
        self.tools_settings_btn.grid(row=3, column=1, padx=(6, 0), pady=4, sticky="ew")

        ttk.Label(
            tools_wrap,
            text=self.tr("tools_tip"),
            style="Card.TLabel",
            foreground="#4b5563",
            wraplength=420,
            justify="left",
        ).grid(row=4, column=0, columnspan=2, sticky="w", pady=(10, 0))

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
        self._build_dictation_panel(self.check_panel)
        self.refresh_dictation_recent_list()
        self._show_dictation_frame(self.dictation_setup_frame)

    def close_dictation_window(self):
        self.close_dictation_mode_picker()
        self._cancel_dictation_timer()
        self._cancel_dictation_feedback_reset()
        if self.dictation_window and self.dictation_window.winfo_exists():
            self.dictation_window.destroy()
        self.dictation_window = None
        self.check_panel = None
        self.dictation_setup_frame = None
        self.dictation_pick_frame = None
        self.dictation_session_frame = None
        self.dictation_result_frame = None
        self.dictation_recent_list = None
        self.dictation_start_word_list = None
        self.dictation_mode_hint_label = None
        self.dictation_input = None
        self.dictation_result_label = None
        self.dictation_progress = None
        self.play_btn_check = None
        self.dictation_speed_buttons = []
        self.dictation_feedback_buttons = []

    def _build_dictation_panel(self, parent):
        self.dictation_setup_frame = ttk.Frame(parent, style="Card.TFrame")
        self.dictation_setup_frame.grid(row=0, column=0, sticky="nsew")
        self.dictation_setup_frame.grid_columnconfigure(0, weight=1)
        self.dictation_setup_frame.grid_rowconfigure(2, weight=1)
        tab_row = ttk.Frame(self.dictation_setup_frame, style="Card.TFrame")
        tab_row.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        tab_row.grid_columnconfigure(0, weight=1)
        tab_row.grid_columnconfigure(1, weight=1)
        self.dictation_all_tab_btn = ttk.Button(
            tab_row,
            textvariable=self.dictation_all_tab_var,
            command=lambda: self.set_dictation_list_mode("all"),
        )
        self.dictation_all_tab_btn.grid(row=0, column=0, padx=(0, 6), sticky="ew")
        self.dictation_recent_tab_btn = ttk.Button(
            tab_row,
            textvariable=self.dictation_recent_tab_var,
            command=lambda: self.set_dictation_list_mode("recent"),
        )
        self.dictation_recent_tab_btn.grid(row=0, column=1, padx=(6, 0), sticky="ew")

        ttk.Label(
            self.dictation_setup_frame,
            textvariable=self.dictation_status_var,
            style="Card.TLabel",
            foreground="#667085",
            wraplength=640,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(0, 8))

        recent_wrap = ttk.Frame(self.dictation_setup_frame, style="Card.TFrame")
        recent_wrap.grid(row=2, column=0, sticky="nsew")
        recent_wrap.grid_columnconfigure(0, weight=1)
        recent_wrap.grid_rowconfigure(0, weight=1)
        self.dictation_recent_list = ttk.Treeview(
            recent_wrap,
            columns=("idx", "word", "meta"),
            show="headings",
            selectmode="browse",
            height=14,
            style="WordList.Treeview",
        )
        self.dictation_recent_list.heading("idx", text="#")
        self.dictation_recent_list.heading("word", text=self.tr("word"))
        self.dictation_recent_list.heading("meta", text=self.tr("notes"))
        self.dictation_recent_list.column("idx", width=60, anchor="center", stretch=False)
        self.dictation_recent_list.column("word", width=430, minwidth=240, anchor="w", stretch=True)
        self.dictation_recent_list.column("meta", width=220, minwidth=120, anchor="w", stretch=True)
        self.dictation_recent_list.grid(row=0, column=0, sticky="nsew")
        self.dictation_recent_list.tag_configure("even", background="#ffffff")
        self.dictation_recent_list.tag_configure("odd", background="#fbfcfe")
        self.dictation_recent_list.bind("<<TreeviewSelect>>", self.on_dictation_list_selected)
        recent_scroll = ttk.Scrollbar(recent_wrap, orient="vertical", command=self.dictation_recent_list.yview)
        recent_scroll.grid(row=0, column=1, sticky="ns")
        self.dictation_recent_list.configure(yscrollcommand=recent_scroll.set)
        self.dictation_recent_list.bind("<Button-3>", self.on_dictation_word_right_click)

        setup_row = ttk.Frame(self.dictation_setup_frame, style="Card.TFrame")
        setup_row.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        setup_row.grid_columnconfigure(0, weight=1)
        setup_row.grid_columnconfigure(1, weight=1)
        setup_row.grid_columnconfigure(2, weight=1)
        ttk.Button(setup_row, text=self.tr("start_from_word"), command=self.open_dictation_start_word_picker).grid(
            row=0, column=0, padx=(0, 6), sticky="ew"
        )
        ttk.Button(setup_row, text=self.tr("start_learning"), command=self.open_dictation_mode_picker).grid(
            row=0, column=1, padx=(6, 0), sticky="ew"
        )
        ttk.Button(setup_row, text=self.tr("add_wrong_word"), command=self.add_manual_wrong_word).grid(
            row=0, column=2, padx=(6, 0), sticky="ew"
        )

        self.dictation_pick_frame = ttk.Frame(parent, style="Card.TFrame")
        self.dictation_pick_frame.grid(row=0, column=0, sticky="nsew")
        self.dictation_pick_frame.grid_remove()
        self.dictation_pick_frame.grid_columnconfigure(0, weight=1)
        self.dictation_pick_frame.grid_rowconfigure(2, weight=1)

        pick_header = ttk.Frame(self.dictation_pick_frame, style="Card.TFrame")
        pick_header.grid(row=0, column=0, sticky="ew")
        pick_header.grid_columnconfigure(1, weight=1)
        ttk.Button(
            pick_header,
            text=self.tr("back"),
            command=lambda: self._show_dictation_frame(self.dictation_setup_frame),
        ).grid(row=0, column=0, padx=(0, 8), sticky="w")
        ttk.Label(
            pick_header,
            textvariable=self.dictation_status_var,
            style="Card.TLabel",
            justify="left",
            wraplength=640,
        ).grid(row=0, column=1, sticky="w")
        ttk.Label(self.dictation_pick_frame, text=self.tr("from_selected_word"), style="Card.TLabel").grid(
            row=1, column=0, sticky="w", pady=(10, 8)
        )

        pick_wrap = ttk.Frame(self.dictation_pick_frame, style="Card.TFrame")
        pick_wrap.grid(row=2, column=0, sticky="nsew")
        pick_wrap.grid_columnconfigure(0, weight=1)
        pick_wrap.grid_rowconfigure(0, weight=1)
        self.dictation_start_word_list = ttk.Treeview(
            pick_wrap,
            columns=("idx", "word"),
            show="headings",
            selectmode="browse",
            height=14,
            style="WordList.Treeview",
        )
        self.dictation_start_word_list.heading("idx", text="#")
        self.dictation_start_word_list.heading("word", text=self.tr("word"))
        self.dictation_start_word_list.column("idx", width=60, anchor="center", stretch=False)
        self.dictation_start_word_list.column("word", width=560, minwidth=260, anchor="w", stretch=True)
        self.dictation_start_word_list.grid(row=0, column=0, sticky="nsew")
        self.dictation_start_word_list.tag_configure("even", background="#ffffff")
        self.dictation_start_word_list.tag_configure("odd", background="#fbfcfe")
        pick_scroll = ttk.Scrollbar(pick_wrap, orient="vertical", command=self.dictation_start_word_list.yview)
        pick_scroll.grid(row=0, column=1, sticky="ns")
        self.dictation_start_word_list.configure(yscrollcommand=pick_scroll.set)
        self.dictation_start_word_list.bind("<Double-1>", self._on_dictation_start_word_double_click)
        self.dictation_start_word_list.bind("<Button-3>", self.on_dictation_word_right_click)

        pick_actions = ttk.Frame(self.dictation_pick_frame, style="Card.TFrame")
        pick_actions.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        pick_actions.grid_columnconfigure(0, weight=1)
        pick_actions.grid_columnconfigure(1, weight=1)
        ttk.Button(
            pick_actions,
            text=self.tr("cancel"),
            command=lambda: self._show_dictation_frame(self.dictation_setup_frame),
        ).grid(row=0, column=0, padx=(0, 6), sticky="ew")
        ttk.Button(pick_actions, text=self.tr("start_here"), command=self._confirm_dictation_start_word).grid(
            row=0, column=1, padx=(6, 0), sticky="ew"
        )

        self.dictation_session_frame = ttk.Frame(parent, style="Card.TFrame")
        self.dictation_session_frame.grid(row=0, column=0, sticky="nsew")
        self.dictation_session_frame.grid_remove()
        self.dictation_session_frame.grid_columnconfigure(0, weight=1)

        session_header = ttk.Frame(self.dictation_session_frame, style="Card.TFrame")
        session_header.grid(row=0, column=0, sticky="ew")
        session_header.grid_columnconfigure(0, weight=1)
        session_header.grid_columnconfigure(1, weight=0)
        ttk.Label(session_header, textvariable=self.dictation_progress_var, style="Card.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(session_header, text=self.tr("answer"), style="Card.TLabel", foreground="#667085").grid(
            row=0, column=1, sticky="e"
        )

        self.dictation_progress = ttk.Progressbar(
            self.dictation_session_frame,
            orient="horizontal",
            mode="determinate",
            maximum=100,
        )
        self.dictation_progress.grid(row=1, column=0, sticky="ew", pady=(8, 10))

        input_card = ttk.Frame(self.dictation_session_frame, style="Card.TFrame")
        input_card.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        input_card.grid_columnconfigure(0, weight=1)
        input_card.grid_columnconfigure(1, weight=0)
        self.dictation_input = tk.Entry(
            input_card,
            font=("Segoe UI", 24, "bold"),
            relief="solid",
            bd=1,
            bg="#f6f6f8",
            fg="#202020",
            insertbackground="#202020",
        )
        self.dictation_input.grid(row=0, column=0, sticky="ew", padx=(0, 8), ipady=20)
        self.dictation_input.bind("<KeyRelease>", self.on_dictation_input_change)
        self.dictation_input.bind("<Return>", self.on_dictation_enter)
        ttk.Label(
            input_card,
            textvariable=self.dictation_timer_var,
            font=("Segoe UI", 18),
            style="Card.TLabel",
            foreground="#6b7280",
        ).grid(row=0, column=1, sticky="ne")

        self.dictation_result_label = ttk.Label(
            self.dictation_session_frame,
            textvariable=self.dictation_status_var,
            style="Card.TLabel",
            foreground="#667085",
        )
        self.dictation_result_label.grid(row=3, column=0, sticky="w", pady=(8, 10))

        ttk.Button(
            self.dictation_session_frame,
            text=self.tr("next_word"),
            style="Primary.TButton",
            command=self.advance_dictation_word,
        ).grid(row=4, column=0, sticky="ew")

        control_row = ttk.Frame(self.dictation_session_frame, style="Card.TFrame")
        control_row.grid(row=5, column=0, sticky="ew", pady=(14, 0))
        control_row.grid_columnconfigure(0, weight=1)
        control_row.grid_columnconfigure(1, weight=1)
        control_row.grid_columnconfigure(2, weight=1)
        control_row.grid_columnconfigure(3, weight=1)
        ttk.Button(control_row, text=self.tr("dictation_settings"), command=lambda: self.open_dictation_mode_picker(auto_start=False)).grid(
            row=0, column=0, padx=(0, 6), sticky="ew"
        )
        ttk.Button(control_row, text=self.tr("previous_word"), command=self.previous_dictation_word).grid(
            row=0, column=1, padx=3, sticky="ew"
        )
        self.play_btn_check = ttk.Button(control_row, text=self.tr("play"), command=self.toggle_dictation_play_pause)
        self.play_btn_check.grid(row=0, column=2, padx=3, sticky="ew")
        ttk.Button(control_row, text=self.tr("replay"), command=self.replay_dictation_word).grid(
            row=0, column=3, padx=(6, 0), sticky="ew"
        )

        self.dictation_result_frame = ttk.Frame(parent, style="Card.TFrame")
        self.dictation_result_frame.grid(row=0, column=0, sticky="nsew")
        self.dictation_result_frame.grid_remove()
        self.dictation_result_frame.grid_columnconfigure(0, weight=1)
        ttk.Label(
            self.dictation_result_frame,
            textvariable=self.dictation_summary_var,
            font=("Segoe UI", 24, "bold"),
            style="Card.TLabel",
            foreground="#5b5cf0",
        ).grid(row=0, column=0, sticky="n", pady=(40, 10))
        ttk.Label(self.dictation_result_frame, text=self.tr("current_session_accuracy"), style="Card.TLabel").grid(
            row=1, column=0, sticky="n"
        )
        ttk.Button(
            self.dictation_result_frame,
            text=self.tr("view_wrong_words"),
            style="Primary.TButton",
            command=self.show_dictation_wrong_words,
        ).grid(row=2, column=0, sticky="ew", padx=80, pady=(30, 10))
        ttk.Button(self.dictation_result_frame, text=self.tr("back_to_list"), command=self.reset_dictation_view).grid(
            row=3, column=0, sticky="ew", padx=80
        )

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
        self.store.clear()
        self.translations = {}
        self.word_pos = {}
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
            self.store.save_to_file(path)
            self.store.add_history(path)
            if was_manual_session and words_snapshot:
                tts_rebind_manual_session_cache_to_source(words_snapshot, path)
            tts_set_preferred_pending_source(path)
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

        if selected_idx is None or selected_idx >= total_words:
            self.detail_word_var.set("No word selected")
            self.detail_translation_var.set("")
            self.detail_note_var.set("Select a word to see notes and translation.")
            self.detail_meta_var.set(f"Words loaded: {total_words}")
            self.review_focus_var.set("Focus: select a word or start playback.")
        else:
            word = self.store.words[selected_idx]
            note = self.store.notes[selected_idx] if selected_idx < len(self.store.notes) else ""
            zh = self.translations.get(word) or ""
            pos_label = str(self.word_pos.get(word) or "").strip()
            self.detail_word_var.set(word)
            detail_line = []
            if pos_label:
                detail_line.append(pos_label)
            if zh:
                detail_line.append(zh)
            self.detail_translation_var.set(" ".join(detail_line) if detail_line else "词性/中文: loading or unavailable")
            self.detail_note_var.set(f"Notes: {note}" if note else "Notes: none")
            meta_parts = [f"Position: {selected_idx + 1}/{total_words}"]
            if self.current_word == word:
                meta_parts.append("Playback focus")
            self.detail_meta_var.set(" | ".join(meta_parts))
            self.review_focus_var.set(
                f"Focus: {word}. Use Dictation for recall, or Tools for sentence and corpus lookup."
            )

        current_source = self.store.get_current_source_path()
        if current_source:
            source_name = os.path.basename(current_source)
        elif self._has_unsaved_manual_words():
            source_name = "manual list (unsaved)"
        else:
            source_name = "manual list"
        self.review_source_var.set(f"Source file: {source_name}")
        self.review_stats_var.set(
            f"Words: {total_words} | Current mode: {self.order_mode.get()} | Current status: {self.play_state}"
        )
        has_selection = selected_idx is not None and selected_idx < total_words
        selection_state = "normal" if has_selection else "disabled"
        for btn in (
            self.tools_sentence_btn,
            self.tools_find_btn,
        ):
            if btn:
                btn.config(state=selection_state)
        if self.tools_passage_btn:
            self.tools_passage_btn.config(state=("normal" if total_words else "disabled"))
        if self.save_as_btn:
            self.save_as_btn.config(state=("normal" if total_words else "disabled"))
        if self.new_list_btn:
            self.new_list_btn.config(state="normal")

    def _format_word_subline(self, word):
        pos_label = str(self.word_pos.get(word) or "").strip()
        zh_text = str(self.translations.get(word) or "").strip()
        parts = [part for part in (pos_label, zh_text) if part]
        if parts:
            return " ".join(parts)
        if zh_text:
            return zh_text
        if pos_label:
            return pos_label
        return "..."

    def _build_word_table_values(self, idx, word, note=None):
        note_value = self.store.notes[idx] if note is None and idx < len(self.store.notes) else (note or "")
        display_text = f"{word}\n{self._format_word_subline(word)}"
        return (f"{idx + 1}.", display_text, note_value)

    def _build_dictation_table_values(self, idx, item):
        word = str(item.get("word") or "").strip()
        subtitle = self._format_word_subline(word)
        note_value = ""
        try:
            note_index = self.store.words.index(word)
        except ValueError:
            note_index = -1
        if 0 <= note_index < len(self.store.notes):
            note_value = str(self.store.notes[note_index] or "").strip()
        if self.dictation_list_mode_var.get() == "recent":
            wrong_count = int(item.get("wrong_count", 0) or 0)
            wrong_input = str(item.get("last_wrong_input") or "").strip()
            wrong_type = str(item.get("last_wrong_type") or "").strip()
            if wrong_count:
                subtitle = f"{subtitle}  |  错过{wrong_count}次"
            note_parts = []
            if wrong_type:
                note_parts.append(wrong_type)
            if wrong_input:
                note_parts.append(f"错写: {wrong_input}")
            note_value = " | ".join(note_parts)
        else:
            if item.get("wrong_count"):
                subtitle = f"{subtitle}  |  错过{int(item.get('wrong_count', 0) or 0)}次"
        return (f"{idx + 1}.", f"{word}\n{subtitle}", note_value)

    def _start_analysis_job(self, words, token):
        if not words:
            return

        def _run():
            analyzed = analyze_words(words)
            self.after(0, lambda: self._apply_pos_analysis(token, analyzed))

        import threading

        threading.Thread(target=_run, daemon=True).start()

    def _apply_pos_analysis(self, token, analyzed):
        if token != self.analysis_token or not self.word_table:
            return
        self.word_pos.update(analyzed)
        for idx, word in enumerate(self.store.words):
            if not self.word_table.exists(str(idx)):
                continue
            note = self.store.notes[idx] if idx < len(self.store.notes) else ""
            tag = "even" if idx % 2 == 0 else "odd"
            self.word_table.item(str(idx), values=self._build_word_table_values(idx, word, note), tags=(tag,))
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
        self.dictation_all_items = [
            {
                "word": word,
                "wrong_count": 0,
                "correct_count": 0,
                "last_wrong_input": "",
                "last_result": "",
                "last_seen": "",
            }
            for word in self.store.words
        ]
        self.dictation_recent_items = self.store.recent_wrong_words(self.store.words, limit=100)
        self.dictation_all_tab_var.set(f"全部({len(self.dictation_all_items)})")
        self.dictation_recent_tab_var.set(f"近期错词({len(self.dictation_recent_items)})")
        self.dictation_recent_list.delete(*self.dictation_recent_list.get_children())
        items = self._get_dictation_source_items()
        if items:
            for idx, item in enumerate(items, start=1):
                tag = "even" if (idx - 1) % 2 == 0 else "odd"
                self.dictation_recent_list.insert(
                    "",
                    tk.END,
                    iid=str(idx - 1),
                    values=self._build_dictation_table_values(idx - 1, item),
                    tags=(tag,),
                )
            self.suppress_dictation_select_action = True
            self.dictation_recent_list.selection_set("0")
            self.dictation_recent_list.focus("0")
        else:
            empty_text = self.tr("dictation_empty_recent") if self.dictation_list_mode_var.get() == "recent" and self.store.words else (
                self.tr("dictation_empty_list") if self.store.words else self.tr("import_words_first")
            )
            self.dictation_recent_list.insert("", tk.END, iid="empty", values=("", empty_text, ""))
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
            self.dictation_pick_frame,
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
        if not self.store.words:
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

    def open_dictation_start_word_picker(self):
        items = self._get_dictation_source_items()
        if not items:
            self.show_info("no_words_available")
            return
        title_text = self.tr("dictation_all_start") if self.dictation_list_mode_var.get() == "all" else self.tr("dictation_recent_start")
        self.dictation_status_var.set(title_text)
        if not self.dictation_start_word_list:
            return
        self.dictation_start_word_list.delete(*self.dictation_start_word_list.get_children())
        for idx, item in enumerate(items, start=1):
            word = item.get("word") or ""
            detail = self._format_word_subline(word)
            tag = "even" if (idx - 1) % 2 == 0 else "odd"
            self.dictation_start_word_list.insert(
                "",
                tk.END,
                iid=str(idx - 1),
                values=(f"{idx}.", f"{word}\n{detail}"),
                tags=(tag,),
            )
        self.dictation_start_word_list.selection_set("0")
        self.dictation_start_word_list.focus("0")
        self._show_dictation_frame(self.dictation_pick_frame)

    def refresh_dictation_start_word_picker(self):
        if not self.dictation_start_word_list:
            return
        self.dictation_start_word_list.delete(*self.dictation_start_word_list.get_children())
        items = self._get_dictation_source_items()
        if not items:
            return
        for idx, item in enumerate(items, start=1):
            word = item.get("word") or ""
            detail = self._format_word_subline(word)
            tag = "even" if (idx - 1) % 2 == 0 else "odd"
            self.dictation_start_word_list.insert(
                "",
                tk.END,
                iid=str(idx - 1),
                values=(f"{idx}.", f"{word}\n{detail}"),
                tags=(tag,),
            )
        self.dictation_start_word_list.selection_set("0")
        self.dictation_start_word_list.focus("0")

    def _on_dictation_start_word_double_click(self, _event=None):
        self._confirm_dictation_start_word()

    def _confirm_dictation_start_word(self):
        if not self.dictation_start_word_list:
            return
        selection = self.dictation_start_word_list.selection()
        if not selection:
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
            words = [item["word"] for item in items if item.get("word") in self.store.words]
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
        if store_index is None or store_index >= len(self.store.words):
            return
        word = self.store.words[store_index]
        self._sync_main_selection_to_index(store_index)
        speak_async(
            word,
            self.volume_var.get() / 100.0,
            rate_ratio=self.speech_rate_var.get(),
            cancel_before=True,
            source_path=self._get_dictation_preview_source_path(),
        )

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

    def start_online_spelling_session(self, start_index=0):
        self.dictation_pool = self._get_dictation_pool()
        if not self.dictation_pool:
            self.show_info("no_words_for_dictation")
            self.reset_dictation_view()
            return
        self.dictation_session_source_path = self._get_dictation_preview_source_path()
        safe_start = max(0, min(int(start_index or 0), max(0, len(self.dictation_pool) - 1)))
        if safe_start > 0:
            self.dictation_pool = self.dictation_pool[safe_start:] + self.dictation_pool[:safe_start]
        self.dictation_index = -1
        self.dictation_wrong_items = []
        self.dictation_correct_count = 0
        self.dictation_current_word = ""
        self.dictation_answer_revealed = False
        self.dictation_running = True
        self.dictation_paused = False
        self.dictation_summary_var.set("")
        self.update_dictation_play_button()
        self._show_dictation_frame(self.dictation_session_frame)
        self.advance_dictation_word(initial=True)

    def play_dictation_current_word(self):
        if not self.dictation_running or not self.dictation_current_word:
            return
        self.dictation_paused = False
        self.update_dictation_play_button()
        self.status_var.set(self.trf("dictation_playing", word=self.dictation_current_word))
        speak_async(
            self.dictation_current_word,
            self.volume_var.get() / 100.0,
            rate_ratio=1.0 if self.dictation_speed_var.get() == "adaptive" else float(self.dictation_speed_var.get()),
            cancel_before=True,
            source_path=self.dictation_session_source_path or self._get_dictation_preview_source_path(),
        )
        self._restart_dictation_timer()
        self._focus_dictation_input()

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
        self._cancel_dictation_feedback_reset()
        self._cancel_dictation_timer()
        tts_cancel_all()
        self.dictation_paused = False
        self.dictation_index = max(-1, self.dictation_index - 2)
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
            self.dictation_timer_var.set("manual")
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
        self.store.record_dictation_result(target, user_text, is_correct)
        if is_correct:
            self.dictation_correct_count += 1
            self._set_dictation_input_color("correct")
            self.dictation_status_var.set(self.tr("dictation_correct"))
            if self.dictation_session_source_path == self._get_recent_wrong_cache_source_path():
                self.store.clear_wrong_word(target)
                tts_cleanup_word_audio_cache(target, source_path=self._get_recent_wrong_cache_source_path())
                self.refresh_dictation_recent_list()
        else:
            self._set_dictation_input_color("wrong")
            self.dictation_status_var.set(self.trf("dictation_wrong_answer", word=target))
            self.dictation_wrong_items.append({"word": target, "input": user_text})
            tts_promote_word_audio_to_recent_wrong(
                target,
                source_path=self.dictation_session_source_path or self.store.get_current_source_path(),
            )
        delay = 1150 if trigger == "input" and is_correct else 1450
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
        self.dictation_timer_var.set(f"{self._dictation_seconds_for_speed()}s" if self._dictation_seconds_for_speed() else "manual")
        if self.dictation_input:
            self.dictation_input.delete(0, tk.END)
            self.dictation_input.focus_set()
        self._set_dictation_input_color("neutral")
        self.update_dictation_play_button()
        self.play_dictation_current_word()

    def finish_dictation_session(self):
        self.dictation_running = False
        self.dictation_paused = False
        self._cancel_dictation_timer()
        self._cancel_dictation_feedback_reset()
        self.update_dictation_play_button()
        total = len(self.dictation_pool)
        accuracy = (self.dictation_correct_count / float(total)) * 100.0 if total else 0.0
        self.dictation_summary_var.set(f"{accuracy:.2f}%")
        self.dictation_status_var.set(self.tr("dictation_session_complete"))
        self.refresh_dictation_recent_list()
        self._show_dictation_frame(self.dictation_result_frame)

    def reset_dictation_view(self):
        self.dictation_running = False
        self.dictation_paused = False
        self.dictation_pool = []
        self.dictation_index = -1
        self.dictation_current_word = ""
        self.dictation_session_source_path = None
        self._cancel_dictation_timer()
        self._cancel_dictation_feedback_reset()
        self.update_dictation_play_button()
        self.dictation_progress_var.set("Spelling (0/0)")
        self.dictation_timer_var.set("5s")
        self.dictation_status_var.set(self.tr("dictation_recent_title"))
        self.dictation_progress["value"] = 0
        if self.dictation_input:
            self.dictation_input.delete(0, tk.END)
        self._set_dictation_input_color("neutral")
        self._show_dictation_frame(self.dictation_setup_frame)
        self.refresh_dictation_recent_list()

    def show_dictation_wrong_words(self):
        if not self.dictation_wrong_items:
            self.show_info("no_wrong_words_session")
            return
        lines = []
        for item in self.dictation_wrong_items:
            lines.append(f"{item['word']}    <-    {item.get('input') or '(blank)'}")
        messagebox.showinfo(self.tr("wrong_words"), "\n".join(lines[:40]))

    def inspect_selected_word_audio_cache(self):
        selected_idx = self._get_context_or_selected_index()
        if selected_idx is None or selected_idx >= len(self.store.words):
            self.show_info("select_word_first")
            return
        word = self.store.words[selected_idx]
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

    def edit_selected_word_meta(self):
        selected_idx = self._get_context_or_selected_index()
        if selected_idx is None or selected_idx >= len(self.store.words):
            messagebox.showinfo("Info", "Please select a word first.")
            return
        word = self.store.words[selected_idx]
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
            if self.word_table and self.word_table.exists(iid):
                note = self.store.notes[selected_idx] if selected_idx < len(self.store.notes) else ""
                tag = "even" if selected_idx % 2 == 0 else "odd"
                self.word_table.item(iid, values=self._build_word_table_values(selected_idx, word, note), tags=(tag,))
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
            filetypes=[("Text files", "*.txt"), ("CSV files", "*.csv")],
        )
        if not path:
            return
        self.cancel_word_edit()
        words = self.store.load_from_file(path)
        self.manual_source_dirty = False
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
            self._mark_manual_words_dirty()
            self.render_words(merged_words)
            self._update_word_stats_from_manual_input(normalized_words)
            status_text = f"Appended {len(normalized_words)} words."
        else:
            self.store.set_words(normalized_words, normalized_notes)
            self._mark_manual_words_dirty()
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
        ttk.Label(wrap, text=self.tr("paste_preview"), style="Card.TLabel").pack(anchor="w")
        ttk.Label(
            wrap,
            text=self.tr("paste_preview_desc"),
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
        self.manual_words_table.heading("en", text=self.tr("english"))
        self.manual_words_table.heading("note", text=self.tr("notes"))
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
            text=self.tr("paste_clipboard"),
            command=self.on_manual_preview_paste,
        ).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(
            row,
            text=self.tr("add_row"),
            command=self._add_manual_preview_row,
        ).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(
            row,
            text=self.tr("delete_selected"),
            command=self._delete_selected_manual_preview_rows,
        ).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(
            row,
            text=self.tr("clear"),
            command=self._clear_manual_preview,
        ).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Button(
            row,
            text=self.tr("replace_list"),
            command=lambda: self._apply_manual_words_from_editor("replace"),
        ).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(
            row,
            text=self.tr("append"),
            command=lambda: self._apply_manual_words_from_editor("append"),
        ).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(
            row,
            text=self.tr("close"),
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

    def _set_word_action_context(self, index, origin="main"):
        try:
            idx = int(index)
        except Exception:
            idx = None
        if idx is None or idx < 0 or idx >= len(self.store.words):
            self.word_action_index = None
            self.word_action_origin = "main"
            return None
        self.word_action_index = idx
        self.word_action_origin = str(origin or "main")
        return idx

    def _clear_word_action_context(self):
        self.word_action_index = None
        self.word_action_origin = "main"

    def _get_context_or_selected_index(self):
        if self.word_action_index is not None and 0 <= self.word_action_index < len(self.store.words):
            return self.word_action_index
        return self._get_selected_index()

    def _get_context_audio_source_path(self):
        if self.word_action_origin == "dictation" and self.dictation_list_mode_var.get() == "recent":
            return self._get_recent_wrong_cache_source_path()
        return self.store.get_current_source_path()

    def _sync_main_selection_to_index(self, index):
        try:
            idx = int(index)
        except Exception:
            return None
        if idx < 0 or idx >= len(self.store.words):
            return None
        self._set_word_action_context(idx)
        if self.word_table and self.word_table.exists(str(idx)):
            try:
                self.word_table.selection_set(str(idx))
                self.word_table.focus(str(idx))
            except Exception:
                pass
        self._refresh_selection_details()
        return idx

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
        if tree is self.dictation_start_word_list:
            return view_index if 0 <= view_index < len(self.store.words) else None
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
            tag = "even" if row_idx % 2 == 0 else "odd"
            self.word_table.item(iid, values=self._build_word_table_values(row_idx, word, note), tags=(tag,))
        self._refresh_selection_details()

    def start_edit_selected_word(self, _event=None):
        return self.start_edit_word_cell(column_id="#2")

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

        if edit_column == "#3":
            old_note = self.store.notes[idx] if idx < len(self.store.notes) else ""
            if idx >= len(self.store.notes):
                self.store.notes.extend([""] * (idx - len(self.store.notes) + 1))
            if new_word == old_note:
                return "break"
            self.store.notes[idx] = new_word
            if self.word_table and self.word_table.exists(iid):
                tag = "even" if idx % 2 == 0 else "odd"
                self.word_table.item(
                    iid,
                    values=self._build_word_table_values(idx, self.store.words[idx], new_word),
                    tags=(tag,),
                )
            saved = False
            source_path = str(self.store.get_current_source_path() or "").lower()
            if source_path.endswith(".csv"):
                saved = self._save_words_back_to_source()
            elif not self.store.has_current_source_file():
                self._mark_manual_words_dirty()
            source_note = " and saved to source file" if saved else ""
            self.status_var.set(f"Updated note for '{self.store.words[idx]}'{source_note}.")
            self._refresh_selection_details()
            return "break"

        old_word = self.store.words[idx]
        if new_word == old_word:
            return "break"

        old_note = self.store.notes[idx] if idx < len(self.store.notes) else ""
        self.store.words[idx] = new_word
        if old_word in self.translations and new_word not in self.translations:
            self.translations.pop(old_word, None)
        self.word_pos.pop(old_word, None)
        if self.word_table and self.word_table.exists(iid):
            tag = "even" if idx % 2 == 0 else "odd"
            self.word_table.item(iid, values=(f"{idx + 1}.", f"{new_word}\nTranslating...", old_note), tags=(tag,))

        saved = self._save_words_back_to_source()
        if not saved and not self.store.has_current_source_file():
            self._mark_manual_words_dirty()
        self._translate_single_word_async(idx, new_word)
        self.analysis_token += 1
        self._start_analysis_job([new_word], self.analysis_token)
        source_note = " and saved to source file" if saved else ""
        self.status_var.set(f"Updated '{old_word}' to '{new_word}'{source_note}.")
        self._refresh_selection_details()
        return "break"

    def delete_selected_word(self):
        selected_idx = self._get_context_or_selected_index()
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
        old_word = self.store.words.pop(selected_idx)
        if selected_idx < len(self.store.notes):
            self.store.notes.pop(selected_idx)
        self.translations.pop(old_word, None)
        self.word_pos.pop(old_word, None)

        saved = self._save_words_back_to_source()
        if not saved and not self.store.has_current_source_file():
            self._mark_manual_words_dirty()

        self._clear_word_action_context()
        self.render_words(list(self.store.words))
        self.refresh_dictation_recent_list()
        self.refresh_dictation_start_word_picker()
        if self.store.words and self.word_table:
            next_idx = min(selected_idx, len(self.store.words) - 1)
            if self.word_table.exists(str(next_idx)):
                self.word_table.selection_set(str(next_idx))
                self.word_table.focus(str(next_idx))
        self.status_var.set(self.trf("word_deleted", word=old_word))

    def render_words(self, words):
        if not self.word_table:
            return
        self.cancel_word_edit()
        self.translation_token += 1
        self.analysis_token += 1
        token = self.translation_token
        analysis_token = self.analysis_token
        cached = get_cached_translations(words)
        cached_pos = get_cached_pos(words)
        self.translations = dict(cached)
        self.word_pos = dict(cached_pos)
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
        self.refresh_dictation_start_word_picker()
        missing_words = [w for w in words if w not in cached]
        missing_pos = [w for w in words if w not in cached_pos]
        if missing_words:
            self._start_translation_job(missing_words, token)
        if missing_pos:
            self._start_analysis_job(missing_pos, analysis_token)
        if words:
            self._start_audio_precache_job(words)

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
            note = self.store.notes[idx] if idx < len(self.store.notes) else ""
            tag = "even" if idx % 2 == 0 else "odd"
            self.word_table.item(str(idx), values=self._build_word_table_values(idx, word, note), tags=(tag,))
        self._refresh_selection_details()

    def update_empty_state(self):
        if self.store.words:
            self.empty_label.grid_remove()
        else:
            self.empty_label.grid()
        self._refresh_selection_details()

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
        self.refresh_dictation_recent_list()

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
        sel = self.history_list.curselection()
        if not sel:
            return
        history = self.store.load_history()
        idx = sel[0]
        if idx >= len(history):
            return
        item = history[idx]
        path = str(item.get("path") or "").strip()
        name = str(item.get("name") or os.path.basename(path) or path)
        if not messagebox.askyesno(self.tr("history"), self.trf("delete_history_confirm", name=name)):
            return
        current_source = os.path.abspath(str(self.store.get_current_source_path() or "").strip()) if self.store.get_current_source_path() else ""
        self.store.remove_history(path)
        removed_count = tts_cleanup_cache_for_source_path(path)
        if current_source and current_source == os.path.abspath(path):
            self.store.detach_current_source()
            self.manual_source_dirty = bool(self.store.words)
            self._refresh_selection_details()
        self.refresh_history()
        self.show_info("history_deleted", name=name, count=removed_count)

    def rename_selected_history_item(self):
        sel = self.history_list.curselection()
        if not sel:
            return
        history = self.store.load_history()
        idx = sel[0]
        if idx >= len(history):
            return
        item = history[idx]
        old_path = os.path.abspath(str(item.get("path") or "").strip())
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
        new_name = str(new_name or "").strip()
        if not new_name:
            return
        if os.path.basename(new_name) != new_name:
            self.show_info("rename_history_invalid")
            return
        old_root, old_ext = os.path.splitext(old_name)
        new_root, new_ext = os.path.splitext(new_name)
        if not new_ext and old_ext:
            new_name = f"{new_name}{old_ext}"
        new_path = os.path.join(os.path.dirname(old_path), new_name)
        if os.path.abspath(new_path) == old_path:
            return
        if os.path.exists(new_path):
            self.show_info("rename_history_exists")
            return
        current_source = os.path.abspath(str(self.store.get_current_source_path() or "").strip()) if self.store.get_current_source_path() else ""
        try:
            os.rename(old_path, new_path)
            migration = tts_rename_cache_source_path(old_path, new_path)
            self.store.rename_history_path(old_path, new_path)
            if current_source and current_source == old_path:
                tts_set_preferred_pending_source(new_path)
            self.refresh_history()
            self._refresh_selection_details()
            self.show_info(
                "rename_history_done",
                name=os.path.basename(new_path),
                count=int((migration or {}).get("migrated", 0) or 0),
                queued=int((migration or {}).get("queued", 0) or 0),
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
        if token not in self.store.words:
            self.store.words.append(token)
            self.store.notes.append("")
            self.manual_source_dirty = True
            self.render_words(self.store.words)
        self.store.add_wrong_word(token)
        tts_queue_word_audio_generation(token, source_path=self._get_recent_wrong_cache_source_path())
        self.refresh_dictation_recent_list()
        self._refresh_selection_details()
        self.show_info("wrong_word_added", word=token)

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

        self.find_window = tk.Toplevel(self)
        self.find_window.title("Find Corpus Sentences")
        self.find_window.configure(bg="#f6f7fb")
        self.find_window.minsize(900, 620)

        wrap = ttk.Frame(self.find_window, style="Card.TFrame")
        wrap.pack(fill="both", expand=True, padx=10, pady=10)
        wrap.grid_columnconfigure(0, weight=3)
        wrap.grid_columnconfigure(1, weight=2)
        wrap.grid_rowconfigure(2, weight=1)

        ttk.Label(wrap, text=self.tr("find_corpus_sentences"), style="Card.TLabel").grid(
            row=0, column=0, columnspan=2, sticky="w"
        )
        ttk.Label(
            wrap,
            text=self.tr("find_desc"),
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
        ttk.Button(search_row, text=self.tr("search"), command=self.run_find_search).grid(row=0, column=1, padx=(0, 6))
        ttk.Label(search_row, text=self.tr("show"), style="Card.TLabel").grid(row=0, column=2, padx=(0, 4))
        limit_combo = ttk.Combobox(
            search_row,
            textvariable=self.find_limit_var,
            state="readonly",
            width=5,
            values=("20", "50", "100"),
        )
        limit_combo.grid(row=0, column=3, padx=(0, 6))
        ttk.Label(search_row, text=self.tr("results"), style="Card.TLabel").grid(row=0, column=4, padx=(0, 6))
        ttk.Button(search_row, text=self.tr("use_selected_word"), command=self.search_selected_word_in_corpus).grid(
            row=0, column=5, padx=(0, 6)
        )
        ttk.Button(search_row, text=self.tr("import_docs"), command=self.import_find_documents).grid(row=0, column=6)

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
        self.find_results_table.heading("sentence", text=self.tr("sentence"))
        self.find_results_table.heading("source", text=self.tr("source"))
        self.find_results_table.column("sentence", width=600, anchor="w")
        self.find_results_table.column("source", width=260, anchor="w")
        find_scroll = ttk.Scrollbar(top, orient="vertical", command=self.find_results_table.yview)
        self.find_results_table.configure(yscrollcommand=find_scroll.set)
        self.find_results_table.grid(row=2, column=0, sticky="nsew")
        find_scroll.grid(row=2, column=1, sticky="ns")
        self.find_results_table.bind("<<TreeviewSelect>>", self._on_find_result_select)

        ttk.Label(top, text=self.tr("preview"), style="Card.TLabel").grid(row=3, column=0, sticky="w", pady=(8, 4))
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
        ttk.Label(side_header, text=self.tr("indexed_documents"), style="Card.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Button(side_header, text=self.tr("clear_filter"), command=self.clear_find_document_filter).grid(
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
        self.find_docs_list.bind("<Button-3>", self.on_find_docs_right_click)
        self.find_docs_context_menu = tk.Menu(self.find_window, tearoff=0)
        self.find_docs_context_menu.add_command(
            label=self.tr("delete_corpus_doc"),
            command=self.delete_selected_corpus_document,
        )

        self._set_find_query_from_selection()
        self.refresh_find_corpus_summary()

        def _on_close():
            self.find_window.destroy()
            self.find_window = None
            self.find_results_table = None
            self.find_preview_text = None
            self.find_docs_list = None
            self.find_docs_context_menu = None
            self.find_doc_items = []
            self.find_result_items = {}

        self.find_window.protocol("WM_DELETE_WINDOW", _on_close)

    def _set_find_query_from_selection(self):
        selected_idx = self._get_context_or_selected_index()
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

        ttk.Label(wrap, text=self.tr("passage_title"), style="Card.TLabel").pack(anchor="w")
        ttk.Label(
            wrap,
            text=self.tr("passage_desc"),
            style="Card.TLabel",
            foreground="#666",
        ).pack(anchor="w", pady=(0, 8))

        ctrl = ttk.Frame(wrap, style="Card.TFrame")
        ctrl.pack(fill="x", pady=(0, 8))

        btn_generate = ttk.Button(ctrl, text=self.tr("generate"), command=self.generate_ielts_passage)
        btn_generate.pack(side=tk.LEFT, padx=(0, 6))
        Tooltip(btn_generate, "Build passage from selected words, or the full list if nothing is selected")

        btn_play = ttk.Button(ctrl, text=self.tr("read_with_gemini"), command=self.play_generated_passage)
        btn_play.pack(side=tk.LEFT, padx=(0, 6))
        Tooltip(btn_play, "Speak current passage")

        btn_stop = ttk.Button(ctrl, text=self.tr("stop"), command=self.stop_passage_playback)
        btn_stop.pack(side=tk.LEFT, padx=(0, 6))
        Tooltip(btn_stop, "Stop speaking")

        btn_practice = ttk.Button(ctrl, text=self.tr("practice"), command=self.start_passage_practice)
        btn_practice.pack(side=tk.LEFT, padx=(0, 6))
        Tooltip(btn_practice, "Hide keywords as blanks")

        btn_check_practice = ttk.Button(ctrl, text=self.tr("check"), command=self.check_passage_practice)
        btn_check_practice.pack(side=tk.LEFT)
        Tooltip(btn_check_practice, "Check your filled answers")

        model_wrap = ttk.Frame(ctrl, style="Card.TFrame")
        model_wrap.pack(side=tk.RIGHT)
        ttk.Label(model_wrap, text=self.tr("model"), style="Card.TLabel").pack(side=tk.LEFT, padx=(0, 4))
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
            text=self.tr("practice_tip"),
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
                elif event_type == "success_tts":
                    self._finish_tts_validation_success(payload or {})
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

    def open_gemini_key_window(self, force_verify=False):
        self.gemini_verified = self.gemini_verified and not force_verify
        if self.gemini_key_window and self.gemini_key_window.winfo_exists():
            self.gemini_key_window.lift()
            self.gemini_key_window.focus_force()
            return

        self.gemini_key_status_var.set("Paste your LLM API key, then test it.")
        self.gemini_key_var.set(get_llm_api_key())
        self.llm_api_provider_var.set(self.tr("provider_gemini"))
        win = tk.Toplevel(self)
        self.gemini_key_window = win
        win.title(self.tr("llm_api"))
        win.configure(bg="#f6f7fb")
        win.resizable(False, False)
        win.transient(self.winfo_toplevel())

        wrap = ttk.Frame(win, style="Card.TFrame")
        wrap.pack(fill="both", expand=True, padx=12, pady=12)

        ttk.Label(wrap, text=self.tr("llm_api_setup"), style="Card.TLabel").pack(anchor="w")
        ttk.Label(
            wrap,
            text=self.tr("llm_key_desc"),
            style="Card.TLabel",
            foreground="#666",
        ).pack(anchor="w", pady=(0, 8))

        provider_row = ttk.Frame(wrap, style="Card.TFrame")
        provider_row.pack(anchor="w", pady=(0, 8), fill="x")
        ttk.Label(provider_row, text=f"{self.tr('api_provider')}:", style="Card.TLabel").pack(side=tk.LEFT)
        provider_combo = ttk.Combobox(
            provider_row,
            textvariable=self.llm_api_provider_var,
            values=[self.tr("provider_gemini")],
            state="readonly",
            width=18,
        )
        provider_combo.pack(side=tk.LEFT, padx=(6, 0))
        provider_combo.bind("<<ComboboxSelected>>", lambda _e: set_llm_api_provider("gemini"))

        entry = ttk.Entry(wrap, textvariable=self.gemini_key_var, width=54, show="*")
        entry.pack(fill="x")
        entry.focus_set()
        entry.icursor(tk.END)
        entry.bind("<Return>", lambda _event: self.test_and_save_gemini_key())

        ttk.Label(
            wrap,
            text=self.tr("gemini_model_desc"),
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
        self.gemini_key_test_btn = ttk.Button(btn_row, text=self.tr("test_and_save"), command=self.test_and_save_gemini_key)
        self.gemini_key_test_btn.pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text=self.tr("exit"), command=self._cancel_gemini_key_setup).pack(side=tk.LEFT)

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
        if self.gemini_key_window and self.gemini_key_window.winfo_exists():
            try:
                self.gemini_key_window.grab_release()
            except Exception:
                pass
            self.gemini_key_window.destroy()
        self.gemini_key_window = None
        self.gemini_key_test_btn = None
        self.status_var.set("LLM API ready.")

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
        if self.tts_key_window and self.tts_key_window.winfo_exists():
            self.tts_key_window.lift()
            self.tts_key_window.focus_force()
            return

        self.tts_key_status_var.set("Paste your TTS API key, then test it.")
        self.tts_key_var.set(get_tts_api_key())
        self.tts_api_provider_var.set(self._tts_provider_label(get_tts_api_provider()))
        win = tk.Toplevel(self)
        self.tts_key_window = win
        win.title(self.tr("tts_api"))
        win.configure(bg="#f6f7fb")
        win.resizable(False, False)
        win.transient(self.winfo_toplevel())

        wrap = ttk.Frame(win, style="Card.TFrame")
        wrap.pack(fill="both", expand=True, padx=12, pady=12)

        ttk.Label(wrap, text=self.tr("tts_api_setup"), style="Card.TLabel").pack(anchor="w")
        ttk.Label(
            wrap,
            text=self.tr("tts_key_desc"),
            style="Card.TLabel",
            foreground="#666",
        ).pack(anchor="w", pady=(0, 8))

        provider_row = ttk.Frame(wrap, style="Card.TFrame")
        provider_row.pack(anchor="w", pady=(0, 8), fill="x")
        ttk.Label(provider_row, text=f"{self.tr('api_provider')}:", style="Card.TLabel").pack(side=tk.LEFT)
        provider_combo = ttk.Combobox(
            provider_row,
            textvariable=self.tts_api_provider_var,
            values=list(self._tts_provider_options().keys()),
            state="readonly",
            width=18,
        )
        provider_combo.pack(side=tk.LEFT, padx=(6, 0))
        provider_combo.bind("<<ComboboxSelected>>", self._on_tts_provider_selected)

        entry = ttk.Entry(wrap, textvariable=self.tts_key_var, width=54, show="*")
        entry.pack(fill="x")
        entry.focus_set()
        entry.icursor(tk.END)
        entry.bind("<Return>", lambda _event: self.test_and_save_tts_key())

        ttk.Label(wrap, textvariable=self.tts_key_status_var, style="Card.TLabel", foreground="#444").pack(
            anchor="w", pady=(10, 0)
        )

        btn_row = ttk.Frame(wrap, style="Card.TFrame")
        btn_row.pack(fill="x", pady=(10, 0))
        self.tts_key_test_btn = ttk.Button(btn_row, text=self.tr("test_and_save"), command=self.test_and_save_tts_key)
        self.tts_key_test_btn.pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text=self.tr("close"), command=self._close_tts_key_window).pack(side=tk.LEFT)

        win.grab_set()
        win.protocol("WM_DELETE_WINDOW", self._close_tts_key_window)

    def _close_tts_key_window(self):
        if self.tts_key_window and self.tts_key_window.winfo_exists():
            try:
                self.tts_key_window.grab_release()
            except Exception:
                pass
            self.tts_key_window.destroy()
        self.tts_key_window = None
        self.tts_key_test_btn = None

    def _on_tts_provider_selected(self, _event=None):
        provider = self._tts_provider_value()
        set_tts_api_provider(provider)
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
        self._close_tts_key_window()
        self.tts_api_provider_var.set(self._tts_provider_label(provider))
        self.refresh_voice_list()
        self.status_var.set("TTS API ready.")

    def _finish_tts_validation_error(self, message):
        self.tts_key_status_var.set("TTS API key test failed. Please paste another key.")
        if self.tts_key_test_btn:
            self.tts_key_test_btn.config(state="normal")
        messagebox.showerror(self.tr("tts_api_key_error"), str(message or "Unknown error"))

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
                api_key=get_llm_api_key(),
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
        runtime = tts_get_runtime_label()
        token = speak_stream_async(
            speech_text,
            self.volume_var.get() / 100.0,
            rate_ratio=self.speech_rate_var.get(),
            cancel_before=False,
            chunk_chars=90,
        )
        self.passage_status_var.set(f"Generating passage audio with {runtime}...")
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
        status = tts_get_gemini_queue_status()
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
        self.gemini_runtime_status_var.set(f"{status_text} | {self.trf('tts_status_queue', count=queue_count)}")

        next_retry_at = float(status.get("next_retry_at") or 0.0)
        if next_retry_at > 0:
            retry_text = time.strftime("%H:%M:%S", time.localtime(next_retry_at))
            self.gemini_retry_status_var.set(self.trf("tts_status_retry_at", time=retry_text))
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
            self.order_tip.text = "In order (from current)" if self.ui_language_var.get() == "en" else "顺序播放（从当前词开始）"
        if self.order_tip_rand:
            self.order_tip_rand.text = "Random (no repeat)" if self.ui_language_var.get() == "en" else "随机不重复"
        if self.order_tip_click:
            self.order_tip_click.text = "Click one word to play one" if self.ui_language_var.get() == "en" else "点一个词读一个词"
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
            "order": True,
            "speed": False,
            "volume": False,
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
            (self.tr("settings_toggle_order"), "order"),
            (self.tr("settings_toggle_speed"), "speed"),
            (self.tr("settings_toggle_volume"), "volume"),
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

        # Order section
        order_section = ttk.Frame(right_panel, style="Card.TFrame")
        ttk.Label(order_section, text=self.tr("order"), style="Card.TLabel").pack(anchor="w")
        ttk.Label(
            order_section,
            text=self.tr("order_desc"),
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
            text=self.tr("stop_after_list"),
            variable=self.stop_at_end_var,
            command=self.on_stop_at_end_toggle,
        )
        self.stop_at_end_check.pack(anchor="w", pady=(6, 0))
        order_sep = ttk.Separator(right_panel, orient="horizontal")
        sections.append({"key": "order", "frame": order_section, "sep": order_sep})

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

        # Volume section
        volume_section = ttk.Frame(right_panel, style="Card.TFrame")
        ttk.Label(volume_section, text=self.tr("volume"), style="Card.TLabel").pack(anchor="w")
        ttk.Label(
            volume_section,
            text=self.tr("volume_desc"),
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

        self.update_order_button()
        self.update_loop_button()
        self.update_speed_buttons()
        self.update_speech_rate_buttons()
        self.on_volume_change()
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
            self.show_info("import_words_first")
            return
        if self.order_mode.get() == "click_to_play":
            if self._get_selected_index() is None:
                self.show_info("click_word_first_mode")
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
        self.status_var.set("Not started")
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
        self._clear_word_action_context()
        # Single-click only updates selection state.
        if not self.store.words:
            return
        if self._get_selected_index() is None:
            self._refresh_selection_details()
            return
        self._refresh_selection_details()

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
        if self._get_selected_index() is not None:
            self.speak_selected()
        return "break"

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
        if not tree or not self.word_context_menu:
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
        if selected_idx is None:
            return "break"
        self._sync_main_selection_to_index(selected_idx)
        self.word_action_origin = "dictation"
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

    def _clear_synonym_event_queue(self):
        try:
            while True:
                self.synonym_event_queue.get_nowait()
        except queue.Empty:
            return

    def _emit_synonym_event(self, event_type, token, payload=None):
        try:
            self.synonym_event_queue.put_nowait((event_type, token, payload))
        except Exception:
            return

    def _poll_synonym_events(self, token):
        if token != self.synonym_lookup_active_token:
            return
        done = False
        try:
            while True:
                event_type, event_token, payload = self.synonym_event_queue.get_nowait()
                if event_token != token:
                    continue
                if event_type == "result":
                    data = payload or {}
                    self._show_synonym_window(
                        data.get("word", ""),
                        data.get("focus", ""),
                        data.get("synonyms") or [],
                    )
                elif event_type == "error":
                    messagebox.showerror(self.tr("synonyms_error"), str(payload or "Unknown error"))
                elif event_type == "done":
                    done = True
        except queue.Empty:
            pass

        if not done and token == self.synonym_lookup_active_token:
            self.after(80, lambda t=token: self._poll_synonym_events(t))

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
        selected_idx = self._get_context_or_selected_index()
        if selected_idx is None or selected_idx >= len(self.store.words):
            self.show_info("select_word_first")
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
                    api_key=get_llm_api_key(),
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

    def lookup_synonyms_for_selected_word(self):
        selected_idx = self._get_context_or_selected_index()
        if selected_idx is None or selected_idx >= len(self.store.words):
            self.show_info("select_word_first")
            return
        word = self.store.words[selected_idx]
        self.status_var.set(f"Looking up local synonyms for '{word}'...")
        self.synonym_lookup_token += 1
        token = self.synonym_lookup_token
        self.synonym_lookup_active_token = token
        self._clear_synonym_event_queue()

        import threading

        def _run():
            try:
                result = get_local_synonyms(word, limit=12)
                self._emit_synonym_event("result", token, result)
            except Exception as exc:
                self._emit_synonym_event("error", token, str(exc))
            self._emit_synonym_event("done", token, None)

        threading.Thread(target=_run, daemon=True).start()
        self.after(80, lambda t=token: self._poll_synonym_events(t))

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
            text=self.tr("read_sentence"),
            command=lambda s=sentence: speak_async(
                s,
                self.volume_var.get() / 100.0,
                rate_ratio=self.speech_rate_var.get(),
                cancel_before=True,
                source_path=self.store.get_current_source_path(),
            ),
        ).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(row, text=self.tr("close"), command=self.sentence_window.destroy).pack(side=tk.LEFT)

    def _show_synonym_window(self, word, focus, synonyms):
        self.status_var.set(self.trf("synonyms_ready", word=word))
        if self.synonym_window and self.synonym_window.winfo_exists():
            self.synonym_window.destroy()

        self.synonym_window = tk.Toplevel(self)
        self.synonym_window.title(f"{self.tr('synonyms_title')} - {word}")
        self.synonym_window.configure(bg="#f6f7fb")
        self.synonym_window.resizable(False, False)

        wrap = ttk.Frame(self.synonym_window, style="Card.TFrame")
        wrap.pack(fill="both", expand=True, padx=10, pady=10)
        ttk.Label(wrap, text=f"Word: {word}", style="Card.TLabel").pack(anchor="w")
        if focus:
            ttk.Label(
                wrap,
                text=self.trf("synonyms_focus", word=focus),
                style="Card.TLabel",
                foreground="#666",
            ).pack(anchor="w", pady=(0, 4))
        ttk.Label(
            wrap,
            text=self.tr("synonyms_source"),
            style="Card.TLabel",
            foreground="#666",
        ).pack(anchor="w", pady=(0, 8))

        text = tk.Text(
            wrap,
            wrap="word",
            height=8,
            width=56,
            bg="#ffffff",
            fg="#222222",
            highlightthickness=1,
            highlightbackground="#d9dbe1",
        )
        text.pack(fill="both", expand=True)
        if synonyms:
            text.insert("1.0", "\n".join(f"- {item}" for item in synonyms))
        else:
            text.insert("1.0", self.tr("no_synonyms_found"))
        text.config(state="disabled")

        row = ttk.Frame(wrap, style="Card.TFrame")
        row.pack(fill="x", pady=(8, 0))
        ttk.Button(row, text=self.tr("close"), command=self.synonym_window.destroy).pack(side=tk.LEFT)

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

