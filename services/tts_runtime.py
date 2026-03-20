from dataclasses import dataclass, field
import threading


@dataclass(frozen=True)
class TtsRuntimeConfig:
    tts_model: str = "gemini-2.5-flash-preview-tts"
    tts_sample_rate: int = 24000
    tts_channels: int = 1
    tts_sample_width: int = 2
    tts_voice_name: str = "Kore"
    tts_style_short: str = "Read the word clearly in a neutral British English accent for IELTS vocabulary practice."
    tts_style_long: str = "Read this passage clearly in a natural British English accent for IELTS listening practice."
    elevenlabs_voice_id: str = "JBFqnCBsd6RMkjVDRZzb"
    elevenlabs_model_id: str = "eleven_multilingual_v2"
    online_tts_request_timeout_seconds: int = 180
    interactive_fast_timeout_short_seconds: int = 12
    interactive_fast_timeout_long_seconds: int = 20
    gemini_queue_request_interval_seconds: int = 40
    gemini_rate_limit_cooldown_seconds: int = 120
    gemini_manual_request_cooldown_seconds: int = 60
    elevenlabs_queue_request_interval_seconds: float = 1.5
    elevenlabs_rate_limit_cooldown_seconds: int = 45
    elevenlabs_manual_request_cooldown_seconds: int = 3
    kokoro_sample_rate: int = 24000
    shared_cache_package_version: int = 2
    shared_cache_package_manifest: str = "manifest.json"
    shared_cache_metadata_file: str = "global/metadata.json"


@dataclass
class TtsRuntimeState:
    lock: threading.Lock = field(default_factory=threading.Lock)
    token: int = 0
    shown_errors: set = field(default_factory=set)
    current_wav: str | None = None
    kokoro: object | None = None
    kokoro_lock: threading.Lock = field(default_factory=threading.Lock)
    piper_voices: dict = field(default_factory=dict)
    piper_lock: threading.Lock = field(default_factory=threading.Lock)
    backend_lock: threading.Lock = field(default_factory=threading.Lock)
    backend_status: dict = field(default_factory=dict)
    pending_gemini_replacements: dict = field(default_factory=dict)
    pending_gemini_lock: threading.Lock = field(default_factory=threading.Lock)
    pending_gemini_worker_running: bool = False
    preferred_pending_source: str | None = None
    manual_session_cache_paths: set = field(default_factory=set)
    manual_session_cache_lock: threading.Lock = field(default_factory=threading.Lock)
    word_audio_override_lock: threading.Lock = field(default_factory=threading.Lock)
    word_audio_override_memory: dict | None = None
    online_queue_manager: object | None = None
    cache_metadata_store: object | None = None
    error_notifier: object | None = None
    runtime_initialized: bool = False
    runtime_init_lock: threading.Lock = field(default_factory=threading.Lock)
