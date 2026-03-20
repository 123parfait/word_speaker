# UI Refactor Map

This document records the current Tk UI split so future refactors do not have to rediscover the boundaries.

## Current Shape

`ui/main_view.py` is no longer the only place that holds UI behavior.

Service-side splits now also include:

- `services/corpus_ingest.py`
- `services/corpus_index_store.py`
- `services/runtime_log.py`
- `services/tts_audio.py`
- `services/tts_queue.py`
- `services/tts_backend_strategy.py`
- `services/tts_persistence.py`
- `services/tts_shared_cache.py`
- `services/tts_synth_cache.py`
- `services/tts_synth_execute.py`

The current split is:

- Controllers
  - `ui/word_list_controller.py`
  - `ui/recent_wrong_controller.py`
  - `ui/dictation_controller.py`
  - `ui/find_controller.py`
- Presenters
  - `ui/detail_presenter.py`
  - `ui/history_presenter.py`
  - `ui/list_presenter.py`
  - `ui/manual_words_presenter.py`
  - `ui/find_presenter.py`
  - `ui/passage_presenter.py`
  - `ui/word_tools_presenter.py`
- Panels / builders
  - `ui/word_list_panel.py`
  - `ui/dictation_panel.py`
  - `ui/sidebar_panels.py`
  - `ui/detail_sidebar.py`
  - `ui/manual_words_panel.py`
  - `ui/find_panel.py`
  - `ui/passage_panel.py`
  - `ui/word_tools_panel.py`
- Async helpers
  - `ui/async_event_helper.py`
  - `ui/find_async.py`
  - `ui/word_tools_async.py`
- Coordinators
  - `ui/main_playback_controller.py`
  - `ui/dictation_session_coordinator.py`
  - `ui/dictation_window_coordinator.py`
  - `ui/find_window_coordinator.py`
  - `ui/word_metadata_coordinator.py`
  - `ui/word_action_coordinator.py`
  - `ui/settings_host_coordinator.py`
  - `ui/main_playback_host.py`
  - `ui/tool_host_coordinator.py`
  - `ui/tts_status_bridge.py`
- Editing helpers
  - `ui/manual_words_editor.py`

`ui/main_view.py` is now mainly a coordinator:

- keeps Tk variables and widget references
- forwards user actions to controllers/helpers
- applies presenter state to widgets
- bridges async events back into the UI thread

## Domain Boundaries

### Word List

Owned by:

- `ui/word_list_controller.py`
- `ui/word_list_panel.py`
- `ui/list_presenter.py`
- `ui/detail_presenter.py`
- `ui/word_metadata_coordinator.py`

Still in `main_view.py`:

- selection routing
- treeview edit lifecycle

### Dictation / Recent Wrong

Owned by:

- `ui/dictation_controller.py`
- `ui/recent_wrong_controller.py`
- `ui/dictation_panel.py`
- `ui/dictation_session_coordinator.py`
- `ui/dictation_window_coordinator.py`
- `ui/list_presenter.py`

Still in `main_view.py`:

- answer review popup lifecycle
- volume popup lifecycle
- a few dictation-only widget state toggles

### Manual Import

Owned by:

- `ui/manual_words_presenter.py`
- `ui/manual_words_panel.py`
- `ui/manual_words_editor.py`
- `ui/word_list_controller.py`

Still in `main_view.py`:

- clipboard access
- opening and closing the tool
- applying imported rows into the current session

### Corpus Find

Owned by:

- `ui/find_controller.py`
- `ui/find_presenter.py`
- `ui/find_panel.py`
- `ui/find_async.py`
- `ui/find_window_coordinator.py`
- `ui/async_event_helper.py`

Still in `main_view.py`:

- host widget references
- method forwarding

### Passage / Sentence / Synonyms

Owned by:

- `ui/passage_presenter.py`
- `ui/passage_panel.py`
- `ui/word_tools_presenter.py`
- `ui/word_tools_panel.py`
- `ui/word_tools_async.py`
- `ui/async_event_helper.py`

Still in `main_view.py`:

- token lifecycle
- TTS handoff
- popup orchestration

## What Not To Split Next

Avoid spending time on micro-extractions that only move one or two lines.

Low-value next steps:

- moving individual `StringVar.set(...)` calls into more helpers
- splitting tiny one-off wrapper methods just for the sake of file count
- reworking stable panel builders that already have clear ownership

## Best Next Cuts

### 1. Word List Selection Host Slice

Candidate responsibility:

- selection change handling
- double-click play behavior
- context reset and detail refresh ordering

Why:

- these actions still sit close to the root host and drive several other domains

### 2. Word List Selection Host Slice

Candidate responsibility:

- selection change handling
- double-click play behavior
- context reset and detail refresh ordering

Why:

- these actions still sit close to the root host and drive several other domains

### 3. Store / Learning Data Split

Candidate responsibility:

- word list content
- learning stats
- recent wrong state
- history/session persistence

Why:

- `data/store.py` is now one of the clearest remaining mixed-responsibility modules

## Safe Working Rule

When continuing this refactor, prefer extracting a full responsibility slice:

- state calculation
- async execution
- or view construction

Do not extract only names without moving ownership.
