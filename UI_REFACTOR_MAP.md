# UI Refactor Map

This document records the current Tk UI split so future refactors do not have to rediscover the boundaries.

## Current Shape

`ui/main_view.py` is no longer the only place that holds UI behavior.

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

Still in `main_view.py`:

- selection and context menu routing
- translation / POS async refresh wiring
- treeview edit lifecycle

### Dictation / Recent Wrong

Owned by:

- `ui/dictation_controller.py`
- `ui/recent_wrong_controller.py`
- `ui/dictation_panel.py`
- `ui/list_presenter.py`

Still in `main_view.py`:

- popup sequencing
- timers
- playback coordination
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
- `ui/async_event_helper.py`

Still in `main_view.py`:

- file dialog interaction
- token lifecycle
- final UI updates

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

### 1. Word Metadata Refresh Coordinator

Candidate responsibility:

- translation token lifecycle
- POS analysis token lifecycle
- table row refresh after async completion

Why:

- this is still one of the densest cross-cutting parts of `main_view.py`
- it touches both data refresh and row rendering

### 2. Dictation Session Coordinator

Candidate responsibility:

- play / pause / replay / timer flow
- current item advance
- input color and feedback reset

Why:

- dictation business state is already mostly out
- what remains is a cohesive interaction state machine

### 3. Playback Coordinator

Candidate responsibility:

- main list playback state
- queue rebuild
- scheduling next word
- passage playback pause handoff

Why:

- it still couples TTS runtime, UI status, and queue state

## Safe Working Rule

When continuing this refactor, prefer extracting a full responsibility slice:

- state calculation
- async execution
- or view construction

Do not extract only names without moving ownership.
