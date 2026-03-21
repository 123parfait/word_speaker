# Word Speaker Guide

## Features

- Import `.txt` / `.csv` word lists
- Type or paste words directly into the app
- Paste two-column tables from Google Docs / Google Sheets into the manual import window
- Main word list uses `# / Word / Notes`
- The `Word` column shows two lines: `English`, then `part of speech + Chinese translation`
- Edit `Word` and `Notes` directly in the list and save changes back to the source file
- Manual pasted lists support `Save As`
- Unsaved manual lists now bind to a temporary backing file first, so translation / POS / phonetic / audio caches can still use a stable source path before the user chooses where to save
- Unsaved manual lists still trigger a save prompt before closing
- If the user discards an unsaved manual list, the temporary file and its related cache entries are cleaned up
- `New List` creates a blank list for building a new vocabulary file
- Main `Play` is now a single sequential loop mode
- If a word is selected, `Play` starts from that word
- If nothing is selected, `Play` starts from the first word
- Playback moves the visible blue selection highlight as it reads
- After the last word, playback loops back to the first word automatically
- `Space` toggles main playback pause / resume when the focus is on the main window
- Single-click and double-click both play pronunciation
- Right-click a word for edit, corpus search, sentence generation, synonyms, and cached-audio inspection
- Right-click also supports:
  - add word
  - replace this word's audio with Piper
  - restore the word back to the default audio backend
- Dictation opens in a dedicated window with `All / Recent Wrong`
- Dictation supports `Start From Word`, `Start Learning`, and answer review
- Dictation supports `Sequential` and `Random (No Repeat)` question order
- The session-end summary now uses the same answer-review layout as the in-session `Answer Review` popup, including accuracy, previous accuracy, buttons, and the comparison table
- Wrong answers are stored locally, sorted by mistake count, and shown in the recent wrong-word list with error causes
- In dictation, `Previous` now reopens the last word for editing and rolls back the previous attempt so the old answer does not count
- Wrong answers in dictation stay visible briefly before the next word instead of disappearing immediately
- Dictation includes a dedicated volume popup with boost up to `600%`
- The dictation volume popup auto-closes when the user returns focus to the typing box
- The session-end result screen plays a short 3-second visual effect based on accuracy:
  - below `50%`: rainbow
  - `50% - 80%`: blooming flower
  - above `80%`: fireworks and confetti
- The app starts maximized by default on Windows
- The project now includes `speaker.png` / `speaker.ico` for runtime and packaged app icons
- History items can be removed or renamed inside the app, with matching cache updates
- Indexed corpus documents can be removed from the app without deleting the original files on disk
- Separate `LLM API` and `TTS API` settings
- User-selectable online TTS provider: ElevenLabs or Gemini
- User-selectable playback source: online TTS, local Kokoro, or local Piper
- Export reusable shared word-audio cache packs as `.zip`
- Import shared word-audio cache packs from another device to reuse generated TTS
- Shared-cache packages now also carry global cached metadata:
  - translation
  - POS
  - phonetics
- Sync an official hosted shared-audio cache manifest and merge only missing/newer shared audio into local `global`
- Export/import clean word resource packs as `.wspack`
- Built-in English -> Chinese translation with Argos Translate
- Built-in part-of-speech tagging with spaCy, cached locally for repeated words
- Built-in UK phonetics with Gemini, cached locally for repeated words
- Part of speech and Chinese translation can be edited by the user and stored locally
- Manual corrections are stored in a separate user dictionary so they can override bad cached results
- Synonym lookup prefers Gemini and falls back to local spaCy + WordNet
- Generate IELTS-listening-style passages from imported words and read them with the selected TTS source
- Generate IELTS-style example sentences for selected words
- Practice mode for generated passages
- `Find` can import `.txt` / `.docx` / `.pdf`, build a local sentence index, and search by word or phrase
- `Tools > Update App` supports online update checks and local update-package import
## API Setup

- `LLM API` is currently Gemini-only
- `TTS API` currently supports `ElevenLabs` and `Gemini`
- Startup uses one combined API setup window for both `LLM API` and `TTS API`
- The unified `Test and Save` button validates both inputs together
- Startup highlights invalid API inputs inline instead of only relying on blocking error popups

## TTS Behavior

- Configure providers in `Settings > LLM API` and `Settings > TTS API`
- Pick the active playback source in `Settings > Source`
- If `TTS API` is set to `ElevenLabs`, online word-audio generation prefers `ElevenLabs`, then `Gemini`, then local fallback
- If `TTS API` is set to `Gemini`, online word-audio generation prefers `Gemini`, then local fallback
- ElevenLabs uses a default British-style voice for IELTS-style playback
- `Kokoro` is an optional offline source and only appears when local model files exist in `data/models/kokoro/`
- `Piper` is a project-local local source using models under `data/models/piper/`
- If the selected online source fails at playback time, the app automatically falls back to a local backend when available
- Passage playback keeps one source for the whole article and does not mix backends inside the same generated passage
- IELTS-oriented TTS text normalization runs before synthesis and expands:
  - numbers
  - years
  - units
  - currencies
  - percentages
- Dynamic queue throttling is provider-specific instead of using one fixed delay for every online source

## Cache Behavior

- Single-word audio is cached under `data/audio_cache/`
- Source-specific caches are stored under `data/audio_cache/sources/`
- Source-specific caches are grouped by source file and then by `a-z` or `other`
- Shared online caches are stored under `data/audio_cache/global/`
- `global` stores the real entity audio files
- `sources` stores lightweight source aliases that point at `global`
- Current file cache, recent-wrong cache, and other source caches all reuse `global`
- The app can export `global` as a shareable cache package and import the same package format later
- `Tools > Update Word Library` downloads the hosted official library bundle and applies:
  - shared cache audio
  - shared metadata
  - official `.wspack`
  - bundled corpus package
- Shared-cache packages now also include:
  - `translation_cache`
  - `pos_cache`
  - `phonetics_cache`
- Shared-cache packages still do not overwrite history, user config, or corpus data
- Each cached word carries metadata for:
  - real backend source
  - desired backend target
  - linked shared/global cache path
- The online replacement queue is persisted in `data/audio_cache/pending_online_tts_replacements.json`
- Queue throttling is conservative by provider:
  - ElevenLabs: dynamic throttling around a `1.5s` base interval, `45s` cooldown after rate-limit errors
  - Gemini TTS: dynamic throttling with a slower base interval and longer cooldown because free-tier limits are stricter
- Online replacement priority is:
  - `ElevenLabs`
  - `Gemini`
  - local placeholder (`Piper`, then `Kokoro`)

## Word Resource Packs

- Word resource packs use the `.wspack` extension
- A resource pack contains only curated vocabulary data:
  - word
  - note
  - manual Chinese translation override
  - manual part-of-speech override
  - phonetic
- Resource packs do not include:
  - audio cache
  - API keys
  - dictation history
  - corpus index
  - the rest of `data/`
- Importing a resource pack replaces the current word list in the app
- Importing a resource pack also writes included translation / POS overrides into the local user dictionary
- Resource packs are the recommended way to share curated vocabulary content with another user

## Dictation

- Main dictation page has two list modes: `All` and `Recent Wrong`
- Two entry buttons are provided: `Start From Word` and `Start Learning`
- `Start From Word` opens an in-window picker so you can jump into dictation from any word in the current list mode
- `Start Learning` opens the study-mode popup; the first implemented mode is `Online Spelling`
- `Online Spelling` supports playback speed presets, countdown timing, replay, pause, previous-word, and live red/green answer feedback
- `Online Spelling` previous-word rollback removes the previous session attempt before you re-answer
- Wrong answers are recorded locally and feed back into the `Recent Wrong` list
- `Recent Wrong` is global, not limited to the currently opened file
- `Recent Wrong` shows error cause instead of normal notes and sorts by mistake count
- Correct answers in recent-wrong study do not auto-remove the word from the recent-wrong list; users clear the list manually
- The answer-review popup shows:
  - accuracy so far
  - last session accuracy
  - session answers
  - wrong-only filtering
- Session-end review rows are ordered with wrong answers first, then correct answers

## Find Corpus

- Open the `Find` window from the main UI
- Import `.txt` / `.docx` / `.pdf`
- The app builds a local sentence index in `data/corpus_index.db`
- During import, the `Import Docs` button is disabled to avoid duplicate clicks
- Import status now shows the first file name so users can tell the task is running
- Import completion now shows an explicit result dialog instead of only refreshing the list silently
- Search by word or phrase
- Filter by the selected document, or search across the full corpus

## Local Runtime Files

- `data/models/piper/`: Piper `*.onnx` voice models and matching `*.onnx.json` config files
- `data/audio_cache/`: generated local audio cache
- `data/audio_cache/sources/`: source-specific word caches grouped by file and first letter
- `data/audio_cache/global/`: shared online-TTS cache reused across files when available
- `data/audio_cache/pending_online_tts_replacements.json`: persisted online replacement queue
- Shared-cache export/import uses a `.zip` package with a manifest and copies only reusable `global` cache entries
- `version.json`: packaged app version used by the updater
- `data/app.instance.lock`: single-instance guard for the desktop app
- `data/pos_cache.json`: cached part-of-speech labels
- `data/translation_cache.json`: cached translations
- `data/phonetics_cache.json`: cached phonetics
- `data/user_dictionary.json`: manual translation / POS overrides that take priority over cached automatic results
- `data/synonyms_cache.json`: cached synonym results
- `data/dictation_stats.json`: wrong-word and dictation statistics
- `data/corpus_index.db`: local corpus sentence index
- `data/nltk_data/`: local WordNet data used by synonym lookup
- `vendor/site-packages/`: optional project-local Python runtime dependencies

## Input Format

- `.txt`: one word per line, or `word<TAB>note`
- `.csv`: first column = English, second column = Notes
- Manual input window:
  - one word per line, or
  - paste a two-column table from Google Docs / Sheets

## Notes

- PDF text extraction is rule-based and works best on text PDFs
- Scanned PDFs are not OCR-enabled yet
- Existing local corpus and runtime cache files are intentionally ignored by Git
- The packaged Windows build is intended to ship with local models and WordNet data so end users do not need extra downloads
- The packaged Windows app is generated at `dist/WordSpeaker/WordSpeaker.exe`; keep the whole `dist/WordSpeaker/` folder when sharing it
- Offline update packages should be zipped from the packaged app folder and must include `version.json`
- The built-in update-package tool expects the packaged app folder (for example `dist/WordSpeaker/`), not the source repository root
- Online update manifests are simple JSON files with `version` and `url`; the app downloads the referenced `.zip` and applies it after exit
- Official shared-audio sync uses the same simple hosted-manifest pattern: a JSON file with `version` and `url` that points at a shared-cache `.zip`
- In this project, the packaged app reads the default GitHub Release manifest URLs from `version.json`, so normal users do not need to type them manually
- The updater preserves local cache/config files by skipping known user-data files during replacement
- Protected local files include audio cache, config, translation cache, POS cache, user dictionary, corpus index, and dictation/history data
- GitHub Releases can be used as the host for the update `.zip`; the app only needs a reachable `manifest.json`
- Users need one first packaged build that already contains the updater; later versions can then be installed from inside the app
- Recommended release workflow:
  1. update `version.json`
  2. clean old artifacts from `dist/` first so old zips do not mix with the new release
  3. rebuild the packaged app folder, for example `dist/WordSpeaker/`
  4. create a full package zip, for example `WordSpeaker-x.x.x-portable.zip`, from the whole packaged folder
  5. create `WordSpeaker-update-x.x.x.zip` from the packaged app folder
  6. generate `manifest.json` for the update package
  7. upload all 3 files to the same GitHub Release:
     - `WordSpeaker-x.x.x-portable.zip`
     - `WordSpeaker-update-x.x.x.zip`
     - `manifest.json`
- Packaging defaults:
  - the packaged app includes built-in models, WordNet data, and `version.json`
- the packaged app does not include your local `audio_cache`, translation cache, POS cache, user dictionary, or dictation history by default
- the packaged app also does not include your local phonetics cache, bundled-corpus state, or per-word audio override state by default
- Release hygiene:
  - do not use a packaged folder that you have already been using for daily study as the source for a public release
  - if you run the packaged app before zipping it, that folder may accumulate local runtime files under `data/`
  - for clean releases, rebuild first and zip the fresh packaged folder before using it as a real app
  - before each release, remove old `WordSpeaker-*.zip`, `manifest.json`, and other stale assets from `dist/`
  - share reusable word audio with `Export Shared Cache` instead of bundling runtime cache folders into the release package
  - if you maintain an official shared-audio library, publish a hosted shared-cache `.zip` plus `shared_audio_manifest.json`
  - share curated vocabulary content with `.wspack` resource packs instead of shipping the whole `data/` folder
- Do not run the packaged app directly from inside a zip file
- Fully extract the packaged folder first, ideally with `7-Zip`, `Bandizip`, or `WinRAR`
- Extract to a short path such as `D:\WS` or `C:\WordSpeaker`
- If Windows Explorer reports path-length extraction errors, some DLLs may be skipped and the app can fail at startup
