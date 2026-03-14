# Word Speaker

A Windows desktop app for vocabulary study, pronunciation practice, dictation, IELTS-style content generation, and local corpus sentence search.

## Features

- Import `.txt` / `.csv` word lists
- Type or paste words directly into the app
- Paste two-column tables from Google Docs / Google Sheets into the manual import window
- Main word list uses `# / Word / Notes`
- The `Word` column shows two lines: `English`, then `part of speech + Chinese translation`
- Edit `Word` and `Notes` directly in the list and save changes back to the source file
- Manual pasted lists support `Save As`
- Unsaved manual lists trigger a save prompt before closing
- `New List` creates a blank list for building a new vocabulary file
- Play in order, random (no repeat), or click-to-play
- Double-click a word to play pronunciation
- Right-click a word for edit, corpus search, sentence generation, and cached-audio inspection
- Dictation opens in a dedicated window with `全部 / 近期错词`
- Dictation supports `从某词开始`, `开始学习`, and manual wrong-word addition
- The first dictation study mode is `在线拼写`, with countdown timing, playback speed control, replay, previous-word, play/pause, and live right/wrong feedback
- Wrong answers are stored locally, sorted by mistake count, and shown in the recent wrong-word list with error causes
- History items can be removed inside the app, and matching audio cache entries are cleaned up at the same time
- User-selectable TTS source: Gemini TTS, local Kokoro, or local Piper
- Automatic fallback to Kokoro when Gemini playback fails and Kokoro is available locally
- Built-in English -> Chinese translation with Argos Translate
- Built-in part-of-speech tagging with spaCy, cached locally for repeated words
- Part of speech and Chinese translation can be edited by the user and stored locally
- Translation cache is stored locally, so repeated words do not need to be translated again
- Single-word audio is cached under `data/audio_cache/words/`
- Cached word audio now stores backend metadata so the app can distinguish `gemini`, `kokoro`, and `piper`
- If Gemini is rate-limited during pre-cache, the app can keep local audio first and later replace it in the background with Gemini audio when quota allows
- Generate IELTS-listening-style passages from imported words and read them with the selected TTS source
- Generate IELTS-style example sentences for selected words
- Practice mode for generated passages
- `Find` can import `.txt` / `.docx` / `.pdf`, build a local sentence index, and search by word or phrase

## Run

Requires **64-bit Python**.

```bat
cd /d path\to\word_speaker
pip install -r requirements.txt
python -m spacy download en_core_web_sm
python app.py
```

If your system uses `py`:

```bat
cd /d path\to\word_speaker
py -3 -m pip install -r requirements.txt
py -3 -m spacy download en_core_web_sm
py -3 app.py
```

When the app opens, paste your Gemini API key into the popup window and click `Test and Save`.

## Dependencies

- `spaCy` is used for English part-of-speech tagging and basic lexical analysis
- `en_core_web_sm` is the expected spaCy model for the current workflow
- `Argos Translate` is used for local English -> Chinese translation
- `Gemini API` is used for passage generation, example sentence generation, and Gemini TTS
- `Kokoro` and `Piper` are optional local TTS backends

## TTS Behavior

- Pick the active source in `Settings > Source`
- `Gemini TTS` is the default online source and requires a valid Gemini API key plus network access
- `Kokoro` is an optional offline source and only appears when local model files exist in `data/models/kokoro/`
- `Piper` is a project-local local source using models under `data/models/piper/`
- If the selected source is Gemini and Gemini playback fails, the app automatically falls back to Kokoro when Kokoro is available
- Single-word audio is cached under `data/audio_cache/words/`
- Each cached word can carry metadata that marks its real backend source and the desired backend target
- Pending Gemini replacements are persisted in a queue, so local fallback audio can still be replaced after restarting the app
- Passage playback keeps one source for the whole article and does not mix backends inside the same generated passage

## Local Runtime Files

- `data/models/kokoro/`: optional offline Kokoro model files
- `data/models/piper/`: Piper `*.onnx` voice models and matching `*.onnx.json` config files
- `data/audio_cache/`: generated local audio cache
- `data/audio_cache/pending_gemini_replacements.json`: persisted queue of local cache files waiting to be replaced by Gemini
- `data/pos_cache.json`: cached part-of-speech labels
- `data/translation_cache.json`: cached translations
- `data/dictation_stats.json`: wrong-word and dictation statistics
- `vendor/site-packages/`: optional project-local Python runtime dependencies

## Input Format

- `.txt`: one word per line, or `word<TAB>note`
- `.csv`: first column = English, second column = Notes
- Manual input window:
  - one word per line, or
  - paste a two-column table from Google Docs / Sheets

## Main UI

- Left side: import, manual paste/type, save-as, new list, study list, playback, settings, and dictation entry
- Right side: current word details plus tabs for `Review / History / Tools`
- Current word details show the selected word, part of speech, translation, notes, and quick actions
- The word list is displayed in a book-style layout with numbering, the English word on the first line, and `part of speech + Chinese translation` on the second line

## Dictation

- Main dictation page has two list modes: `全部` and `近期错词`
- Two entry buttons are provided: `从某词开始` and `开始学习`
- `从某词开始` opens an in-window picker so you can jump into dictation from any word in the current list mode
- `开始学习` opens the study-mode popup; the first implemented mode is `在线拼写`
- `在线拼写` supports playback speed presets, countdown timing, replay, pause, previous-word, and live red/green answer feedback
- Wrong answers are recorded locally and feed back into the `近期错词` list
- `近期错词` shows error cause instead of normal notes, sorts by mistake count, and supports manual additions

## Find Corpus

- Open the `Find` window from the main UI
- Import `.txt` / `.docx` / `.pdf`
- The app builds a local sentence index in `data/corpus_index.db`
- Search by word or phrase
- Filter by the selected document, or search across the full corpus

## Notes

- PDF text extraction is rule-based and works best on text PDFs
- Scanned PDFs are not OCR-enabled yet
- Existing local corpus and runtime cache files are intentionally ignored by Git

## Credits

- Speech synthesis is powered by Gemini TTS and optional local Kokoro / Piper playback
- Part-of-speech tagging is powered by [spaCy](https://spacy.io/)
- English -> Chinese translation is powered by [argosopentech/argos-translate](https://github.com/argosopentech/argos-translate)
