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
- Right-click a word for edit, corpus search, sentence generation, synonyms, and cached-audio inspection
- Dictation opens in a dedicated window with `All / Recent Wrong`
- Dictation supports `Start From Word`, `Start Learning`, and manual wrong-word addition
- The first dictation study mode is `Online Spelling`, with countdown timing, playback speed control, replay, previous-word, play/pause, and live right/wrong feedback
- Wrong answers are stored locally, sorted by mistake count, and shown in the recent wrong-word list with error causes
- History items can be removed inside the app, and matching audio cache entries are cleaned up at the same time
- Indexed corpus documents can be removed from the app without deleting the original files on disk
- Separate `LLM API` and `TTS API` settings
- User-selectable online TTS provider: ElevenLabs or Gemini
- User-selectable playback source: online TTS, local Kokoro, or local Piper
- Built-in English -> Chinese translation with Argos Translate
- Built-in part-of-speech tagging with spaCy, cached locally for repeated words
- Part of speech and Chinese translation can be edited by the user and stored locally
- Local synonym lookup is powered by spaCy + WordNet
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

When the app opens:

- configure `LLM API` if you want passage generation and example-sentence generation
- configure `TTS API` if you want online word audio generation

`LLM API` currently supports Gemini. `TTS API` currently supports Gemini and ElevenLabs.

## Dependencies

- `spaCy` is used for English part-of-speech tagging and basic lexical analysis
- `en_core_web_sm` is the expected spaCy model for the current workflow
- `spacy-wordnet` and local WordNet data are used for synonym lookup
- `Argos Translate` is used for local English -> Chinese translation
- `Gemini API` is used for LLM features such as passage generation and example sentence generation
- `Gemini TTS` is supported as an online TTS provider
- `ElevenLabs` is supported as an online TTS provider
- `Kokoro` and `Piper` are optional local TTS backends

## TTS Behavior

- Configure providers in `Settings > LLM API` and `Settings > TTS API`
- Pick the active playback source in `Settings > Source`
- `LLM API` is currently Gemini-only
- `TTS API` currently supports `ElevenLabs` and `Gemini`
- If `TTS API` is set to `ElevenLabs`, online word-audio generation prefers `ElevenLabs`, then `Gemini`, then local fallback
- If `TTS API` is set to `Gemini`, online word-audio generation prefers `Gemini`, then local fallback
- ElevenLabs uses a default British-style voice for IELTS-style playback
- `Kokoro` is an optional offline source and only appears when local model files exist in `data/models/kokoro/`
- `Piper` is a project-local local source using models under `data/models/piper/`
- If the selected online source fails at playback time, the app automatically falls back to a local backend when available
- Single-word audio is cached under `data/audio_cache/sources/`
- Source-specific caches are grouped by source file and then by `a-z` or `other`
- Regular list entries prefer lightweight metadata links instead of duplicating the same online `wav`
- Recent wrong words keep dedicated cache entries so dictation review can reuse stable local audio
- Each cached word can carry metadata that marks its real backend source and the desired backend target
- The online replacement queue is persisted, so local fallback audio can still be replaced after restarting the app
- Queue throttling is conservative by provider:
  - ElevenLabs: about `1.5s` per request, `45s` cooldown after rate-limit errors
  - Gemini TTS: slower paced requests and longer cooldowns because the free-tier limit is stricter
- Passage playback keeps one source for the whole article and does not mix backends inside the same generated passage

## Local Runtime Files

- `data/models/kokoro/`: optional offline Kokoro model files
- `data/models/piper/`: Piper `*.onnx` voice models and matching `*.onnx.json` config files
- `data/audio_cache/`: generated local audio cache
- `data/audio_cache/sources/`: source-specific word caches grouped by file and first letter
- `data/audio_cache/global/`: shared online-TTS cache reused across files when available
- `data/audio_cache/pending_gemini_replacements.json`: persisted online replacement queue
- `data/pos_cache.json`: cached part-of-speech labels
- `data/translation_cache.json`: cached translations
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

## Main UI

- Left side: import, manual paste/type, save-as, new list, playback, settings, and dictation entry
- Right side: current word details plus tabs for `Review / History / Tools`
- Current word details show the selected word, part of speech, translation, and notes
- The word list is displayed in a book-style layout with numbering, the English word on the first line, and `part of speech + Chinese translation` on the second line

## Dictation

- Main dictation page has two list modes: `All` and `Recent Wrong`
- Two entry buttons are provided: `Start From Word` and `Start Learning`
- `Start From Word` opens an in-window picker so you can jump into dictation from any word in the current list mode
- `Start Learning` opens the study-mode popup; the first implemented mode is `Online Spelling`
- `Online Spelling` supports playback speed presets, countdown timing, replay, pause, previous-word, and live red/green answer feedback
- Wrong answers are recorded locally and feed back into the `Recent Wrong` list
- `Recent Wrong` shows error cause instead of normal notes, sorts by mistake count, and supports manual additions
- Recent wrong audio has its own dedicated cache source and can reuse or promote audio from the active word list

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

- Speech synthesis is powered by ElevenLabs, Gemini TTS, and optional local Kokoro / Piper playback
- Part-of-speech tagging is powered by [spaCy](https://spacy.io/)
- Synonym lookup is powered by [spaCy WordNet](https://github.com/argilla-io/spacy-wordnet) and WordNet data
- English -> Chinese translation is powered by [argosopentech/argos-translate](https://github.com/argosopentech/argos-translate)
