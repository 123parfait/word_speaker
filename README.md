# Word Speaker

A Windows desktop app for vocabulary study, pronunciation, translation, IELTS-style content generation, and local corpus sentence search.

## Features
- Import `.txt` / `.csv` word lists
- Type or paste words directly into the app
- Paste two-column tables from Google Docs / Google Sheets into the manual import window
- Main word table uses three columns: `English / Notes / Chinese`
- Edit `English` and `Notes` directly in the table and save changes back to the source file
- Play in order, random (no repeat), or click-to-play
- Double-click a word to edit it
- Double-click a word or use the speaker button to play pronunciation
- Dictation check mode
- User-selectable TTS source: Gemini TTS or local Kokoro
- Automatic fallback to Kokoro when Gemini playback fails and local Kokoro models are available
- Built-in English -> Chinese translation with Argos Translate
- Translation cache is stored locally, so repeated words do not need to be translated again
- Word audio is cached locally; if words come from a source file, cached wav files are stored beside that file
- Generate IELTS-listening-style passages from imported words and read them with the selected TTS source
- Generate IELTS-style example sentences for selected words
- Practice mode for generated passages
- Gemini API is used for article and sentence generation
- The app asks for a Gemini API key at startup and tests it before enabling AI features
- `Find` window can import `.txt` / `.docx` / `.pdf`, build a local sentence index, and search by word or phrase
- `Find` supports result highlighting, result count selection (`20 / 50 / 100`), and filtering by the selected document on the right
- Local corpus data is stored in `data/corpus_index.db`

## Run (CMD)

Requires **64-bit Python**. 32-bit Python is not supported.

```bat
cd /d path\to\word_speaker
pip install -r requirements.txt
python -m spacy download en_core_web_sm
python app.py
```

If your system uses py to run Python:

```bat
cd /d path\to\word_speaker
py -3 -m pip install -r requirements.txt
py -3 -m spacy download en_core_web_sm
py -3 app.py
```

When the app opens, paste your own Gemini API key into the popup window and click `Test and Save`.

## TTS Behavior

- Pick the active source in `Settings > Source`.
- `Gemini TTS` is the default online source and requires a valid Gemini API key plus network access.
- `Kokoro` is an optional offline source and only appears when both local model files exist in `data/models/kokoro/`.
- If the selected source is Gemini and Gemini playback fails, the app automatically falls back to Kokoro when Kokoro is available locally.
- Single-word audio is cached locally. If the current list came from a source file, the cache is created beside that file. Otherwise, audio is stored under `data/audio_cache/words/`.
- Passage playback keeps one source for the whole article. It does not mix Gemini and Kokoro inside the same generated passage.

## Local Runtime Files

- `data/models/kokoro/`: optional offline Kokoro model files
- `data/audio_cache/`: generated local audio cache
- `vendor/site-packages/`: optional project-local Python runtime dependencies for local packaging or isolated runs

## Input format
- `.txt`: one word per line
- `.csv`: first column = English, second column = Notes
- Manual input window:
  - one word per line, or
  - paste a two-column table from Google Docs / Sheets

## Find Corpus
- Open the `Find` window from the `Find` button
- Import `.txt` / `.docx` / `.pdf`
- The app builds a local sentence index in `data/corpus_index.db`
- Search by word or phrase
- If you select a document on the right, search is limited to that document
- If no document is selected, search runs across the full local corpus
- Use `Clear Filter` to cancel the current document filter

## Notes
- PDF text extraction is rule-based and works best on text PDFs
- Scanned PDFs are not OCR-enabled yet
- Existing local corpus files are intentionally ignored by Git

## Credits
- Speech synthesis is powered by Gemini TTS and optional local Kokoro playback
- English -> Chinese translation is powered by [argosopentech/argos-translate](https://github.com/argosopentech/argos-translate)
