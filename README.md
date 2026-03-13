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
- Kokoro TTS for more natural English pronunciation
- Switch accent between English (US) and English (UK) in Settings > Source
- First run auto-downloads Kokoro model files; later playback is offline
- Built-in English -> Chinese translation with Argos Translate
- Translation cache is stored locally, so repeated words do not need to be translated again
- Generate IELTS-listening-style passages from imported words and read them with Kokoro
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
- TTS is powered by [thewh1teagle/kokoro-onnx](https://github.com/thewh1teagle/kokoro-onnx)
- English -> Chinese translation is powered by [argosopentech/argos-translate](https://github.com/argosopentech/argos-translate)
