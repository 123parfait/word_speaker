# Word Speaker

A tiny Windows desktop app to import a word list (txt/csv), listen to pronunciation, and generate IELTS-style content for self-study.

## Features
- Import .txt/.csv word lists
- Type or paste words directly into the app
- Play in order, random (no repeat), or click-to-play
- Dictation check mode
- Kokoro TTS for more natural English pronunciation (64-bit Python)
- Switch accent between English (US) and English (UK) in Settings > Source
- First run auto-downloads Kokoro model files; later playback is offline
- Built-in English -> Chinese translation with Argos Translate (first run downloads model)
- Generate IELTS-listening-style passages from imported words and read them with Kokoro
- Generate IELTS-style example sentences for selected words
- Gemini API is used for article and sentence generation
- The app asks for a Gemini API key at startup and tests it before enabling AI features

## Run (CMD)

Requires **64-bit Python** (Kokoro + onnxruntime). 32-bit Python is not supported.

```bat
cd /d path\to\word_speaker
pip install -r requirements.txt
python app.py
```

If your system uses py to run Python:

```bat
cd /d path\to\word_speaker
py -3 -m pip install -r requirements.txt
py -3 app.py
```

When the app opens, paste your own Gemini API key into the popup window and click `Test and Save`.

## Input format
- .txt: one word per line
- .csv: use first column as word

## Credits
- TTS is powered by [thewh1teagle/kokoro-onnx](https://github.com/thewh1teagle/kokoro-onnx)
- English -> Chinese translation is powered by [argosopentech/argos-translate](https://github.com/argosopentech/argos-translate)
