# Word Speaker

Word Speaker is a Windows desktop app for vocabulary study, pronunciation practice, dictation, IELTS-style content generation, and local corpus search.

## Project

- Study from imported `.txt` / `.csv` word lists
- Play word audio with online TTS or local TTS
- Practice dictation with wrong-word review
- Generate IELTS-style passages and example sentences
- Search imported corpus documents locally
- Configure separate `LLM API` and `TTS API`
- Export/import shared word-audio cache packs to reuse generated TTS across devices
- Export/import clean word resource packs as `.wspack` instead of sharing the whole `data` folder
- Update the packaged app from `Tools > Update App` via an online manifest or a local update zip
- Build an update zip and optional online manifest from a packaged app folder

For the full feature list, TTS/cache behavior, dictation workflow, and runtime file layout, see [GUIDE.md](GUIDE.md).

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

- configure `LLM API` if you want AI generation features
- configure `TTS API` if you want online TTS

Current API support:

- `LLM API`: Gemini
- `TTS API`: ElevenLabs, Gemini

Windows packaging is also supported. The packaged app includes local models and WordNet data so end users do not need to download extra runtime assets separately.

Packaged output:

- `dist/WordSpeaker/WordSpeaker.exe`
- distribute the whole `dist/WordSpeaker/` folder together

When sharing the packaged build:

- compress and share the whole `dist/WordSpeaker/` folder
- the other user must fully extract it before running
- do not run `WordSpeaker.exe` from inside the zip preview window
- prefer `7-Zip`, `Bandizip`, or `WinRAR` instead of Windows Explorer extraction if possible
- extract to a short path such as `D:\WS` or `C:\WordSpeaker`

If extraction skips files because of Windows path-length errors, the app may fail at startup with missing `numpy`/DLL errors because required files were not extracted completely.

## Distribution Notes

- `Tools > Export Shared Cache` creates a reusable audio-cache package for generated word audio only
- `Tools > Export Resource Pack` creates a `.wspack` file containing:
  - word
  - note
  - manually corrected part of speech
  - manually corrected Chinese translation
- `.wspack` is the recommended format for sharing curated vocabulary content
- do not share the whole `data/` folder between users
- packaged online updates can use a GitHub Release or any other hosted `.zip`, as long as the app can read a matching `manifest.json`
- users still need one initial packaged build that already contains the updater; after that, they can update from inside the app

## Credits

- Speech synthesis is powered by ElevenLabs, Gemini TTS, and optional local Kokoro / Piper playback
- Part-of-speech tagging is powered by [spaCy](https://spacy.io/)
- Synonym lookup is powered by [spaCy WordNet](https://github.com/argilla-io/spacy-wordnet) and WordNet data
- Online synonym lookup prefers Gemini and falls back to local spaCy + WordNet
- English -> Chinese translation is powered by [argosopentech/argos-translate](https://github.com/argosopentech/argos-translate)
