# Word Speaker

Word Speaker is a Windows desktop app for vocabulary study, pronunciation practice, dictation, IELTS-style content generation, and local corpus search.

## Project

- Study from imported `.txt` / `.csv` word lists
- Play word audio with online TTS or local TTS
- Practice dictation with wrong-word review
- Generate IELTS-style passages and example sentences
- Search imported corpus documents locally
- Configure separate `LLM API` and `TTS API`

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

## Credits

- Speech synthesis is powered by ElevenLabs, Gemini TTS, and optional local Kokoro / Piper playback
- Part-of-speech tagging is powered by [spaCy](https://spacy.io/)
- Synonym lookup is powered by [spaCy WordNet](https://github.com/argilla-io/spacy-wordnet) and WordNet data
- Online synonym lookup prefers Gemini and falls back to local spaCy + WordNet
- English -> Chinese translation is powered by [argosopentech/argos-translate](https://github.com/argosopentech/argos-translate)
