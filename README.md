# Word Speaker

A tiny Windows desktop app to import a word list (txt/csv) and listen to pronunciation.

## Features
- Import .txt/.csv word lists
- Play in order or random (no repeat)
- Dictation check mode
- Offline speech via pyttsx3 (Windows SAPI5). No network required.

## Run (CMD)

`at
cd /d path\to\word_speaker
pip install -r requirements.txt
python app.py
`

If your system uses py to run Python:

`at
cd /d path\to\word_speaker
py -3 -m pip install -r requirements.txt
py -3 app.py
`

## Input format
- .txt: one word per line
- .csv: use first column as word
