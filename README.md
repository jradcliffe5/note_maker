# note_maker

AI-powered tool for generating and supplementing meeting notes from Zoom recordings.

## Usage

### Generate notes from scratch

```bash
python3 meeting_notes.py \
    --captions closed_caption.txt \
    --chat chat.txt \
    --audio audio.m4a \
    --output meeting_notes.docx
```

Feeds sources to Claude, which returns structured JSON (title, attendees, summary, agenda items, decisions, action items). The result is written to a formatted DOCX.

### Supplement existing notes

```bash
python3 meeting_notes.py EVNPC_agenda.docx \
    --captions closed_caption.txt \
    --chat chat.txt \
    --audio audio1.m4a audio2.m4a
```

Claude reads the existing notes alongside supplementary sources and identifies new information not already captured. Insertions are added inline after the relevant paragraph, colour-coded by source (captions = blue, chat = purple, audio = green). Falls back to appending full transcripts if Claude returns nothing.

## Options

| Flag | Default | Description |
|---|---|---|
| `notes` (positional, optional) | — | Existing DOCX to supplement; omit to generate from scratch |
| `--captions FILE` | — | Closed captions text file |
| `--chat FILE` | — | Chat log text file |
| `--audio FILE [FILE ...]` | — | Audio file(s) to transcribe |
| `--output FILE` | `meeting_notes.docx` / `<input>_supplemented.docx` | Output DOCX path |
| `--title TITLE` | `Meeting Notes` | Meeting title (generate mode only) |
| `--date DATE` | today | Meeting date (generate mode only) |
| `--model MODEL` | `mlx-community/whisper-large-v3-turbo` | Whisper model for transcription |
| `--no-ai` | — | Skip Claude; write raw transcripts into document |

## Requirements

```bash
pip install python-docx mlx-whisper
```

- **Claude Code CLI** (`claude`) must be installed and on `PATH` — used for AI synthesis.
- `mlx-whisper` is only required when transcribing audio files (Apple Silicon recommended).

## Audio transcript caching

Transcripts are cached alongside the audio file as `<filename>.transcript.txt` or `<filename>_transcript.txt`. On subsequent runs the cached file is used automatically, avoiding re-transcription of large audio files.
