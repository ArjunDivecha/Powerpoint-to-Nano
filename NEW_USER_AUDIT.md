# New User Setup Audit (Current)

This file captures the actual setup risks and checks for a new user cloning this repo today.

## Current Status Snapshot
- README is aligned with current renderer behavior.
- CLI default renderer is LibreOffice.
- Streamlit PPTX conversion uses LibreOffice.
- Python dependency list includes `python-docx`, `markdown`, and `pypdf`.
- Stale slide-output issues were fixed by cleaning render folders before export.

## Critical Setup Blockers

### 1. Missing Gemini API key
If `GEMINI_API_KEY` is not set in `.env`, generation fails.

### 2. Missing LibreOffice
PPTX rendering depends on LibreOffice unless using explicit Keynote mode in CLI.

### 3. Wrong Python environment
If you run with system Python instead of the repo venv, you may hit `google-genai` version mismatches (missing Interactions API).

## Recommended Setup Flow

```bash
git clone https://github.com/ArjunDivecha/Powerpoint-to-Nano.git
cd Powerpoint-to-Nano

# System deps
# macOS:
brew install --cask libreoffice
brew install poppler  # optional but recommended

# Linux (Ubuntu/Debian):
# sudo apt update && sudo apt install libreoffice-impress poppler-utils

# Python env
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
# add GEMINI_API_KEY to .env

# optional: raise default LibreOffice timeout for large decks
export PPTX_LIBREOFFICE_TIMEOUT_SECONDS=300
```

## Verification Checklist

```bash
# 1) Python deps load
source .venv/bin/activate
python -c "import google.genai, docx, markdown, pypdf, streamlit; print('ok')"

# 2) LibreOffice exists
which soffice || which libreoffice

# 3) Quick CLI smoke test (1 slide)
python pptx2nano.py /path/to/deck.pptx --max-slides 1 --workers 1

# 4) Unit/integration tests
pytest -q
```

## Test Reality
- `pytest -q` runs a portable unit test by default.
- A real LibreOffice integration test is included but optional.
- Enable integration test with:

```bash
export P2N_TEST_PPTX=/absolute/path/to/sample.pptx
pytest -q
```

## Known Limitations (Still True)

### 1. First conversion can feel slow
LibreOffice startup cost can make first run noticeably slower.
For large decks, increase timeout via `--libreoffice-timeout-seconds` or `PPTX_LIBREOFFICE_TIMEOUT_SECONDS`.

### 2. Keynote mode is macOS-only
`--pptx-method keynote` and `--pptx-method auto` Keynote fallback are only applicable on macOS.

### 3. Streamlit file picker is macOS-specific
The Streamlit "Choose file..." flow calls AppleScript (`osascript`). Non-macOS users may need an alternate upload/input flow.

## Fast Troubleshooting Table

| Problem | Likely Cause | Fix |
|---|---|---|
| `GEMINI_API_KEY is not set` | `.env` missing key | Add key to `.env` |
| `LibreOffice not found` | System package missing | Install LibreOffice and re-open shell |
| Interactions API error | Wrong interpreter/package version | Use `.venv/bin/python` and `pip install -r requirements.txt --upgrade` |
| `pytest` integration test skipped | `P2N_TEST_PPTX` unset | Set env var to a real `.pptx` file |

## Suggested Next Hardening Steps
1. Add CI workflow to run `pytest -q` on every push.
2. Add a startup preflight command (`make doctor` or script) to validate API key, LibreOffice, and model SDK version.
3. Add Windows-specific setup notes once validated end-to-end.
