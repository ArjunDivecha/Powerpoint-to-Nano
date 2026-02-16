# PPTX Rendering Guide (Current Behavior)

This guide documents how PPTX rendering works in the repo today.

## Current Reality
- CLI (`pptx2nano.py`) defaults to LibreOffice rendering.
- CLI supports `--pptx-method libreoffice|keynote|auto`.
- Streamlit (`streamlit_app.py`) uses LibreOffice for PPTX conversion.
- Keynote is optional and mainly useful for quality comparison on macOS.

## 1. Install System Dependencies

### macOS
```bash
brew install --cask libreoffice
# Optional but recommended for faster PDF -> PNG conversion
brew install poppler
```

### Linux (Ubuntu/Debian)
```bash
sudo apt update
sudo apt install libreoffice-impress poppler-utils
```

## 2. Install Python Dependencies

```bash
cd Powerpoint-to-Nano
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
# Add GEMINI_API_KEY to .env
```

## 3. Quick CLI Smoke Test (Default = LibreOffice)

```bash
source .venv/bin/activate
python pptx2nano.py /path/to/deck.pptx --max-slides 1 --workers 1 --style minimalist --write-run-log
```

Expected output includes:
- `[INFO] Rendering slides with libreoffice: ...`
- `[DONE] Created: ...pdf`

### Timeout tuning for large decks

```bash
# One-off timeout override (seconds)
python pptx2nano.py /path/to/deck.pptx --libreoffice-timeout-seconds 300

# Global fallback for converter helpers
export PPTX_LIBREOFFICE_TIMEOUT_SECONDS=300
```

## 4. Renderer Comparison (Keynote vs LibreOffice)

Use the comparison helper:

```bash
source .venv/bin/activate
python test_converter.py /path/to/deck.pptx
```

This writes:
- `pptx_converter_test/keynote_output/`
- `pptx_converter_test/libreoffice_output/`

On macOS, you can open both:

```bash
open pptx_converter_test/keynote_output/
open pptx_converter_test/libreoffice_output/
```

## 5. Explicit Renderer Selection in CLI

```bash
# Default path
python pptx2nano.py /path/to/deck.pptx --pptx-method libreoffice

# Optional Keynote path on macOS
python pptx2nano.py /path/to/deck.pptx --pptx-method keynote

# Auto: try Keynote first on macOS, then LibreOffice
python pptx2nano.py /path/to/deck.pptx --pptx-method auto
```

## 6. Streamlit Check

```bash
source .venv/bin/activate
streamlit run streamlit_app.py
```

For PPTX inputs, Streamlit currently uses LibreOffice conversion.
The built-in "Choose file..." button currently uses `osascript`, so that file-picker flow is macOS-specific.

Useful Streamlit controls for speed/accuracy:
- `LibreOffice timeout (seconds)` for PPTX conversion.
- `Generate ALL workers` for parallel generation.
- `Text extraction mode (PPTX)`: `off`, `strict`, `assisted`.
- `Deduplicate extracted text lines` toggle.
- `Preview max pages` to cap in-browser PDF preview rendering.

## 7. Optional Integration Test (pytest)

`test_streamlit_libreoffice.py` contains:
- one portable unit test (always runnable)
- one optional real conversion test

Enable the real conversion test by setting a fixture file path:

```bash
export P2N_TEST_PPTX=/absolute/path/to/sample.pptx
pytest -q
```

Without `P2N_TEST_PPTX`, the integration test is skipped.

## Troubleshooting

### `RuntimeError: LibreOffice not found`
Install LibreOffice and verify:

```bash
which soffice
# or
which libreoffice
```

### Interactions API errors (google-genai version mismatch)
Run using the project venv interpreter:

```bash
.venv/bin/python pptx2nano.py /path/to/deck.pptx --max-slides 1
```

### First conversion is slow
This is normal due to LibreOffice startup time. Subsequent conversions are faster.
