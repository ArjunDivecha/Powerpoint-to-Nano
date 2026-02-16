# Powerpoint-to-Nano

Convert a PowerPoint deck (`.pptx`) into a brand-new set of "Nano Banana" style slides using Google's Gemini image model, and export the result as a **single multi-page PDF**.

## What this does (plain English)
- **Step 1:** The tool renders your PowerPoint slides to images (CLI: LibreOffice default, with optional Keynote/auto on macOS; Streamlit PPTX path: LibreOffice).
- **Step 2:** Each slide picture is sent to Gemini (image model), asking it to redesign the slide as a clean infographic.
- **Step 3:** All redesigned slide images are combined into one PDF.

## Requirements
- LibreOffice installed (default renderer)
- Keynote installed (optional, only if you choose `--pptx-method keynote`)
- Python 3.10+ recommended
- Gemini API key
- google-genai SDK version that supports Interactions API (see `requirements.txt`)

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
# edit .env and add GEMINI_API_KEY
```

If you already installed dependencies earlier, upgrade to the required SDK version:

```bash
pip install -r requirements.txt --upgrade
```

## First-run macOS permissions (only for Keynote mode)
If you run with `--pptx-method keynote`, macOS may prompt you to allow automation:
- “Terminal” (or your IDE) wants to control “Keynote”

Click **Allow**. If it fails, check:
- System Settings -> Privacy & Security -> Automation

## Usage

```bash
.venv/bin/python pptx2nano.py /path/to/deck.pptx --workers 4
```

Optional style:

```bash
.venv/bin/python pptx2nano.py /path/to/deck.pptx --style minimalist --workers 4
```

Select PPTX render method:

```bash
# default is libreoffice
.venv/bin/python pptx2nano.py /path/to/deck.pptx --pptx-method libreoffice

# optional Keynote path on macOS
.venv/bin/python pptx2nano.py /path/to/deck.pptx --pptx-method keynote

# auto: try Keynote first on macOS, then fall back to LibreOffice
.venv/bin/python pptx2nano.py /path/to/deck.pptx --pptx-method auto
```

Set a custom LibreOffice timeout (useful for large decks):

```bash
# one-off CLI override
.venv/bin/python pptx2nano.py /path/to/deck.pptx --libreoffice-timeout-seconds 300

# or environment default used by converter helpers
export PPTX_LIBREOFFICE_TIMEOUT_SECONDS=300
```

List built-in styles:

```bash
# NOTE: current CLI parser still requires a pptx_path argument
.venv/bin/python pptx2nano.py /path/to/deck.pptx --list-styles
```

### Built-in styles (same names as the upstream repo)

- **lego**: Bright primary colors, blocky shapes, toy-like 3D appearance, snap-together aesthetic
- **ghibli**: Hand-drawn feel, soft watercolor palette, whimsical organic shapes, Studio Ghibli anime aesthetic
- **cyberpunk**: Neon colors, dark background, glowing elements, futuristic tech aesthetic, grid patterns
- **minimalist**: Clean white/gray palette, thin lines, lots of whitespace, simple sans-serif fonts
- **blueprint**: Technical drawing style, blue background, white lines, grid paper, architectural feel
- **hand-drawn**: Sketchy lines, imperfect shapes, notebook paper feel, casual doodle aesthetic

You can also use a custom style name like `retro`, `corporate`, `vaporwave`, etc.

## Outputs
By default it writes to `pptx2nano_output/`:

- Rendered slide images (from LibreOffice or Keynote)
  - `pptx2nano_output/<deck_name>/rendered/*.png`
- Generated slide images (from Gemini)
  - `pptx2nano_output/<deck_name>/generated/slide_###.png`
- Final PDF
  - `pptx2nano_output/<deck_name>.pdf`

## Notes
- This tool keeps the **original slide aspect ratio** by using PPTX rendering output and by telling the model to keep the same aspect ratio.
- Parallel generation uses threads because the slow part is network calls to Gemini.
- If you see an Interactions API error while `python3` works differently from your app, run via `.venv/bin/python` to ensure the expected `google-genai` version is used.

## Streamlit App
To use the local web UI:

```bash
source .venv/bin/activate
streamlit run streamlit_app.py
```

Current behavior:
- PPTX conversion in Streamlit uses LibreOffice.
- Streamlit also supports PDF, GIF, TXT, DOCX, image files, and Markdown as inputs.
- The current native file-picker button uses AppleScript (`osascript`), so that picker flow is macOS-specific.
- Streamlit speed/accuracy controls include:
  - LibreOffice timeout (seconds) for PPTX conversion
  - Generate ALL worker count
  - PPTX text extraction mode (`off`, `strict`, `assisted`) and optional dedupe
  - Preview max pages (limits PDF pages rendered in-browser)

## Tests
Run pytest:

```bash
source .venv/bin/activate
pytest -q
```

Notes:
- `test_streamlit_logic_uses_libreoffice` is a portable unit test.
- `test_libreoffice_conversion` is an optional integration test.
  - Enable by setting `P2N_TEST_PPTX=/absolute/path/to/sample.pptx`.
- `test_speed_accuracy_improvements.py` validates timeout resolution and pagination/markdown token behavior.
