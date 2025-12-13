# Powerpoint-to-Nano

Convert a PowerPoint deck (`.pptx`) into a brand-new set of "Nano Banana" style slides using Google's Gemini image model, and export the result as a **single multi-page PDF**.

## What this does (plain English)
- **Step 1:** Keynote opens your PowerPoint and exports every slide as a picture.
- **Step 2:** Each slide picture is sent to Gemini (image model), asking it to redesign the slide as a clean infographic.
- **Step 3:** All redesigned slide images are combined into one PDF.

## Requirements
- macOS
- Keynote installed
- Python 3.10+ recommended
- Gemini API key

## Setup

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
# edit .env and add GEMINI_API_KEY
```

## First-run macOS permissions (important)
The first time you run this, macOS may prompt you to allow automation:
- “Terminal” (or your IDE) wants to control “Keynote”

Click **Allow**.

If it fails, check:
- System Settings -> Privacy & Security -> Automation

## Usage

```bash
python pptx2nano.py /path/to/deck.pptx --workers 4
```

Optional style:

```bash
python pptx2nano.py /path/to/deck.pptx --style minimalist --workers 4
```

List built-in styles:

```bash
python pptx2nano.py --list-styles
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

- Rendered slide images (from Keynote)
  - `pptx2nano_output/<deck_name>/rendered/*.png`
- Generated slide images (from Gemini)
  - `pptx2nano_output/<deck_name>/generated/slide_###.png`
- Final PDF
  - `pptx2nano_output/<deck_name>.pdf`

## Notes
- This tool keeps the **original slide aspect ratio** by using Keynote's rendering and by telling the model to keep the same aspect ratio.
- Parallel generation uses threads because the slow part is network calls to Gemini.
