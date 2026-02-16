# Testing LibreOffice PPTX Conversion

This guide helps you test LibreOffice as a replacement for Keynote before committing to it.

## Why Consider LibreOffice?

| Feature | Keynote | LibreOffice |
|---------|---------|-------------|
| **Platform** | macOS only | Windows, macOS, Linux |
| **Cloud Deploy** | ❌ No | ✅ Yes |
| **Cost** | Free (macOS) | Free |
| **Quality** | Excellent | Good (may vary) |
| **Speed** | Fast | Moderate |

## Quick Test (Command Line)

### 1. Install LibreOffice

**macOS:**
```bash
brew install --cask libreoffice
# Optional: install poppler for faster PDF→PNG conversion
brew install poppler
```

**Linux (Ubuntu/Debian):**
```bash
sudo apt install libreoffice-impress poppler-utils
```

### 2. Run the Comparison Test

```bash
cd Powerpoint-to-Nano
python test_converter.py /path/to/your/test.pptx
```

This will:
- Convert the PPTX using **Keynote** → `pptx_converter_test/keynote_output/`
- Convert the PPTX using **LibreOffice** → `pptx_converter_test/libreoffice_output/`
- Show a comparison summary

### 3. Visually Compare Results

```bash
# Open both output folders
open pptx_converter_test/keynote_output/
open pptx_converter_test/libreoffice_output/
```

Compare the same slide numbers side-by-side. Check for:
- Font rendering differences
- Image quality
- Layout shifts
- Missing elements (charts, shapes)
- Color accuracy

## Test in Streamlit App

### 1. Install LibreOffice (if not done above)

### 2. Run the Streamlit App

```bash
streamlit run streamlit_app.py
```

### 3. Select Conversion Method

When you load a PPTX file, you'll see a new toggle:

```
[🍎 Keynote] [📄 LibreOffice]
```

- Click **LibreOffice** to test the new method
- Click **Keynote** to use the original method

### 4. Compare Results

Process the same PPTX with both methods and compare the generated slides.

## Code Usage

### In Python Scripts

```python
from pptx_converter import export_slides

# Use Keynote (macOS only)
slides = export_slides("deck.pptx", output_dir, method="keynote")

# Use LibreOffice (cross-platform)
slides = export_slides("deck.pptx", output_dir, method="libreoffice")

# Auto-detect (uses Keynote on macOS if available, else LibreOffice)
slides = export_slides("deck.pptx", output_dir, method="auto")
```

### In Your Modified Streamlit App

The app now automatically shows the conversion method selector when:
1. You're on macOS
2. LibreOffice is installed
3. You've selected a PPTX file

The selected method is stored in `st.session_state["pptx_conversion_method"]`.

## Making the Switch

If LibreOffice quality is acceptable:

### Option A: Environment Variable (Recommended)

Set an environment variable to default to LibreOffice:

```bash
export PPTX_CONVERSION_METHOD=libreoffice
streamlit run streamlit_app.py
```

Then modify `streamlit_app.py` to read this:

```python
import os
# ...
default_method = os.getenv("PPTX_CONVERSION_METHOD", "keynote")
st.session_state.setdefault("pptx_conversion_method", default_method)
```

### Option B: Hardcode the Change

In `streamlit_app.py`, change the default:

```python
# Change this line
conversion_method = st.session_state.get("pptx_conversion_method", "keynote")
# To:
conversion_method = st.session_state.get("pptx_conversion_method", "libreoffice")
```

### Option C: Remove Keynote Entirely

Replace the import and usage:

```python
# Old
import pptx2nano
rendered_paths = pptx2nano.export_slides_with_keynote(...)

# New
from pptx_converter import export_slides
rendered_paths = export_slides(..., method="libreoffice")
```

## Known Differences

### LibreOffice May Produce:
- Slightly different font rendering (substitutes missing fonts)
- Different default slide dimensions
- Slightly lower image quality (configurable via DPI)
- Slower conversion (especially first run)

### To Improve LibreOffice Quality:

1. **Increase DPI** (default is 150):
   ```python
   slides = export_slides("deck.pptx", output_dir, method="libreoffice", dpi=300)
   ```

2. **Install Microsoft Fonts** (Linux):
   ```bash
   sudo apt install ttf-mscorefonts-installer
   ```

3. **Embed Fonts in PPTX** when saving from PowerPoint

## Troubleshooting

### "LibreOffice not found"

Install LibreOffice and ensure it's in your PATH:
```bash
which soffice
# or
which libreoffice
```

### "pdftoppm not found"

Install poppler:
```bash
# macOS
brew install poppler

# Linux
sudo apt install poppler-utils
```

The converter will fall back to `pdf2image` if `pdftoppm` is not available.

### Conversion is Slow

- First conversion is slow (LibreOffice startup)
- Subsequent conversions are faster
- Increase timeout in code if needed:
  ```python
  # In pptx_converter.py, increase timeout
  timeout=300  # 5 minutes instead of 2
  ```

## Deployment Ready

Once you've verified LibreOffice works for your PPTX files, you can deploy to Railway (or any cloud platform) using the provided `Dockerfile` which includes LibreOffice.

See the deployment guide for details.
