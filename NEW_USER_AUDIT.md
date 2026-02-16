# New User Setup Audit

This document identifies all potential problems a friend might encounter when cloning and running this repository.

## 🔴 Critical Issues (Will Break Setup)

### 1. README is Outdated

**Problem:** The README still mentions Keynote and macOS requirements, but the code now uses LibreOffice.

**Impact:** HIGH - User will be confused about requirements

**Fix Needed:** Update README to reflect LibreOffice requirement

---

### 2. LibreOffice Not Installed

**Problem:** LibreOffice is required but not listed in `requirements.txt` (it's a system package, not Python)

**Error they'll see:**
```
RuntimeError: LibreOffice not found. Install it:
  macOS: brew install --cask libreoffice
  Linux: sudo apt install libreoffice-impress
```

**Impact:** HIGH - App won't work without it

**Fix:** Install LibreOffice BEFORE running the app

---

### 3. poppler Not Installed (Optional but Recommended)

**Problem:** `pdftoppm` from poppler is used for faster PDF→PNG conversion. Falls back to slower `pdf2image` if not installed.

**Impact:** LOW - Works without it, but slower

**Fix:** `brew install poppler` (macOS) or `sudo apt install poppler-utils` (Linux)

---

### 4. Gemini API Key Required

**Problem:** App won't work without `GEMINI_API_KEY` in `.env` file

**Error they'll see:**
```
RuntimeError: GEMINI_API_KEY is not set. Put it in a .env file
```

**Impact:** HIGH - Core functionality requires it

**Fix:** 
1. Copy `.env.example` to `.env`
2. Get API key from https://aistudio.google.com/app/apikey
3. Add to `.env` file

---

## 🟡 Medium Issues (Will Cause Confusion)

### 5. Python Version

**Problem:** Code uses Python 3.10+ features (type hints like `str | None`)

**Error they'll see:**
```
TypeError: unsupported operand type(s) for |: 'type' and 'NoneType'
```

**Impact:** MEDIUM - Won't run on older Python

**Fix:** Use Python 3.10 or newer

---

### 6. Virtual Environment Not Activated

**Problem:** If they forget to activate venv, imports will fail

**Error they'll see:**
```
ModuleNotFoundError: No module named 'streamlit'
```

**Impact:** MEDIUM - Common beginner mistake

**Fix:** Remind to run `source .venv/bin/activate` (macOS/Linux) or `.venv\Scripts\activate` (Windows)

---

### 7. Missing System Dependencies for pdf2image

**Problem:** `pdf2image` Python package requires poppler installed on system

**Error they'll see:**
```
pdf2image.exceptions.PDFInfoNotInstalledError: Unable to get page count
```

**Impact:** MEDIUM - Only affects PDF uploads

**Fix:** Install poppler (see #3)

---

## 🟢 Minor Issues (Annoyances)

### 8. Streamlit Warnings About `use_container_width`

**Problem:** Deprecation warnings spam the console

**What they'll see:**
```
Please replace `use_container_width` with `width`.
`use_container_width` will be removed after 2025-12-31.
```

**Impact:** LOW - Works fine, just noisy

**Fix:** Update code to use `width='stretch'` instead

---

### 9. Output Directories Already Exist

**Problem:** Previous runs leave `pptx2nano_output_streamlit/` folders

**Impact:** LOW - Might confuse user about which files are new

**Fix:** Add cleanup instructions or auto-cleanup

---

### 10. First LibreOffice Run is Slow

**Problem:** LibreOffice takes ~40 seconds on first conversion (startup time)

**Impact:** LOW - User might think it's stuck

**Fix:** Add progress indicator or warning about first-run slowness

---

## 📋 Platform-Specific Issues

### macOS
- ✅ Primary development platform
- ✅ LibreOffice available via Homebrew
- ⚠️ Apple Silicon (M1/M2/M3) - should work fine

### Linux
- ✅ LibreOffice available via apt
- ⚠️ May need `libreoffice-impress` specifically
- ⚠️ Font differences may affect PPTX rendering

### Windows
- ❌ NOT TESTED
- ⚠️ Path handling may differ
- ⚠️ LibreOffice paths are different
- ⚠️ Activation script is `.venv\Scripts\activate.bat`

---

## 🚀 Step-by-Step Setup Guide for Friends

### Prerequisites Checklist

Before starting, ensure you have:
- [ ] Python 3.10+ installed (`python --version`)
- [ ] Git installed (`git --version`)
- [ ] Gemini API key from https://aistudio.google.com/app/apikey

### Setup Steps

```bash
# 1. Clone the repo
git clone https://github.com/ArjunDivecha/Powerpoint-to-Nano.git
cd Powerpoint-to-Nano

# 2. Install LibreOffice (system dependency)
# macOS:
brew install --cask libreoffice
brew install poppler  # optional but recommended

# Linux (Ubuntu/Debian):
sudo apt install libreoffice-impress poppler-utils

# 3. Create virtual environment
python -m venv .venv

# 4. Activate virtual environment
# macOS/Linux:
source .venv/bin/activate

# Windows:
# .venv\Scripts\activate

# 5. Install Python dependencies
pip install -r requirements.txt

# 6. Set up environment variables
cp .env.example .env
# Edit .env and add your GEMINI_API_KEY

# 7. Run the app
streamlit run streamlit_app.py
```

### Troubleshooting

| Problem | Solution |
|---------|----------|
| `LibreOffice not found` | Install LibreOffice (see step 2) |
| `GEMINI_API_KEY is not set` | Copy `.env.example` to `.env` and add your key |
| `ModuleNotFoundError` | Activate virtual environment (step 4) |
| `Python version error` | Upgrade to Python 3.10+ |
| Conversion is very slow | Normal for first run; subsequent runs are faster |

---

## 🔧 Recommended Fixes for Repo Owner

### Immediate Actions

1. **Update README.md**
   - Remove Keynote references
   - Add LibreOffice installation instructions
   - Update requirements section

2. **Add `runtime.txt`** or `.python-version`
   - Specify Python 3.10+ requirement

3. **Create `setup.sh` script**
   - Automate the setup process
   - Check for LibreOffice
   - Create .env file

4. **Add better error messages**
   - Detect missing LibreOffice early
   - Provide clear installation instructions

### Code Improvements

5. **Fix deprecation warnings**
   - Replace `use_container_width` with `width`

6. **Add startup check**
   - Verify LibreOffice is installed on app start
   - Show helpful message if not

7. **Add progress bar for LibreOffice**
   - First conversion takes ~40s, user needs feedback

---

## Summary

| Severity | Count | Issues |
|----------|-------|--------|
| 🔴 Critical | 3 | Outdated README, Missing LibreOffice, Missing API key |
| 🟡 Medium | 4 | Python version, venv activation, poppler, pdf2image |
| 🟢 Minor | 3 | Deprecation warnings, output folders, first-run slowness |

**Biggest Risk:** User clones repo, follows outdated README, gets confused about Keynote vs LibreOffice requirements.
