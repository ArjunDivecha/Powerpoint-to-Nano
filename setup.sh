#!/bin/bash
# Setup script for Powerpoint-to-Nano
# Run this after cloning the repository

set -e  # Exit on error

echo "=========================================="
echo "Powerpoint-to-Nano Setup"
echo "=========================================="
echo ""

# Check Python version
echo "Checking Python version..."
PYTHON_VERSION=$(python3 --version 2>&1 | awk '{print $2}')
REQUIRED_VERSION="3.10"

if [ "$(printf '%s\n' "$REQUIRED_VERSION" "$PYTHON_VERSION" | sort -V | head -n1)" != "$REQUIRED_VERSION" ]; then 
    echo "❌ Python 3.10+ required. Found: $PYTHON_VERSION"
    exit 1
fi
echo "✅ Python $PYTHON_VERSION"

# Check for LibreOffice
echo ""
echo "Checking for LibreOffice..."
if command -v soffice &> /dev/null || command -v libreoffice &> /dev/null; then
    echo "✅ LibreOffice found"
else
    echo "❌ LibreOffice not found!"
    echo ""
    echo "Please install LibreOffice:"
    echo "  macOS:   brew install --cask libreoffice"
    echo "  Linux:   sudo apt install libreoffice-impress"
    echo "  Windows: https://www.libreoffice.org/download/"
    exit 1
fi

# Check for poppler (optional)
echo ""
echo "Checking for poppler (optional)..."
if command -v pdftoppm &> /dev/null; then
    echo "✅ poppler found (faster PDF processing)"
else
    echo "⚠️  poppler not found (optional but recommended)"
    echo "   Install for faster PDF processing:"
    echo "   macOS: brew install poppler"
    echo "   Linux: sudo apt install poppler-utils"
fi

# Create virtual environment
echo ""
echo "Creating virtual environment..."
if [ -d ".venv" ]; then
    echo "⚠️  .venv already exists. Skipping creation."
else
    python3 -m venv .venv
    echo "✅ Virtual environment created"
fi

# Activate virtual environment
echo ""
echo "Activating virtual environment..."
source .venv/bin/activate

# Install Python dependencies
echo ""
echo "Installing Python dependencies..."
pip install --upgrade pip
pip install -r requirements.txt
echo "✅ Dependencies installed"

# Set up .env file
echo ""
echo "Setting up environment variables..."
if [ -f ".env" ]; then
    echo "⚠️  .env already exists. Skipping creation."
else
    cp .env.example .env
    echo "✅ Created .env file"
    echo ""
    echo "🔴 IMPORTANT: Edit .env and add your GEMINI_API_KEY"
    echo "   Get your API key from: https://aistudio.google.com/app/apikey"
fi

echo ""
echo "=========================================="
echo "Setup Complete!"
echo "=========================================="
echo ""
echo "Next steps:"
echo "1. Edit .env and add your GEMINI_API_KEY"
echo "2. Run: source .venv/bin/activate"
echo "3. Run: streamlit run streamlit_app.py"
echo ""
echo "For troubleshooting, see NEW_USER_AUDIT.md"
