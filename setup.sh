#!/usr/bin/env bash
# Bootstrap script for Powerpoint-to-Nano.
# Goal: after clone, a user should only need to add .env and run this script.

set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$REPO_ROOT"

info() { echo "[INFO] $*"; }
ok() { echo "[OK]   $*"; }
warn() { echo "[WARN] $*"; }
err() { echo "[ERR]  $*" >&2; }

run_sudo_if_available() {
  if command -v sudo >/dev/null 2>&1; then
    sudo "$@"
  else
    "$@"
  fi
}

require_python() {
  if ! command -v python3 >/dev/null 2>&1; then
    err "python3 not found. Install Python 3.10+ first."
    exit 1
  fi

  local py_version
  py_version="$(python3 --version 2>&1 | awk '{print $2}')"
  local required="3.10"
  if [ "$(printf '%s\n' "$required" "$py_version" | sort -V | head -n1)" != "$required" ]; then
    err "Python 3.10+ required. Found: $py_version"
    exit 1
  fi
  ok "Python $py_version"
}

ensure_libreoffice() {
  if command -v soffice >/dev/null 2>&1 || command -v libreoffice >/dev/null 2>&1; then
    ok "LibreOffice already installed"
    return
  fi

  info "LibreOffice not found. Attempting install..."
  local os_name
  os_name="$(uname -s)"

  if [ "$os_name" = "Darwin" ]; then
    if command -v brew >/dev/null 2>&1; then
      brew install --cask libreoffice
    else
      err "Homebrew not found. Install Homebrew, then run: brew install --cask libreoffice"
      exit 1
    fi
  elif [ "$os_name" = "Linux" ]; then
    if command -v apt-get >/dev/null 2>&1; then
      run_sudo_if_available apt-get update
      run_sudo_if_available apt-get install -y libreoffice-impress
    else
      err "Unsupported Linux package manager. Install LibreOffice manually."
      exit 1
    fi
  else
    err "Unsupported OS: $os_name. Install LibreOffice manually: https://www.libreoffice.org/download/"
    exit 1
  fi

  if command -v soffice >/dev/null 2>&1 || command -v libreoffice >/dev/null 2>&1; then
    ok "LibreOffice installed"
  else
    err "LibreOffice installation appears to have failed."
    exit 1
  fi
}

ensure_poppler_optional() {
  if command -v pdftoppm >/dev/null 2>&1; then
    ok "Poppler already installed (pdftoppm found)"
    return
  fi

  info "Poppler not found. Attempting optional install for faster PDF conversion..."
  local os_name
  os_name="$(uname -s)"

  if [ "$os_name" = "Darwin" ]; then
    if command -v brew >/dev/null 2>&1; then
      if ! brew install poppler; then
        warn "Could not auto-install Poppler. Continuing without it."
      fi
    else
      warn "Homebrew not found. Skipping Poppler install."
    fi
  elif [ "$os_name" = "Linux" ]; then
    if command -v apt-get >/dev/null 2>&1; then
      if ! run_sudo_if_available apt-get install -y poppler-utils; then
        warn "Could not auto-install Poppler. Continuing without it."
      fi
    else
      warn "Unsupported Linux package manager. Skipping Poppler install."
    fi
  else
    warn "Unsupported OS for auto Poppler install. Skipping."
  fi
}

ensure_venv() {
  if [ ! -d ".venv" ]; then
    info "Creating virtual environment (.venv)..."
    python3 -m venv .venv
    ok "Created .venv"
  else
    ok "Using existing .venv"
  fi
}

requirements_hash() {
  if command -v shasum >/dev/null 2>&1; then
    shasum -a 256 requirements.txt | awk '{print $1}'
  elif command -v sha256sum >/dev/null 2>&1; then
    sha256sum requirements.txt | awk '{print $1}'
  else
    cat requirements.txt
  fi
}

ensure_python_deps() {
  local marker=".venv/.requirements_hash"
  local desired_hash
  desired_hash="$(requirements_hash)"
  local current_hash=""

  if [ -f "$marker" ]; then
    current_hash="$(cat "$marker")"
  fi

  if [ "$desired_hash" = "$current_hash" ]; then
    ok "Python dependencies already up-to-date"
    return
  fi

  info "Installing/updating Python dependencies..."
  .venv/bin/python -m pip install --upgrade pip
  .venv/bin/python -m pip install -r requirements.txt
  printf '%s\n' "$desired_hash" > "$marker"
  ok "Python dependencies installed"
}

ensure_env_file() {
  if [ ! -f ".env" ]; then
    cp .env.example .env
    warn "Created .env from .env.example. Fill GEMINI_API_KEY before running generation."
  else
    ok ".env file found"
  fi
}

echo "=========================================="
echo "Powerpoint-to-Nano Bootstrap"
echo "=========================================="

require_python
ensure_libreoffice
ensure_poppler_optional
ensure_venv
ensure_python_deps
ensure_env_file

echo ""
ok "Bootstrap complete."
echo "Next step: edit .env and set GEMINI_API_KEY (if not already set)."
