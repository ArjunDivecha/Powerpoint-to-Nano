#!/usr/bin/env bash
# One-command launcher for local app usage.
# Intended workflow after clone: add .env, then run this script.

set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$REPO_ROOT"

./setup.sh

if [ ! -f ".env" ]; then
  echo "[ERR] .env file not found. Create it (or copy .env.example) and set GEMINI_API_KEY." >&2
  exit 1
fi

api_key_line="$(grep -E '^GEMINI_API_KEY=' .env | tail -n 1 || true)"
api_key_value="${api_key_line#GEMINI_API_KEY=}"

if [ -z "$api_key_value" ] || [ "$api_key_value" = "your_real_api_key_here" ]; then
  echo "[ERR] GEMINI_API_KEY is missing in .env." >&2
  echo "      Edit .env and set GEMINI_API_KEY, then re-run ./run_streamlit.sh" >&2
  exit 1
fi

exec .venv/bin/python -m streamlit run streamlit_app.py "$@"
