#!/usr/bin/env python3
from __future__ import annotations

from pathlib import Path

import pytest

import pptx_converter

pytest.importorskip("streamlit")
import streamlit_app


def test_resolve_libreoffice_timeout_explicit_overrides_env(monkeypatch) -> None:
    monkeypatch.setenv("PPTX_LIBREOFFICE_TIMEOUT_SECONDS", "999")
    assert pptx_converter._resolve_libreoffice_timeout_seconds(45) == 45


def test_resolve_libreoffice_timeout_from_env(monkeypatch) -> None:
    monkeypatch.setenv("PPTX_LIBREOFFICE_TIMEOUT_SECONDS", "300")
    assert pptx_converter._resolve_libreoffice_timeout_seconds(None) == 300


def test_resolve_libreoffice_timeout_invalid_env_falls_back(monkeypatch) -> None:
    monkeypatch.setenv("PPTX_LIBREOFFICE_TIMEOUT_SECONDS", "not-a-number")
    assert pptx_converter._resolve_libreoffice_timeout_seconds(None) == 120


def test_markdown_to_tokens_preserves_headers_and_list_markers() -> None:
    md = """# Title

Paragraph text.

- First item
- Second item
"""
    tokens = streamlit_app._markdown_to_tokens(md)
    assert any(is_header and text == "Title" for text, is_header in tokens)
    assert any((not is_header) and text.startswith("• ") for text, is_header in tokens)


def test_text_token_pagination_creates_multiple_pages(tmp_path: Path) -> None:
    tokens = [("Long body line " * 12, False) for _ in range(120)]
    pages = streamlit_app._render_text_tokens_to_pages(tokens, tmp_path, prefix="page")
    assert len(pages) >= 2
    for p in pages:
        assert p.exists()
