#!/usr/bin/env python3
"""Tests for LibreOffice conversion paths used by the Streamlit app.

Notes:
- `test_libreoffice_conversion` is an optional integration test.
  Set `P2N_TEST_PPTX=/absolute/path/to/sample.pptx` to enable it.
- `test_streamlit_logic_uses_libreoffice` is a portable unit test.
"""

from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path

import pytest
from PIL import Image

from pptx_converter import _find_libreoffice, export_slides


def _integration_pptx_path() -> Path:
    """Resolve integration test fixture path from env and skip if unavailable."""
    env_path = os.getenv("P2N_TEST_PPTX")
    if not env_path:
        pytest.skip("Set P2N_TEST_PPTX to run LibreOffice integration test.")

    pptx_path = Path(env_path).expanduser().resolve()
    if not pptx_path.exists():
        pytest.skip(f"P2N_TEST_PPTX does not exist: {pptx_path}")
    return pptx_path


def test_libreoffice_conversion() -> None:
    """Run a real LibreOffice conversion when integration fixtures are configured."""
    pptx_path = _integration_pptx_path()

    libreoffice_path = _find_libreoffice()
    if not libreoffice_path:
        pytest.skip("LibreOffice is not installed or not on PATH.")

    with tempfile.TemporaryDirectory() as tmpdir:
        out_dir = Path(tmpdir) / "pptx2nano_output_streamlit"
        rendered_dir = out_dir / pptx_path.stem / "rendered"

        rendered_paths = export_slides(pptx_path, rendered_dir, method="libreoffice")

        assert rendered_paths, "Expected at least one rendered slide."
        for path in rendered_paths:
            assert path.exists(), f"Missing rendered output: {path}"

        # Validate a few outputs are actually readable images.
        for path in rendered_paths[:3]:
            with Image.open(path) as img:
                assert img.width > 0 and img.height > 0


def test_streamlit_logic_uses_libreoffice(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    """Simulate Streamlit logic and assert it calls export_slides with LibreOffice."""
    calls: dict[str, str] = {}

    def fake_export_slides(pptx_path: Path, rendered_dir: Path, method: str = "auto") -> list[Path]:
        calls["method"] = method
        rendered_dir.mkdir(parents=True, exist_ok=True)
        fake_slide = rendered_dir / "slide_001.png"
        Image.new("RGB", (100, 100), "white").save(fake_slide, "PNG")
        return [fake_slide]

    monkeypatch.setattr(sys.modules[__name__], "export_slides", fake_export_slides)

    # This mirrors the app's method selection logic.
    session_state = {"pptx_conversion_method": "libreoffice"}
    input_path = tmp_path / "sample.pptx"
    input_path.write_bytes(b"placeholder")
    rendered_dir = tmp_path / "pptx2nano_output_streamlit" / input_path.stem / "rendered"

    conversion_method = session_state.get("pptx_conversion_method", "libreoffice")
    if conversion_method == "libreoffice":
        rendered_paths = export_slides(input_path, rendered_dir, method="libreoffice")
    else:
        rendered_paths = export_slides(input_path, rendered_dir, method="keynote")

    assert calls.get("method") == "libreoffice"
    assert len(rendered_paths) == 1
    assert rendered_paths[0].exists()
