"""pptx_converter.py

Unified PPTX to image converter supporting both Keynote (macOS) and LibreOffice (cross-platform).
This allows testing both approaches before committing to one.

Usage:
    # Using Keynote (current, macOS only)
    from pptx_converter import export_slides
    paths = export_slides("deck.pptx", output_dir, method="keynote")
    
    # Using LibreOffice (cross-platform, cloud-compatible)
    paths = export_slides("deck.pptx", output_dir, method="libreoffice")
    
    # Auto-detect (uses Keynote on macOS if available, falls back to LibreOffice)
    paths = export_slides("deck.pptx", output_dir, method="auto")
"""

from __future__ import annotations

import re
import subprocess
import tempfile
import shutil
from pathlib import Path
from typing import Iterable, Literal


def _extract_last_int(text: str) -> int | None:
    """Return the last integer found in text."""
    matches = re.findall(r"(\d+)", text)
    if not matches:
        return None
    try:
        return int(matches[-1])
    except ValueError:
        return None


def _sort_slide_files(paths: Iterable[Path]) -> list[Path]:
    """Sort slide image files in natural slide order."""
    def sort_key(p: Path):
        n = _extract_last_int(p.name)
        return (n is None, n if n is not None else p.name.lower())
    return sorted(list(paths), key=sort_key)


def _find_libreoffice() -> str | None:
    """Find LibreOffice executable path."""
    for cmd in ["soffice", "libreoffice"]:
        if shutil.which(cmd):
            return cmd
    # Common macOS installation paths
    mac_paths = [
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/libreoffice",
    ]
    for path in mac_paths:
        if Path(path).exists():
            return path
    return None


def export_slides_with_keynote(pptx_path: Path, rendered_dir: Path) -> list[Path]:
    """Render PPTX slides into PNG images using Keynote via AppleScript."""
    rendered_dir.mkdir(parents=True, exist_ok=True)

    applescript = r'''on run argv
  set inputPosix to item 1 of argv
  set outputPosix to item 2 of argv

  set inputFile to POSIX file inputPosix as alias
  set outputFolder to POSIX file outputPosix as alias

  set inputName to name of (info for inputFile)
  set baseName to inputName
  try
    if (offset of "." in inputName) > 1 then
      set baseName to text 1 thru ((offset of "." in inputName) - 1) of inputName
    end if
  end try

  tell application "Keynote"
    activate

    set oldDocIds to {}
    repeat with d in documents
      try
        set end of oldDocIds to (id of d)
      end try
    end repeat

    open inputFile

    set theDoc to missing value
    repeat with i from 1 to 600
      repeat with d in documents
        try
          if (id of d) is not in oldDocIds then
            set theDoc to d
            exit repeat
          end if
        end try
      end repeat

      if theDoc is missing value then
        repeat with d in documents
          try
            if (name of d) contains baseName then
              set theDoc to d
              exit repeat
            end if
          end try
        end repeat
      end if

      if theDoc is not missing value then exit repeat
      delay 0.5
    end repeat

    if theDoc is missing value then
      error "Timed out waiting for Keynote to open/import the PPTX."
    end if

    export theDoc as slide images to outputFolder with properties {image format:PNG}
    close theDoc saving no
  end tell
end run
'''

    try:
        subprocess.run(
            [
                "osascript",
                "-e",
                applescript,
                str(pptx_path),
                str(rendered_dir.resolve()),
            ],
            check=True,
            capture_output=True,
            text=True,
        )
    except subprocess.CalledProcessError as e:
        stderr = (e.stderr or "").strip()
        stdout = (e.stdout or "").strip()
        msg = "Keynote export failed via osascript.\n"
        if stdout:
            msg += f"osascript stdout:\n{stdout}\n"
        if stderr:
            msg += f"osascript stderr:\n{stderr}\n"
        raise RuntimeError(msg) from e

    pngs = list(rendered_dir.glob("*.png"))
    if not pngs:
        jpgs = list(rendered_dir.glob("*.jpg")) + list(rendered_dir.glob("*.jpeg"))
        if jpgs:
            return _sort_slide_files(jpgs)
        raise RuntimeError(f"Keynote export produced no images in: {rendered_dir}")

    return _sort_slide_files(pngs)


def export_slides_with_libreoffice(pptx_path: Path, rendered_dir: Path, dpi: int = 150) -> list[Path]:
    """Render PPTX slides into PNG images using LibreOffice.
    
    Args:
        pptx_path: Path to the PPTX file
        rendered_dir: Directory to save rendered images
        dpi: Resolution for export (default 150, higher = better quality but slower)
    
    Returns:
        List of PNG paths in slide order
    """
    rendered_dir.mkdir(parents=True, exist_ok=True)
    
    libreoffice = _find_libreoffice()
    if not libreoffice:
        raise RuntimeError(
            "LibreOffice not found. Install it:\n"
            "  macOS: brew install --cask libreoffice\n"
            "  Linux: sudo apt install libreoffice-impress\n"
            "  Or download from: https://www.libreoffice.org/download/"
        )
    
    # LibreOffice exports to a temp directory, then we move files
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        
        # Convert PPTX to PDF first (more reliable than direct PNG export)
        pdf_path = tmp_path / f"{pptx_path.stem}.pdf"
        
        try:
            # Step 1: Convert PPTX to PDF using LibreOffice
            result = subprocess.run(
                [
                    libreoffice,
                    "--headless",
                    "--convert-to", "pdf",
                    "--outdir", str(tmp_path),
                    str(pptx_path)
                ],
                capture_output=True,
                text=True,
                timeout=120  # 2 minute timeout
            )
            
            if result.returncode != 0:
                stderr = result.stderr.strip()
                raise RuntimeError(f"LibreOffice PPTX to PDF conversion failed:\n{stderr}")
            
            # Find the generated PDF
            pdfs = list(tmp_path.glob("*.pdf"))
            if not pdfs:
                raise RuntimeError("LibreOffice did not produce a PDF file")
            
            pdf_file = pdfs[0]
            
            # Step 2: Convert PDF to PNG images using pdftoppm or pdf2image
            try:
                # Try pdftoppm first (faster, no Python deps)
                _convert_pdf_to_png_pdftoppm(pdf_file, rendered_dir, dpi)
            except (RuntimeError, FileNotFoundError):
                # Fallback to pdf2image
                _convert_pdf_to_png_pdf2image(pdf_file, rendered_dir, dpi)
                
        except subprocess.TimeoutExpired:
            raise RuntimeError("LibreOffice conversion timed out (took > 2 minutes)")
    
    # Collect and sort the output files
    pngs = list(rendered_dir.glob("*.png"))
    if not pngs:
        raise RuntimeError(f"LibreOffice export produced no images in: {rendered_dir}")
    
    return _sort_slide_files(pngs)


def _convert_pdf_to_png_pdftoppm(pdf_path: Path, output_dir: Path, dpi: int) -> None:
    """Convert PDF to PNG using pdftoppm (poppler-utils)."""
    if not shutil.which("pdftoppm"):
        raise FileNotFoundError("pdftoppm not found")
    
    result = subprocess.run(
        [
            "pdftoppm",
            "-png",
            "-r", str(dpi),
            "-progress",
            str(pdf_path),
            str(output_dir / "slide")
        ],
        capture_output=True,
        text=True
    )
    
    if result.returncode != 0:
        raise RuntimeError(f"pdftoppm failed: {result.stderr}")
    
    # pdftoppm names files like slide-1.png, slide-2.png
    # Rename to consistent format
    for f in output_dir.glob("slide-*.png"):
        match = re.search(r"slide-(\d+)\.png$", f.name)
        if match:
            slide_num = int(match.group(1))
            new_name = f"slide_{slide_num:03d}.png"
            f.rename(output_dir / new_name)


def _convert_pdf_to_png_pdf2image(pdf_path: Path, output_dir: Path, dpi: int) -> None:
    """Convert PDF to PNG using pdf2image (Python library)."""
    try:
        from pdf2image import convert_from_path
    except ImportError:
        raise RuntimeError("pdf2image not installed. Run: pip install pdf2image")
    
    images = convert_from_path(pdf_path, dpi=dpi)
    
    for i, image in enumerate(images, start=1):
        output_path = output_dir / f"slide_{i:03d}.png"
        image.save(output_path, "PNG")


def export_slides(
    pptx_path: Path | str,
    rendered_dir: Path | str,
    method: Literal["auto", "keynote", "libreoffice"] = "auto",
    dpi: int = 150
) -> list[Path]:
    """Export PPTX slides to PNG images using specified method.
    
    Args:
        pptx_path: Path to the PPTX file
        rendered_dir: Directory to save rendered images
        method: Conversion method - "keynote", "libreoffice", or "auto"
        dpi: DPI for LibreOffice method (default 150)
    
    Returns:
        List of PNG file paths in slide order
    """
    pptx_path = Path(pptx_path)
    rendered_dir = Path(rendered_dir)
    
    if not pptx_path.exists():
        raise FileNotFoundError(f"PPTX file not found: {pptx_path}")
    
    if method == "auto":
        # On macOS, try Keynote first
        import platform
        if platform.system() == "Darwin":
            try:
                return export_slides_with_keynote(pptx_path, rendered_dir)
            except RuntimeError:
                # Fall back to LibreOffice
                return export_slides_with_libreoffice(pptx_path, rendered_dir, dpi)
        else:
            method = "libreoffice"
    
    if method == "keynote":
        return export_slides_with_keynote(pptx_path, rendered_dir)
    elif method == "libreoffice":
        return export_slides_with_libreoffice(pptx_path, rendered_dir, dpi)
    else:
        raise ValueError(f"Unknown method: {method}")


def compare_methods(pptx_path: Path | str, output_base_dir: Path | str | None = None) -> dict:
    """Compare Keynote vs LibreOffice output for the same PPTX.
    
    This is useful for testing quality differences before switching.
    
    Returns:
        Dict with comparison results
    """
    pptx_path = Path(pptx_path)
    
    if output_base_dir is None:
        output_base_dir = Path("pptx_converter_test")
    else:
        output_base_dir = Path(output_base_dir)
    
    results = {
        "input_file": str(pptx_path),
        "keynote": {"success": False, "output_dir": None, "slide_count": 0, "error": None},
        "libreoffice": {"success": False, "output_dir": None, "slide_count": 0, "error": None},
    }
    
    # Test Keynote (macOS only)
    import platform
    if platform.system() == "Darwin":
        keynote_dir = output_base_dir / "keynote_output"
        try:
            keynote_slides = export_slides_with_keynote(pptx_path, keynote_dir)
            results["keynote"]["success"] = True
            results["keynote"]["output_dir"] = str(keynote_dir)
            results["keynote"]["slide_count"] = len(keynote_slides)
            results["keynote"]["slides"] = [str(s) for s in keynote_slides]
        except Exception as e:
            results["keynote"]["error"] = str(e)
    else:
        results["keynote"]["error"] = "Keynote only available on macOS"
    
    # Test LibreOffice
    libreoffice_dir = output_base_dir / "libreoffice_output"
    try:
        libreoffice_slides = export_slides_with_libreoffice(pptx_path, libreoffice_dir)
        results["libreoffice"]["success"] = True
        results["libreoffice"]["output_dir"] = str(libreoffice_dir)
        results["libreoffice"]["slide_count"] = len(libreoffice_slides)
        results["libreoffice"]["slides"] = [str(s) for s in libreoffice_slides]
    except Exception as e:
        results["libreoffice"]["error"] = str(e)
    
    return results


if __name__ == "__main__":
    # Simple CLI for testing
    import sys
    import json
    
    if len(sys.argv) < 2:
        print("Usage: python pptx_converter.py <pptx_file> [method] [output_dir]")
        print("  method: keynote, libreoffice, or compare (default: compare)")
        sys.exit(1)
    
    pptx_file = Path(sys.argv[1])
    method = sys.argv[2] if len(sys.argv) > 2 else "compare"
    output_dir = Path(sys.argv[3]) if len(sys.argv) > 3 else Path("pptx_converter_test")
    
    if method == "compare":
        print(f"Comparing conversion methods for: {pptx_file}")
        print("-" * 60)
        results = compare_methods(pptx_file, output_dir)
        print(json.dumps(results, indent=2))
        
        # Print summary
        print("\n" + "=" * 60)
        print("SUMMARY")
        print("=" * 60)
        
        for m in ["keynote", "libreoffice"]:
            r = results[m]
            status = "✅ SUCCESS" if r["success"] else "❌ FAILED"
            print(f"\n{m.upper()}: {status}")
            if r["success"]:
                print(f"  Slides: {r['slide_count']}")
                print(f"  Output: {r['output_dir']}")
            else:
                print(f"  Error: {r['error']}")
    else:
        output_path = output_dir / f"{method}_output"
        print(f"Converting with {method}: {pptx_file}")
        slides = export_slides(pptx_file, output_path, method=method)
        print(f"Success! Generated {len(slides)} slides:")
        for s in slides:
            print(f"  - {s}")
