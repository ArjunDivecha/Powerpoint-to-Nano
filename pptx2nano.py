#!/usr/bin/env python3
"""pptx2nano.py

Purpose
- Turn a PowerPoint deck (.pptx) into a brand-new "Nano Banana" style deck by:
  1) Rendering each slide to an image using Keynote (macOS)
  2) Sending each slide image to Gemini Image (Nano Banana) to generate a redesigned graphic
  3) Combining all generated slide images into a single multi-page PDF

INPUT FILES (prominent)
- PPTX deck (provided via CLI)
  - Example: /Users/you/path/to/deck.pptx
  - Format: Microsoft PowerPoint .pptx

OUTPUT FILES (prominent)
- Output folder (default: ./pptx2nano_output)
  - Rendered slide images (from Keynote)
    - {out_dir}/{deck_stem}/rendered/*.png
  - Generated slide images (from Gemini image model)
    - {out_dir}/{deck_stem}/generated/slide_###.png
  - Final multi-page PDF (all generated slides)
    - {out_dir}/{deck_stem}.pdf

Version History
- v0.1.0 (2025-12-13): Initial version (PPTX -> Keynote render -> Gemini image per slide -> PDF)

Last Updated
- 2025-12-13

Notes (for a 10th grader)
- Keynote is used because it can "open" PowerPoint files and export each slide as a picture.
- Then we send each slide picture to Gemini and ask it to redraw the slide as a clean infographic.
- Finally we put all the new slide pictures into a single PDF.

Requirements
- macOS with Keynote installed
- A Gemini API key in .env (GEMINI_API_KEY=...)

"""

from __future__ import annotations

import argparse
import concurrent.futures
import json
import mimetypes
import os
import re
import subprocess
import sys
import time
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Iterable

from dotenv import find_dotenv, load_dotenv
from google import genai
from google.genai import types
from PIL import Image
import base64


DEFAULT_IMAGE_MODEL = "gemini-3-pro-image-preview"


BUILTIN_STYLES: dict[str, str] = {
    "lego": "Bright primary colors, blocky shapes, toy-like 3D appearance, snap-together aesthetic",
    "ghibli": "Hand-drawn feel, soft watercolor palette, whimsical organic shapes, Studio Ghibli anime aesthetic",
    "cyberpunk": "Neon colors, dark background, glowing elements, futuristic tech aesthetic, grid patterns",
    "minimalist": "Clean white/gray palette, thin lines, lots of whitespace, simple sans-serif fonts",
    "blueprint": "Technical drawing style, blue background, white lines, grid paper, architectural feel",
    "hand-drawn": "Sketchy lines, imperfect shapes, notebook paper feel, casual doodle aesthetic",
    # Data & Business Visualization (Nano Banana Pro)
    "minimalist-infographic": "Swiss-style infographic: ample whitespace, clean neutral palette, simple sans-serif typography, uncluttered data summary",
    "corporate-dashboard": "SaaS/analytics dashboard look: dark mode UI, structured charts, high-contrast neon accents (electric blue / emerald)",
    "timeline-roadmap": "Roadmap timeline with milestone nodes (horizontal/vertical), progressive flow, clear phases and dates",
    "technical-blueprint": "Engineering schematic: white lines on blue grid, right-angle connectors, geometric precision (architectural/technical)",
    "whiteboard-strategy": "Hand-drawn sketchnote on whiteboard: marker textures, arrows, friendly low-fidelity strategy sketches",
    "comparison-split": "Symmetric split layout (Option A vs Option B): clear side-by-side contrast, distinct color coding, balanced structure",
    "process-flow": "Flowchart/system diagram: rectangles for steps, diamonds for decisions, directional arrows, if/then logic",
    "multi-layer-venn": "3+ set Venn with transparent overlaps (multiply), labeled intersections, clean hierarchy",
    # Artistic & Illustration Styles
    "lego-brick-builder": "Plastic brick build: bright primary colors, snap-together LEGO-like forms, toy proportions",
    "lego-diorama": "A photorealistic isometric diorama made entirely of LEGO bricks, organized into distinct color-coded vertical lanes on a large baseplate. Miniature LEGO machinery, conveyor belts, robotic arms, and server racks connected by arrows. LEGO minifigures posed working in each zone. Small white tiles display legible text labels and filenames. Macro photography, tilt-shift effect with shallow depth of field, bright studio lighting, sharp plastic textures, high contrast, 8k resolution, organized industrial aesthetic",
    "studio-ghibli-anime": "Soft watercolor anime: whimsical organic shapes, lush green backgrounds, 80s/90s Studio Ghibli vibe",
    "ghibli-pride": "Studio Ghibli anime aesthetic with MAXIMUM rainbow vibrancy: explosive pride flag colors (hot pink, electric purple, bright cyan, sunshine yellow), joyful celebratory energy, sparkles and glitter effects, whimsical flowing ribbons, radiant gradients, love and joy radiating from every element, soft watercolor meets neon brilliance, euphoric and fabulous",
    "neon-cyberpunk": "Neon night cyberpunk: pink/purple/cyan LEDs, rain-slick streets, glowing tech, futuristic cityscape",
    "3d-claymation": "Clay/plasticine look: rounded edges, subtle fingerprint textures, miniature tilt-shift lighting",
    "pixel-art-8bit": "Retro 8-bit pixel art: low-res jagged edges, limited NES/SNES-like palette, game UI sensibility",
    "comic-book-hero": "Comic ink style: heavy outlines, cross-hatching shadows, Ben-Day dots, dramatic angles",
    "vintage-travel-poster": "Art Deco travel poster: flat bold colors, geometric composition, integrated blocky title text",
    "graffiti-street-art": "Spray paint on concrete: wild-style lettering, drips, vibrant chaotic color schemes",
    "paper-cutout-origami": "Layered paper craft: colored sheets, drop-shadows for depth, cut edges, origami-like forms",
    # Photorealistic & Cinematic Styles
    "cinematic-realism": "High-end cinema camera realism: controlled lighting (golden/blue hour), shallow depth of field, bokeh",
    "analog-film": "35mm film emulation: subtle grain, light leaks, vignette, softer nostalgic color grading",
    "product-hero-shot": "Studio product hero shot: pristine background, controlled lighting, sharp reflections, e-commerce look",
    "macro-close-up": "Extreme macro detail: visible textures/fibers, very shallow DOF, strongly blurred background",
    "knitted-doll-amigurumi": "Crocheted yarn doll: visible fibers, button eyes, soft fuzzy lighting, handcrafted texture",
    # Text-Heavy Formats
    "editorial-magazine": "Magazine layout: central hero image, bold serif headline, multi-column body text, editorial grid",
    "chalkboard-menu": "Chalkboard/restaurant menu: slate texture, hand lettering/calligraphy, chalk strokes",
    "instructional-manual": "IKEA-style manual: black-and-white line art, simple figures, arrows, assembly instructions",
}


@dataclass(frozen=True)
class SlideJob:
    index_1based: int
    total_slides: int
    rendered_path: Path
    generated_path: Path


def _parse_slides_arg(value: str) -> list[int]:
    """Parse a slide selection string into sorted unique 1-based slide indices.

    Supported formats:
    - "3"            -> [3]
    - "3,4,5"        -> [3,4,5]
    - "3-5"          -> [3,4,5]
    - "1,3-5,9"      -> [1,3,4,5,9]
    """

    raw = (value or "").strip()
    if not raw:
        raise ValueError("--slides cannot be empty")

    out: set[int] = set()
    parts = [p.strip() for p in raw.split(",") if p.strip()]
    for part in parts:
        if "-" in part:
            a, b = [x.strip() for x in part.split("-", 1)]
            start = int(a)
            end = int(b)
            if start < 1 or end < 1:
                raise ValueError("Slide numbers must be >= 1")
            if end < start:
                raise ValueError(f"Invalid slide range: {part}")
            for i in range(start, end + 1):
                out.add(i)
        else:
            i = int(part)
            if i < 1:
                raise ValueError("Slide numbers must be >= 1")
            out.add(i)

    return sorted(out)


def _slides_label(slides: list[int]) -> str:
    """Create a compact label for filenames."""

    if not slides:
        return "none"
    if slides == list(range(slides[0], slides[-1] + 1)):
        return f"{slides[0]:03d}-{slides[-1]:03d}"
    return "_".join(f"{s:03d}" for s in slides)


def _print_styles() -> None:
    """Print built-in styles and exit."""

    print("Available built-in styles:")
    for name, desc in BUILTIN_STYLES.items():
        print(f"- {name}: {desc}")


def _human_seconds(seconds: float) -> str:
    if seconds < 60:
        return f"{seconds:.1f}s"
    minutes = int(seconds // 60)
    rem = seconds - minutes * 60
    return f"{minutes}m {rem:.0f}s"


def _extract_last_int(text: str) -> int | None:
    """Return the last integer found in text.

    Keynote export filenames often look like:
    - _render_test4.011.png
    We want the *last* number (11), not the earlier test folder number (4).
    """

    matches = re.findall(r"(\d+)", text)
    if not matches:
        return None
    try:
        return int(matches[-1])
    except ValueError:
        return None


def _load_env() -> None:
    """Load environment variables from .env.

    Search order:
    - Current working directory (walk up) via find_dotenv(usecwd=True)
    - This script's folder (Powerpoint-to-Nano/)
    - Parent folder (Powerpoint to Nano/) to support the user's existing layout
    """

    env_from_cwd = find_dotenv(usecwd=True)
    if env_from_cwd:
        load_dotenv(env_from_cwd, override=False)

    script_dir = Path(__file__).resolve().parent
    load_dotenv(script_dir / ".env", override=False)
    load_dotenv(script_dir.parent / ".env", override=False)


def create_client() -> genai.Client:
    """Create a Gemini client using GEMINI_API_KEY from environment (.env supported)."""
    _load_env()
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise RuntimeError(
            "GEMINI_API_KEY is not set. Put it in a .env file (project folder or parent folder)."
        )
    client = genai.Client(api_key=api_key)
    # Interactions API is only available in newer google-genai versions.
    if not hasattr(client, "interactions"):
        raise RuntimeError(
            "This script requires the Interactions API (google-genai>=1.55.0). "
            "Upgrade your environment with: pip install -r requirements.txt --upgrade"
        )
    return client


def export_slides_with_keynote(pptx_path: Path, rendered_dir: Path) -> list[Path]:
    """Render PPTX slides into PNG images using Keynote via AppleScript.

    This function:
    - Opens the PPTX in Keynote
    - Exports the presentation as "slide images" (PNG) into rendered_dir
    - Closes the document without saving

    Returns
    - A list of PNG paths in slide order.

    Raises
    - RuntimeError if export produced no PNGs.
    """

    rendered_dir.mkdir(parents=True, exist_ok=True)

    # AppleScript: use argv to avoid quoting issues.
    # IMPORTANT: macOS may prompt you to allow automation the first time.
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

    -- Capture currently open documents so we can identify the new one.
    set oldDocIds to {}
    repeat with d in documents
      try
        set end of oldDocIds to (id of d)
      end try
    end repeat

    open inputFile

    -- PPTX import can take time; wait up to ~5 minutes.
    set theDoc to missing value
    repeat with i from 1 to 600
      -- First preference: find a newly opened document (id not in oldDocIds)
      repeat with d in documents
        try
          if (id of d) is not in oldDocIds then
            set theDoc to d
            exit repeat
          end if
        end try
      end repeat

      -- Fallback: match by document name (Keynote often drops .pptx extension)
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
      error "Timed out waiting for Keynote to open/import the PPTX. Try opening the PPTX once manually in Keynote (and closing other docs) and re-run."
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
        msg += (
            "\nCommon fixes:\n"
            "- Open Keynote once manually and accept any first-run dialogs\n"
            "- In System Settings -> Privacy & Security -> Automation, allow the calling app to control Keynote\n"
            "- If the PPTX is in Dropbox/iCloud, try copying it to Desktop and re-run\n"
        )
        raise RuntimeError(msg) from e

    pngs = list(rendered_dir.glob("*.png"))
    if not pngs:
        # Some Keynote versions may export JPG depending on settings.
        jpgs = list(rendered_dir.glob("*.jpg")) + list(rendered_dir.glob("*.jpeg"))
        if jpgs:
            return _sort_slide_files(jpgs)
        raise RuntimeError(
            f"Keynote export produced no images in: {rendered_dir}. "
            "Check Keynote permissions and try again."
        )

    return _sort_slide_files(pngs)


def _sort_slide_files(paths: Iterable[Path]) -> list[Path]:
    """Sort slide image files in natural slide order.

    Keynote typically exports filenames like:
    - Slide 1.png, Slide 2.png, ...

    We sort by the last integer in the filename.
    """

    def sort_key(p: Path):
        n = _extract_last_int(p.name)
        return (n is None, n if n is not None else p.name.lower())

    return sorted(list(paths), key=sort_key)


def build_image_model_prompt(
    slide_index_1based: int,
    total_slides: int,
    source_width: int,
    source_height: int,
    style: str | None,
) -> str:
    """Prompt to redraw a slide as a clean infographic while preserving content."""

    style_instruction = ""
    if style:
        style_instruction = f"""

VISUAL STYLE REQUIREMENT
Create the redesigned slide in a '{style}' visual style. Apply this style consistently across:
- Overall aesthetic and color palette
- Typography and text rendering
- Icons and visual elements
- Shapes and layout

Examples of how to interpret styles:
- 'lego': Bright primary colors, blocky shapes, toy-like 3D appearance, snap-together aesthetic
- 'ghibli': Hand-drawn feel, soft watercolor palette, whimsical organic shapes, Studio Ghibli anime aesthetic
- 'cyberpunk': Neon colors, dark background, glowing elements, futuristic tech aesthetic, grid patterns
- 'minimalist': Clean white/gray palette, thin lines, lots of whitespace, simple sans-serif fonts
- 'blueprint': Technical drawing style, blue background, white lines, grid paper, architectural feel
- 'hand-drawn': Sketchy lines, imperfect shapes, notebook paper feel, casual doodle aesthetic
"""

    return f"""
You are an expert presentation designer.

TASK
Redesign the provided slide image as a clean, professional infographic slide.

CONSTRAINTS (VERY IMPORTANT)
- Keep the SAME aspect ratio as the input slide.
- Target pixel size should closely match the input: {source_width} x {source_height}.
- Preserve ALL meaningful content from the slide (text, numbers, labels, relationships).
- Do NOT add speaker notes.
- Do NOT invent new facts. If something is unreadable, keep it minimal and do not hallucinate.
- Improve layout, spacing, and readability.

SLIDE CONTEXT
- Slide {slide_index_1based} of {total_slides}

OUTPUT
- Generate exactly ONE image.
{style_instruction}
""".strip()


def _extract_image_bytes_from_interaction(interaction) -> tuple[bytes, str]:
    """Extract image bytes from an Interactions API response.

    Per Interactions API docs, interaction.outputs contains typed outputs.
    Image outputs have:
    - type == "image"
    - data as base64
    - mime_type like "image/png"
    """

    outputs = getattr(interaction, "outputs", None)
    if not outputs:
        raise RuntimeError("Interactions API returned no outputs.")

    for out in outputs:
        out_type = getattr(out, "type", None)
        if out_type == "image":
            mime_type = getattr(out, "mime_type", None) or "application/octet-stream"
            data_b64 = getattr(out, "data", None)
            if not data_b64:
                raise RuntimeError("Image output had no data.")
            return base64.b64decode(data_b64), mime_type

    raise RuntimeError("No image output found in Interactions API outputs.")


def generate_one_slide(
    job: SlideJob,
    out_dir: Path,
    image_model: str,
    style: str | None,
) -> dict:
    """Worker function: generate one redesigned slide image."""

    t0 = time.time()

    # Load source slide image to get size (helps preserve original aspect ratio).
    with Image.open(job.rendered_path) as im:
        source_width, source_height = im.size

    prompt = build_image_model_prompt(
        slide_index_1based=job.index_1based,
        total_slides=job.total_slides,
        source_width=source_width,
        source_height=source_height,
        style=style,
    )

    # We want the *real* image bytes sent to the model.
    image_bytes = job.rendered_path.read_bytes()
    mime_type = mimetypes.guess_type(job.rendered_path.name)[0] or "application/octet-stream"

    # Interactions API requires google-genai>=1.55.0.
    # We pass the slide image as base64 and request IMAGE modality.
    client = create_client()
    interaction = client.interactions.create(
        model=image_model,
        input=[
            {"type": "image", "data": base64.b64encode(image_bytes).decode("utf-8"), "mime_type": mime_type},
            {"type": "text", "text": prompt},
        ],
        response_modalities=["IMAGE"],
    )

    out_image_bytes, _out_mime = _extract_image_bytes_from_interaction(interaction)

    job.generated_path.parent.mkdir(parents=True, exist_ok=True)
    img = Image.open(BytesIO(out_image_bytes))
    img.save(str(job.generated_path))

    dt = time.time() - t0
    return {
        "slide_index": job.index_1based,
        "rendered_path": str(job.rendered_path),
        "generated_path": str(job.generated_path),
        "seconds": dt,
    }


def images_to_pdf(image_paths: list[Path], pdf_path: Path):
    """Combine images into a single multi-page PDF."""

    if not image_paths:
        raise ValueError("No images provided to build PDF")

    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    pil_images = []
    for p in image_paths:
        im = Image.open(p)
        if im.mode != "RGB":
            im = im.convert("RGB")
        pil_images.append(im)

    first, rest = pil_images[0], pil_images[1:]
    first.save(str(pdf_path), save_all=True, append_images=rest)

    for im in pil_images:
        try:
            im.close()
        except Exception:
            pass


def main():
    parser = argparse.ArgumentParser(
        description="Convert a PPTX into Nano Banana-style slides and export a PDF."
    )
    parser.add_argument("pptx_path", help="Path to a .pptx file")
    parser.add_argument(
        "--out-dir",
        default="pptx2nano_output",
        help="Output directory (default: pptx2nano_output)",
    )
    parser.add_argument(
        "--image-model",
        default=DEFAULT_IMAGE_MODEL,
        help=f"Gemini image model to use (default: {DEFAULT_IMAGE_MODEL})",
    )
    parser.add_argument(
        "--style",
        default=None,
        help=(
            "Visual style to apply. Built-in styles: "
            + ", ".join(sorted(BUILTIN_STYLES.keys()))
            + ". You may also provide a custom style name like 'retro' or 'corporate'."
        ),
    )
    parser.add_argument(
        "--list-styles",
        action="store_true",
        help="Print built-in styles and exit",
    )
    parser.add_argument(
        "--workers",
        type=int,
        default=4,
        help="Number of parallel workers for slide generation (default: 4)",
    )
    parser.add_argument(
        "--slides",
        default=None,
        help="Only process these slide numbers (e.g., '3,4,5' or '3-5')",
    )
    parser.add_argument(
        "--max-slides",
        type=int,
        default=None,
        help="Only process the first N slides (useful for quick testing)",
    )
    parser.add_argument(
        "--write-run-log",
        action="store_true",
        help="Write a JSON log describing what was generated",
    )
    args = parser.parse_args()

    if args.list_styles:
        _print_styles()
        return

    pptx_path = Path(args.pptx_path).expanduser().resolve()
    if not pptx_path.exists():
        raise FileNotFoundError(f"PPTX not found: {pptx_path}")
    if pptx_path.suffix.lower() != ".pptx":
        raise ValueError("Input must be a .pptx file")

    out_dir = Path(args.out_dir).expanduser().resolve()
    deck_stem = pptx_path.stem

    rendered_dir = out_dir / deck_stem / "rendered"
    generated_dir = out_dir / deck_stem / "generated"
    pdf_path = out_dir / f"{deck_stem}.pdf"

    selected_slides: list[int] | None = None
    if args.slides is not None:
        selected_slides = _parse_slides_arg(args.slides)

    if selected_slides is not None and args.max_slides is not None:
        raise ValueError("Use either --slides or --max-slides, not both")

    # 1) Render slides using Keynote
    print(f"[INFO] Rendering slides with Keynote: {pptx_path}")
    rendered_paths = export_slides_with_keynote(pptx_path, rendered_dir)
    full_total = len(rendered_paths)

    if args.max_slides is not None:
        if args.max_slides < 1:
            raise ValueError("--max-slides must be >= 1")
        rendered_paths = rendered_paths[: args.max_slides]

    if selected_slides is not None:
        if not selected_slides:
            raise ValueError("--slides produced an empty selection")
        if selected_slides[-1] > full_total:
            raise ValueError(
                f"Requested slide {selected_slides[-1]} but deck only has {full_total} slides"
            )
        rendered_paths = [rendered_paths[i - 1] for i in selected_slides]
        pdf_path = out_dir / f"{deck_stem}_slides_{_slides_label(selected_slides)}.pdf"

    total = len(rendered_paths)
    print(f"[INFO] Rendered {total} slides to: {rendered_dir}")

    # 2) Create jobs
    jobs: list[SlideJob] = []
    for local_i, rendered_path in enumerate(rendered_paths, start=1):
        source_slide_index = local_i
        if selected_slides is not None:
            source_slide_index = selected_slides[local_i - 1]
        generated_path = generated_dir / f"slide_{source_slide_index:03d}.png"
        jobs.append(
            SlideJob(
                index_1based=source_slide_index,
                total_slides=full_total if selected_slides is not None else total,
                rendered_path=rendered_path,
                generated_path=generated_path,
            )
        )

    # 3) Generate new slide images (parallel)
    print(
        f"[INFO] Generating redesigned slides with {args.image_model} (workers={args.workers})"
    )

    t_all = time.time()
    results: list[dict] = []

    # Track progress
    done = 0
    durations: list[float] = []

    # NOTE: ThreadPool is appropriate here because the work is mostly I/O (network to Gemini).
    with concurrent.futures.ThreadPoolExecutor(max_workers=max(1, args.workers)) as ex:
        future_to_job = {
            ex.submit(generate_one_slide, job, out_dir, args.image_model, args.style): job
            for job in jobs
        }

        for fut in concurrent.futures.as_completed(future_to_job):
            job = future_to_job[fut]
            try:
                res = fut.result()
            except Exception as e:
                print(
                    f"[ERROR] Slide {job.index_1based}/{total} failed: {e}",
                    file=sys.stderr,
                )
                raise

            results.append(res)
            done += 1
            durations.append(float(res.get("seconds", 0.0)))

            avg = sum(durations) / len(durations) if durations else 0.0
            remaining = max(0, total - done)
            eta = remaining * avg
            pct = (done / total) * 100

            print(
                f"[PROGRESS] {done}/{total} ({pct:.0f}%) "
                f"last={_human_seconds(durations[-1])} avg={_human_seconds(avg)} "
                f"ETA={_human_seconds(eta)} -> {job.generated_path.name}"
            )

    # Ensure results are ordered by slide index
    results_sorted = sorted(results, key=lambda r: int(r["slide_index"]))
    generated_paths = [Path(r["generated_path"]) for r in results_sorted]

    # 4) Build PDF
    print(f"[INFO] Building multi-page PDF: {pdf_path}")
    images_to_pdf(generated_paths, pdf_path)

    total_time = time.time() - t_all
    print(f"[DONE] Created: {pdf_path}")
    print(f"[DONE] Total time: {_human_seconds(total_time)}")

    if args.write_run_log:
        log_suffix = ""
        if selected_slides is not None:
            log_suffix = f"_slides_{_slides_label(selected_slides)}"
        log_path = out_dir / f"{deck_stem}{log_suffix}_run_log.json"
        payload = {
            "pptx_path": str(pptx_path),
            "rendered_dir": str(rendered_dir),
            "generated_dir": str(generated_dir),
            "pdf_path": str(pdf_path),
            "image_model": args.image_model,
            "style": args.style,
            "workers": args.workers,
            "max_slides": args.max_slides,
            "slides_selected": selected_slides,
            "slides": results_sorted,
            "total_seconds": total_time,
        }
        log_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        print(f"[INFO] Wrote run log: {log_path}")


if __name__ == "__main__":
    main()
