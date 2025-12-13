#!/usr/bin/env python3
"""precompute_style_examples.py

Purpose
- One-off helper to precompute and cache "style example" images for every built-in style.
- These cached images are used by the Streamlit UI so style previews load instantly.

INPUT FILES (prominent)
- Base image (your reference image to restyle)
  - Provided via: --base-image
  - Example: /Users/macbook2024/Desktop/Generated Image November 24, 2025 - 12_14AM.jpeg

OUTPUT FILES (prominent)
- Cached style example images (PNG)
  - Folder: {repo}/style_examples_cache/
  - Files: {style_key}.png (sanitized)

Version History
- v0.1.0 (2025-12-13): Initial version.

Last Updated
- 2025-12-13

Notes (for a 10th grader)
- We take one "base" image and ask Gemini to redraw it in many different art styles.
- We save each result to disk so the web app can show style thumbnails instantly.

"""

from __future__ import annotations

import argparse
import base64
import concurrent.futures
import mimetypes
import os
import time
from pathlib import Path

import pptx2nano


def _safe_style_filename(style_key: str) -> str:
    safe = "".join(ch for ch in style_key.lower() if ch.isalnum() or ch in ("-", "_"))
    return safe if safe else "style"


def _cache_path_for_style(cache_dir: Path, style_key: str) -> Path:
    return cache_dir / f"{_safe_style_filename(style_key)}.png"


def _build_style_example_prompt(style_key: str, style_desc: str) -> str:
    # We keep this short and explicit. The base image already contains structure/content.
    # We want a consistent "same base, different style" thumbnail.
    return f"""You are an expert presentation designer.

TASK
Redraw the provided image in the visual style '{style_key}'.

STYLE DEFINITION
{style_desc}

CONSTRAINTS (VERY IMPORTANT)
- Preserve the overall layout and content of the provided image.
- Keep the SAME aspect ratio as the input image.
- Do NOT invent new facts. If something is unreadable, keep it minimal.
- Improve readability where possible while preserving meaning.

OUTPUT
- Generate exactly ONE image.
""".strip()


def _generate_one_style(
    *,
    base_image_bytes: bytes,
    base_mime_type: str,
    image_model: str,
    style_key: str,
    style_desc: str,
    cache_dir: Path,
    overwrite: bool,
) -> tuple[str, Path, float, str | None]:
    """Generate one style example and write it to cache.

    Returns:
        (style_key, output_path, seconds, error_message)
    """

    out_path = _cache_path_for_style(cache_dir, style_key)
    if out_path.exists() and not overwrite:
        return style_key, out_path, 0.0, None

    t0 = time.time()
    prompt = _build_style_example_prompt(style_key, style_desc)

    # Create client inside worker to avoid any potential thread-safety issues.
    client = pptx2nano.create_client()

    interaction = client.interactions.create(
        model=image_model,
        input=[
            {
                "type": "image",
                "data": base64.b64encode(base_image_bytes).decode("utf-8"),
                "mime_type": base_mime_type,
            },
            {"type": "text", "text": prompt},
        ],
        response_modalities=["IMAGE"],
    )

    img_bytes, _mime = pptx2nano._extract_image_bytes_from_interaction(interaction)

    cache_dir.mkdir(parents=True, exist_ok=True)
    out_path.write_bytes(img_bytes)

    return style_key, out_path, (time.time() - t0), None


def main() -> None:
    parser = argparse.ArgumentParser(description="Precompute and cache style example images.")
    parser.add_argument(
        "--base-image",
        required=True,
        help="Path to the base image used for all style examples.",
    )
    parser.add_argument(
        "--cache-dir",
        default=str(Path(__file__).resolve().parent / "style_examples_cache"),
        help="Directory to write cached PNG thumbnails (default: ./style_examples_cache).",
    )
    parser.add_argument(
        "--model",
        default=pptx2nano.DEFAULT_IMAGE_MODEL,
        help=f"Gemini image model (default: {pptx2nano.DEFAULT_IMAGE_MODEL})",
    )
    parser.add_argument(
        "--workers",
        type=int,
        default=4,
        help="Number of parallel workers (default: 4)",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite any existing cached images.",
    )

    args = parser.parse_args()

    base_path = Path(args.base_image).expanduser().resolve()
    if not base_path.exists():
        raise FileNotFoundError(f"Base image not found: {base_path}")

    base_bytes = base_path.read_bytes()
    base_mime = mimetypes.guess_type(base_path.name)[0] or "application/octet-stream"

    cache_dir = Path(args.cache_dir).expanduser().resolve()
    image_model = args.model

    style_items = sorted(pptx2nano.BUILTIN_STYLES.items(), key=lambda kv: kv[0])
    total = len(style_items)
    if total == 0:
        print("[INFO] No built-in styles found. Nothing to do.")
        return

    print("[INFO] Precomputing style examples")
    print(f"[INFO] Base image: {base_path}")
    print(f"[INFO] Cache dir:  {cache_dir}")
    print(f"[INFO] Model:      {image_model}")
    print(f"[INFO] Workers:    {args.workers}")
    print(f"[INFO] Overwrite:  {bool(args.overwrite)}")

    t_all = time.time()

    done = 0
    durations: list[float] = []
    errors: list[tuple[str, str]] = []

    with concurrent.futures.ThreadPoolExecutor(max_workers=max(1, args.workers)) as ex:
        futures = []
        for style_key, style_desc in style_items:
            futures.append(
                ex.submit(
                    _generate_one_style,
                    base_image_bytes=base_bytes,
                    base_mime_type=base_mime,
                    image_model=image_model,
                    style_key=style_key,
                    style_desc=style_desc,
                    cache_dir=cache_dir,
                    overwrite=args.overwrite,
                )
            )

        for fut in concurrent.futures.as_completed(futures):
            style_key, out_path, dt, err = fut.result()
            done += 1
            if dt > 0:
                durations.append(dt)

            avg = (sum(durations) / len(durations)) if durations else 0.0
            remaining = max(0, total - done)
            eta = remaining * avg
            pct = int((done / total) * 100)

            if err:
                errors.append((style_key, err))
                print(f"[PROGRESS] {done}/{total} ({pct}%) ERROR style={style_key} ETA={eta:.0f}s")
            else:
                if dt == 0.0:
                    print(
                        f"[PROGRESS] {done}/{total} ({pct}%) cached style={style_key} ETA={eta:.0f}s"
                    )
                else:
                    print(
                        f"[PROGRESS] {done}/{total} ({pct}%) last={dt:.1f}s avg={avg:.1f}s ETA={eta:.0f}s -> {out_path.name}"
                    )

    total_time = time.time() - t_all
    print(f"[DONE] Finished in {total_time:.1f}s")

    if errors:
        print("[WARN] Some styles failed:")
        for style_key, err in errors:
            print(f"- {style_key}: {err}")
        raise SystemExit(1)


if __name__ == "__main__":
    main()
