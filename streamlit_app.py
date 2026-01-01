#!/usr/bin/env python3

"""streamlit_app.py

Purpose
- Local Streamlit web UI for the PowerPoint-to-Nano pipeline.
- Lets you:
  - Pick a PPTX using a native macOS file picker
  - Render slides via Keynote
  - Generate / regenerate one slide at a time in a chosen style (Interactions API)
  - Preview a PDF in-browser before saving
  - Save the PDF next to the input PPTX as: <stem>_nano.pdf (overwrite)

INPUT FILES (prominent)
- PPTX deck selected via the native file picker

OUTPUT FILES (prominent)
- Cached style example images (generated once per style):
  - {repo}/style_examples_cache/<style>.png
- Rendered slide images (Keynote exports):
  - {repo}/pptx2nano_output_streamlit/{deck_stem}/rendered/*.png
- Final PDF (written only when you click Save):
  - {pptx_folder}/{pptx_stem}_nano.pdf

Version History
- v0.1.0 (2025-12-13): Initial Streamlit UI.

Last Updated
- 2025-12-13
"""

from __future__ import annotations

import base64
import subprocess
import time
from io import BytesIO
from pathlib import Path

import streamlit as st
from PIL import Image

import pptx2nano


def _pick_pptx_path() -> Path | None:
    # IMPORTANT: On macOS, Streamlit runs this script in a worker thread.
    # tkinter uses AppKit and will crash with:
    #   "NSWindow should only be instantiated on the main thread!"
    # Instead we use AppleScript's "choose file" via osascript.

    # Return POSIX path directly, and handle user cancel (AppleScript error -128)
    # by returning an empty string.
    applescript = r'''try
  set f to choose file with prompt "Select a PowerPoint (.pptx)"
  return POSIX path of f
on error number -128
  return ""
end try'''

    proc = subprocess.run(
        ["osascript", "-e", applescript],
        check=True,
        capture_output=True,
        text=True,
    )

    selected = (proc.stdout or "").strip()
    if not selected:
        return None
    return Path(selected).expanduser().resolve()


def _ensure_session_defaults() -> None:
    st.session_state.setdefault("pptx_path", None)
    st.session_state.setdefault("rendered_paths", None)
    st.session_state.setdefault("generated_images", {})
    st.session_state.setdefault("interaction_ids", {})
    st.session_state.setdefault("last_preview_pdf", None)
    st.session_state.setdefault("style_example_set", "slide1")


def _reset_for_new_pptx(pptx_path: Path) -> None:
    st.session_state["pptx_path"] = pptx_path
    st.session_state["rendered_paths"] = None
    st.session_state["generated_images"] = {}
    st.session_state["interaction_ids"] = {}
    st.session_state["last_preview_pdf"] = None


def _style_options() -> list[str]:
    return sorted(list(pptx2nano.BUILTIN_STYLES.keys())) + ["custom"]


def _get_selected_style(style_choice: str, custom_style: str) -> str | None:
    if style_choice == "custom":
        s = (custom_style or "").strip()
        return s if s else None
    return style_choice


def _style_example_cache_root() -> Path:
    return Path(__file__).resolve().parent / "style_examples_cache"


def _style_example_cache_base(style_example_set: str) -> Path:
    root = _style_example_cache_root()
    if style_example_set == "slide2":
        return root / "slide2"
    if style_example_set == "slide1":
        return root / "slide1"
    return root


def _style_example_cache_path(style: str, *, style_example_set: str) -> Path:
    safe = "".join(ch for ch in style.lower() if ch.isalnum() or ch in ("-", "_"))
    if not safe:
        safe = "style"
    return _style_example_cache_base(style_example_set) / f"{safe}.png"


def _get_or_create_style_example(style: str, image_model: str, *, style_example_set: str) -> bytes:
    cache_path = _style_example_cache_path(style, style_example_set=style_example_set)
    if cache_path.exists():
        return cache_path.read_bytes()

    # Backward-compatible fallback: if slide1 set is selected but hasn't been copied yet,
    # fall back to the legacy root cache.
    if style_example_set == "slide1":
        legacy_path = _style_example_cache_root() / f"{_safe_style_filename(style)}.png"
        if legacy_path.exists():
            return legacy_path.read_bytes()

    # We do not auto-generate missing thumbnails for slide1/slide2 sets, because
    # those sets are meant to be precomputed from specific deck slides.
    if style_example_set in {"slide1", "slide2"}:
        raise RuntimeError(
            f"Missing cached example for style='{style}' in set='{style_example_set}'. "
            "Generate the cache first (style-sample-generator)."
        )

    # Legacy fallback (root cache): preserve old behavior (generate-on-miss).
    client = pptx2nano.create_client()
    prompt = f"""You are an expert presentation designer.

TASK
Create a single-page example exhibit in the visual style '{style}'.

CONSTRAINTS (VERY IMPORTANT)
- Make it look like a real, clean infographic exhibit.
- Use clear typography and spacing.
- Include a simple chart and a small table with fake placeholder values.
- Do NOT mention this is a sample.

OUTPUT
- Generate exactly ONE image.
""".strip()

    interaction = client.interactions.create(
        model=image_model,
        input=prompt,
        response_modalities=["IMAGE"],
    )

    img_bytes, _mime = pptx2nano._extract_image_bytes_from_interaction(interaction)
    img = Image.open(BytesIO(img_bytes))
    out = BytesIO()
    img.save(out, format="PNG")
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    cache_path.write_bytes(out.getvalue())
    return cache_path.read_bytes()


def _call_slide_image_model(
    slide_index_1based: int,
    rendered_path: Path,
    image_model: str,
    style: str | None,
    total_slides: int,
    previous_interaction_id: str | None,
) -> tuple[bytes, str]:
    client = pptx2nano.create_client()

    with Image.open(rendered_path) as im:
        source_width, source_height = im.size

    if previous_interaction_id and style:
        followup = (
            f"Restyle the slide in a '{style}' visual style. "
            "Preserve all meaningful content from the slide and keep the same aspect ratio. "
            "Generate exactly ONE image."
        )
        interaction = client.interactions.create(
            model=image_model,
            input=followup,
            previous_interaction_id=previous_interaction_id,
            response_modalities=["IMAGE"],
        )
        try:
            img_bytes, _mime = pptx2nano._extract_image_bytes_from_interaction(interaction)
            return img_bytes, interaction.id
        except Exception:
            pass

    prompt = pptx2nano.build_image_model_prompt(
        slide_index_1based=slide_index_1based,
        total_slides=total_slides,
        source_width=source_width,
        source_height=source_height,
        style=style,
    )

    image_bytes = rendered_path.read_bytes()
    mime_type = pptx2nano.mimetypes.guess_type(rendered_path.name)[0] or "application/octet-stream"

    interaction = client.interactions.create(
        model=image_model,
        input=[
            {
                "type": "image",
                "data": base64.b64encode(image_bytes).decode("utf-8"),
                "mime_type": mime_type,
            },
            {"type": "text", "text": prompt},
        ],
        response_modalities=["IMAGE"],
    )

    img_bytes, _mime = pptx2nano._extract_image_bytes_from_interaction(interaction)
    return img_bytes, interaction.id


def _build_pdf_bytes(slide_numbers: list[int]) -> bytes:
    images: list[Image.Image] = []
    for n in slide_numbers:
        img_bytes = st.session_state["generated_images"].get(n)
        if not img_bytes:
            continue
        im = Image.open(BytesIO(img_bytes))
        if im.mode != "RGB":
            im = im.convert("RGB")
        images.append(im)

    if not images:
        raise RuntimeError("No generated slides available to build a PDF.")

    first, rest = images[0], images[1:]
    buf = BytesIO()
    first.save(buf, format="PDF", save_all=True, append_images=rest)
    pdf_bytes = buf.getvalue()

    for im in images:
        try:
            im.close()
        except Exception:
            pass

    return pdf_bytes


def _render_pdf_inline(pdf_bytes: bytes, height: int = 800) -> None:
    # Convert PDF to images for reliable display in Chrome
    try:
        from pdf2image import convert_from_bytes
        import io
        
        # Convert PDF to images using pdf2image
        images = convert_from_bytes(pdf_bytes, dpi=150, size=(800, None))
        
        if images:
            st.info(f"PDF Preview (showing {len(images)} pages):")
            for i, img in enumerate(images, 1):
                if i > 10:  # Limit to first 10 pages
                    st.info(f"... and {len(images) - 10} more pages")
                    break
                st.image(img, caption=f"Page {i}", use_container_width=True)
        else:
            raise Exception("No pages found")
            
    except Exception as e:
        # Ultimate fallback
        st.warning("PDF preview not available in this browser. Use the download button below to view the PDF.")
        st.caption(f"Preview error: {str(e)}")


def main() -> None:
    st.set_page_config(page_title="PowerPoint to Nano", layout="wide")
    _ensure_session_defaults()

    st.markdown(
        """
        <div style="padding: 1rem 0 0.5rem 0;">
          <h1 style="margin: 0;">PowerPoint to Nano</h1>
          <p style="margin: 0.25rem 0 0 0; color: #6b7280;">
            Pick a deck, pick a style, generate one slide at a time, preview the PDF, then save.
          </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    left, right = st.columns([1, 1])

    with left:
        st.subheader("1) Choose a PowerPoint")
        if st.button("Choose .pptx file…", use_container_width=True):
            try:
                picked = _pick_pptx_path()
                if picked is None:
                    st.info("No file selected.")
                else:
                    if picked.suffix.lower() != ".pptx":
                        st.error("Please select a .pptx file.")
                    else:
                        _reset_for_new_pptx(picked)
            except Exception as e:
                st.error(str(e))

        pptx_path = st.session_state.get("pptx_path")
        if pptx_path:
            st.success(f"Selected: {pptx_path}")

        st.subheader("2) Render slides")
        if st.button("Render slides with Keynote", use_container_width=True, disabled=not bool(pptx_path)):
            try:
                with st.spinner("Rendering slides using Keynote…"):
                    out_dir = Path(__file__).resolve().parent / "pptx2nano_output_streamlit"
                    rendered_dir = out_dir / Path(pptx_path).stem / "rendered"
                    rendered_paths = pptx2nano.export_slides_with_keynote(Path(pptx_path), rendered_dir)
                    st.session_state["rendered_paths"] = rendered_paths
                st.success(f"Rendered {len(rendered_paths)} slides.")
            except Exception as e:
                st.error(str(e))

        rendered_paths = st.session_state.get("rendered_paths")
        if rendered_paths:
            st.caption(f"Slides rendered: {len(rendered_paths)}")

    with right:
        st.subheader("Style")
        st.session_state["style_example_set"] = st.selectbox(
            "Example set",
            options=["slide1", "slide2"],
            index=0 if st.session_state.get("style_example_set", "slide1") == "slide1" else 1,
            help="Switch which cached style thumbnails you want to see.",
        )

        style_choice = st.selectbox("Choose a style", _style_options(), index=0)
        custom_style = ""
        if style_choice == "custom":
            custom_style = st.text_input("Custom style", value="")
        style = _get_selected_style(style_choice, custom_style)

        image_model = pptx2nano.DEFAULT_IMAGE_MODEL

        if style_choice != "custom":
            st.caption(pptx2nano.BUILTIN_STYLES.get(style_choice, ""))

        if style:
            try:
                with st.spinner("Loading style example…"):
                    example_bytes = _get_or_create_style_example(
                        style,
                        image_model=image_model,
                        style_example_set=st.session_state.get("style_example_set", "slide1"),
                    )
                st.image(example_bytes, caption=f"Example style: {style}")
            except Exception as e:
                st.warning(f"Could not load style example: {e}")

    st.divider()

    pptx_path = st.session_state.get("pptx_path")
    rendered_paths = st.session_state.get("rendered_paths")
    if not pptx_path or not rendered_paths:
        st.info("Select a PPTX and render slides to begin.")
        return

    total_slides = len(rendered_paths)
    slide_numbers = list(range(1, total_slides + 1))

    st.subheader("3) Generate / Regenerate a slide")
    c1, c2, c3 = st.columns([1, 1, 2])

    with c1:
        gen_target = st.selectbox("Slide", ["ALL"] + slide_numbers, index=1)
        overwrite_existing = st.checkbox(
            "Overwrite slides that were already generated",
            value=False,
            help="When unchecked, Generate ALL will skip slides that already have a generated image so you can keep different styles per slide.",
        )

    # When ALL is selected, we don't show a single rendered_path preview.
    slide_n = None if gen_target == "ALL" else int(gen_target)
    rendered_path = None if slide_n is None else rendered_paths[slide_n - 1]

    with c2:
        if slide_n is None:
            st.caption("Generating ALL will iterate through the full deck.")
        else:
            st.image(str(rendered_path), caption=f"Original slide {slide_n}")

    with c3:
        if slide_n is None:
            if st.button("Generate ALL slides (skip existing unless overwrite checked)", use_container_width=True):
                try:
                    t0 = time.time()
                    durations: list[float] = []
                    prog = st.progress(0)
                    status = st.empty()

                    total = len(slide_numbers)
                    done = 0
                    for n in slide_numbers:
                        already = n in st.session_state["generated_images"]
                        if already and not overwrite_existing:
                            done += 1
                            prog.progress(int((done / total) * 100))
                            status.caption(f"Skipped existing slide {n} ({done}/{total})")
                            continue

                        rp = rendered_paths[n - 1]
                        prev_id = st.session_state["interaction_ids"].get(n)

                        status.caption(f"Generating slide {n} ({done + 1}/{total})")
                        t_slide = time.time()
                        img_bytes, new_id = _call_slide_image_model(
                            slide_index_1based=n,
                            rendered_path=rp,
                            image_model=image_model,
                            style=style,
                            total_slides=total_slides,
                            previous_interaction_id=prev_id,
                        )
                        dt = time.time() - t_slide
                        durations.append(dt)

                        st.session_state["generated_images"][n] = img_bytes
                        st.session_state["interaction_ids"][n] = new_id
                        st.session_state["last_preview_pdf"] = None

                        done += 1
                        avg = sum(durations) / len(durations) if durations else 0.0
                        remaining = max(0, total - done)
                        eta = remaining * avg
                        prog.progress(int((done / total) * 100))
                        status.caption(
                            f"Generated slide {n} ({done}/{total}) | last={dt:.1f}s avg={avg:.1f}s ETA={eta:.0f}s"
                        )

                    prog.progress(100)
                    total_time = time.time() - t0
                    status.success(f"Done. Generate ALL finished in {total_time:.1f}s")
                except Exception as e:
                    st.error(str(e))
        else:
            prev_id = st.session_state["interaction_ids"].get(slide_n)
            if st.button("Generate / Regenerate this slide", use_container_width=True):
                try:
                    prog = st.progress(0)
                    with st.spinner("Calling Gemini…"):
                        prog.progress(20)
                        img_bytes, new_id = _call_slide_image_model(
                            slide_index_1based=slide_n,
                            rendered_path=rendered_path,
                            image_model=image_model,
                            style=style,
                            total_slides=total_slides,
                            previous_interaction_id=prev_id,
                        )
                        prog.progress(90)
                        st.session_state["generated_images"][slide_n] = img_bytes
                        st.session_state["interaction_ids"][slide_n] = new_id
                        st.session_state["last_preview_pdf"] = None
                        prog.progress(100)
                    st.success(f"Generated slide {slide_n}.")
                except Exception as e:
                    st.error(str(e))

            gen_bytes = st.session_state["generated_images"].get(slide_n)
            if gen_bytes:
                st.image(gen_bytes, caption=f"Generated slide {slide_n}")
            else:
                st.caption("No generated image for this slide yet.")

    st.divider()

    st.subheader("4) Preview and Save PDF")

    preview_mode = st.selectbox(
        "Include slides",
        options=["ALL", "Custom"],
        index=0,
    )
    if preview_mode == "ALL":
        include_slides_sorted = slide_numbers
        st.caption("Preview will include all slides (that have generated images).")
    else:
        include_slides = st.multiselect(
            "Slides to include in PDF",
            options=slide_numbers,
            default=slide_numbers,
        )
        include_slides_sorted = sorted(include_slides)

    b1, b2 = st.columns([1, 1])

    with b1:
        if st.button("Build PDF preview", use_container_width=True):
            try:
                with st.spinner("Building PDF preview…"):
                    pdf_bytes = _build_pdf_bytes(include_slides_sorted)
                    st.session_state["last_preview_pdf"] = pdf_bytes
                st.success("Preview ready.")
            except Exception as e:
                st.error(str(e))

    with b2:
        preview_bytes = st.session_state.get("last_preview_pdf")
        can_save = bool(preview_bytes) and bool(pptx_path)
        if st.button("Save PDF next to PPTX (overwrite)", use_container_width=True, disabled=not can_save):
            try:
                pptx_p = Path(pptx_path)
                out_path = pptx_p.parent / f"{pptx_p.stem}_nano.pdf"
                out_path.write_bytes(preview_bytes)
                st.success(f"Saved: {out_path}")
            except Exception as e:
                st.error(str(e))

    preview_bytes = st.session_state.get("last_preview_pdf")
    if preview_bytes:
        _render_pdf_inline(preview_bytes)
        st.download_button(
            label="Download preview PDF",
            data=preview_bytes,
            file_name=f"{Path(pptx_path).stem}_nano.pdf",
            mime="application/pdf",
            use_container_width=True,
        )


if __name__ == "__main__":
    main()
