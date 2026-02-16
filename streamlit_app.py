#!/usr/bin/env python3

"""streamlit_app.py

Purpose
- Local Streamlit web UI for the PowerPoint-to-Nano pipeline.
- Lets you:
  - Pick various input files using native macOS file picker (PPTX, PDF, GIF, TXT, DOCX, images, Markdown)
  - Render slides/pages via Keynote (PPTX) or conversion (other formats)
  - Generate / regenerate one slide at a time in a chosen style (Interactions API)
  - Auto-detect text-heavy slides for enhanced text extraction accuracy
  - Preview a PDF in-browser before saving
  - Save the PDF next to the input file as: <stem>_nano.pdf (overwrite)

INPUT FILES (prominent)
- Multiple formats supported via native file picker:
  - PPTX (PowerPoint presentations)
  - PDF (multi-page documents)
  - GIF (animated graphics)
  - TXT (plain text files)
  - DOCX (Word documents)
  - Images (PNG, JPG, JPEG, WebP, etc.)
  - Markdown (.md files)

OUTPUT FILES (prominent)
- Cached style example images (generated once per style):
  - {repo}/style_examples_cache/slide1/<style>.png
  - {repo}/style_examples_cache/slide2/<style>.png
- Rendered slide images (Keynote exports or conversions):
  - {repo}/pptx2nano_output_streamlit/{deck_stem}/rendered/*.png
- Final PDF (written only when you click Save):
  - {input_folder}/{input_stem}_nano.pdf

Version History
- v0.1.0 (2025-12-13): Initial Streamlit UI with PPTX support only.
- v0.2.0 (2026-01-03): Added multi-format support, text extraction, auto-detection, custom styles

Last Updated
- 2026-01-03

Features
- 33+ built-in styles including ghibli-pride
- Custom style support with on-the-fly generation
- Text extraction for PPTX files with auto-detection
- Text deduplication to prevent duplicate titles
- Fresh regenerations (always from original, not previous generation)
- Debug logging for text extraction decisions
"""

from __future__ import annotations

import base64
import subprocess
import time
from io import BytesIO
from pathlib import Path

import streamlit as st
from PIL import Image
from pptx import Presentation

import pptx2nano
from pptx_converter import export_slides


def _pick_input_file() -> Path | None:
    # IMPORTANT: On macOS, Streamlit runs this script in a worker thread.
    # tkinter uses AppKit and will crash with:
    #   "NSWindow should only be instantiated on the main thread!"
    # Instead we use AppleScript's "choose file" via osascript.

    # Return POSIX path directly, and handle user cancel (AppleScript error -128)
    # by returning an empty string.
    applescript = r'''try
  set f to choose file with prompt "Select a file (.pptx, .pdf, .gif, .txt, .docx, .png, .jpg, .jpeg, .webp, .md)"
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
    st.session_state.setdefault("input_path", None)
    st.session_state.setdefault("input_type", None)
    st.session_state.setdefault("rendered_paths", None)
    st.session_state.setdefault("generated_images", {})
    st.session_state.setdefault("interaction_ids", {})
    st.session_state.setdefault("last_preview_pdf", None)
    st.session_state.setdefault("style_example_set", "slide1")
    st.session_state.setdefault("selected_style", "lego")
    st.session_state.setdefault("pptx_conversion_method", "libreoffice")  # Default to LibreOffice
    st.session_state.setdefault("selected_slide", 1)


def _reset_for_new_input(input_path: Path, input_type: str) -> None:
    st.session_state["input_path"] = input_path
    st.session_state["input_type"] = input_type
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
        safe = "".join(ch for ch in style.lower() if ch.isalnum() or ch in ("-", "_"))
        if not safe:
            safe = "style"
        legacy_path = _style_example_cache_root() / f"{safe}.png"
        if legacy_path.exists():
            return legacy_path.read_bytes()

    # We do not auto-generate missing thumbnails for slide1/slide2 sets for BUILTIN styles,
    # because those sets are meant to be precomputed from specific deck slides.
    # However, for custom styles, we allow on-the-fly generation.
    is_builtin_style = style in pptx2nano.BUILTIN_STYLES
    if style_example_set in {"slide1", "slide2"} and is_builtin_style:
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


def _process_pdf_input(pdf_path: Path, output_dir: Path) -> list[Path]:
    """Extract pages from PDF as images."""
    from pdf2image import convert_from_path
    
    output_dir.mkdir(parents=True, exist_ok=True)
    images = convert_from_path(pdf_path, dpi=150)
    
    rendered_paths = []
    for i, img in enumerate(images, 1):
        out_path = output_dir / f"page_{i:03d}.png"
        img.save(out_path, "PNG")
        rendered_paths.append(out_path)
    
    return rendered_paths


def _process_gif_input(gif_path: Path, output_dir: Path) -> list[Path]:
    """Extract frames from GIF as images."""
    output_dir.mkdir(parents=True, exist_ok=True)
    
    with Image.open(gif_path) as img:
        rendered_paths = []
        frame_num = 0
        
        try:
            while True:
                img.seek(frame_num)
                out_path = output_dir / f"frame_{frame_num + 1:03d}.png"
                # Convert to RGB if necessary (GIFs can be in palette mode)
                frame = img.convert("RGB")
                frame.save(out_path, "PNG")
                rendered_paths.append(out_path)
                frame_num += 1
        except EOFError:
            pass
    
    return rendered_paths


def _process_text_input(text_path: Path, output_dir: Path) -> list[Path]:
    """Convert text file to images (one page per paragraph or chunk)."""
    from PIL import ImageDraw, ImageFont
    
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Read text file
    text_content = text_path.read_text(encoding='utf-8')
    
    # Split into chunks (by double newline or every ~500 chars)
    paragraphs = [p.strip() for p in text_content.split('\n\n') if p.strip()]
    if not paragraphs:
        # Fallback: split by single newlines
        paragraphs = [p.strip() for p in text_content.split('\n') if p.strip()]
    
    rendered_paths = []
    
    # Create images for each paragraph
    for i, para in enumerate(paragraphs, 1):
        # Create a white image
        img_width, img_height = 1920, 1080
        img = Image.new('RGB', (img_width, img_height), color='white')
        draw = ImageDraw.Draw(img)
        
        # Try to use a system font, fallback to default
        try:
            font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 40)
        except:
            font = ImageFont.load_default()
        
        # Add text with word wrapping
        margin = 100
        max_width = img_width - 2 * margin
        y_text = margin
        
        words = para.split()
        lines = []
        current_line = []
        
        for word in words:
            test_line = ' '.join(current_line + [word])
            bbox = draw.textbbox((0, 0), test_line, font=font)
            if bbox[2] - bbox[0] <= max_width:
                current_line.append(word)
            else:
                if current_line:
                    lines.append(' '.join(current_line))
                current_line = [word]
        
        if current_line:
            lines.append(' '.join(current_line))
        
        # Draw lines
        for line in lines:
            draw.text((margin, y_text), line, fill='black', font=font)
            y_text += 60
        
        # Save image
        out_path = output_dir / f"text_{i:03d}.png"
        img.save(out_path, "PNG")
        rendered_paths.append(out_path)
    
    return rendered_paths


def _process_docx_input(docx_path: Path, output_dir: Path) -> list[Path]:
    """Extract text from DOCX and convert to images."""
    from docx import Document
    from PIL import ImageDraw, ImageFont
    
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Read DOCX file
    doc = Document(docx_path)
    
    # Extract paragraphs
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    
    if not paragraphs:
        # Create a single empty page
        paragraphs = ["(Empty document)"]
    
    rendered_paths = []
    
    # Create images for each paragraph
    for i, para in enumerate(paragraphs, 1):
        img_width, img_height = 1920, 1080
        img = Image.new('RGB', (img_width, img_height), color='white')
        draw = ImageDraw.Draw(img)
        
        try:
            font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 40)
        except:
            font = ImageFont.load_default()
        
        margin = 100
        max_width = img_width - 2 * margin
        y_text = margin
        
        words = para.split()
        lines = []
        current_line = []
        
        for word in words:
            test_line = ' '.join(current_line + [word])
            bbox = draw.textbbox((0, 0), test_line, font=font)
            if bbox[2] - bbox[0] <= max_width:
                current_line.append(word)
            else:
                if current_line:
                    lines.append(' '.join(current_line))
                current_line = [word]
        
        if current_line:
            lines.append(' '.join(current_line))
        
        for line in lines:
            draw.text((margin, y_text), line, fill='black', font=font)
            y_text += 60
        
        out_path = output_dir / f"docx_{i:03d}.png"
        img.save(out_path, "PNG")
        rendered_paths.append(out_path)
    
    return rendered_paths


def _process_image_input(image_path: Path, output_dir: Path) -> list[Path]:
    """Process single image file (PNG, JPG, JPEG, WebP, etc.)."""
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Open and convert image to RGB if necessary
    with Image.open(image_path) as img:
        # Convert to RGB if needed (handles RGBA, palette mode, etc.)
        if img.mode not in ('RGB', 'L'):
            img = img.convert('RGB')
        
        # Save as PNG
        out_path = output_dir / "image_001.png"
        img.save(out_path, "PNG")
    
    return [out_path]


def _process_markdown_input(md_path: Path, output_dir: Path) -> list[Path]:
    """Convert Markdown file to images."""
    import markdown
    from PIL import ImageDraw, ImageFont
    
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Read markdown file
    md_content = md_path.read_text(encoding='utf-8')
    
    # Convert markdown to plain text (strip HTML tags)
    # For simplicity, we'll just use the raw markdown text
    # You could use markdown library to convert to HTML then strip tags
    
    # Split by headers or double newlines
    sections = []
    current_section = []
    
    for line in md_content.split('\n'):
        if line.startswith('#') or (not line.strip() and current_section):
            if current_section:
                sections.append('\n'.join(current_section))
                current_section = []
            if line.strip():
                current_section.append(line)
        else:
            if line.strip():
                current_section.append(line)
    
    if current_section:
        sections.append('\n'.join(current_section))
    
    if not sections:
        sections = ["(Empty markdown)"]
    
    rendered_paths = []
    
    for i, section in enumerate(sections, 1):
        img_width, img_height = 1920, 1080
        img = Image.new('RGB', (img_width, img_height), color='white')
        draw = ImageDraw.Draw(img)
        
        try:
            font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 40)
            header_font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 50)
        except:
            font = ImageFont.load_default()
            header_font = font
        
        margin = 100
        max_width = img_width - 2 * margin
        y_text = margin
        
        # Process each line
        for line in section.split('\n'):
            # Check if it's a header
            is_header = line.startswith('#')
            if is_header:
                line = line.lstrip('#').strip()
                current_font = header_font
            else:
                current_font = font
            
            # Word wrap
            words = line.split()
            lines = []
            current_line = []
            
            for word in words:
                test_line = ' '.join(current_line + [word])
                bbox = draw.textbbox((0, 0), test_line, font=current_font)
                if bbox[2] - bbox[0] <= max_width:
                    current_line.append(word)
                else:
                    if current_line:
                        lines.append(' '.join(current_line))
                    current_line = [word]
            
            if current_line:
                lines.append(' '.join(current_line))
            
            for wrapped_line in lines:
                draw.text((margin, y_text), wrapped_line, fill='black', font=current_font)
                y_text += 70 if is_header else 60
        
        out_path = output_dir / f"md_{i:03d}.png"
        img.save(out_path, "PNG")
        rendered_paths.append(out_path)
    
    return rendered_paths


def _extract_text_from_pptx_slide(pptx_path: Path, slide_index: int) -> str:
    """Extract text directly from a PPTX slide using python-pptx."""
    try:
        prs = Presentation(pptx_path)
        if slide_index >= len(prs.slides):
            return ""
        
        slide = prs.slides[slide_index]
        text_parts = []
        
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                text_parts.append(shape.text.strip())
        
        return "\n".join(text_parts)
    except Exception as e:
        st.warning(f"Could not extract text from slide {slide_index + 1}: {e}")
        return ""


def _clean_text_with_gemini(raw_text: str) -> str:
    """Use gemini-3-flash-preview to clean and structure extracted text."""
    if not raw_text.strip():
        return ""
    
    try:
        client = pptx2nano.create_client()
        response = client.models.generate_content(
            model="gemini-3-flash-preview",
            contents=f"""Clean and structure the following slide text. Fix any formatting issues, 
preserve all content, and maintain the original meaning. Return only the cleaned text:

{raw_text}"""
        )
        return response.text.strip()
    except Exception as e:
        st.warning(f"Could not clean text with gemini-3-flash-preview: {e}")
        return raw_text


def _should_use_text_extraction(rendered_path: Path, pptx_path: Path | None, slide_index: int) -> bool:
    """Auto-detect if a slide needs text extraction for better accuracy.
    
    Criteria:
    - Slide has lots of text (detected from PPTX)
    - Image has small dimensions (likely small fonts)
    """
    if not pptx_path or not pptx_path.exists():
        return False
    
    try:
        # Check 1: Extract text and see if it's text-heavy
        extracted_text = _extract_text_from_pptx_slide(pptx_path, slide_index)
        word_count = len(extracted_text.split())
        
        # Check 2: Check image dimensions (small images = small fonts)
        with Image.open(rendered_path) as im:
            width, height = im.size
            is_small = width < 1200 or height < 900
        
        # Debug logging
        st.write(f"🔍 Text extraction check for slide {slide_index + 1}:")
        st.write(f"  - Word count: {word_count}")
        st.write(f"  - Image size: {width}x{height}")
        st.write(f"  - Is small: {is_small}")
        st.write(f"  - Will use text extraction: {word_count > 50 or is_small}")
        
        # Use text extraction if:
        # - More than 20 words (lowered threshold for better detection)
        # - OR small image dimensions (likely small fonts)
        return word_count > 20 or is_small
    except Exception as e:
        st.warning(f"Text extraction check failed: {e}")
        return False


def _call_slide_image_model(
    slide_index_1based: int,
    rendered_path: Path,
    image_model: str,
    style: str | None,
    total_slides: int,
    previous_interaction_id: str | None,
    pptx_path: Path | None = None,
) -> tuple[bytes, str]:
    client = pptx2nano.create_client()

    with Image.open(rendered_path) as im:
        source_width, source_height = im.size

    # Don't use previous_interaction_id for restyling - always start fresh from original slide
    # This ensures dramatic style changes are properly applied

    # Auto-detect if text extraction is needed
    use_text_extraction = _should_use_text_extraction(rendered_path, pptx_path, slide_index_1based - 1)
    
    if use_text_extraction and pptx_path:
        # Extract and clean text with gemini-3-flash-preview
        raw_text = _extract_text_from_pptx_slide(pptx_path, slide_index_1based - 1)
        
        # Remove duplicate lines from extracted text
        lines = raw_text.split('\n')
        seen = set()
        unique_lines = []
        for line in lines:
            line_stripped = line.strip()
            if line_stripped and line_stripped not in seen:
                seen.add(line_stripped)
                unique_lines.append(line)
        deduplicated_text = '\n'.join(unique_lines)
        
        cleaned_text = _clean_text_with_gemini(deduplicated_text)
        
        # Build enhanced prompt with extracted text
        prompt = f"""You are an expert presentation designer.

TASK
Redesign the provided slide image as a clean, professional infographic slide.

EXTRACTED TEXT CONTENT (use this for accurate text - the image may have small/blurry text):
{cleaned_text}

CONSTRAINTS (VERY IMPORTANT)
- Use the EXTRACTED TEXT above for accurate text content - DO NOT REPEAT text elements
- If the same text appears multiple times in the extracted content, only show it ONCE in the final design
- Use the IMAGE for layout, visual elements, charts, and spatial relationships
- Keep the SAME aspect ratio as the input slide
- Target pixel size: {source_width} x {source_height}
- Apply the '{style}' visual style
- Do NOT add speaker notes
- Do NOT invent new facts
- Do NOT duplicate titles, headings, or any text elements
- Improve layout, spacing, and readability

SLIDE CONTEXT
- Slide {slide_index_1based} of {total_slides}

OUTPUT
- Generate exactly ONE image with NO repeated text.
""".strip()
        
        st.info(f"🔍 Auto-detected text-heavy slide {slide_index_1based} - using enhanced text extraction for better accuracy")
    else:
        # Standard prompt without text extraction
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
            Pick a file (PPTX, PDF, GIF, TXT, DOCX, PNG, JPG, MD), pick a style, generate pages, preview the PDF, then save.
          </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    left, right = st.columns([1, 1])

    with left:
        st.subheader("1) Choose Input File")
        if st.button("Choose file…", use_container_width=True):
            try:
                picked = _pick_input_file()
                if picked is None:
                    st.info("No file selected.")
                else:
                    suffix = picked.suffix.lower()
                    supported_formats = [".pptx", ".pdf", ".gif", ".txt", ".docx", ".png", ".jpg", ".jpeg", ".webp", ".md"]
                    if suffix not in supported_formats:
                        st.error(f"Unsupported file type. Please select: {', '.join(supported_formats)}")
                    else:
                        # Normalize image extensions
                        if suffix in [".png", ".jpg", ".jpeg", ".webp"]:
                            input_type = "image"
                        else:
                            input_type = suffix[1:]  # Remove the dot
                        _reset_for_new_input(picked, input_type)
            except Exception as e:
                st.error(str(e))

        input_path = st.session_state.get("input_path")
        input_type = st.session_state.get("input_type")
        if input_path:
            st.success(f"Selected: {input_path}")
            st.caption(f"Type: {input_type.upper()}")

        # PPTX conversion info (LibreOffice is now the only method)
        if input_type == "pptx":
            st.caption("📄 Using LibreOffice for PPTX conversion")

        st.subheader("2) Process Input")
        # Render slides label
        pptx_label = "Render slides with LibreOffice" if input_type == "pptx" else "Process file"
        
        process_label = {
            "pptx": pptx_label,
            "pdf": "Extract PDF pages",
            "gif": "Extract GIF frames",
            "txt": "Convert text to images",
            "docx": "Extract DOCX text to images",
            "image": "Process image",
            "md": "Convert Markdown to images",
        }.get(input_type, "Process file")
        
        if st.button(process_label, use_container_width=True, disabled=not bool(input_path)):
            try:
                out_dir = Path(__file__).resolve().parent / "pptx2nano_output_streamlit"
                rendered_dir = out_dir / Path(input_path).stem / "rendered"
                
                if input_type == "pptx":
                    with st.spinner("Rendering slides using LibreOffice…"):
                        rendered_paths = export_slides(Path(input_path), rendered_dir, method="libreoffice")
                elif input_type == "pdf":
                    with st.spinner("Extracting PDF pages…"):
                        rendered_paths = _process_pdf_input(Path(input_path), rendered_dir)
                elif input_type == "gif":
                    with st.spinner("Extracting GIF frames…"):
                        rendered_paths = _process_gif_input(Path(input_path), rendered_dir)
                elif input_type == "txt":
                    with st.spinner("Converting text to images…"):
                        rendered_paths = _process_text_input(Path(input_path), rendered_dir)
                elif input_type == "docx":
                    with st.spinner("Extracting DOCX text…"):
                        rendered_paths = _process_docx_input(Path(input_path), rendered_dir)
                elif input_type == "image":
                    with st.spinner("Processing image…"):
                        rendered_paths = _process_image_input(Path(input_path), rendered_dir)
                elif input_type == "md":
                    with st.spinner("Converting Markdown…"):
                        rendered_paths = _process_markdown_input(Path(input_path), rendered_dir)
                else:
                    st.error(f"Unsupported file type: {input_type}")
                    return
                
                st.session_state["rendered_paths"] = rendered_paths
                st.success(f"Processed {len(rendered_paths)} pages/slides/frames.")
            except Exception as e:
                st.error(str(e))

        rendered_paths = st.session_state.get("rendered_paths")
        if rendered_paths:
            st.caption(f"Pages rendered: {len(rendered_paths)}")

    with left:
        st.subheader("Style Example")
        st.session_state["style_example_set"] = st.selectbox(
            "Example set",
            options=["slide1", "slide2"],
            index=0 if st.session_state.get("style_example_set", "slide1") == "slide1" else 1,
            help="Switch which cached style thumbnails you want to see.",
        )

        # Style selection setup
        style_options = _style_options()
        current_style = st.session_state.get("selected_style", "lego")
        
        style_choice = current_style
        custom_style = ""
        if style_choice == "custom":
            custom_style = st.text_input("Custom style", value="")
        style = _get_selected_style(style_choice, custom_style)

        image_model = pptx2nano.DEFAULT_IMAGE_MODEL

        if style_choice != "custom":
            st.caption(pptx2nano.BUILTIN_STYLES.get(style_choice, ""))

        # Display style example image
        if style:
            try:
                with st.spinner("Loading style example…"):
                    example_bytes = _get_or_create_style_example(
                        style,
                        image_model=image_model,
                        style_example_set=st.session_state.get("style_example_set", "slide1"),
                    )
                st.image(example_bytes, caption=f"Example: {style}", use_container_width=True)
            except Exception as e:
                st.warning(f"Could not load style example: {e}")

    with right:
        st.subheader("Style")
        st.caption("Choose a style:")
        cols = st.columns(2)
        
        for i, style_name in enumerate(style_options):
            with cols[i % 2]:
                # Use buttons instead of radio buttons for better control
                button_type = "primary" if style_name == current_style else "secondary"
                if st.button(style_name, key=f"style_btn_{i}", type=button_type, use_container_width=True):
                    st.session_state["selected_style"] = style_name
                    st.rerun()

    st.divider()

    input_path = st.session_state.get("input_path")
    rendered_paths = st.session_state.get("rendered_paths")
    if not input_path or not rendered_paths:
        st.info("Select a file and process it to begin.")
        return

    total_slides = len(rendered_paths)
    slide_numbers = list(range(1, total_slides + 1))

    st.subheader("3) Generate / Regenerate a slide")
    c1, c2, c3 = st.columns([1, 1, 2])

    with c1:
        # Get current slide index to maintain scroll position
        slide_options = ["ALL"] + slide_numbers
        current_slide = st.session_state.get("selected_slide", 1)
        try:
            if current_slide == "ALL":
                current_slide_index = 0
            else:
                current_slide_index = slide_options.index(current_slide)
        except ValueError:
            current_slide_index = 1
        
        gen_target = st.selectbox("Slide", slide_options, index=current_slide_index)
        st.session_state["selected_slide"] = gen_target
        
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
                            pptx_path=input_path if input_type == "pptx" else None,
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
                            pptx_path=input_path if input_type == "pptx" else None,
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
        can_save = bool(preview_bytes) and bool(input_path)
        if st.button("Save PDF next to input file (overwrite)", use_container_width=True, disabled=not can_save):
            try:
                input_p = Path(input_path)
                out_path = input_p.parent / f"{input_p.stem}_nano.pdf"
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
            file_name=f"{Path(input_path).stem}_nano.pdf",
            mime="application/pdf",
            use_container_width=True,
        )


if __name__ == "__main__":
    main()
