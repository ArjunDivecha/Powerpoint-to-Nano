#!/usr/bin/env python3
"""generate_style_examples.py

Purpose
- Generate style examples for ALL builtin styles using the India Mom PDF sample.
- Creates cached examples in both slide1 and slide2 directories.
- Useful for precomputing all style examples at once.

INPUT FILES (prominent)
- Sample slide image:
  - {repo}/style_samples/sample_slide.png
  - Extracted from India Mom PDF (first page)
  - Dimensions: 3245x2534 pixels

OUTPUT FILES (prominent)
- Cached style example images (33+ styles):
  - {repo}/style_examples_cache/slide1/<style>.png
  - {repo}/style_examples_cache/slide2/<style>.png

Version History
- v0.1.0 (2026-01-03): Created for bulk style example generation.

Last Updated
- 2026-01-03

Notes (for a 10th grader)
- This script generates examples for ALL built-in styles at once.
- It uses the India Mom PDF sample as the base image for all styles.
- The Gemini API redraws this sample in each different style.
- This can take 5-10 minutes as it calls the API for each style.
- Results are cached so the Streamlit app can show them instantly.
"""

from pathlib import Path
import pptx2nano
from PIL import Image
import base64

def generate_all_style_examples():
    """Generate style examples for all builtin styles using the India Mom sample."""
    
    # Load the sample slide
    sample_path = Path(__file__).parent / "style_samples" / "sample_slide.png"
    if not sample_path.exists():
        print(f"Error: Sample slide not found at {sample_path}")
        return
    
    print(f"Using sample slide: {sample_path}")
    
    # Get image dimensions
    with Image.open(sample_path) as im:
        width, height = im.size
    print(f"Sample dimensions: {width}x{height}")
    
    # Create cache directories
    cache_root = Path(__file__).parent / "style_examples_cache"
    slide1_dir = cache_root / "slide1"
    slide2_dir = cache_root / "slide2"
    slide1_dir.mkdir(parents=True, exist_ok=True)
    slide2_dir.mkdir(parents=True, exist_ok=True)
    
    # Get all builtin styles
    styles = list(pptx2nano.BUILTIN_STYLES.keys())
    print(f"\nGenerating examples for {len(styles)} styles...")
    
    # Create Gemini client
    client = pptx2nano.create_client()
    
    # Read sample image
    image_bytes = sample_path.read_bytes()
    
    for i, style in enumerate(styles, 1):
        print(f"\n[{i}/{len(styles)}] Generating: {style}")
        
        try:
            # Build prompt
            prompt = pptx2nano.build_image_model_prompt(
                slide_index_1based=1,
                total_slides=1,
                source_width=width,
                source_height=height,
                style=style,
            )
            
            # Call Gemini
            interaction = client.interactions.create(
                model=pptx2nano.DEFAULT_IMAGE_MODEL,
                input=[
                    {
                        "type": "image",
                        "data": base64.b64encode(image_bytes).decode("utf-8"),
                        "mime_type": "image/png",
                    },
                    {"type": "text", "text": prompt},
                ],
                response_modalities=["IMAGE"],
            )
            
            # Extract image
            img_bytes, _ = pptx2nano._extract_image_bytes_from_interaction(interaction)
            
            # Save to both slide1 and slide2 caches
            safe_name = style.replace(" ", "_").replace("/", "-")
            slide1_path = slide1_dir / f"{safe_name}.png"
            slide2_path = slide2_dir / f"{safe_name}.png"
            
            slide1_path.write_bytes(img_bytes)
            slide2_path.write_bytes(img_bytes)
            
            print(f"  ✓ Saved to {slide1_path.name}")
            
        except Exception as e:
            print(f"  ✗ Error: {e}")
            continue
    
    print(f"\n✓ Done! Generated examples in:")
    print(f"  - {slide1_dir}")
    print(f"  - {slide2_dir}")

if __name__ == "__main__":
    generate_all_style_examples()
