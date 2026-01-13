#!/usr/bin/env python3
"""copy_lego_diorama.py

Purpose
- Helper script to copy a LEGO diorama image to style example cache directories.
- Used to prepopulate the style cache with a custom example image.

Usage:
    python copy_lego_diorama.py /path/to/lego-diorama-image.png

INPUT FILES (prominent)
- LEGO diorama image file
  - Any image format supported by Pillow (PNG, JPG, etc.)
  - Example: /Users/you/Desktop/lego-diorama.png

OUTPUT FILES (prominent)
- Cached style example images:
  - {repo}/style_examples_cache/slide1/lego-diorama.png
  - {repo}/style_examples_cache/slide2/lego-diorama.png

Version History
- v0.1.0 (2025-12-13): Initial version.

Last Updated
- 2025-12-13

Notes (for a 10th grader)
- This script copies the same image to two cache directories.
- The Streamlit app uses these cached images to show style previews instantly.
- Without caching, the app would need to generate style previews on-demand, which is slower.
"""

import sys
import shutil
from pathlib import Path

def main():
    if len(sys.argv) != 2:
        print("Usage: python copy_lego_diorama.py /path/to/lego-diorama-image.png")
        sys.exit(1)
    
    source_image = Path(sys.argv[1])
    
    if not source_image.exists():
        print(f"Error: Image file not found: {source_image}")
        sys.exit(1)
    
    # Get the script directory
    script_dir = Path(__file__).parent
    
    # Define target paths
    slide1_path = script_dir / "style_examples_cache" / "slide1" / "lego-diorama.png"
    slide2_path = script_dir / "style_examples_cache" / "slide2" / "lego-diorama.png"
    
    # Copy to both locations
    print(f"Copying {source_image} to:")
    print(f"  - {slide1_path}")
    shutil.copy2(source_image, slide1_path)
    print(f"  - {slide2_path}")
    shutil.copy2(source_image, slide2_path)
    
    print("\n✓ Successfully copied LEGO diorama image to both slide1 and slide2 caches!")
    print("The 'lego-diorama' style will now show this image as the example.")

if __name__ == "__main__":
    main()
