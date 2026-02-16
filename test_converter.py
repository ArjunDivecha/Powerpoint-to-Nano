#!/usr/bin/env python3
"""Test script to compare Keynote vs LibreOffice PPTX conversion.

This script lets you test LibreOffice conversion before committing to it.
Run this to see side-by-side comparison of both methods.

Setup:
    1. Install LibreOffice:
       macOS: brew install --cask libreoffice
       Or download from: https://www.libreoffice.org/download/
    
    2. Install poppler (for pdftoppm - faster PDF to PNG):
       macOS: brew install poppler
    
    3. Run the test:
       python test_converter.py /path/to/your/deck.pptx

Output:
    - Creates pptx_converter_test/ directory
    - keynote_output/ - slides from Keynote
    - libreoffice_output/ - slides from LibreOffice
    - You can visually compare the results
"""

import sys
from pathlib import Path

# Add the current directory to path for imports
sys.path.insert(0, str(Path(__file__).parent))

from pptx_converter import compare_methods


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        print("\nUsage: python test_converter.py <pptx_file>")
        sys.exit(1)
    
    pptx_file = Path(sys.argv[1])
    
    if not pptx_file.exists():
        print(f"❌ File not found: {pptx_file}")
        sys.exit(1)
    
    if not pptx_file.suffix.lower() == ".pptx":
        print(f"⚠️  Warning: File doesn't have .pptx extension: {pptx_file}")
    
    print(f"🧪 Testing PPTX conversion: {pptx_file.name}")
    print("=" * 70)
    
    # Run comparison
    results = compare_methods(pptx_file)
    
    # Print detailed results
    print("\n📊 RESULTS:")
    print("-" * 70)
    
    for method in ["keynote", "libreoffice"]:
        r = results[method]
        print(f"\n{method.upper()}:")
        
        if r["success"]:
            print(f"  ✅ Success - {r['slide_count']} slides generated")
            print(f"  📁 Output: {r['output_dir']}")
            
            # Show first few slide files
            slides = r.get("slides", [])
            if slides:
                print(f"  📄 Files:")
                for s in slides[:3]:
                    print(f"     - {Path(s).name}")
                if len(slides) > 3:
                    print(f"     ... and {len(slides) - 3} more")
        else:
            print(f"  ❌ Failed")
            print(f"  💥 Error: {r['error']}")
    
    # Quality comparison tips
    print("\n" + "=" * 70)
    print("🔍 NEXT STEPS - Compare Quality:")
    print("-" * 70)
    print("""
1. Open both output directories:
   
   open pptx_converter_test/keynote_output/
   open pptx_converter_test/libreoffice_output/

2. Compare the same slide number from both methods

3. Check for:
   - Font rendering differences
   - Image quality and resolution
   - Layout shifts or alignment issues
   - Missing elements (charts, shapes, etc.)
   - Color accuracy

4. If LibreOffice quality is acceptable, you can switch!

To use LibreOffice in your app:
   from pptx_converter import export_slides
   slides = export_slides("deck.pptx", output_dir, method="libreoffice")
""")


if __name__ == "__main__":
    main()
