#!/usr/bin/env python3
"""Test LibreOffice conversion through the same code path as Streamlit app."""

import sys
from pathlib import Path
import tempfile
import time

# Add current directory to path
sys.path.insert(0, str(Path(__file__).parent))

from pptx_converter import export_slides, _find_libreoffice


def test_libreoffice_conversion():
    """Test LibreOffice conversion the same way Streamlit app would use it."""
    
    # Test file
    pptx_path = Path("/Users/arjundivecha/Dropbox/GMO/India CFA.pptx")
    if not pptx_path.exists():
        print(f"❌ Test file not found: {pptx_path}")
        return False
    
    # Check LibreOffice is available
    libreoffice_path = _find_libreoffice()
    if not libreoffice_path:
        print("❌ LibreOffice not found!")
        return False
    
    print(f"✅ LibreOffice found: {libreoffice_path}")
    print(f"📁 Test file: {pptx_path.name}")
    print("=" * 70)
    
    # Create output directory (like Streamlit app does)
    with tempfile.TemporaryDirectory() as tmpdir:
        out_dir = Path(tmpdir) / "pptx2nano_output_streamlit"
        rendered_dir = out_dir / pptx_path.stem / "rendered"
        
        print(f"\n🚀 Starting LibreOffice conversion...")
        print(f"   Output: {rendered_dir}")
        
        start_time = time.time()
        
        try:
            # This is the exact call the Streamlit app makes
            rendered_paths = export_slides(
                pptx_path, 
                rendered_dir, 
                method="libreoffice"
            )
            
            elapsed = time.time() - start_time
            
            print(f"\n✅ SUCCESS! Converted {len(rendered_paths)} slides in {elapsed:.1f}s")
            print(f"\n📊 Output files:")
            for i, path in enumerate(rendered_paths[:5], 1):
                size_kb = path.stat().st_size / 1024
                print(f"   {i}. {path.name} ({size_kb:.1f} KB)")
            
            if len(rendered_paths) > 5:
                print(f"   ... and {len(rendered_paths) - 5} more")
            
            # Verify files are valid images
            print(f"\n🔍 Validating output...")
            from PIL import Image
            for path in rendered_paths[:3]:
                try:
                    with Image.open(path) as img:
                        print(f"   ✅ {path.name}: {img.size[0]}x{img.size[1]} {img.mode}")
                except Exception as e:
                    print(f"   ❌ {path.name}: Invalid image - {e}")
            
            return True
            
        except Exception as e:
            elapsed = time.time() - start_time
            print(f"\n❌ FAILED after {elapsed:.1f}s")
            print(f"   Error: {e}")
            return False


def test_streamlit_app_logic():
    """Simulate the exact logic flow from streamlit_app.py"""
    
    print("\n" + "=" * 70)
    print("🧪 TESTING STREAMLIT APP LOGIC")
    print("=" * 70)
    
    # Simulate session state (like Streamlit)
    session_state = {
        "pptx_conversion_method": "libreoffice"  # New default
    }
    
    # Test file
    input_path = Path("/Users/arjundivecha/Dropbox/GMO/India CFA.pptx")
    input_type = "pptx"
    
    print(f"\n📋 Session State:")
    print(f"   pptx_conversion_method: {session_state['pptx_conversion_method']}")
    print(f"\n📁 Input:")
    print(f"   Path: {input_path}")
    print(f"   Type: {input_type}")
    
    # Simulate the button click logic from streamlit_app.py
    with tempfile.TemporaryDirectory() as tmpdir:
        out_dir = Path(tmpdir) / "pptx2nano_output_streamlit"
        rendered_dir = out_dir / input_path.stem / "rendered"
        
        # This is the exact code from streamlit_app.py lines 799-806
        conversion_method = session_state.get("pptx_conversion_method", "libreoffice")
        
        print(f"\n🚀 Calling export_slides with method='{conversion_method}'...")
        
        start_time = time.time()
        
        if conversion_method == "libreoffice":
            print("   Using LibreOffice path...")
            rendered_paths = export_slides(
                Path(input_path), 
                rendered_dir, 
                method="libreoffice"
            )
        else:
            print("   Using Keynote path...")
            rendered_paths = export_slides(
                Path(input_path), 
                rendered_dir, 
                method="keynote"
            )
        
        elapsed = time.time() - start_time
        
        print(f"\n✅ SUCCESS!")
        print(f"   Slides: {len(rendered_paths)}")
        print(f"   Time: {elapsed:.1f}s")
        print(f"   Method: {conversion_method}")
        
        return True


if __name__ == "__main__":
    print("=" * 70)
    print("STREAMLIT + LIBREOFFICE INTEGRATION TEST")
    print("=" * 70)
    
    # Test 1: Direct conversion
    success1 = test_libreoffice_conversion()
    
    # Test 2: Streamlit app logic simulation
    try:
        success2 = test_streamlit_app_logic()
    except Exception as e:
        print(f"\n❌ Streamlit logic test failed: {e}")
        success2 = False
    
    # Summary
    print("\n" + "=" * 70)
    print("SUMMARY")
    print("=" * 70)
    print(f"Direct conversion test: {'✅ PASS' if success1 else '❌ FAIL'}")
    print(f"Streamlit logic test:   {'✅ PASS' if success2 else '❌ FAIL'}")
    
    if success1 and success2:
        print("\n🎉 All tests passed! LibreOffice integration is ready.")
        print("\nYou can now run the Streamlit app:")
        print("   streamlit run streamlit_app.py")
        print("\nThe app will default to LibreOffice for PPTX conversion.")
    else:
        print("\n⚠️  Some tests failed. Check the errors above.")
        sys.exit(1)
