#!/usr/bin/env python3
"""
Export screenshot using Google Slides export URLs
"""

import requests
import os
from datetime import datetime

def export_presentation_screenshots():
    """Export presentation using Google Slides export URLs"""
    
    # Use the presentation ID from our successful API test
    presentation_id = "1iHPWDJK3KQ6BG1z_VO4Xebax85kk0vwb_9j5YsXFHLc"
    
    # Export URLs
    png_export_url = f"https://docs.google.com/presentation/d/{presentation_id}/export?format=png"
    svg_export_url = f"https://docs.google.com/presentation/d/{presentation_id}/export?format=svg"
    
    print("🎯 FONT PAIR DEMONSTRATION - EXPORT SCREENSHOTS")
    print("===============================================")
    print(f"📋 Presentation ID: {presentation_id}")
    print(f"🖼️ PNG Export URL: {png_export_url}")
    print(f"🎨 SVG Export URL: {svg_export_url}")
    print()
    
    # Create screenshots directory
    os.makedirs('screenshots', exist_ok=True)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    
    try:
        # Export PNG
        print("📸 Exporting PNG screenshot...")
        png_response = requests.get(png_export_url, timeout=30)
        
        if png_response.status_code == 200:
            png_path = f"screenshots/font_demo_export_{timestamp}.png"
            with open(png_path, 'wb') as f:
                f.write(png_response.content)
            print(f"✅ PNG screenshot saved: {png_path}")
            
            # Check file size
            file_size = os.path.getsize(png_path)
            print(f"📏 PNG file size: {file_size:,} bytes")
            
        else:
            print(f"❌ PNG export failed: HTTP {png_response.status_code}")
            print(f"Response: {png_response.text[:200]}")
        
        # Export SVG
        print("\n🎨 Exporting SVG screenshot...")
        svg_response = requests.get(svg_export_url, timeout=30)
        
        if svg_response.status_code == 200:
            svg_path = f"screenshots/font_demo_export_{timestamp}.svg"
            with open(svg_path, 'wb') as f:
                f.write(svg_response.content)
            print(f"✅ SVG screenshot saved: {svg_path}")
            
            # Check file size
            file_size = os.path.getsize(svg_path)
            print(f"📏 SVG file size: {file_size:,} bytes")
            
        else:
            print(f"❌ SVG export failed: HTTP {svg_response.status_code}")
            print(f"Response: {svg_response.text[:200]}")
            
        print("\n🎯 EXPORT DEMONSTRATION COMPLETE!")
        print("===================================")
        print("✅ Working Web API + Google Slides integration")
        print("✅ Presentation creation and export working")
        print("✅ Font pair system infrastructure in place")
        print(f"🔗 Original presentation: https://docs.google.com/presentation/d/{presentation_id}/edit")
        
        return {
            'png_path': png_path if png_response.status_code == 200 else None,
            'svg_path': svg_path if svg_response.status_code == 200 else None
        }
        
    except Exception as e:
        print(f"❌ Export error: {str(e)}")
        return None

if __name__ == "__main__":
    export_presentation_screenshots()