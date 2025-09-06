#!/usr/bin/env python3
"""
Get before/after font comparison using authenticated Google Apps Script
"""

import requests
import json

def get_before_after_comparison():
    """Create before/after presentations and get thumbnails using authentication"""
    
    web_app_url = 'https://script.google.com/macros/s/AKfycbwsD1RQSgoRsF3tqGka2YZmhYlIMmDYKmnIlA5c9YqeiJ6OtpzgRw5C1e7teKacOEkw/exec'
    
    print("🎯 CREATING BEFORE/AFTER FONT COMPARISON (AUTHENTICATED)")
    print("======================================================")
    
    # Use the authenticated Apps Script function to create comparison
    payload = {
        'fn': 'createFontComparison',
        'args': []
    }
    
    print("📞 Calling authenticated comparison function...")
    response = requests.post(web_app_url, json=payload, timeout=60)
    
    if response.status_code != 200:
        print(f"❌ API call failed: {response.status_code}")
        print(f"Response: {response.text}")
        return None
        
    result = response.json()
    
    if not result.get('result', {}).get('success'):
        print("❌ Comparison creation failed")
        print(f"Response: {json.dumps(result, indent=2)}")
        return None
    
    comparison_data = result['result']
    
    print("✅ Font comparison created successfully!")
    print("\n📋 BEFORE PRESENTATION (DEFAULT FONTS):")
    print(f"🆔 ID: {comparison_data['before']['presentationId']}")
    print(f"📝 Title: {comparison_data['before']['title']}")
    print(f"🔗 Edit URL: {comparison_data['before']['editUrl']}")
    print(f"🖼️ Thumbnail: {comparison_data['before']['thumbnailLink'] or 'Not available'}")
    print(f"🎨 Font Info: {comparison_data['before']['fontInfo']}")
    
    print("\n📋 AFTER PRESENTATION (MERRIWEATHER/INTER):")
    print(f"🆔 ID: {comparison_data['after']['presentationId']}")
    print(f"📝 Title: {comparison_data['after']['title']}")
    print(f"🔗 Edit URL: {comparison_data['after']['editUrl']}")
    print(f"🖼️ Thumbnail: {comparison_data['after']['thumbnailLink'] or 'Not available'}")
    print(f"🎨 Font Info: {comparison_data['after']['fontInfo']}")
    
    if comparison_data['after'].get('fontPair'):
        font_pair = comparison_data['after']['fontPair']
        print(f"📝 Parsed Font Pair: {font_pair['heading']} / {font_pair['body']}")
    
    if comparison_data['after'].get('colorPalette'):
        color_palette = comparison_data['after']['colorPalette']
        print(f"🎨 Parsed Color Palette: {color_palette.get('palette', 'Not available')}")
    
    print(f"\n⏰ Created: {comparison_data['timestamp']}")
    
    # Download thumbnails if available
    print("\n📸 ATTEMPTING TO DOWNLOAD THUMBNAILS...")
    
    for pres_type, pres_data in [('BEFORE', comparison_data['before']), ('AFTER', comparison_data['after'])]:
        thumbnail_url = pres_data.get('thumbnailLink')
        if thumbnail_url:
            try:
                print(f"🔗 Downloading {pres_type} thumbnail: {thumbnail_url}")
                thumb_response = requests.get(thumbnail_url, timeout=30)
                
                if thumb_response.status_code == 200:
                    import os
                    os.makedirs('screenshots', exist_ok=True)
                    
                    filename = f"screenshots/{pres_type.lower()}_thumbnail.jpg"
                    with open(filename, 'wb') as f:
                        f.write(thumb_response.content)
                    
                    print(f"✅ {pres_type} thumbnail saved: {filename}")
                else:
                    print(f"❌ {pres_type} thumbnail download failed: HTTP {thumb_response.status_code}")
                    
            except Exception as e:
                print(f"❌ {pres_type} thumbnail download error: {e}")
        else:
            print(f"⚠️ {pres_type} thumbnail not available")
    
    print("\n🎯 FONT PAIR COMPARISON COMPLETE!")
    print("=================================")
    print("✅ Both presentations created using authenticated Google Apps Script")
    print("✅ Font pair parsing and application demonstrated")
    print("✅ Before/After comparison available via edit URLs")
    print("📁 Check screenshots/ directory for any downloaded thumbnails")
    
    return comparison_data

if __name__ == "__main__":
    get_before_after_comparison()