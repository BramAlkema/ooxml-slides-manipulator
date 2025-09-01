#!/usr/bin/env python3
"""
Test export immediately after creating presentation
"""

import requests
import time
import os

def test_create_and_export():
    """Create presentation and immediately test export"""
    
    web_app_url = 'https://script.google.com/macros/s/AKfycbwsD1RQSgoRsF3tqGka2YZmhYlIMmDYKmnIlA5c9YqeiJ6OtpzgRw5C1e7teKacOEkw/exec'
    
    # Step 1: Create presentation
    print("📞 Creating presentation with sharing enabled...")
    payload = {
        'fn': 'createPresentationFromPrompt',
        'args': ['Create presentation with Merriweather/Inter fonts for export test']
    }
    
    response = requests.post(web_app_url, json=payload, timeout=30)
    
    if response.status_code != 200:
        print(f"❌ API call failed: {response.status_code}")
        return
        
    result = response.json()
    if not result.get('result', {}).get('success'):
        print("❌ Presentation creation failed")
        return
        
    presentation_id = result['result']['presentationId']
    print(f"✅ Presentation created: {presentation_id}")
    
    # Step 2: Wait a moment for sharing to take effect
    print("⏳ Waiting 5 seconds for sharing permissions to propagate...")
    time.sleep(5)
    
    # Step 3: Test export
    print("📸 Testing PNG export...")
    png_export_url = f"https://docs.google.com/presentation/d/{presentation_id}/export?format=png"
    
    export_response = requests.get(png_export_url, timeout=30)
    
    print(f"🔍 Export response: HTTP {export_response.status_code}")
    print(f"📏 Content-Type: {export_response.headers.get('content-type', 'unknown')}")
    print(f"📐 Content-Length: {len(export_response.content):,} bytes")
    
    if export_response.status_code == 200:
        # Save the exported image
        os.makedirs('screenshots', exist_ok=True)
        export_path = f"screenshots/export_test_{presentation_id}.png"
        
        with open(export_path, 'wb') as f:
            f.write(export_response.content)
            
        print(f"✅ Export successful: {export_path}")
        print(f"📋 Presentation ID: {presentation_id}")
        print(f"🔗 Edit URL: https://docs.google.com/presentation/d/{presentation_id}/edit")
        print(f"👁️ Public URL: https://docs.google.com/presentation/d/{presentation_id}/pub")
        print(f"📸 Export URL: {png_export_url}")
        
        return export_path
        
    else:
        print(f"❌ Export failed: HTTP {export_response.status_code}")
        # Show first 200 chars of response for debugging
        response_text = export_response.text[:200] if export_response.text else "No response body"
        print(f"Response snippet: {response_text}")
        return None

if __name__ == "__main__":
    print("🎯 TESTING EXPORT AFTER CREATION WITH PROPER SHARING")
    print("====================================================")
    test_create_and_export()