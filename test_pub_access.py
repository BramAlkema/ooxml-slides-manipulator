#!/usr/bin/env python3
"""
Test if /pub URL is accessible after proper publishing
"""

import requests
import time

def test_pub_access():
    """Test if presentation is properly published"""
    
    web_app_url = 'https://script.google.com/macros/s/AKfycbwsD1RQSgoRsF3tqGka2YZmhYlIMmDYKmnIlA5c9YqeiJ6OtpzgRw5C1e7teKacOEkw/exec'
    
    # Create presentation with new Drive API publishing
    print("📞 Creating presentation with Drive API publishing...")
    payload = {
        'fn': 'createPresentationFromPrompt',
        'args': ['Test presentation with proper Drive API publishing']
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
    
    # Test /pub URL accessibility
    pub_url = f"https://docs.google.com/presentation/d/{presentation_id}/pub"
    
    print("⏳ Waiting 10 seconds for publishing to take effect...")
    time.sleep(10)
    
    print(f"🔗 Testing public URL: {pub_url}")
    
    try:
        pub_response = requests.get(pub_url, timeout=30)
        print(f"📊 /pub URL Response: HTTP {pub_response.status_code}")
        print(f"📏 Content-Type: {pub_response.headers.get('content-type', 'unknown')}")
        print(f"📐 Content-Length: {len(pub_response.content):,} bytes")
        
        # Check if we get actual presentation content vs error page
        if pub_response.status_code == 200:
            content_preview = pub_response.text[:200].replace('\n', ' ')
            print(f"📝 Content preview: {content_preview}...")
            
            # Look for signs it's a real presentation vs error page
            if 'presentation' in pub_response.text.lower() or 'slides' in pub_response.text.lower():
                print("✅ /pub URL appears to be working - contains presentation content")
            elif 'not published' in pub_response.text.lower():
                print("❌ /pub URL shows 'not published' error")
            else:
                print("⚠️ /pub URL accessible but content unclear")
        else:
            print(f"❌ /pub URL not accessible: HTTP {pub_response.status_code}")
            
        print(f"\n🎯 FINAL RESULT:")
        print(f"📋 Presentation ID: {presentation_id}")
        print(f"🔗 Edit URL: https://docs.google.com/presentation/d/{presentation_id}/edit") 
        print(f"👁️ Public URL: {pub_url}")
        print(f"📸 Export URL: https://docs.google.com/presentation/d/{presentation_id}/export?format=png")
        
    except Exception as e:
        print(f"❌ Error testing /pub URL: {e}")

if __name__ == "__main__":
    print("🎯 TESTING DRIVE API PUBLISHING METHOD")
    print("=====================================")
    test_pub_access()