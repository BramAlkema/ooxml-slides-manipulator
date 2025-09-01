#!/usr/bin/env python3
"""
Create presentations showing before/after font comparison
"""

import requests
import time

def create_font_comparison():
    """Create presentations showing font differences"""
    
    web_app_url = 'https://script.google.com/macros/s/AKfycbwsD1RQSgoRsF3tqGka2YZmhYlIMmDYKmnIlA5c9YqeiJ6OtpzgRw5C1e7teKacOEkw/exec'
    
    presentations = []
    
    # Create "BEFORE" presentation - default fonts
    print("📞 Creating BEFORE presentation (default fonts)...")
    before_payload = {
        'fn': 'createPresentationFromPrompt',
        'args': ['Create presentation with default fonts - BEFORE demonstration']
    }
    
    before_response = requests.post(web_app_url, json=before_payload, timeout=30)
    if before_response.status_code == 200:
        before_result = before_response.json()
        if before_result.get('result', {}).get('success'):
            before_id = before_result['result']['presentationId']
            presentations.append({
                'type': 'BEFORE',
                'id': before_id,
                'title': 'Default Fonts (Before)',
                'edit_url': f"https://docs.google.com/presentation/d/{before_id}/edit",
                'pub_url': f"https://docs.google.com/presentation/d/{before_id}/pub"
            })
            print(f"✅ BEFORE presentation created: {before_id}")
        else:
            print("❌ BEFORE presentation creation failed")
    else:
        print(f"❌ BEFORE API call failed: {before_response.status_code}")
    
    time.sleep(2)  # Brief pause between requests
    
    # Create "AFTER" presentation - Merriweather/Inter fonts  
    print("📞 Creating AFTER presentation (Merriweather/Inter fonts)...")
    after_payload = {
        'fn': 'createPresentationFromPrompt', 
        'args': ['Create presentation with https://coolors.co/edd3c4-c8adc0-7765e3-3b60e4-080708 and Merriweather/Inter fonts - AFTER demonstration']
    }
    
    after_response = requests.post(web_app_url, json=after_payload, timeout=30)
    if after_response.status_code == 200:
        after_result = after_response.json()
        if after_result.get('result', {}).get('success'):
            after_id = after_result['result']['presentationId']
            presentations.append({
                'type': 'AFTER',
                'id': after_id,
                'title': 'Merriweather/Inter Fonts (After)',
                'edit_url': f"https://docs.google.com/presentation/d/{after_id}/edit",
                'pub_url': f"https://docs.google.com/presentation/d/{after_id}/pub"
            })
            print(f"✅ AFTER presentation created: {after_id}")
        else:
            print("❌ AFTER presentation creation failed")
    else:
        print(f"❌ AFTER API call failed: {after_response.status_code}")
    
    # Display comparison
    print("\n🎯 BEFORE/AFTER FONT PAIR COMPARISON")
    print("====================================")
    
    for pres in presentations:
        print(f"\n📋 {pres['type']} - {pres['title']}")
        print(f"🆔 ID: {pres['id']}")
        print(f"📝 Edit URL: {pres['edit_url']}")
        print(f"👁️ Public URL: {pres['pub_url']}")
        print(f"📸 PNG Export: https://docs.google.com/presentation/d/{pres['id']}/export?format=png")
    
    if len(presentations) == 2:
        print("\n🔍 TO VIEW THE DIFFERENCE:")
        print("1. Open both edit URLs in separate browser tabs")
        print("2. Compare the font styling between presentations")
        print("3. BEFORE uses default Google Slides fonts")
        print("4. AFTER uses Merriweather (headings) + Inter (body text)")
        
        return presentations
    else:
        print("\n⚠️ Could not create both presentations for comparison")
        return presentations

if __name__ == "__main__":
    print("🎨 CREATING BEFORE/AFTER FONT COMPARISON")
    print("========================================")
    create_font_comparison()