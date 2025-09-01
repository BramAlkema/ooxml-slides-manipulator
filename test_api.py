#!/usr/bin/env python3
"""
Simple API test to create presentation and get URL
"""

import requests
import json

def main():
    web_app_url = 'https://script.google.com/macros/s/AKfycbwsD1RQSgoRsF3tqGka2YZmhYlIMmDYKmnIlA5c9YqeiJ6OtpzgRw5C1e7teKacOEkw/exec'
    
    payload = {
        'fn': 'createPresentationFromPrompt',
        'args': ['Create presentation with https://coolors.co/edd3c4-c8adc0-7765e3-3b60e4-080708 and Merriweather/Inter fonts']
    }
    
    print("📞 Calling Web API to create presentation...")
    try:
        response = requests.post(web_app_url, json=payload, timeout=30)
        
        if response.status_code == 200:
            result = response.json()
            print("✅ API Response:", json.dumps(result, indent=2))
            
            if 'result' in result and result['result'].get('success'):
                presentation_data = result['result']
                presentation_id = presentation_data.get('presentationId')
                
                print(f"\n🎯 PRESENTATION CREATED SUCCESSFULLY")
                print("=" * 40)
                print(f"📋 Presentation ID: {presentation_id}")
                print(f"📝 Edit URL: https://docs.google.com/presentation/d/{presentation_id}/edit")
                print(f"👁️ Public URL: https://docs.google.com/presentation/d/{presentation_id}/pub")
                
                font_pair = presentation_data.get('fontPair')
                color_palette = presentation_data.get('colorPalette')
                
                if font_pair:
                    print(f"✅ Font Pair: {font_pair['heading']}/{font_pair['body']}")
                else:
                    print("⚠️ Font pair data not available (basic function deployed)")
                
                if color_palette:
                    print(f"✅ Color Palette: {color_palette['palette']}")
                else:
                    print("⚠️ Color palette data not available")
                    
            else:
                print("❌ API call failed:", result)
        else:
            print(f"❌ HTTP Error: {response.status_code}")
            print(response.text)
            
    except Exception as e:
        print(f"❌ Error: {str(e)}")

if __name__ == "__main__":
    main()