#!/usr/bin/env python3
"""
Simple Python script to screenshot Google Slides presentations
"""

import requests
import json
from selenium import webdriver
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
from datetime import datetime

def create_presentation_via_api():
    """Call the Google Apps Script Web API to create a presentation"""
    
    web_app_url = 'https://script.google.com/macros/s/AKfycbwsD1RQSgoRsF3tqGka2YZmhYlIMmDYKmnIlA5c9YqeiJ6OtpzgRw5C1e7teKacOEkw/exec'
    
    payload = {
        'fn': 'createPresentationFromPrompt',
        'args': ['Create presentation with https://coolors.co/edd3c4-c8adc0-7765e3-3b60e4-080708 and Merriweather/Inter fonts']
    }
    
    print("📞 Calling Web API to create presentation...")
    response = requests.post(web_app_url, json=payload)
    
    if response.status_code == 200:
        result = response.json()
        print("✅ API Response:", json.dumps(result, indent=2))
        
        if 'result' in result and result['result'].get('success'):
            presentation_data = result['result']
            presentation_id = presentation_data.get('presentationId')
            
            print(f"✅ Presentation created with ID: {presentation_id}")
            return {
                'presentation_id': presentation_id,
                'edit_url': f"https://docs.google.com/presentation/d/{presentation_id}/edit",
                'pub_url': f"https://docs.google.com/presentation/d/{presentation_id}/pub",
                'font_pair': presentation_data.get('fontPair'),
                'color_palette': presentation_data.get('colorPalette')
            }
        else:
            print("❌ API call failed:", result)
            return None
    else:
        print(f"❌ HTTP Error: {response.status_code}")
        return None

def screenshot_presentation(presentation_data):
    """Screenshot the Google Slides presentation"""
    
    if not presentation_data:
        print("❌ No presentation data to screenshot")
        return
    
    # Setup Chrome options for headless mode
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--window-size=1920,1080')
    chrome_options.add_argument('--disable-gpu')
    
    driver = None
    try:
        # Initialize Chrome driver
        driver = webdriver.Chrome(options=chrome_options)
        
        # Create screenshots directory
        os.makedirs('screenshots', exist_ok=True)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        # Screenshot public URL first (no auth required)
        print(f"📸 Taking screenshot of public URL...")
        pub_url = presentation_data['pub_url']
        
        driver.get(pub_url)
        time.sleep(5)  # Allow time for loading
        
        pub_screenshot_path = f"screenshots/font_demo_public_{timestamp}.png"
        driver.save_screenshot(pub_screenshot_path)
        print(f"✅ Public view screenshot saved: {pub_screenshot_path}")
        
        # Try edit URL (may require auth)
        print(f"📸 Attempting screenshot of edit URL...")
        edit_url = presentation_data['edit_url']
        
        driver.get(edit_url)
        time.sleep(8)  # Allow more time for editor to load
        
        edit_screenshot_path = f"screenshots/font_demo_edit_{timestamp}.png"
        driver.save_screenshot(edit_screenshot_path)
        print(f"✅ Edit view screenshot saved: {edit_screenshot_path}")
        
        # Print summary
        print("\n🎯 FONT PAIR DEMONSTRATION COMPLETE")
        print("=" * 40)
        print(f"📋 Presentation ID: {presentation_data['presentation_id']}")
        print(f"🎨 Font Pair Data: {presentation_data.get('font_pair', 'Not available')}")
        print(f"🌈 Color Palette: {presentation_data.get('color_palette', 'Not available')}")
        print(f"📁 Screenshots saved in: screenshots/")
        print(f"🔗 Edit URL: {edit_url}")
        print(f"👁️ Public URL: {pub_url}")
        
        return {
            'pub_screenshot': pub_screenshot_path,
            'edit_screenshot': edit_screenshot_path
        }
        
    except Exception as e:
        print(f"❌ Screenshot error: {str(e)}")
        return None
        
    finally:
        if driver:
            driver.quit()

def main():
    """Main function to run the font pair demonstration"""
    
    print("🎨 Starting Font Pair Demonstration with Python Screenshots")
    print("=" * 60)
    
    # Step 1: Create presentation via API
    presentation_data = create_presentation_via_api()
    
    if not presentation_data:
        print("❌ Failed to create presentation")
        return
    
    # Step 2: Screenshot the presentation
    screenshot_results = screenshot_presentation(presentation_data)
    
    if screenshot_results:
        print("\n✅ Font pair demonstration completed successfully!")
        print("📁 Check the screenshots/ directory for results")
    else:
        print("\n❌ Screenshot capture failed")

if __name__ == "__main__":
    main()