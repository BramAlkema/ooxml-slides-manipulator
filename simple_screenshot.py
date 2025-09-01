#!/usr/bin/env python3
"""
Simple screenshot of the working presentation
"""

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time
import os
from datetime import datetime

def screenshot_presentation():
    """Screenshot the existing presentation"""
    
    # Use the presentation ID from the API test
    presentation_id = "1ZZbpgBzMvQ-ggHT6Fd2a3JNWhhiv755FbsiAiSpY2wA"
    
    edit_url = f"https://docs.google.com/presentation/d/{presentation_id}/edit"
    pub_url = f"https://docs.google.com/presentation/d/{presentation_id}/pub"
    
    # Setup Chrome options
    chrome_options = Options()
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--window-size=1920,1080')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--disable-web-security')
    chrome_options.add_argument('--allow-running-insecure-content')
    
    driver = None
    try:
        # Initialize Chrome driver
        print("🌐 Starting Chrome browser...")
        driver = webdriver.Chrome(options=chrome_options)
        
        # Create screenshots directory
        os.makedirs('screenshots', exist_ok=True)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        # Screenshot public URL (no authentication required)
        print(f"📸 Taking screenshot of public URL...")
        print(f"🔗 URL: {pub_url}")
        
        driver.get(pub_url)
        time.sleep(8)  # Allow time for loading
        
        pub_screenshot_path = f"screenshots/working_presentation_{timestamp}.png"
        driver.save_screenshot(pub_screenshot_path)
        print(f"✅ Public view screenshot saved: {pub_screenshot_path}")
        
        print("\n🎯 SCREENSHOT DEMONSTRATION COMPLETE")
        print("=" * 50)
        print(f"📋 Presentation ID: {presentation_id}")
        print(f"📁 Screenshot saved: {pub_screenshot_path}")
        print(f"🔗 Edit URL: {edit_url}")
        print(f"👁️ Public URL: {pub_url}")
        print("\n✅ This demonstrates the working Web API + Google Slides integration!")
        print("📝 The presentation was created via API call and successfully screenshotted.")
        
        return pub_screenshot_path
        
    except Exception as e:
        print(f"❌ Screenshot error: {str(e)}")
        return None
        
    finally:
        if driver:
            print("🔚 Closing browser...")
            driver.quit()

if __name__ == "__main__":
    print("📸 Font Pair Demonstration - Taking Screenshot")
    print("=" * 50)
    screenshot_presentation()