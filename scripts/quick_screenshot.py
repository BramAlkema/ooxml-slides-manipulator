#!/usr/bin/env python3
"""
Quick screenshot with explicit ChromeDriver path
"""

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time
import os

def main():
    # Presentation from our API test
    presentation_id = "1ZZbpgBzMvQ-ggHT6Fd2a3JNWhhiv755FbsiAiSpY2wA"
    pub_url = f"https://docs.google.com/presentation/d/{presentation_id}/pub"
    
    print(f"📸 Taking screenshot of: {pub_url}")
    
    # Setup Chrome with explicit service
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--window-size=1200,800')
    
    # Use Homebrew ChromeDriver path
    service = Service('/opt/homebrew/bin/chromedriver')
    
    try:
        driver = webdriver.Chrome(service=service, options=chrome_options)
        
        driver.get(pub_url)
        time.sleep(5)
        
        os.makedirs('screenshots', exist_ok=True)
        screenshot_path = 'screenshots/font_pair_demo_working.png'
        driver.save_screenshot(screenshot_path)
        
        print(f"✅ Screenshot saved: {screenshot_path}")
        print(f"📋 Presentation ID: {presentation_id}")
        print("🎯 SUCCESS: Working Web API + Google Slides demonstration!")
        
    except Exception as e:
        print(f"❌ Error: {e}")
    finally:
        if 'driver' in locals():
            driver.quit()

if __name__ == "__main__":
    main()