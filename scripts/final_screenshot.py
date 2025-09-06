#!/usr/bin/env python3
"""
Final screenshot attempt with webdriver-manager
"""

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time
import os

def main():
    # Presentation from our API test
    presentation_id = "1ZZbpgBzMvQ-ggHT6Fd2a3JNWhhiv755FbsiAiSpY2wA"
    pub_url = f"https://docs.google.com/presentation/d/{presentation_id}/pub"
    
    print(f"📸 Taking screenshot of: {pub_url}")
    
    # Setup Chrome with webdriver-manager
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--window-size=1200,800')
    chrome_options.add_argument('--disable-gpu')
    
    try:
        # Use webdriver-manager to handle ChromeDriver
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        
        driver.get(pub_url)
        time.sleep(5)
        
        os.makedirs('screenshots', exist_ok=True)
        screenshot_path = 'screenshots/working_presentation_demo.png'
        driver.save_screenshot(screenshot_path)
        
        print(f"✅ Screenshot saved: {screenshot_path}")
        print(f"📋 Presentation ID: {presentation_id}")
        print("🎯 SUCCESS: Working Web API + Google Slides demonstration captured!")
        
        driver.quit()
        return screenshot_path
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return None

if __name__ == "__main__":
    main()