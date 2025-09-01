#!/usr/bin/env python3
"""
Debug screenshot with visible browser
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
    
    print(f"📸 Debug: Opening {pub_url}")
    
    # Setup Chrome with visible browser
    chrome_options = Options()
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--window-size=1200,800')
    
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        
        print("🌐 Opening URL...")
        driver.get(pub_url)
        
        print("⏳ Waiting 10 seconds for page to load...")
        time.sleep(10)
        
        print("📄 Page title:", driver.title)
        print("🔗 Current URL:", driver.current_url)
        
        os.makedirs('screenshots', exist_ok=True)
        screenshot_path = 'screenshots/debug_presentation.png'
        driver.save_screenshot(screenshot_path)
        
        print(f"✅ Screenshot saved: {screenshot_path}")
        
        # Keep browser open for manual inspection
        print("🔍 Browser will close in 5 seconds...")
        time.sleep(5)
        
        driver.quit()
        return screenshot_path
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    main()