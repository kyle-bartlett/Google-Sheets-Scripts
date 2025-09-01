#!/usr/bin/env python3
"""
Quick page inspector to see what's actually on the voting page
"""

import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import config

def inspect_page():
    """See what's actually on the voting page"""
    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    
    driver = webdriver.Chrome(options=chrome_options)
    
    try:
        print(f"🌐 Going to: {config.MARATHON_URL}")
        driver.get(config.MARATHON_URL)
        time.sleep(5)
        
        print(f"📄 Page Title: {driver.title}")
        
        # Check if voting is open
        page_text = driver.page_source.lower()
        
        if "voting" in page_text:
            print("✅ Found 'voting' text on page")
        else:
            print("❌ No 'voting' text found")
            
        if "closed" in page_text:
            print("⚠️ Found 'closed' text - voting might be closed")
            
        # Look for any clickable elements that might be candidates
        print("\n🔍 Looking for clickable elements...")
        
        # Scroll through page and collect all text
        all_text = []
        for scroll in range(5):
            driver.execute_script(f"window.scrollTo(0, {scroll * 800});")
            time.sleep(2)
            
            # Get buttons
            buttons = driver.find_elements(By.XPATH, "//button")
            for btn in buttons:
                if btn.text.strip():
                    all_text.append(f"BUTTON: {btn.text.strip()}")
            
            # Get clickable divs
            divs = driver.find_elements(By.XPATH, "//div[@onclick or @role='button']")
            for div in divs:
                if div.text.strip():
                    all_text.append(f"CLICKABLE DIV: {div.text.strip()}")
                    
            # Get labels
            labels = driver.find_elements(By.XPATH, "//label")
            for label in labels:
                if label.text.strip():
                    all_text.append(f"LABEL: {label.text.strip()}")
        
        # Show unique elements
        unique_text = list(set(all_text))[:20]  # First 20 unique items
        
        print(f"\n📋 Found {len(unique_text)} unique clickable elements:")
        for i, text in enumerate(unique_text, 1):
            print(f"   {i}. {text}")
            
        # Look specifically for our candidates
        print(f"\n🎯 Searching for our candidates:")
        for category, candidate in config.VOTING_RESPONSES.items():
            if candidate.lower() in page_text:
                print(f"   ✅ Found '{candidate}' on page")
            else:
                print(f"   ❌ '{candidate}' not found")
        
        print(f"\n⏸️ Keeping browser open for 15 seconds to inspect...")
        time.sleep(15)
        
    finally:
        driver.quit()

if __name__ == "__main__":
    inspect_page()

