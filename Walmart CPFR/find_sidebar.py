#!/usr/bin/env python3
"""
Find what's actually in the left sidebar
"""

import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import config

def find_sidebar_content():
    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    
    driver = webdriver.Chrome(options=chrome_options)
    
    try:
        print(f"🌐 Going to: {config.MARATHON_URL}")
        driver.get(config.MARATHON_URL)
        time.sleep(5)
        
        print("🔍 Looking for ALL elements in the LEFT sidebar (x < 400px)...")
        print("=" * 60)
        
        # Get all elements that might be in sidebar
        all_elements = driver.find_elements(By.XPATH, "//*")
        
        sidebar_texts = []
        for elem in all_elements:
            try:
                text = elem.text.strip()
                if text and len(text) > 2:  # Skip empty or very short text
                    location = elem.location
                    if location['x'] < 400 and location['x'] > 0:  # In left sidebar area
                        # Check if it's clickable
                        tag = elem.tag_name
                        clickable = tag in ['a', 'button'] or elem.get_attribute('onclick') or elem.get_attribute('role') == 'button'
                        
                        sidebar_texts.append({
                            'text': text[:50],  # First 50 chars
                            'x': location['x'],
                            'y': location['y'],
                            'tag': tag,
                            'clickable': clickable
                        })
            except:
                continue
        
        # Remove duplicates and sort by Y position
        unique_texts = []
        seen = set()
        for item in sorted(sidebar_texts, key=lambda x: x['y']):
            if item['text'] not in seen:
                seen.add(item['text'])
                unique_texts.append(item)
        
        print(f"Found {len(unique_texts)} unique items in sidebar:\n")
        
        for i, item in enumerate(unique_texts[:30], 1):  # Show first 30
            click_indicator = "🖱️" if item['clickable'] else "  "
            print(f"{click_indicator} {i:2}. [{item['tag']:6}] x={item['x']:3} y={item['y']:4} : {item['text']}")
        
        print("\n" + "=" * 60)
        print("🔍 Looking specifically for category-like items...")
        
        # Look for items that might be categories
        for item in unique_texts:
            text_lower = item['text'].lower()
            if any(word in text_lower for word in ['business', 'realtor', 'real estate', 'vacation', 'rental', 'best']):
                print(f"   📍 Found: '{item['text']}' (tag: {item['tag']}, clickable: {item['clickable']})")
        
        print("\n⏸️ Browser staying open for 20 seconds - check the sidebar yourself...")
        time.sleep(20)
        
    finally:
        driver.quit()

if __name__ == "__main__":
    find_sidebar_content()

