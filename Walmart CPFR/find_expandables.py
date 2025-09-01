#!/usr/bin/env python3
"""
Find expandable elements and arrows in sidebar
"""

import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import config

def find_expandables():
    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    
    driver = webdriver.Chrome(options=chrome_options)
    
    try:
        print(f"🌐 Going to: {config.MARATHON_URL}")
        driver.get(config.MARATHON_URL)
        time.sleep(5)
        
        print("🔍 Looking for expandable elements (arrows, icons, etc)...")
        print("=" * 60)
        
        # Scroll down to see more of the page
        driver.execute_script("window.scrollTo(0, 500);")
        time.sleep(2)
        
        # Look for common expand/collapse indicators
        expand_selectors = [
            "//i[contains(@class, 'arrow')]",
            "//i[contains(@class, 'chevron')]",
            "//i[contains(@class, 'expand')]",
            "//i[contains(@class, 'plus')]",
            "//i[contains(@class, 'fa-')]",  # Font Awesome icons
            "//svg",  # SVG icons
            "//*[@aria-expanded]",  # Elements with aria-expanded
            "//*[contains(@class, 'dropdown')]",
            "//*[contains(@class, 'accordion')]",
            "//*[contains(@class, 'collapse')]"
        ]
        
        found_expandables = []
        
        for selector in expand_selectors:
            try:
                elements = driver.find_elements(By.XPATH, selector)
                for elem in elements:
                    location = elem.location
                    if location['x'] < 500:  # Left side of page
                        # Get parent text to understand context
                        try:
                            parent = elem.find_element(By.XPATH, "./..")
                            parent_text = parent.text.strip()[:50] if parent.text else "No text"
                            
                            found_expandables.append({
                                'type': selector.split('/')[-1],
                                'x': location['x'],
                                'y': location['y'],
                                'parent_text': parent_text,
                                'class': elem.get_attribute('class') or 'no-class'
                            })
                        except:
                            continue
            except:
                continue
        
        # Remove duplicates
        unique_expandables = []
        seen = set()
        for item in sorted(found_expandables, key=lambda x: x['y']):
            key = f"{item['x']}-{item['y']}"
            if key not in seen:
                seen.add(key)
                unique_expandables.append(item)
        
        print(f"Found {len(unique_expandables)} expandable elements:\n")
        
        for i, item in enumerate(unique_expandables[:20], 1):
            print(f"{i:2}. x={item['x']:3} y={item['y']:4} [{item['type']}] class='{item['class'][:30]}' parent='{item['parent_text']}'")
        
        print("\n" + "=" * 60)
        print("💡 TRYING DIFFERENT APPROACH - Looking in main content area...")
        
        # Maybe the categories are in the main content, not sidebar
        driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(2)
        
        # Look for any text containing business, realtor, etc in main area
        main_selectors = [
            "//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'business')]",
            "//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'realtor')]",
            "//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'real estate')]"
        ]
        
        for selector in main_selectors[:1]:  # Just check first one
            elements = driver.find_elements(By.XPATH, selector)
            for elem in elements[:5]:  # First 5 matches
                try:
                    print(f"   Found: '{elem.text[:60]}' at x={elem.location['x']}")
                except:
                    continue
        
        print("\n⏸️ Browser staying open for 30 seconds...")
        print("📝 Please look at the page and tell me:")
        print("   1. Where exactly is 'The businesses' located?")
        print("   2. What does the expansion arrow look like?")
        print("   3. Is it in a sidebar or main content area?")
        time.sleep(30)
        
    finally:
        driver.quit()

if __name__ == "__main__":
    find_expandables()

