"""
Comprehensive Sidebar Search - Find the category navigation you described
Looks everywhere on the page for category links
"""

import time
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import config

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')
logger = logging.getLogger(__name__)

def comprehensive_search(url, page_name):
    """Comprehensive search for category navigation"""
    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    
    driver = webdriver.Chrome(options=chrome_options)
    driver.maximize_window()
    
    try:
        logger.info(f"🌐 Comprehensive search on {page_name}...")
        driver.get(url)
        time.sleep(8)  # Longer wait for dynamic content
        
        logger.info(f"📄 Page: {driver.title}")
        
        # Wait for any JavaScript to finish loading
        logger.info("⏳ Waiting for dynamic content to load...")
        time.sleep(5)
        
        # Execute JavaScript to scroll and wait for content
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight/4);")
        time.sleep(2)
        driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(2)
        
        logger.info("🔍 COMPREHENSIVE SEARCH FOR CATEGORY LINKS...")
        
        # Search 1: Look for ANY elements containing category-related words
        category_keywords = [
            'realtor', 'real estate', 'vacation rental', 'business', 'office',
            'restaurant', 'bar', 'hotel', 'salon', 'gym', 'doctor', 'dentist',
            'contractor', 'photographer', 'florist', 'jewelry', 'spa', 'bank'
        ]
        
        found_category_elements = []
        
        for keyword in category_keywords:
            try:
                # Look for text containing keyword
                elements = driver.find_elements(By.XPATH, 
                    f"//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{keyword}')]")
                
                for elem in elements:
                    try:
                        text = elem.text.strip()
                        location = elem.location
                        tag = elem.tag_name
                        
                        if (len(text) > 2 and len(text) < 100 and  # Reasonable text length
                            location['x'] < 500):  # Left half of page
                            
                            found_category_elements.append({
                                'keyword': keyword,
                                'text': text,
                                'tag': tag,
                                'location': location,
                                'element': elem
                            })
                    except:
                        continue
            except:
                continue
        
        logger.info(f"📊 Found {len(found_category_elements)} potential category elements")
        
        # Remove duplicates and sort by position
        unique_elements = []
        seen_texts = set()
        
        for item in found_category_elements:
            if item['text'] not in seen_texts:
                unique_elements.append(item)
                seen_texts.add(item['text'])
        
        unique_elements.sort(key=lambda x: (x['location']['x'], x['location']['y']))
        
        # Show results grouped by location
        left_side = [e for e in unique_elements if e['location']['x'] < 300]
        center = [e for e in unique_elements if 300 <= e['location']['x'] < 600]
        right_side = [e for e in unique_elements if e['location']['x'] >= 600]
        
        logger.info(f"📍 LEFT SIDE ({len(left_side)} elements):")
        for elem in left_side[:10]:
            logger.info(f"   🎯 '{elem['text']}' ({elem['tag']}) at x={elem['location']['x']}")
        
        logger.info(f"📍 CENTER ({len(center)} elements):")
        for elem in center[:5]:
            logger.info(f"   📋 '{elem['text']}' ({elem['tag']}) at x={elem['location']['x']}")
        
        logger.info(f"📍 RIGHT SIDE ({len(right_side)} elements):")
        for elem in right_side[:5]:
            logger.info(f"   📄 '{elem['text']}' ({elem['tag']}) at x={elem['location']['x']}")
        
        # Search 2: Look for clickable category elements specifically
        logger.info("🖱️ Looking for clickable category elements...")
        
        clickable_categories = []
        for item in unique_elements:
            elem = item['element']
            try:
                # Check if element or parent is clickable
                if (elem.tag_name in ['a', 'button'] or 
                    elem.find_element(By.XPATH, "./ancestor-or-self::a") or
                    'click' in elem.get_attribute('onclick') or ''):
                    
                    clickable_categories.append(item)
            except:
                # Check if parent is clickable
                try:
                    parent = elem.find_element(By.XPATH, "./..")
                    if parent.tag_name in ['a', 'button']:
                        clickable_categories.append(item)
                except:
                    pass
        
        logger.info(f"🖱️ Found {len(clickable_categories)} clickable category elements")
        
        for item in clickable_categories[:10]:
            logger.info(f"   ✅ Clickable: '{item['text']}' at x={item['location']['x']}")
        
        # Search 3: Look for navigation lists or menus
        logger.info("📋 Looking for navigation lists...")
        
        nav_containers = driver.find_elements(By.CSS_SELECTOR, 
            "nav, ul, ol, .menu, .navigation, .nav, .sidebar, .categories")
        
        for i, container in enumerate(nav_containers):
            try:
                location = container.location
                size = container.size
                if location['x'] < 400 and size['height'] > 100:  # Left side, decent size
                    links = container.find_elements(By.TAG_NAME, "a")
                    logger.info(f"   📋 Nav container {i+1}: {len(links)} links at x={location['x']}")
                    
                    for link in links[:5]:  # Show first 5 links
                        try:
                            text = link.text.strip()
                            if len(text) > 2:
                                logger.info(f"      🔗 '{text}'")
                        except:
                            pass
            except:
                continue
        
        # Test interaction with most promising element
        if clickable_categories:
            logger.info("🖱️ Testing interaction with most promising category...")
            test_item = clickable_categories[0]
            
            try:
                elem = test_item['element']
                logger.info(f"🎯 Testing: '{test_item['text']}'")
                
                # Try to find clickable parent
                clickable_elem = elem
                if elem.tag_name not in ['a', 'button']:
                    try:
                        clickable_elem = elem.find_element(By.XPATH, "./ancestor::a[1]")
                    except:
                        try:
                            clickable_elem = elem.find_element(By.XPATH, "./..")
                        except:
                            pass
                
                # Scroll to and click
                driver.execute_script("arguments[0].scrollIntoView(true);", clickable_elem)
                time.sleep(2)
                
                original_url = driver.current_url
                clickable_elem.click()
                time.sleep(5)
                
                new_url = driver.current_url
                logger.info(f"✅ Click successful! URL changed: {original_url != new_url}")
                
            except Exception as e:
                logger.error(f"❌ Click test failed: {e}")
        
        logger.info("⏸️ Keeping browser open for detailed manual inspection (45 seconds)...")
        logger.info("👀 Look for the category sidebar you mentioned!")
        time.sleep(45)
        
        return {
            'total_found': len(unique_elements),
            'left_side': len(left_side),
            'clickable': len(clickable_categories),
            'nav_containers': len(nav_containers)
        }
        
    except Exception as e:
        logger.error(f"❌ Search failed: {e}")
        return {}
    finally:
        driver.quit()

def main():
    print("🔍 COMPREHENSIVE CATEGORY SIDEBAR SEARCH")
    print("This will find ALL category-related elements on the page")
    print("Take a close look at the browser to spot the sidebar!")
    
    choice = input("\nWhich page?\n1. Marathon\n2. Key West\nEnter 1 or 2: ").strip()
    
    if choice == "1":
        results = comprehensive_search(config.MARATHON_URL, "Marathon")
    elif choice == "2":
        results = comprehensive_search(config.KEY_WEST_URL, "Key West")
    else:
        print("Invalid choice")
        return
    
    print(f"\n✅ Search complete!")
    print(f"📊 Found {results.get('total_found', 0)} category elements")
    print(f"📍 {results.get('left_side', 0)} on left side")
    print(f"🖱️ {results.get('clickable', 0)} clickable")

if __name__ == "__main__":
    main()
