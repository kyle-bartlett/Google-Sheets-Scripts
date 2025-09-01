"""
Generic Sidebar Explorer - Find ALL category links in sidebar
This will show us the pattern so we can adapt our automation
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

def explore_sidebar_pattern(url, page_name):
    """Explore sidebar to understand the link pattern"""
    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    
    driver = webdriver.Chrome(options=chrome_options)
    driver.maximize_window()
    
    try:
        logger.info(f"🌐 Exploring sidebar pattern on {page_name}...")
        driver.get(url)
        time.sleep(5)
        
        logger.info(f"📄 Page: {driver.title}")
        
        # Find ALL links on the left side of the page
        logger.info("🔍 Finding all links on left side...")
        
        all_links = driver.find_elements(By.TAG_NAME, "a")
        left_side_links = []
        
        for link in all_links:
            try:
                location = link.location
                size = link.size
                text = link.text.strip()
                href = link.get_attribute('href')
                
                # Filter for left side links with meaningful text
                if (location['x'] < 400 and  # Left side
                    len(text) > 2 and  # Has meaningful text
                    len(text) < 50 and  # Not too long
                    href and  # Has href
                    'javascript:' not in href):  # Not a JS function
                    
                    left_side_links.append({
                        'text': text,
                        'href': href,
                        'x': location['x'],
                        'y': location['y']
                    })
            except:
                continue
        
        # Sort by vertical position
        left_side_links.sort(key=lambda x: x['y'])
        
        logger.info(f"📊 Found {len(left_side_links)} left-side links:")
        
        # Group by rough categories
        category_like_links = []
        nav_links = []
        
        for link in left_side_links:
            text_lower = link['text'].lower()
            
            # Look for category-like words
            category_words = ['bar', 'restaurant', 'business', 'office', 'store', 'service', 
                            'rental', 'realtor', 'hotel', 'salon', 'gym', 'doctor', 'best']
            
            if any(word in text_lower for word in category_words):
                category_like_links.append(link)
            else:
                nav_links.append(link)
        
        logger.info(f"🎯 Category-like links ({len(category_like_links)}):")
        for link in category_like_links[:10]:  # Show first 10
            logger.info(f"   📋 '{link['text']}' at x={link['x']}")
        
        logger.info(f"🧭 Navigation links ({len(nav_links)}):")
        for link in nav_links[:10]:  # Show first 10
            logger.info(f"   🔗 '{link['text']}' at x={link['x']}")
        
        # Test clicking a category-like link if found
        if category_like_links:
            test_link_info = category_like_links[0]
            logger.info(f"🖱️ Testing click on '{test_link_info['text']}'...")
            
            try:
                # Find the link element again (fresh reference)
                test_link = driver.find_element(By.XPATH, f"//a[text()='{test_link_info['text']}']")
                
                # Scroll to and click
                driver.execute_script("arguments[0].scrollIntoView(true);", test_link)
                time.sleep(2)
                
                original_url = driver.current_url
                test_link.click()
                time.sleep(5)
                
                new_url = driver.current_url
                logger.info(f"📍 URL changed: {original_url != new_url}")
                logger.info(f"📍 New URL: {new_url}")
                
                # Look for changes in content
                inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='text'], textarea")
                buttons = driver.find_elements(By.CSS_SELECTOR, "button, input[type='submit']")
                
                logger.info(f"📝 After click: {len(inputs)} inputs, {len(buttons)} buttons")
                
            except Exception as e:
                logger.error(f"❌ Click test failed: {e}")
        
        logger.info("⏸️ Keeping browser open for manual inspection (30 seconds)...")
        time.sleep(30)
        
        return {
            'category_links': category_like_links,
            'nav_links': nav_links,
            'total_left_links': len(left_side_links)
        }
        
    except Exception as e:
        logger.error(f"❌ Exploration failed: {e}")
        return {}
    finally:
        driver.quit()

def main():
    print("🔍 GENERIC SIDEBAR PATTERN EXPLORER")
    print("This will show ALL category links in the sidebar")
    
    choice = input("\nWhich page?\n1. Marathon\n2. Key West\nEnter 1 or 2: ").strip()
    
    if choice == "1":
        results = explore_sidebar_pattern(config.MARATHON_URL, "Marathon")
    elif choice == "2":
        results = explore_sidebar_pattern(config.KEY_WEST_URL, "Key West")
    else:
        print("Invalid choice")
        return
    
    print(f"\n✅ Found {results.get('total_left_links', 0)} total left-side links")
    print(f"🎯 {len(results.get('category_links', []))} appeared to be category links")

if __name__ == "__main__":
    main()
