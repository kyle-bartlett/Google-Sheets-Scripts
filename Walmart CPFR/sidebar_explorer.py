"""
Sidebar Explorer - Investigate the left sidebar navigation
This should be much more reliable than scrolling through main content
"""

import time
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import config

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')
logger = logging.getLogger(__name__)

class SidebarExplorer:
    def __init__(self):
        self.driver = None
        self.wait = None
        
    def setup_driver(self):
        """Setup Chrome driver"""
        chrome_options = Options()
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        
        try:
            self.driver = webdriver.Chrome(options=chrome_options)
            self.driver.maximize_window()
            self.wait = WebDriverWait(self.driver, 20)
            logger.info("✅ Browser ready for sidebar exploration!")
            return True
        except Exception as e:
            logger.error(f"❌ Browser setup failed: {e}")
            return False
    
    def navigate_and_explore_sidebar(self, url, page_name):
        """Navigate to page and explore the sidebar"""
        try:
            logger.info(f"🌐 Navigating to {page_name}...")
            self.driver.get(url)
            time.sleep(5)  # Let page load completely
            
            logger.info(f"📄 Page loaded: {self.driver.title}")
            
            # Look for sidebar elements on the left side
            logger.info("🔍 Looking for sidebar navigation on the left side...")
            
            # Common sidebar selectors
            sidebar_selectors = [
                "aside",
                ".sidebar", 
                ".navigation",
                ".nav",
                "#sidebar",
                ".left-nav",
                ".categories",
                ".category-list",
                "[class*='sidebar']",
                "[class*='nav']",
                "[id*='sidebar']",
                "[id*='nav']"
            ]
            
            sidebar_found = False
            sidebar_element = None
            
            for selector in sidebar_selectors:
                try:
                    elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    if elements:
                        logger.info(f"✅ Found potential sidebar with selector: {selector}")
                        for elem in elements:
                            # Check if it's on the left side (rough estimate)
                            location = elem.location
                            size = elem.size
                            if location['x'] < 300 and size['height'] > 200:  # Left side and decent height
                                logger.info(f"🎯 Found left sidebar at position x={location['x']}, height={size['height']}")
                                sidebar_element = elem
                                sidebar_found = True
                                break
                        if sidebar_found:
                            break
                except Exception as e:
                    logger.debug(f"Selector {selector} failed: {e}")
            
            if not sidebar_found:
                logger.info("🔍 No obvious sidebar found, scanning left side of page...")
                # Look for any elements on the left side that might contain links
                all_elements = self.driver.find_elements(By.XPATH, "//*")
                left_side_elements = []
                
                for elem in all_elements:
                    try:
                        location = elem.location
                        size = elem.size
                        if location['x'] < 300 and size['width'] > 50 and size['height'] > 20:
                            left_side_elements.append(elem)
                    except:
                        continue
                
                logger.info(f"📊 Found {len(left_side_elements)} elements on left side")
            
            # Look for category links specifically
            logger.info("🔍 Looking for category links...")
            
            # Search for your specific categories in links
            target_categories = list(config.VOTING_RESPONSES.keys())
            found_categories = {}
            
            for category in target_categories:
                logger.info(f"🔍 Searching for '{category}' in sidebar links...")
                
                # Look for links containing the category name
                link_selectors = [
                    f"//a[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{category.lower()}')]",
                    f"//a[contains(@href, '{category.lower().replace(' ', '')}')]",
                    f"//a[contains(@href, '{category.lower().replace(' ', '-')}')]",
                    f"//li[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{category.lower()}')]/a",
                    f"//div[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{category.lower()}')]/a"
                ]
                
                for selector in link_selectors:
                    try:
                        links = self.driver.find_elements(By.XPATH, selector)
                        if links:
                            for link in links:
                                try:
                                    # Check if link is on left side
                                    location = link.location
                                    if location['x'] < 400:  # Left side of page
                                        href = link.get_attribute('href')
                                        text = link.text.strip()
                                        logger.info(f"✅ Found '{category}' link: '{text}' -> {href}")
                                        found_categories[category] = {
                                            'element': link,
                                            'text': text,
                                            'href': href,
                                            'location': location
                                        }
                                        break
                                except:
                                    continue
                        if category in found_categories:
                            break
                    except Exception as e:
                        logger.debug(f"Selector failed for {category}: {e}")
            
            # Show summary
            logger.info("📊 SIDEBAR EXPLORATION SUMMARY:")
            logger.info(f"✅ Found {len(found_categories)}/{len(target_categories)} target categories in sidebar")
            
            for category, info in found_categories.items():
                logger.info(f"   🎯 {category}: '{info['text']}' at x={info['location']['x']}")
            
            # Test clicking one if found
            if found_categories:
                test_category = list(found_categories.keys())[0]
                logger.info(f"🖱️ Testing click on '{test_category}' category...")
                
                try:
                    test_link = found_categories[test_category]['element']
                    
                    # Scroll to link and click
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", test_link)
                    time.sleep(2)
                    test_link.click()
                    
                    logger.info("✅ Successfully clicked category link!")
                    time.sleep(5)  # Wait for navigation
                    
                    # Check if we moved to the right section
                    current_url = self.driver.current_url
                    logger.info(f"📍 Current URL after click: {current_url}")
                    
                    # Look for voting form or category content
                    inputs = self.driver.find_elements(By.CSS_SELECTOR, "input[type='text'], textarea")
                    logger.info(f"📝 Found {len(inputs)} input fields after navigation")
                    
                except Exception as e:
                    logger.error(f"❌ Click test failed: {e}")
            
            # Keep browser open for inspection
            logger.info("⏸️ Keeping browser open for 20 seconds for manual inspection...")
            time.sleep(20)
            
            return found_categories
            
        except Exception as e:
            logger.error(f"❌ Sidebar exploration failed: {e}")
            return {}
    
    def explore_both_pages(self):
        """Explore sidebar on both pages"""
        results = {}
        
        # Explore Marathon page
        logger.info("🏃 EXPLORING MARATHON PAGE SIDEBAR")
        logger.info("=" * 50)
        try:
            marathon_results = self.navigate_and_explore_sidebar(config.MARATHON_URL, "Marathon")
            results['marathon'] = marathon_results
        except Exception as e:
            logger.error(f"Marathon exploration failed: {e}")
            results['marathon'] = {}
        
        time.sleep(2)
        
        # Explore Key West page
        logger.info("\n🏝️ EXPLORING KEY WEST PAGE SIDEBAR")
        logger.info("=" * 50)
        try:
            keywest_results = self.navigate_and_explore_sidebar(config.KEY_WEST_URL, "Key West")
            results['keywest'] = keywest_results
        except Exception as e:
            logger.error(f"Key West exploration failed: {e}")
            results['keywest'] = {}
        
        return results

def main():
    print("🔍 SIDEBAR NAVIGATION EXPLORER")
    print("This will find the category links in the left sidebar")
    print("Much more reliable than scrolling through main content!")
    
    choice = input("\nWhich page to explore?\n1. Marathon (your target)\n2. Key West (active)\n3. Both pages\nEnter 1-3: ").strip()
    
    explorer = SidebarExplorer()
    if not explorer.setup_driver():
        return
    
    try:
        if choice == "1":
            results = explorer.navigate_and_explore_sidebar(config.MARATHON_URL, "Marathon")
        elif choice == "2":
            results = explorer.navigate_and_explore_sidebar(config.KEY_WEST_URL, "Key West")
        elif choice == "3":
            results = explorer.explore_both_pages()
        else:
            print("Invalid choice")
            return
        
        print("\n✅ Sidebar exploration complete!")
        print("Check the browser window and logs for details")
        
    finally:
        if explorer.driver:
            explorer.driver.quit()

if __name__ == "__main__":
    main()
