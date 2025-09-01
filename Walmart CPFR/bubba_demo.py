"""
Bubba's Key West Demo - Custom demo showing real interaction
Navigates to gay bar category, clicks "SEE ALL ENTRANTS", pauses, then finds another category
"""

import time
import logging
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import config

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')
logger = logging.getLogger(__name__)

class BubbaDemo:
    def __init__(self):
        self.driver = None
        self.wait = None
        
    def setup_driver(self):
        """Setup Chrome driver for demo"""
        chrome_options = Options()
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        # Visible browser for demo
        
        try:
            self.driver = webdriver.Chrome(options=chrome_options)
            self.driver.maximize_window()
            self.wait = WebDriverWait(self.driver, 20)
            logger.info("✅ Demo browser opened and maximized!")
            return True
        except Exception as e:
            logger.error(f"❌ Failed to setup browser: {e}")
            return False
    
    def navigate_to_bubba(self):
        """Navigate to Bubba's Key West page"""
        try:
            logger.info("🌐 Navigating to Bubba's Best of Key West...")
            self.driver.get(config.KEY_WEST_URL)
            time.sleep(5)  # Let page load
            
            logger.info(f"📄 Page loaded: {self.driver.title}")
            logger.info("👀 You should see the Bubba's Best of Key West page!")
            time.sleep(3)
            return True
            
        except Exception as e:
            logger.error(f"❌ Failed to navigate: {e}")
            return False
    
    def find_gay_bar_category(self):
        """Scroll down and find the gay bar category"""
        logger.info("🔍 Looking for the 'gay bar' category...")
        
        # Scroll down gradually looking for gay bar
        for scroll_attempt in range(15):
            logger.info(f"📜 Scrolling... attempt {scroll_attempt + 1}")
            
            # Look for text containing "gay bar" (case insensitive)
            try:
                gay_bar_elements = self.driver.find_elements(By.XPATH, 
                    "//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'gay bar')]")
                
                if gay_bar_elements:
                    logger.info("✅ Found 'gay bar' category!")
                    # Scroll element into view
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", gay_bar_elements[0])
                    time.sleep(2)
                    logger.info("🎯 Scrolled to gay bar category section")
                    return True
                    
            except Exception as e:
                logger.debug(f"Search attempt {scroll_attempt + 1} failed: {e}")
            
            # Scroll down
            self.driver.execute_script("window.scrollBy(0, 400);")
            time.sleep(1)
        
        logger.warning("⚠️ Could not find 'gay bar' category after scrolling")
        return False
    
    def click_see_all_entrants(self, category_context="gay bar"):
        """Find and click 'SEE ALL ENTRANTS' button"""
        logger.info(f"🔍 Looking for 'SEE ALL ENTRANTS' button near {category_context}...")
        
        try:
            # Look for "SEE ALL ENTRANTS" text
            see_all_elements = self.driver.find_elements(By.XPATH, 
                "//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'see all entrants')]")
            
            if see_all_elements:
                logger.info(f"✅ Found {len(see_all_elements)} 'SEE ALL ENTRANTS' buttons")
                
                # Try to click the first one
                target_button = see_all_elements[0]
                logger.info("🖱️ Clicking 'SEE ALL ENTRANTS' button...")
                
                # Scroll to button and click
                self.driver.execute_script("arguments[0].scrollIntoView(true);", target_button)
                time.sleep(1)
                target_button.click()
                
                logger.info("✅ Clicked 'SEE ALL ENTRANTS' successfully!")
                time.sleep(3)  # Wait for content to load
                return True
            else:
                logger.warning("⚠️ Could not find 'SEE ALL ENTRANTS' button")
                return False
                
        except Exception as e:
            logger.error(f"❌ Error clicking 'SEE ALL ENTRANTS': {e}")
            return False
    
    def show_entrants(self):
        """Display the entrants that are now visible"""
        logger.info("👥 Looking for the entrants list...")
        
        try:
            # Look for common patterns that might show entrants
            # Check for lists, divs with names, etc.
            possible_entrants = []
            
            # Look for elements that might contain entrant names
            entrant_selectors = [
                "div[class*='entrant']",
                "li[class*='entry']", 
                "div[class*='nominee']",
                "div[class*='candidate']",
                ".entry-title",
                ".entrant-name"
            ]
            
            for selector in entrant_selectors:
                elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                if elements:
                    logger.info(f"✅ Found {len(elements)} elements with selector: {selector}")
                    for i, elem in enumerate(elements[:5]):  # Show first 5
                        try:
                            text = elem.text.strip()
                            if text and len(text) < 100:  # Reasonable name length
                                possible_entrants.append(text)
                                logger.info(f"   {i+1}. {text}")
                        except:
                            pass
            
            if possible_entrants:
                logger.info(f"🎉 Found {len(possible_entrants)} potential entrants!")
                logger.info("⏸️ PAUSING for 20 seconds so you can see the entrants...")
                time.sleep(20)
                return True
            else:
                logger.info("ℹ️ Entrants may be visible but in a different format")
                logger.info("⏸️ PAUSING for 20 seconds so you can see what's displayed...")
                time.sleep(20)
                return True
                
        except Exception as e:
            logger.error(f"❌ Error showing entrants: {e}")
            logger.info("⏸️ PAUSING for 20 seconds anyway...")
            time.sleep(20)
            return False
    
    def find_another_category(self):
        """Scroll up and find another category with 'SEE ALL ENTRANTS'"""
        logger.info("🔄 Now scrolling up to find another category...")
        
        # Scroll up gradually
        for scroll_attempt in range(10):
            logger.info(f"⬆️ Scrolling up... attempt {scroll_attempt + 1}")
            
            self.driver.execute_script("window.scrollBy(0, -500);")
            time.sleep(2)
            
            # Look for other "SEE ALL ENTRANTS" buttons
            try:
                see_all_elements = self.driver.find_elements(By.XPATH, 
                    "//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'see all entrants')]")
                
                if len(see_all_elements) > 1:
                    logger.info(f"✅ Found {len(see_all_elements)} 'SEE ALL ENTRANTS' buttons")
                    
                    # Try to click a different one (second in list)
                    target_button = see_all_elements[1] if len(see_all_elements) > 1 else see_all_elements[0]
                    
                    # Get nearby text to understand what category this is
                    try:
                        parent = target_button.find_element(By.XPATH, "./../..")
                        category_text = parent.text[:100]  # First 100 chars
                        logger.info(f"🎯 Found another category: {category_text}")
                    except:
                        logger.info("🎯 Found another 'SEE ALL ENTRANTS' button")
                    
                    # Scroll to and click
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", target_button)
                    time.sleep(1)
                    target_button.click()
                    
                    logger.info("✅ Clicked second 'SEE ALL ENTRANTS' button!")
                    time.sleep(5)
                    return True
                    
            except Exception as e:
                logger.debug(f"Scroll up attempt {scroll_attempt + 1} failed: {e}")
        
        logger.warning("⚠️ Could not find another 'SEE ALL ENTRANTS' button")
        return False
    
    def run_bubba_demo(self):
        """Run the complete Bubba's demo"""
        logger.info("🎬 BUBBA'S KEY WEST CUSTOM DEMO")
        logger.info("=" * 60)
        logger.info("🎯 This demo will:")
        logger.info("   1. Navigate to Bubba's Best of Key West")
        logger.info("   2. Find the 'gay bar' category")
        logger.info("   3. Click 'SEE ALL ENTRANTS'")
        logger.info("   4. Show the 3 names and pause 20 seconds")
        logger.info("   5. Scroll up and find another category")
        logger.info("   6. Click another 'SEE ALL ENTRANTS'")
        logger.info("=" * 60)
        
        try:
            # Step 1: Setup browser
            if not self.setup_driver():
                return False
            
            # Step 2: Navigate to page
            if not self.navigate_to_bubba():
                return False
            
            # Step 3: Find gay bar category
            if not self.find_gay_bar_category():
                logger.info("⚠️ Continuing demo even though gay bar category wasn't found...")
            
            # Step 4: Click SEE ALL ENTRANTS
            if self.click_see_all_entrants("gay bar"):
                # Step 5: Show entrants and pause
                self.show_entrants()
            else:
                logger.info("⚠️ Trying to find any 'SEE ALL ENTRANTS' button...")
                # Try to find any "SEE ALL ENTRANTS" button
                see_all_elements = self.driver.find_elements(By.XPATH, 
                    "//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'see all')]")
                if see_all_elements:
                    see_all_elements[0].click()
                    logger.info("✅ Clicked a 'SEE ALL' button")
                    time.sleep(20)
            
            # Step 6: Find another category
            self.find_another_category()
            
            logger.info("🎉 Bubba's demo complete!")
            logger.info("⏸️ Keeping browser open for 10 more seconds...")
            time.sleep(10)
            
        except Exception as e:
            logger.error(f"❌ Demo error: {e}")
        finally:
            if self.driver:
                logger.info("🔚 Closing browser...")
                self.driver.quit()

def main():
    print("🎬 BUBBA'S KEY WEST CUSTOM DEMO")
    print("👀 This will show real interaction with the voting page!")
    print("⏳ Demo will take about 60-90 seconds")
    print("🎯 You'll see it find categories and click 'SEE ALL ENTRANTS'")
    
    input("\nPress Enter to start the Bubba's demo...")
    
    demo = BubbaDemo()
    demo.run_bubba_demo()
    
    print("\n🎉 Demo complete!")
    print("💡 This shows the same type of interaction your automation will do on Marathon!")

if __name__ == "__main__":
    main()
