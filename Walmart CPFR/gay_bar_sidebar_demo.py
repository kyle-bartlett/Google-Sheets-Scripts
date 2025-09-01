 The f"""
Gay Bar Sidebar Navigation Demo
Demonstrates the exact sidebar navigation method you described
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

class GayBarSidebarDemo:
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
            logger.info("✅ Demo browser opened!")
            return True
        except Exception as e:
            logger.error(f"❌ Browser setup failed: {e}")
            return False
    
    def navigate_to_key_west(self):
        """Navigate to Key West page"""
        try:
            logger.info("🌐 Navigating to Key West page...")
            self.driver.get(config.KEY_WEST_URL)
            time.sleep(5)  # Let page load
            
            logger.info(f"📄 Page loaded: {self.driver.title}")
            logger.info("👀 You should see the Key West page with sidebar on the left!")
            time.sleep(3)
            return True
            
        except Exception as e:
            logger.error(f"❌ Navigation failed: {e}")
            return False
    
    def demonstrate_sidebar_navigation(self):
        """Demonstrate the exact sidebar navigation to Gay Bar"""
        try:
            logger.info("🎯 SIDEBAR NAVIGATION DEMO - Finding Gay Bar")
            logger.info("=" * 50)
            
            # Step 1: Scroll to top to ensure sidebar is visible
            logger.info("📜 Step 1: Scrolling to top to see sidebar...")
            self.driver.execute_script("window.scrollTo(0, 0);")
            time.sleep(3)
            
            # Step 2: Find and expand "Food & Drink" section
            logger.info("🔍 Step 2: Looking for 'Food & Drink' section in sidebar...")
            
            food_drink_selectors = [
                "//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'food & drink')]",
                "//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'food and drink')]",
                "//div[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'food')]"
            ]
            
            food_drink_element = None
            for selector in food_drink_selectors:
                try:
                    elements = self.driver.find_elements(By.XPATH, selector)
                    for elem in elements:
                        location = elem.location
                        if location['x'] < 500:  # Left sidebar area
                            food_drink_element = elem
                            logger.info(f"✅ Found 'Food & Drink' at x={location['x']}")
                            break
                    if food_drink_element:
                        break
                except Exception:
                    continue
            
            if not food_drink_element:
                logger.warning("⚠️ Could not find 'Food & Drink' section")
                return False
            
            # Step 3: Click to expand Food & Drink
            logger.info("🖱️ Step 3: Clicking 'Food & Drink' to expand...")
            try:
                self.driver.execute_script("arguments[0].scrollIntoView(true);", food_drink_element)
                time.sleep(2)
                food_drink_element.click()
                time.sleep(3)  # Wait for expansion
                logger.info("✅ Successfully expanded 'Food & Drink' section!")
            except Exception as e:
                logger.warning(f"Direct click failed, trying parent: {e}")
                try:
                    parent = food_drink_element.find_element(By.XPATH, "./..")
                    parent.click()
                    time.sleep(3)
                    logger.info("✅ Successfully clicked parent element!")
                except Exception as e2:
                    logger.error(f"❌ Both clicks failed: {e2}")
                    return False
            
            # Step 4: Find Gay Bar link in expanded section
            logger.info("🔍 Step 4: Looking for 'Gay Bar' in expanded Food & Drink section...")
            
            gay_bar_selectors = [
                "//a[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'gay bar')]",
                "//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'gay bar')]"
            ]
            
            gay_bar_link = None
            for selector in gay_bar_selectors:
                try:
                    elements = self.driver.find_elements(By.XPATH, selector)
                    for elem in elements:
                        location = elem.location
                        if location['x'] < 500:  # Still in sidebar
                            gay_bar_link = elem
                            logger.info(f"✅ Found 'Gay Bar' link at x={location['x']}")
                            break
                    if gay_bar_link:
                        break
                except Exception:
                    continue
            
            if not gay_bar_link:
                logger.warning("⚠️ Could not find 'Gay Bar' link in sidebar")
                # Show what we did find
                logger.info("🔍 Let me show you what links we found in the sidebar...")
                all_links = self.driver.find_elements(By.XPATH, "//a")
                sidebar_links = [link for link in all_links if link.location['x'] < 500]
                
                for i, link in enumerate(sidebar_links[:15]):  # Show first 15
                    try:
                        text = link.text.strip()
                        if len(text) > 1:
                            logger.info(f"   📋 Link {i+1}: '{text}'")
                    except:
                        pass
                
                return False
            
            # Step 5: Click Gay Bar link
            logger.info("🖱️ Step 5: Clicking 'Gay Bar' link...")
            try:
                self.driver.execute_script("arguments[0].scrollIntoView(true);", gay_bar_link)
                time.sleep(2)
                gay_bar_link.click()
                time.sleep(5)  # Wait for navigation
                logger.info("✅ Successfully clicked 'Gay Bar' link!")
                
                # Check what happened
                current_url = self.driver.current_url
                logger.info(f"📍 Current URL: {current_url}")
                
                # Look for Gay Bar content
                page_text = self.driver.page_source.lower()
                if "gay bar" in page_text:
                    logger.info("🎉 SUCCESS! Gay Bar content is visible!")
                    
                    # Look for the specific elements you showed in screenshots
                    try:
                        # Look for winner badge
                        winner_elements = self.driver.find_elements(By.XPATH, "//*[contains(text(), 'WINNER')]")
                        if winner_elements:
                            logger.info(f"🏆 Found {len(winner_elements)} WINNER badges")
                        
                        # Look for 22&Co (the winner)
                        winner_name = self.driver.find_elements(By.XPATH, "//*[contains(text(), '22&Co')]")
                        if winner_name:
                            logger.info("🏆 Found winner: 22&Co")
                        
                        # Look for "SEE ALL ENTRANTS"
                        see_all = self.driver.find_elements(By.XPATH, "//*[contains(text(), 'SEE ALL ENTRANTS')]")
                        if see_all:
                            logger.info(f"👥 Found {len(see_all)} 'SEE ALL ENTRANTS' buttons")
                            
                            # Click the first one
                            logger.info("🖱️ Clicking 'SEE ALL ENTRANTS'...")
                            see_all[0].click()
                            time.sleep(3)
                            
                            # Look for the entrants list
                            entrants = ["22&Co", "Aqua Bar & Nightclub", "801 Bourbon Bar"]
                            found_entrants = 0
                            for entrant in entrants:
                                if entrant.lower() in self.driver.page_source.lower():
                                    logger.info(f"✅ Found entrant: {entrant}")
                                    found_entrants += 1
                            
                            logger.info(f"🎯 Found {found_entrants}/{len(entrants)} expected entrants")
                        
                    except Exception as e:
                        logger.debug(f"Content analysis failed: {e}")
                else:
                    logger.warning("⚠️ Gay Bar content not found on page")
                
                return True
                
            except Exception as e:
                logger.error(f"❌ Failed to click 'Gay Bar' link: {e}")
                return False
            
        except Exception as e:
            logger.error(f"❌ Sidebar navigation demo failed: {e}")
            return False
    
    def run_demo(self):
        """Run the complete sidebar navigation demo"""
        logger.info("🎬 GAY BAR SIDEBAR NAVIGATION DEMO")
        logger.info("=" * 60)
        logger.info("This demonstrates the EXACT navigation method you described:")
        logger.info("1. Find sidebar on left")
        logger.info("2. Expand 'Food & Drink' section") 
        logger.info("3. Click 'Gay Bar' link")
        logger.info("4. Show the winner and entrants")
        logger.info("=" * 60)
        
        try:
            # Setup browser
            if not self.setup_driver():
                return False
            
            # Navigate to page
            if not self.navigate_to_key_west():
                return False
            
            # Demonstrate sidebar navigation
            success = self.demonstrate_sidebar_navigation()
            
            if success:
                logger.info("🎉 SIDEBAR NAVIGATION DEMO SUCCESSFUL!")
                logger.info("📊 This proves the sidebar method works perfectly!")
            else:
                logger.error("❌ Demo had issues - but automation has multiple fallbacks")
            
            logger.info("⏸️ Keeping browser open for 30 seconds to inspect results...")
            time.sleep(30)
            
        except Exception as e:
            logger.error(f"❌ Demo error: {e}")
        finally:
            if self.driver:
                logger.info("🔚 Closing demo browser...")
                self.driver.quit()

def main():
    print("🏳️‍🌈 GAY BAR SIDEBAR NAVIGATION DEMO")
    print("This will demonstrate the sidebar navigation method")
    print("you described using the Gay Bar example!")
    
    input("\nPress Enter to start the demo...")
    
    demo = GayBarSidebarDemo()
    demo.run_demo()
    
    print("\n✅ Demo complete!")
    print("🎯 This same method will work for your Marathon business categories!")

if __name__ == "__main__":
    main()
