"""
Page Explorer for Keys Weekly - Analyze page structure
This helps us understand what's available on the voting pages
"""

import time
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import config

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class PageExplorer:
    def __init__(self):
        self.driver = None
        self.wait = None
        
    def setup_driver(self):
        """Initialize Chrome WebDriver"""
        try:
            chrome_options = Options()
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--disable-blink-features=AutomationControlled")
            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            chrome_options.add_experimental_option('useAutomationExtension', False)
            
            # Try to use system Chrome first, then fall back to webdriver-manager
            try:
                # For M2 Mac, try using system Chrome driver path
                self.driver = webdriver.Chrome(options=chrome_options)
                logger.info("Using system Chrome driver")
            except Exception:
                # Fall back to webdriver-manager
                try:
                    service = Service(ChromeDriverManager().install())
                    self.driver = webdriver.Chrome(service=service, options=chrome_options)
                    logger.info("Using webdriver-manager Chrome driver")
                except Exception:
                    # Final fallback - try manual path
                    service = Service("/usr/local/bin/chromedriver")
                    self.driver = webdriver.Chrome(service=service, options=chrome_options)
                    logger.info("Using manual Chrome driver path")
            
            self.wait = WebDriverWait(self.driver, config.WAIT_TIMEOUT)
            logger.info("WebDriver initialized successfully")
            
        except Exception as e:
            logger.error(f"Failed to setup WebDriver: {e}")
            raise
    
    def login_and_explore(self, url):
        """Login and explore the page structure"""
        try:
            logger.info(f"🌐 Exploring: {url}")
            self.driver.get(url)
            time.sleep(3)
            
            # Check current page title and URL
            logger.info(f"📄 Page title: {self.driver.title}")
            logger.info(f"🔗 Current URL: {self.driver.current_url}")
            
            # Try to login
            self.attempt_login()
            
            # Explore page structure
            self.explore_page_structure()
            
            # Look for voting-related elements
            self.look_for_voting_elements()
            
            return True
            
        except Exception as e:
            logger.error(f"Exploration failed: {e}")
            return False
        finally:
            if self.driver:
                input("Press Enter to close browser and continue...")
                self.driver.quit()
    
    def attempt_login(self):
        """Try to login"""
        try:
            # Look for sign in button
            sign_in_element = self.driver.find_element(By.LINK_TEXT, "Sign in")
            logger.info("🔐 Found 'Sign in' button, clicking...")
            sign_in_element.click()
            time.sleep(2)
            
            # Try to enter email
            email_input = self.wait.until(
                EC.presence_of_element_located((By.NAME, "username"))
            )
            email_input.clear()
            email_input.send_keys(config.EMAIL)
            logger.info(f"✉️ Entered email: {config.EMAIL}")
            
            # Look for and click login button
            login_buttons = self.driver.find_elements(By.XPATH, 
                "//input[@type='submit'] | //button[@type='submit'] | //button[contains(text(), 'Log')]")
            
            if login_buttons:
                login_buttons[0].click()
                logger.info("🚀 Clicked login button")
                time.sleep(3)
            
        except Exception as e:
            logger.info(f"Login attempt: {e} (this might be expected)")
    
    def explore_page_structure(self):
        """Analyze the overall page structure"""
        logger.info("\n🔍 EXPLORING PAGE STRUCTURE:")
        
        # Scroll and analyze
        scroll_height = self.driver.execute_script("return document.body.scrollHeight")
        logger.info(f"📏 Total page height: {scroll_height}px")
        
        # Find all forms
        forms = self.driver.find_elements(By.TAG_NAME, "form")
        logger.info(f"📝 Found {len(forms)} forms on page")
        
        # Find all input elements
        inputs = self.driver.find_elements(By.TAG_NAME, "input")
        logger.info(f"⌨️ Found {len(inputs)} input elements")
        
        # Find text inputs specifically
        text_inputs = self.driver.find_elements(By.CSS_SELECTOR, "input[type='text'], input[type='email'], textarea")
        logger.info(f"📝 Found {len(text_inputs)} text input fields")
        
        # Find buttons
        buttons = self.driver.find_elements(By.TAG_NAME, "button")
        logger.info(f"🔘 Found {len(buttons)} buttons")
        
        # Find divs with common voting-related classes/ids
        voting_containers = self.driver.find_elements(By.CSS_SELECTOR, 
            "[class*='vote'], [class*='category'], [class*='poll'], [id*='vote'], [id*='category']")
        logger.info(f"🗳️ Found {len(voting_containers)} potential voting containers")
    
    def look_for_voting_elements(self):
        """Look for elements related to voting"""
        logger.info("\n🗳️ LOOKING FOR VOTING ELEMENTS:")
        
        # Scroll through the page
        for scroll_num in range(8):
            logger.info(f"📜 Scrolling section {scroll_num + 1}...")
            self.driver.execute_script("window.scrollBy(0, window.innerHeight);")
            time.sleep(2)
            
            # Look for text containing our categories
            for category in config.VOTING_RESPONSES.keys():
                try:
                    elements = self.driver.find_elements(By.XPATH, 
                        f"//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{category.lower()}')]")
                    if elements:
                        logger.info(f"✅ Found text containing '{category}' in {len(elements)} elements")
                        
                        # Check if there are nearby input fields
                        for element in elements[:3]:  # Check first 3 matches
                            try:
                                parent = element.find_element(By.XPATH, "./..")
                                nearby_inputs = parent.find_elements(By.CSS_SELECTOR, "input, textarea")
                                if nearby_inputs:
                                    logger.info(f"  📝 Found {len(nearby_inputs)} input fields near '{category}' text")
                            except:
                                pass
                except Exception as e:
                    continue
        
        # Look for common voting patterns
        voting_patterns = [
            "vote", "submit", "nomination", "poll", "survey", "contest", "winner", "finalist"
        ]
        
        for pattern in voting_patterns:
            elements = self.driver.find_elements(By.XPATH, 
                f"//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{pattern}')]")
            if elements:
                logger.info(f"🔍 Found {len(elements)} elements containing '{pattern}'")
    
    def explore_both_pages(self):
        """Explore both Marathon and Key West pages"""
        logger.info("🎯 EXPLORING BOTH VOTING PAGES")
        
        # Explore Key West page
        logger.info("\n" + "="*50)
        logger.info("🏝️ EXPLORING KEY WEST PAGE (BUBBA'S BEST)")
        logger.info("="*50)
        self.setup_driver()
        self.login_and_explore(config.KEY_WEST_URL)
        
        # Small break
        time.sleep(2)
        
        # Explore Marathon page  
        logger.info("\n" + "="*50)
        logger.info("🏃 EXPLORING MARATHON PAGE (BEST OF MARATHON)")
        logger.info("="*50)
        self.setup_driver()
        self.login_and_explore(config.MARATHON_URL)

if __name__ == "__main__":
    explorer = PageExplorer()
    
    print("🔍 Page Structure Explorer")
    print("This will help us understand what's available on the voting pages")
    print("\nOptions:")
    print("1. Explore Key West page (Bubba's Best)")
    print("2. Explore Marathon page (Best of Marathon)")
    print("3. Explore both pages")
    
    choice = input("\nChoose option (1-3): ").strip()
    
    explorer.setup_driver()
    
    if choice == "1":
        explorer.login_and_explore(config.KEY_WEST_URL)
    elif choice == "2":
        explorer.login_and_explore(config.MARATHON_URL)
    elif choice == "3":
        explorer.explore_both_pages()
    else:
        print("Invalid choice")
        
    print("\n✅ Exploration complete! Check the logs above for details.")
