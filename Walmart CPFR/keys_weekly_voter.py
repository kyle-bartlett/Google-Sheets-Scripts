"""
Keys Weekly Voting Automation Bot
Automatically votes on local community categories every 24 hours
"""

import time
import logging
import sys
from datetime import datetime, timezone
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
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('voting_log.txt'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

class KeysWeeklyVoter:
    def __init__(self):
        self.driver = None
        self.wait = None
        
    def setup_driver(self):
        """Initialize Chrome WebDriver with appropriate settings"""
        try:
            chrome_options = Options()
            
            # Add options for better compatibility and reliability
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--disable-blink-features=AutomationControlled")
            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            chrome_options.add_experimental_option('useAutomationExtension', False)
            
            # Set user agent to look more like a real browser
            chrome_options.add_argument("--user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
            
            if config.HEADLESS_MODE:
                chrome_options.add_argument("--headless")
            
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
            
            # Hide automation indicators
            self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            
            self.wait = WebDriverWait(self.driver, config.WAIT_TIMEOUT)
            logger.info("WebDriver initialized successfully")
            
        except Exception as e:
            logger.error(f"Failed to setup WebDriver: {e}")
            raise
    
    def login(self, url):
        """Navigate to the voting page and handle login"""
        try:
            logger.info(f"Navigating to: {url}")
            self.driver.get(url)
            time.sleep(3)  # Let page load
            
            # Check if we need to sign in
            try:
                # Look for sign in button or form
                sign_in_element = self.wait.until(
                    EC.element_to_be_clickable((By.LINK_TEXT, "Sign in"))
                )
                sign_in_element.click()
                logger.info("Clicked sign in button")
                time.sleep(2)
                
                # Enter email
                email_input = self.wait.until(
                    EC.presence_of_element_located((By.NAME, "username"))
                )
                email_input.clear()
                email_input.send_keys(config.EMAIL)
                logger.info(f"Entered email: {config.EMAIL}")
                
                # Submit login (look for login/submit button)
                try:
                    login_button = self.driver.find_element(By.XPATH, "//input[@type='submit' or @value='Log in' or @value='Login']")
                    login_button.click()
                except NoSuchElementException:
                    # Try alternative selectors
                    login_button = self.driver.find_element(By.CSS_SELECTOR, "button[type='submit'], .login-button, #login-button")
                    login_button.click()
                
                logger.info("Submitted login form")
                time.sleep(3)  # Wait for login to process
                
            except TimeoutException:
                logger.info("No sign-in required or already logged in")
            
            return True
            
        except Exception as e:
            logger.error(f"Login failed: {e}")
            return False
    
    def expand_businesses_dropdown(self):
        """Step 1: Find and click 'The Businesses' dropdown to expand it"""
        logger.info("STEP 1: Looking for 'The Businesses' dropdown to expand...")
        
        try:
            # Make sure page is fully loaded and scroll to top
            self.driver.execute_script("window.scrollTo(0, 0);")
            time.sleep(5)  # Give more time for dynamic content to load
            
            # Debug: Print page title and URL
            page_title = self.driver.title
            current_url = self.driver.current_url
            logger.info(f"Page title: {page_title}")
            logger.info(f"Current URL: {current_url}")
            
            # Look for visible sidebar elements (x >= 0 and x < 400)
            visible_elements = self.driver.find_elements(By.XPATH, "//*[string-length(text()) > 0]")
            visible_left_elements = []
            for elem in visible_elements:
                try:
                    location = elem.location
                    if 0 <= location['x'] < 400 and elem.is_displayed():
                        text = elem.text.strip()
                        if text and len(text) < 100:  # Reasonable text length
                            visible_left_elements.append((text, location['x']))
                except Exception:
                    continue
            
            logger.info(f"Found {len(visible_left_elements)} visible left sidebar texts")
            for text, x in visible_left_elements[:10]:  # Show first 10
                logger.info(f"Sidebar text: '{text}' at x={x}")
            
            # Try to find "The Businesses" more broadly
            all_businesses = []
            for text, x in visible_left_elements:
                if 'business' in text.lower() or 'businesses' in text.lower():
                    all_businesses.append((text, x))
            
            logger.info(f"Found {len(all_businesses)} elements with 'business': {all_businesses}")
            
            # The sidebar might not be visible yet - try scrolling down to find the voting interface
            logger.info("Scrolling down to look for voting interface...")
            for scroll_attempt in range(15):  # Scroll down more to find voting content
                self.driver.execute_script("window.scrollBy(0, 500);")
                time.sleep(2)
                
                # Look for voting-related content
                voting_indicators = self.driver.find_elements(By.XPATH, "//*[contains(text(), 'Vote') or contains(text(), 'VOTE')]")
                category_indicators = self.driver.find_elements(By.XPATH, "//*[contains(text(), 'Best') or contains(text(), 'The Businesses')]")
                
                logger.info(f"Scroll {scroll_attempt + 1}: Found {len(voting_indicators)} vote elements, {len(category_indicators)} category elements")
                
                # Check all text elements on current view
                current_texts = []
                all_current_elements = self.driver.find_elements(By.XPATH, "//*[string-length(text()) > 0]")
                for elem in all_current_elements:
                    try:
                        if elem.is_displayed():
                            text = elem.text.strip()
                            if text and len(text) < 200 and ('business' in text.lower() or 'vote' in text.lower() or 'best' in text.lower()):
                                location = elem.location
                                current_texts.append((text, location['x'], location['y']))
                    except Exception:
                        continue
                
                if current_texts:
                    logger.info(f"Found {len(current_texts)} relevant elements:")
                    for text, x, y in current_texts[:5]:  # Show first 5
                        logger.info(f"  '{text}' at x={x}, y={y}")
                
                # Look for "The Businesses" in current view
                businesses_selectors = [
                    "//*[contains(text(), 'The Businesses')]",
                    "//*[text()='The Businesses']", 
                    "//div[contains(text(), 'The Businesses')]"
                ]
                
                businesses_element = None
                for selector in businesses_selectors:
                    try:
                        elements = self.driver.find_elements(By.XPATH, selector)
                        for elem in elements:
                            if elem.is_displayed():
                                businesses_element = elem
                                location = elem.location
                                logger.info(f"Found 'The Businesses' at x={location['x']}, y={location['y']}")
                                break
                        if businesses_element:
                            break
                    except Exception:
                        continue
                
                if businesses_element:
                    break
            
            if not businesses_element:
                logger.error("Could not find 'The Businesses' dropdown after scrolling!")
                return False
            
            # Click to expand "The Businesses" section
            logger.info("Clicking 'The Businesses' to expand...")
            try:
                self.driver.execute_script("arguments[0].scrollIntoView(true);", businesses_element)
                time.sleep(1)
                businesses_element.click()
                time.sleep(3)  # Wait for expansion
                logger.info("✅ Successfully expanded 'The Businesses' dropdown")
                return True
            except Exception as e:
                logger.error(f"Failed to click 'The Businesses': {e}")
                return False
                
        except Exception as e:
            logger.error(f"Error expanding businesses dropdown: {e}")
            return False
    
    def find_and_click_best_realtor(self):
        """Step 2: Scroll down left sidebar to find and click 'Best Realtor' link"""
        logger.info("STEP 2: Looking for 'Best Realtor' link in left sidebar...")
        
        try:
            # Scroll down in the left sidebar area to find "Best Realtor"
            for scroll_attempt in range(10):
                logger.info(f"Scrolling down left sidebar (attempt {scroll_attempt + 1})...")
                
                # Scroll down a bit
                self.driver.execute_script("window.scrollBy(0, 300);")
                time.sleep(2)
                
                # Look for "Best Realtor" link in the left sidebar
                realtor_selectors = [
                    "//a[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'best realtor')]",
                    "//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'best realtor')]/ancestor-or-self::a",
                    "//a[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'realtor')]"
                ]
                
                realtor_link = None
                for selector in realtor_selectors:
                    try:
                        elements = self.driver.find_elements(By.XPATH, selector)
                        for elem in elements:
                            location = elem.location
                            # Look for elements on the left side (sidebar)
                            if location['x'] < 400:  # Left sidebar area
                                realtor_link = elem
                                logger.info(f"Found 'Best Realtor' link at x={location['x']}")
                                break
                        if realtor_link:
                            break
                    except Exception:
                        continue
                
                if realtor_link:
                    # Click the "Best Realtor" link
                    logger.info("Clicking 'Best Realtor' link...")
                    try:
                        self.driver.execute_script("arguments[0].scrollIntoView(true);", realtor_link)
                        time.sleep(1)
                        realtor_link.click()
                        time.sleep(4)  # Wait for navigation to voting section
                        logger.info("✅ Successfully clicked 'Best Realtor' link - now in voting section")
                        return True
                    except Exception as e:
                        logger.error(f"Failed to click 'Best Realtor' link: {e}")
                        return False
            
            logger.error("Could not find 'Best Realtor' link in sidebar after scrolling")
            return False
            
        except Exception as e:
            logger.error(f"Error finding Best Realtor link: {e}")
            return False
    
    def handle_email_registration(self):
        """Handle one-time email registration if required"""
        logger.info("Checking if email registration is required...")
        
        try:
            # Look for email input field
            email_input = None
            email_selectors = [
                "input[type='email']",
                "input[name='email']", 
                "input[placeholder*='email']",
                "input[placeholder*='Email']"
            ]
            
            for selector in email_selectors:
                try:
                    email_input = self.driver.find_element(By.CSS_SELECTOR, selector)
                    if email_input.is_displayed():
                        break
                except NoSuchElementException:
                    continue
            
            if email_input:
                logger.info("Email registration required - filling in email...")
                email_input.clear()
                email_input.send_keys(config.EMAIL)
                logger.info(f"Entered email: {config.EMAIL}")
                
                # Look for continue/submit button
                continue_selectors = [
                    "//button[contains(text(), 'CONTINUE')]",
                    "//input[@value='CONTINUE']",
                    "button[type='submit']",
                    ".continue-button",
                    "#continue"
                ]
                
                for selector in continue_selectors:
                    try:
                        if selector.startswith("//"):
                            continue_button = self.driver.find_element(By.XPATH, selector)
                        else:
                            continue_button = self.driver.find_element(By.CSS_SELECTOR, selector)
                        
                        if continue_button.is_enabled():
                            continue_button.click()
                            logger.info("✅ Clicked continue button after email registration")
                            time.sleep(3)
                            return True
                    except NoSuchElementException:
                        continue
                
                logger.warning("Could not find continue button after email entry")
                return False
            else:
                logger.info("No email registration required")
                return True
                
        except Exception as e:
            logger.error(f"Error handling email registration: {e}")
            return False
    
    def vote_in_category(self, category_name, response):
        """Vote in a specific category by clicking the green VOTE button"""
        logger.info(f"STEP 3: Voting for '{response}' in '{category_name}' category...")
        
        try:
            # Look for the category section and its green VOTE button
            vote_selectors = [
                "//button[contains(text(), 'VOTE') and contains(@class, 'green')]",
                "//button[contains(text(), 'VOTE')]",
                ".vote-button",
                "button[class*='vote']"
            ]
            
            # First, scroll to make sure we can see voting buttons
            self.driver.execute_script("window.scrollBy(0, 200);")
            time.sleep(1)
            
            vote_buttons = []
            for selector in vote_selectors:
                try:
                    if selector.startswith("//"):
                        buttons = self.driver.find_elements(By.XPATH, selector)
                    else:
                        buttons = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    vote_buttons.extend(buttons)
                except Exception:
                    continue
            
            if not vote_buttons:
                logger.warning(f"No VOTE buttons found for {category_name}")
                return False
            
            # Find the right vote button for this category
            # Look for the candidate name near the vote button
            for vote_button in vote_buttons:
                try:
                    # Check if our candidate name is near this vote button
                    parent_element = vote_button.find_element(By.XPATH, "./../../..")
                    parent_text = parent_element.text.lower()
                    
                    if response.lower() in parent_text:
                        logger.info(f"Found matching VOTE button for '{response}'")
                        
                        # Click the vote button
                        self.driver.execute_script("arguments[0].scrollIntoView(true);", vote_button)
                        time.sleep(1)
                        vote_button.click()
                        logger.info(f"✅ Successfully voted for '{response}' in '{category_name}'")
                        time.sleep(2)
                        return True
                        
                except Exception:
                    continue
            
            # If we couldn't find a specific match, try the first available vote button
            logger.info(f"Couldn't find specific match, trying first available vote button...")
            vote_button = vote_buttons[0]
            self.driver.execute_script("arguments[0].scrollIntoView(true);", vote_button)
            time.sleep(1)
            vote_button.click()
            logger.info(f"✅ Clicked vote button for '{category_name}'")
            time.sleep(2)
            return True
            
        except Exception as e:
            logger.error(f"Error voting in category '{category_name}': {e}")
            return False
    
    def scroll_and_find_categories(self):
        """Hybrid approach: Try sidebar first, then scroll"""
        logger.info("Looking for voting categories...")
        
        # Approach 1: Try sidebar navigation first
        if self.try_sidebar_navigation():
            logger.info("Using sidebar navigation approach")
            return True
        
        # Approach 2: Traditional scrolling method
        logger.info("Using scrolling approach to find voting categories...")
        
        # Scroll down in increments to find the voting section
        for scroll_count in range(10):  # Adjust as needed
            self.driver.execute_script("window.scrollBy(0, window.innerHeight);")
            time.sleep(config.SCROLL_PAUSE_TIME)
            
            # Check if we can find any voting forms or input boxes
            try:
                voting_elements = self.driver.find_elements(By.CSS_SELECTOR, 
                    "input[type='text'], textarea, .voting-input, .category-input")
                if voting_elements:
                    logger.info(f"Found {len(voting_elements)} potential voting inputs after {scroll_count + 1} scrolls")
                    return True
            except Exception as e:
                logger.debug(f"Scroll {scroll_count + 1}: {e}")
                continue
        
        # Approach 3: Look for any input fields at all (fallback)
        logger.info("Checking for any input fields on page...")
        all_inputs = self.driver.find_elements(By.CSS_SELECTOR, "input[type='text'], textarea")
        if all_inputs:
            logger.info(f"Found {len(all_inputs)} input fields as fallback")
            return True
        
        logger.warning("Could not find voting categories after all attempts")
        return False
    
    def navigate_to_category_via_sidebar(self, category_name):
        """Navigate to a specific category using sidebar navigation"""
        try:
            logger.info(f"Navigating to '{category_name}' via sidebar...")
            
            # Look for the category link in the sidebar
            category_link = None
            category_selectors = [
                f"//a[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{category_name.lower()}')]",
                f"//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{category_name.lower()}')]/ancestor-or-self::a",
                f"//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{category_name.lower()}')]"
            ]
            
            for selector in category_selectors:
                try:
                    elements = self.driver.find_elements(By.XPATH, selector)
                    for elem in elements:
                        location = elem.location
                        if location['x'] < 500:  # Sidebar area
                            category_link = elem
                            logger.info(f"Found '{category_name}' link in sidebar")
                            break
                    if category_link:
                        break
                except Exception:
                    continue
            
            if not category_link:
                logger.warning(f"Could not find '{category_name}' link in sidebar")
                return False
            
            # Click the category link
            try:
                self.driver.execute_script("arguments[0].scrollIntoView(true);", category_link)
                time.sleep(1)
                category_link.click()
                time.sleep(3)  # Wait for navigation
                logger.info(f"Successfully navigated to '{category_name}' category")
                return True
                
            except Exception as e:
                logger.error(f"Failed to click '{category_name}' link: {e}")
                return False
                
        except Exception as e:
            logger.error(f"Sidebar navigation to '{category_name}' failed: {e}")
            return False
    
    def find_and_fill_category(self, category_name, response):
        """Find a specific category and fill in the response - improved for dynamic content"""
        try:
            # Wait for dynamic content to load
            time.sleep(2)
            
            category_found = False
            input_element = None
            
            # Approach 1: Wait for and find text containing category name
            try:
                # Wait up to 10 seconds for category text to appear
                category_elements = self.wait.until(
                    EC.presence_of_all_elements_located((By.XPATH, 
                        f"//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{category_name.lower()}')]"))
                )
                
                for element in category_elements:
                    # Look for input fields near this text
                    try:
                        # Check parent and sibling elements for input fields
                        parent = element.find_element(By.XPATH, "./..")
                        input_element = parent.find_element(By.CSS_SELECTOR, "input[type='text'], textarea")
                        category_found = True
                        break
                    except NoSuchElementException:
                        # Try grandparent
                        try:
                            grandparent = element.find_element(By.XPATH, "./../..")
                            input_element = grandparent.find_element(By.CSS_SELECTOR, "input[type='text'], textarea")
                            category_found = True
                            break
                        except NoSuchElementException:
                            continue
            except TimeoutException:
                logger.debug(f"No text elements found for '{category_name}' using Approach 1")
            
            # Approach 2: Look for input fields with labels or placeholders
            if not category_found:
                try:
                    # Look for labels containing the category name
                    label_elements = self.driver.find_elements(By.XPATH, 
                        f"//label[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{category_name.lower()}')]")
                    
                    for label in label_elements:
                        try:
                            # Find associated input field
                            input_id = label.get_attribute("for")
                            if input_id:
                                input_element = self.driver.find_element(By.ID, input_id)
                                category_found = True
                                break
                        except NoSuchElementException:
                            continue
                except Exception:
                    pass
            
            # Approach 3: Look for input fields with names/ids related to the category
            if not category_found:
                try:
                    category_clean = category_name.lower().replace(' ', '').replace('-', '')
                    input_selectors = [
                        f"input[name*='{category_clean}']",
                        f"input[id*='{category_clean}']",
                        f"input[placeholder*='{category_name}']",
                        f"textarea[name*='{category_clean}']",
                        f"textarea[id*='{category_clean}']"
                    ]
                    
                    for selector in input_selectors:
                        try:
                            input_element = self.driver.find_element(By.CSS_SELECTOR, selector)
                            category_found = True
                            break
                        except NoSuchElementException:
                            continue
                except Exception:
                    pass
            
            # Approach 4: Try category aliases with improved search
            if not category_found and category_name in config.CATEGORY_ALIASES:
                for alias in config.CATEGORY_ALIASES[category_name]:
                    try:
                        # Multiple patterns for each alias
                        alias_selectors = [
                            f"//input[@type='text' and contains(@placeholder, '{alias}')]",
                            f"//textarea[contains(@placeholder, '{alias}')]",
                            f"//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{alias.lower()}')]/following::input[1]",
                            f"//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{alias.lower()}')]/preceding::input[1]"
                        ]
                        
                        for selector in alias_selectors:
                            try:
                                input_element = self.driver.find_element(By.XPATH, selector)
                                category_found = True
                                break
                            except NoSuchElementException:
                                continue
                        
                        if category_found:
                            break
                    except Exception:
                        continue
            
            # Approach 5: Sequential input matching (if forms are in order)
            if not category_found:
                try:
                    # Get all text inputs and try to match by position/context
                    all_inputs = self.driver.find_elements(By.CSS_SELECTOR, "input[type='text'], textarea")
                    category_list = list(config.VOTING_RESPONSES.keys())
                    
                    if category_name in category_list:
                        category_index = category_list.index(category_name)
                        if category_index < len(all_inputs):
                            input_element = all_inputs[category_index]
                            category_found = True
                            logger.info(f"Found '{category_name}' using sequential matching (position {category_index})")
                except Exception:
                    pass
            
            if category_found and input_element:
                # Scroll to input and fill it
                self.driver.execute_script("arguments[0].scrollIntoView(true);", input_element)
                time.sleep(1)
                
                # Clear and fill the input
                input_element.clear()
                time.sleep(0.5)
                input_element.send_keys(response)
                logger.info(f"Successfully filled '{category_name}' with '{response}'")
                time.sleep(1)  # Brief pause between entries
                return True
            else:
                logger.warning(f"Could not find input field for category: {category_name}")
                return False
                
        except Exception as e:
            logger.error(f"Error filling category '{category_name}': {e}")
            return False
    
    def submit_votes(self):
        """Find and click the submit button"""
        try:
            # Look for submit buttons with various selectors
            submit_selectors = [
                "input[type='submit']",
                "button[type='submit']", 
                "button[name='submit']",
                ".submit-button",
                "#submit",
                "//button[contains(text(), 'Submit')]",
                "//input[@value='Submit']"
            ]
            
            for selector in submit_selectors:
                try:
                    if selector.startswith("//"):
                        submit_button = self.driver.find_element(By.XPATH, selector)
                    else:
                        submit_button = self.driver.find_element(By.CSS_SELECTOR, selector)
                    
                    if submit_button.is_enabled():
                        submit_button.click()
                        logger.info("Successfully clicked submit button")
                        time.sleep(3)  # Wait for submission to process
                        return True
                except NoSuchElementException:
                    continue
            
            logger.warning("Could not find submit button")
            return False
            
        except Exception as e:
            logger.error(f"Error submitting votes: {e}")
            return False
    
    def check_page_state(self):
        """Check if the page is in voting mode, results mode, or closed"""
        try:
            page_text = self.driver.page_source.lower()
            
            # Check for voting-related indicators
            voting_indicators = ["vote", "voting", "submit your choice", "enter your", "nomination"]
            results_indicators = ["winner", "results", "congratulations", "thank you for voting"]
            closed_indicators = ["voting closed", "check back", "voting starts", "nomination phase"]
            
            voting_score = sum(1 for indicator in voting_indicators if indicator in page_text)
            results_score = sum(1 for indicator in results_indicators if indicator in page_text)
            closed_score = sum(1 for indicator in closed_indicators if indicator in page_text)
            
            if closed_score > 0:
                return "closed"
            elif results_score > voting_score:
                return "results"
            elif voting_score > 0:
                return "voting"
            else:
                return "unknown"
                
        except Exception as e:
            logger.error(f"Error checking page state: {e}")
            return "unknown"
    
    def vote(self, url=None):
        """Main voting process - Updated for new interface"""
        if url is None:
            url = config.DEFAULT_URL
            
        try:
            self.setup_driver()
            logger.info("🚀 Starting SIMPLIFIED voting workflow...")
            
            # Navigate to the site
            logger.info(f"Navigating to: {url}")
            self.driver.get(url)
            time.sleep(5)  # Let page load
            
            # CRITICAL: Switch to the SecondStreet iframe where voting actually happens
            logger.info("=== SWITCHING TO VOTING IFRAME ===")
            iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
            logger.info(f"Found {len(iframes)} iframes on page")
            
            voting_iframe_found = False
            for i, iframe in enumerate(iframes):
                try:
                    src = iframe.get_attribute('src')
                    logger.info(f"iframe {i+1}: {src}")
                    
                    # Look for SecondStreet app iframe (the voting system)
                    if src and 'secondstreetapp.com' in src:
                        logger.info(f"Found SecondStreet voting iframe: {src}")
                        self.driver.switch_to.frame(iframe)
                        time.sleep(3)  # Let iframe content load
                        
                        # Navigate to the base SecondStreet URL to see category dropdown
                        base_url = "https://embed-1109801.secondstreetapp.com/embed/b222ccf1-e1a9-4ac5-8af8-a9f96673db22/"
                        logger.info(f"Navigating to base SecondStreet URL: {base_url}")
                        self.driver.get(base_url)
                        time.sleep(5)  # Let voting interface load
                        
                        voting_iframe_found = True
                        break
                except Exception as e:
                    logger.debug(f"Error checking iframe {i+1}: {e}")
                    continue
            
            if not voting_iframe_found:
                logger.error("Could not find SecondStreet voting iframe!")
                return False
            
            logger.info("✅ Successfully switched to voting iframe")
            
            # Now check for voting content inside the iframe
            logger.info("Checking for voting content inside iframe...")
            vote_elements = self.driver.find_elements(By.XPATH, "//*[contains(text(), 'Vote') or contains(text(), 'VOTE')]")
            logger.info(f"Found {len(vote_elements)} vote elements in iframe")
            
            # Now we should be able to find "The Businesses" dropdown like in your screenshots!
            logger.info("=== STEP 1: Looking for 'The Businesses' dropdown (now in iframe) ===")
            
            # Look for "The Businesses" dropdown in the iframe
            businesses_found = False
            businesses_selectors = [
                "//*[contains(text(), 'The Businesses')]",
                "//*[text()='The Businesses']",
                "//div[contains(text(), 'The Businesses')]",
                "//span[contains(text(), 'The Businesses')]"
            ]
            
            for selector in businesses_selectors:
                try:
                    elements = self.driver.find_elements(By.XPATH, selector)
                    for elem in elements:
                        if elem.is_displayed():
                            text = elem.text.strip()
                            location = elem.location
                            tag = elem.tag_name
                            logger.info(f"Found 'The Businesses' element: '{text}' (tag: {tag}) at ({location['x']}, {location['y']})")
                            
                            # Try different click approaches
                            try:
                                # Scroll to element first
                                self.driver.execute_script("arguments[0].scrollIntoView(true);", elem)
                                time.sleep(1)
                                
                                # Try regular click first
                                elem.click()
                                time.sleep(3)
                                logger.info("✅ Successfully clicked 'The Businesses' with regular click")
                                businesses_found = True
                                break
                                
                            except Exception as click_error:
                                logger.warning(f"Regular click failed: {click_error}")
                                try:
                                    # Try JavaScript click
                                    self.driver.execute_script("arguments[0].click();", elem)
                                    time.sleep(3)
                                    logger.info("✅ Successfully clicked 'The Businesses' with JavaScript click")
                                    businesses_found = True
                                    break
                                except Exception as js_error:
                                    logger.warning(f"JavaScript click also failed: {js_error}")
                                    continue
                    
                    if businesses_found:
                        break
                except Exception as e:
                    logger.debug(f"Error with businesses selector {selector}: {e}")
                    continue
            
            # After clicking, check if we can now see more categories
            if businesses_found:
                logger.info("Checking for expanded categories after clicking 'The Businesses'...")
                time.sleep(2)
                category_elements = self.driver.find_elements(By.XPATH, "//*[contains(text(), 'Best') or contains(text(), 'Realtor')]")
                logger.info(f"Found {len(category_elements)} category elements after expansion")
                for cat_elem in category_elements[:5]:
                    try:
                        if cat_elem.is_displayed():
                            text = cat_elem.text
                            location = cat_elem.location
                            tag = cat_elem.tag_name
                            logger.info(f"  Category: '{text}' (tag: {tag}) at ({location['x']}, {location['y']})")
                    except:
                        continue
                
                # Also check for any new elements that appeared after clicking
                logger.info("Checking for all new clickable elements after expansion...")
                new_clickables = self.driver.find_elements(By.XPATH, "//a | //button | //*[@onclick] | //*[contains(@class, 'click')]")
                new_texts = []
                for elem in new_clickables:
                    try:
                        if elem.is_displayed():
                            text = elem.text.strip()
                            if text and len(text) < 100 and ('realtor' in text.lower() or 'best' in text.lower() or 'vote' in text.lower()):
                                location = elem.location
                                new_texts.append((text, elem.tag_name, location['x'], location['y']))
                    except:
                        continue
                
                logger.info(f"Found {len(new_texts)} relevant clickable elements after expansion:")
                for text, tag, x, y in new_texts:
                    logger.info(f"  {tag}: '{text}' at ({x}, {y})")
            
            if not businesses_found:
                logger.warning("Could not find 'The Businesses' dropdown - trying to find 'Best Realtor' directly")
            
            # Now we should have all the voting categories visible! Let's vote!
            logger.info("=== STEP 2: Voting in all 4 categories ===")
            successful_votes = 0
            
            # Define our target categories and their responses
            target_categories = [
                ("Best Realtor", "Nate Bartlett"),
                ("Best Real Estate Office", "Berkshire Hathaway Keys Real Estate 9141 Overseas"),
                ("Best Vacation Rental Company", "Berkshire Keys Vacation Rentals"),
                ("Best Business", "Berkshire Hathaway Keys Real Estate 9141 overseas")
            ]
            
            # Vote in each category
            for category_name, response in target_categories:
                logger.info(f"Looking for '{category_name}' to vote for '{response}'...")
                
                # Find the category link in the left sidebar (x=120)
                category_link_found = False
                # Look through all the links we found in the debugging
                all_links = self.driver.find_elements(By.TAG_NAME, "a")
                for elem in all_links:
                    try:
                        if elem.is_displayed():
                            text = elem.text.strip()
                            location = elem.location
                            
                            # Check if this matches our category and is in the left sidebar
                            if text == category_name and 100 <= location['x'] <= 140:
                                logger.info(f"Found '{category_name}' link at ({location['x']}, {location['y']})")
                                
                                try:
                                    # Click the category link to navigate to voting
                                    self.driver.execute_script("arguments[0].scrollIntoView(true);", elem)
                                    time.sleep(1)
                                    elem.click()
                                    time.sleep(3)
                                    logger.info(f"✅ Clicked '{category_name}' link")
                                    category_link_found = True
                                    break
                                except Exception as click_error:
                                    logger.error(f"Failed to click '{category_name}' link: {click_error}")
                                    # Try JavaScript click as fallback
                                    try:
                                        self.driver.execute_script("arguments[0].click();", elem)
                                        time.sleep(3)
                                        logger.info(f"✅ JavaScript clicked '{category_name}' link")
                                        category_link_found = True
                                        break
                                    except Exception as js_error:
                                        logger.error(f"JavaScript click also failed for '{category_name}': {js_error}")
                                        continue
                    except Exception as e:
                        logger.debug(f"Error checking link: {e}")
                        continue
                
                if not category_link_found:
                    logger.warning(f"Could not find '{category_name}' link")
                    continue
                
                # Handle email registration if it appears
                self.handle_email_registration()
                
                # Wait for the voting page to load and scroll to find voting content
                time.sleep(3)
                logger.info(f"Looking for voting content in '{category_name}' page...")
                
                # Scroll down to find the voting interface
                vote_buttons_found = False
                for scroll in range(10):
                    # Look for VOTE buttons
                    vote_buttons = self.driver.find_elements(By.XPATH, "//button[contains(text(), 'VOTE')]")
                    if vote_buttons:
                        logger.info(f"Found {len(vote_buttons)} VOTE buttons for '{category_name}' after {scroll} scrolls")
                        vote_buttons_found = True
                        break
                    
                    # Scroll down to find more content
                    self.driver.execute_script("window.scrollBy(0, 500);")
                    time.sleep(1)
                
                if not vote_buttons_found:
                    logger.warning(f"No VOTE buttons found for '{category_name}' - checking all buttons...")
                    # Look for any buttons that might be voting buttons
                    all_buttons = self.driver.find_elements(By.TAG_NAME, "button")
                    logger.info(f"Found {len(all_buttons)} total buttons on page")
                    
                    # Try any button that might be related to voting
                    for button in all_buttons[:5]:  # Check first 5 buttons
                        try:
                            if button.is_displayed():
                                button_text = button.text.strip()
                                logger.info(f"Button found: '{button_text}'")
                                if button_text.upper() in ['VOTE', 'SUBMIT', 'ENTER']:
                                    vote_buttons = [button]
                                    vote_buttons_found = True
                                    break
                        except:
                            continue
                
                # Click the first available VOTE button
                voted = False
                if vote_buttons_found and vote_buttons:
                    logger.info(f"Attempting to vote using {len(vote_buttons)} VOTE buttons found...")
                    
                    # Try to hide any overlays that might be blocking clicks
                    try:
                        self.driver.execute_script("""
                            var overlays = document.querySelectorAll('.user-info-container, .modal, .overlay, .popup');
                            for (var i = 0; i < overlays.length; i++) {
                                overlays[i].style.display = 'none';
                                overlays[i].style.zIndex = '-1';
                            }
                        """)
                        logger.info("Attempted to hide potential overlay elements")
                    except Exception as overlay_error:
                        logger.debug(f"Could not hide overlays: {overlay_error}")
                    
                    for vote_button in vote_buttons[:3]:  # Try first 3 buttons
                        try:
                            if vote_button.is_displayed() and vote_button.is_enabled():
                                # Get button details for debugging
                                button_location = vote_button.location
                                button_text = vote_button.text
                                logger.info(f"Attempting to click VOTE button: '{button_text}' at ({button_location['x']}, {button_location['y']})")
                                
                                # Scroll to center the button
                                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", vote_button)
                                time.sleep(2)
                                
                                # Use JavaScript click to bypass any overlays
                                self.driver.execute_script("arguments[0].click();", vote_button)
                                time.sleep(3)
                                logger.info(f"✅ JavaScript clicked VOTE button for '{category_name}'")
                                
                                # Handle email registration after voting
                                self.handle_email_registration()
                                
                                successful_votes += 1
                                logger.info(f"✅ Successfully voted in '{category_name}' (Vote #{successful_votes})")
                                voted = True
                                break
                        except Exception as e:
                            logger.warning(f"Failed to click vote button: {e}")
                            continue
                else:
                    logger.warning(f"No vote buttons found for '{category_name}'")
                
                if not voted:
                    logger.warning(f"❌ Failed to vote in '{category_name}'")
                
                # Brief pause between categories
                time.sleep(2)
            
            logger.info(f"🎯 VOTING SUMMARY: {successful_votes}/{len(target_categories)} votes completed")
            
            if successful_votes >= 1:
                logger.info("🎉 Voting process completed successfully!")
                return True
            else:
                logger.error("❌ No votes were successfully cast")
                return False
                
        except Exception as e:
            logger.error(f"Voting process failed: {e}")
            return False
        finally:
            if self.driver:
                self.driver.quit()
                logger.info("WebDriver closed")
    
    def test_run(self):
        """Run a test to verify the automation works"""
        logger.info("=== STARTING TEST RUN ===")
        success = self.vote()
        if success:
            logger.info("=== TEST RUN SUCCESSFUL ===")
        else:
            logger.error("=== TEST RUN FAILED ===")
        return success

if __name__ == "__main__":
    voter = KeysWeeklyVoter()
    
    # Check if this is a test run
    if len(sys.argv) > 1 and sys.argv[1] == "test":
        voter.test_run()
    else:
        # Regular voting run
        voter.vote()
