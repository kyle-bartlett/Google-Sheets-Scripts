"""
Comprehensive Test Suite for Keys Weekly Voting Automation
Tests everything we can before voting actually opens
"""

import time
import logging
import sys
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import config

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('comprehensive_test_log.txt'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

class ComprehensiveTest:
    def __init__(self):
        self.driver = None
        self.wait = None
        self.test_results = {}
        
    def setup_driver(self):
        """Initialize Chrome WebDriver"""
        try:
            chrome_options = Options()
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--disable-blink-features=AutomationControlled")
            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            chrome_options.add_experimental_option('useAutomationExtension', False)
            
            # Run visible for testing
            # chrome_options.add_argument("--headless")  # Commented out for testing
            
            try:
                self.driver = webdriver.Chrome(options=chrome_options)
                logger.info("✅ WebDriver initialized successfully")
                self.test_results['webdriver_setup'] = True
            except Exception as e:
                logger.error(f"❌ WebDriver setup failed: {e}")
                self.test_results['webdriver_setup'] = False
                return False
            
            self.wait = WebDriverWait(self.driver, config.WAIT_TIMEOUT)
            return True
            
        except Exception as e:
            logger.error(f"❌ Driver setup failed: {e}")
            self.test_results['webdriver_setup'] = False
            return False
    
    def test_website_access(self, url, page_name):
        """Test basic website access"""
        try:
            logger.info(f"🌐 Testing access to {page_name}...")
            self.driver.get(url)
            time.sleep(3)
            
            # Check if page loaded
            if self.driver.title:
                logger.info(f"✅ {page_name} loaded successfully")
                logger.info(f"   📄 Title: {self.driver.title}")
                self.test_results[f'{page_name.lower().replace(" ", "_")}_access'] = True
                return True
            else:
                logger.error(f"❌ {page_name} failed to load properly")
                self.test_results[f'{page_name.lower().replace(" ", "_")}_access'] = False
                return False
                
        except Exception as e:
            logger.error(f"❌ Error accessing {page_name}: {e}")
            self.test_results[f'{page_name.lower().replace(" ", "_")}_access'] = False
            return False
    
    def test_login_process(self):
        """Test the login process"""
        try:
            logger.info("🔐 Testing login process...")
            
            # Look for sign in button
            try:
                sign_in_element = self.wait.until(
                    EC.element_to_be_clickable((By.LINK_TEXT, "Sign in"))
                )
                sign_in_element.click()
                logger.info("✅ Found and clicked 'Sign in' button")
                time.sleep(2)
                
                # Try to find email input
                try:
                    email_input = self.wait.until(
                        EC.presence_of_element_located((By.NAME, "username"))
                    )
                    email_input.clear()
                    email_input.send_keys(config.EMAIL)
                    logger.info(f"✅ Successfully entered email: {config.EMAIL}")
                    
                    # Look for submit button
                    try:
                        submit_buttons = self.driver.find_elements(By.XPATH, 
                            "//input[@type='submit'] | //button[@type='submit'] | //button[contains(text(), 'Log')]")
                        
                        if submit_buttons:
                            submit_buttons[0].click()
                            logger.info("✅ Clicked login submit button")
                            time.sleep(3)
                            
                            # Check if login was successful
                            current_url = self.driver.current_url
                            if "sign" not in current_url.lower():
                                logger.info("✅ Login appears successful (redirected away from sign-in)")
                                self.test_results['login_process'] = True
                                return True
                            else:
                                logger.warning("⚠️ Still on sign-in page, login may have failed")
                                self.test_results['login_process'] = False
                                return False
                        else:
                            logger.warning("⚠️ Could not find submit button")
                            self.test_results['login_process'] = False
                            return False
                            
                    except Exception as e:
                        logger.warning(f"⚠️ Submit button interaction failed: {e}")
                        self.test_results['login_process'] = False
                        return False
                        
                except TimeoutException:
                    logger.warning("⚠️ Could not find email input field")
                    self.test_results['login_process'] = False
                    return False
                    
            except TimeoutException:
                logger.warning("⚠️ Could not find 'Sign in' button")
                self.test_results['login_process'] = False
                return False
                
        except Exception as e:
            logger.error(f"❌ Login test failed: {e}")
            self.test_results['login_process'] = False
            return False
    
    def test_page_navigation(self):
        """Test scrolling and page navigation"""
        try:
            logger.info("📜 Testing page navigation and scrolling...")
            
            initial_height = self.driver.execute_script("return window.pageYOffset")
            
            # Test scrolling
            for i in range(3):
                self.driver.execute_script("window.scrollBy(0, window.innerHeight);")
                time.sleep(1)
                current_height = self.driver.execute_script("return window.pageYOffset")
                
                if current_height > initial_height:
                    logger.info(f"✅ Scroll {i+1} successful - moved from {initial_height}px to {current_height}px")
                    initial_height = current_height
                else:
                    logger.warning(f"⚠️ Scroll {i+1} may have reached bottom")
                    break
            
            self.test_results['page_navigation'] = True
            return True
            
        except Exception as e:
            logger.error(f"❌ Page navigation test failed: {e}")
            self.test_results['page_navigation'] = False
            return False
    
    def test_element_detection(self):
        """Test our ability to detect page elements"""
        try:
            logger.info("🔍 Testing element detection capabilities...")
            
            # Test form detection
            forms = self.driver.find_elements(By.TAG_NAME, "form")
            logger.info(f"✅ Found {len(forms)} forms on page")
            
            # Test input detection  
            text_inputs = self.driver.find_elements(By.CSS_SELECTOR, "input[type='text'], textarea")
            logger.info(f"✅ Found {len(text_inputs)} text input fields")
            
            # Test button detection
            buttons = self.driver.find_elements(By.TAG_NAME, "button")
            submit_inputs = self.driver.find_elements(By.CSS_SELECTOR, "input[type='submit']")
            logger.info(f"✅ Found {len(buttons)} buttons and {len(submit_inputs)} submit inputs")
            
            # Test our category detection logic
            categories_detected = 0
            for category in config.VOTING_RESPONSES.keys():
                elements = self.driver.find_elements(By.XPATH, 
                    f"//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{category.lower()}')]")
                if elements:
                    logger.info(f"✅ Found elements containing '{category}'")
                    categories_detected += 1
                else:
                    logger.info(f"ℹ️ No elements found for '{category}' (expected if voting not open)")
            
            logger.info(f"📊 Element detection summary: {len(forms)} forms, {len(text_inputs)} inputs, {categories_detected} categories")
            self.test_results['element_detection'] = True
            return True
            
        except Exception as e:
            logger.error(f"❌ Element detection test failed: {e}")
            self.test_results['element_detection'] = False
            return False
    
    def test_page_state_detection(self):
        """Test our page state detection logic"""
        try:
            logger.info("🎯 Testing page state detection...")
            
            page_text = self.driver.page_source.lower()
            
            # Check for different state indicators
            voting_indicators = ["vote", "voting", "submit your choice", "enter your", "nomination"]
            results_indicators = ["winner", "results", "congratulations", "thank you for voting"]
            closed_indicators = ["voting closed", "check back", "voting starts", "nomination phase"]
            
            voting_count = sum(1 for indicator in voting_indicators if indicator in page_text)
            results_count = sum(1 for indicator in results_indicators if indicator in page_text)
            closed_count = sum(1 for indicator in closed_indicators if indicator in page_text)
            
            logger.info(f"📊 State indicators found:")
            logger.info(f"   🗳️ Voting indicators: {voting_count}")
            logger.info(f"   🏆 Results indicators: {results_count}")
            logger.info(f"   🚫 Closed indicators: {closed_count}")
            
            if closed_count > 0:
                state = "CLOSED"
            elif results_count > voting_count:
                state = "RESULTS"
            elif voting_count > 0:
                state = "VOTING"
            else:
                state = "UNKNOWN"
            
            logger.info(f"✅ Detected page state: {state}")
            self.test_results['page_state_detection'] = True
            self.test_results['detected_state'] = state
            return True
            
        except Exception as e:
            logger.error(f"❌ Page state detection failed: {e}")
            self.test_results['page_state_detection'] = False
            return False
    
    def run_comprehensive_test(self):
        """Run all tests"""
        logger.info("🚀 STARTING COMPREHENSIVE TEST SUITE")
        logger.info("=" * 60)
        
        # Test 1: Driver setup
        if not self.setup_driver():
            logger.error("❌ Critical failure: Could not setup WebDriver")
            return False
        
        try:
            # Test 2: Marathon page access
            self.test_website_access(config.MARATHON_URL, "Marathon Page")
            
            # Test 3: Login process
            self.test_login_process()
            
            # Test 4: Page navigation
            self.test_page_navigation()
            
            # Test 5: Element detection
            self.test_element_detection()
            
            # Test 6: Page state detection
            self.test_page_state_detection()
            
            # Give user a chance to see the page
            logger.info("👀 Browser will stay open for 10 seconds so you can see the page...")
            time.sleep(10)
            
        finally:
            if self.driver:
                self.driver.quit()
                logger.info("🔚 Browser closed")
        
        # Print summary
        self.print_test_summary()
        return True
    
    def print_test_summary(self):
        """Print a summary of all test results"""
        logger.info("\n" + "=" * 60)
        logger.info("📊 COMPREHENSIVE TEST SUMMARY")
        logger.info("=" * 60)
        
        total_tests = len(self.test_results) - (1 if 'detected_state' in self.test_results else 0)
        passed_tests = sum(1 for k, v in self.test_results.items() if k != 'detected_state' and v)
        
        logger.info(f"🎯 Overall Score: {passed_tests}/{total_tests} tests passed")
        
        for test_name, result in self.test_results.items():
            if test_name == 'detected_state':
                continue
            status = "✅ PASS" if result else "❌ FAIL"
            logger.info(f"   {status} - {test_name.replace('_', ' ').title()}")
        
        if 'detected_state' in self.test_results:
            logger.info(f"\n🎯 Marathon Page State: {self.test_results['detected_state']}")
        
        # Recommendations
        logger.info("\n🎯 RECOMMENDATIONS:")
        if passed_tests >= total_tests * 0.8:  # 80% pass rate
            logger.info("✅ System looks ready! You can confidently start the scheduler.")
            if self.test_results.get('detected_state') == 'CLOSED':
                logger.info("✅ Marathon voting is closed as expected. Scheduler will wait for it to open.")
        else:
            logger.info("⚠️ Some tests failed. Review the issues above before starting automation.")
        
        logger.info("=" * 60)

if __name__ == "__main__":
    tester = ComprehensiveTest()
    tester.run_comprehensive_test()
