"""
Quick Page Check - Non-interactive version to check both voting pages
"""

import time
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import config

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def setup_driver():
    """Initialize Chrome WebDriver"""
    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--headless")  # Run in background
    
    try:
        driver = webdriver.Chrome(options=chrome_options)
        logger.info("Using system Chrome driver")
    except Exception:
        try:
            service = Service("/usr/local/bin/chromedriver")
            driver = webdriver.Chrome(service=service, options=chrome_options)
            logger.info("Using manual Chrome driver path")
        except Exception as e:
            logger.error(f"Failed to setup WebDriver: {e}")
            return None
    
    return driver

def check_page(url, page_name):
    """Check a single page"""
    driver = setup_driver()
    if not driver:
        return
    
    try:
        logger.info(f"\n🌐 Checking {page_name}: {url}")
        driver.get(url)
        time.sleep(5)  # Let page load
        
        # Get basic info
        logger.info(f"📄 Page title: {driver.title}")
        
        # Check for key indicators
        page_source = driver.page_source.lower()
        
        # Look for voting indicators
        voting_keywords = ["vote now", "voting", "submit your vote", "enter your choice"]
        results_keywords = ["congratulations", "winners", "results", "thank you for voting"]
        closed_keywords = ["voting closed", "check back", "voting starts", "nomination phase"]
        
        voting_found = any(keyword in page_source for keyword in voting_keywords)
        results_found = any(keyword in page_source for keyword in results_keywords)
        closed_found = any(keyword in page_source for keyword in closed_keywords)
        
        logger.info(f"🗳️ Voting indicators found: {voting_found}")
        logger.info(f"🏆 Results indicators found: {results_found}")
        logger.info(f"🚫 Closed indicators found: {closed_found}")
        
        # Look for forms and inputs
        forms = driver.find_elements(By.TAG_NAME, "form")
        text_inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='text'], textarea")
        
        logger.info(f"📝 Forms found: {len(forms)}")
        logger.info(f"⌨️ Text inputs found: {len(text_inputs)}")
        
        # Look for our specific categories
        for category in config.VOTING_RESPONSES.keys():
            try:
                elements = driver.find_elements(By.XPATH, 
                    f"//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{category.lower()}')]")
                if elements:
                    logger.info(f"✅ Found '{category}' mentioned {len(elements)} times")
                else:
                    logger.info(f"❌ No mention of '{category}' found")
            except Exception:
                pass
        
        # Determine page state
        if closed_found:
            state = "CLOSED - Voting not yet open"
        elif results_found and not voting_found:
            state = "RESULTS - Voting completed, showing winners"
        elif voting_found:
            state = "VOTING - Active voting in progress"
        else:
            state = "UNKNOWN - Unable to determine state"
        
        logger.info(f"🎯 Page State: {state}")
        
    except Exception as e:
        logger.error(f"Error checking {page_name}: {e}")
    finally:
        driver.quit()

def main():
    logger.info("🔍 Quick Page Status Check")
    logger.info("=" * 50)
    
    # Check Key West page
    check_page(config.KEY_WEST_URL, "Key West (Bubba's Best)")
    
    logger.info("\n" + "=" * 50)
    
    # Check Marathon page
    check_page(config.MARATHON_URL, "Marathon (Best of Marathon)")
    
    logger.info("\n✅ Page check complete!")

if __name__ == "__main__":
    main()
