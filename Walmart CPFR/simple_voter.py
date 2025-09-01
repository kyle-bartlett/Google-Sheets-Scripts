#!/usr/bin/env python3
"""
Simple Daily Voting Bot
Goes to Marathon voting page and votes for 4 categories. That's it.
"""

import time
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import config

# Simple logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')
logger = logging.getLogger(__name__)

def vote_daily():
    """Simple function that votes once and exits"""
    driver = None
    
    try:
        # Setup browser
        logger.info("🚀 Starting daily voting...")
        chrome_options = Options()
        if config.HEADLESS_MODE:
            chrome_options.add_argument("--headless")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        
        driver = webdriver.Chrome(options=chrome_options)
        wait = WebDriverWait(driver, config.WAIT_TIMEOUT)
        
        # Go to voting page
        logger.info(f"📍 Going to: {config.MARATHON_URL}")
        driver.get(config.MARATHON_URL)
        time.sleep(5)
        
        # Vote for each category by clicking on candidate names
        votes_cast = 0
        for category, candidate in config.VOTING_RESPONSES.items():
            logger.info(f"🗳️ Looking for {candidate} in {category}")
            
            try:
                # Find category section (scroll through page to find it)
                found = False
                for scroll in range(5):  # Try scrolling down to find categories
                    driver.execute_script(f"window.scrollTo(0, {scroll * 800});")
                    time.sleep(2)
                    
                    # Look for the candidate name as clickable element
                    # Try different ways to find the candidate
                    candidate_selectors = [
                        f"//*[contains(text(), '{candidate}')]",
                        f"//button[contains(text(), '{candidate}')]",
                        f"//label[contains(text(), '{candidate}')]",
                        f"//div[contains(text(), '{candidate}')]",
                        f"//span[contains(text(), '{candidate}')]"
                    ]
                    
                    for selector in candidate_selectors:
                        try:
                            elements = driver.find_elements(By.XPATH, selector)
                            for element in elements:
                                # Check if this element is clickable and in the right category area
                                if element.is_displayed() and element.is_enabled():
                                    # Click the candidate
                                    driver.execute_script("arguments[0].scrollIntoView(true);", element)
                                    time.sleep(1)
                                    element.click()
                                    logger.info(f"✅ Clicked {candidate}")
                                    votes_cast += 1
                                    found = True
                                    break
                            if found:
                                break
                        except Exception:
                            continue
                    
                    if found:
                        break
                
                if not found:
                    logger.warning(f"⚠️ Could not find clickable option for: {candidate}")
                    
            except Exception as e:
                logger.error(f"❌ Error voting for {category}: {e}")
        
        # Submit votes if any were cast
        if votes_cast > 0:
            logger.info(f"📤 Submitting {votes_cast} votes...")
            try:
                # Look for submit button
                submit_buttons = driver.find_elements(By.XPATH, "//input[@type='submit']")
                submit_buttons.extend(driver.find_elements(By.XPATH, "//button[contains(text(), 'Submit')]"))
                
                if submit_buttons:
                    submit_buttons[0].click()
                    time.sleep(3)
                    logger.info("✅ Votes submitted successfully!")
                else:
                    logger.warning("⚠️ Could not find submit button")
                    
            except Exception as e:
                logger.error(f"❌ Error submitting votes: {e}")
        else:
            logger.warning("⚠️ No votes were cast")
            
        logger.info(f"🎯 Daily voting complete! Cast {votes_cast} votes.")
        
    except Exception as e:
        logger.error(f"❌ Voting failed: {e}")
        
    finally:
        if driver:
            driver.quit()
            logger.info("🔚 Browser closed")

if __name__ == "__main__":
    vote_daily()
