#!/usr/bin/env python3
"""
Working Voting Bot - Uses the sidebar navigation method
"""

import time
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import config

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')
logger = logging.getLogger(__name__)

def vote_using_sidebar():
    """Navigate using sidebar and vote"""
    driver = None
    
    try:
        # Setup browser
        logger.info("🚀 Starting voting with sidebar navigation...")
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
        
        # Step 1: Find and click "The businesses" in left sidebar
        logger.info("🔍 Looking for 'The businesses' in left sidebar...")
        
        # Try different selectors for "The businesses"
        business_found = False
        selectors = [
            "//div[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'the businesses')]",
            "//span[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'the businesses')]",
            "//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'businesses')]"
        ]
        
        for selector in selectors:
            try:
                elements = driver.find_elements(By.XPATH, selector)
                for elem in elements:
                    # Check if it's in the left sidebar (x position < 400px)
                    if elem.location['x'] < 400:
                        logger.info(f"✅ Found 'The businesses' at x={elem.location['x']}")
                        
                        # Look for expansion arrow near this element
                        parent = elem.find_element(By.XPATH, "./..")
                        
                        # Try to click the element or its parent
                        try:
                            driver.execute_script("arguments[0].scrollIntoView(true);", elem)
                            time.sleep(1)
                            elem.click()
                            business_found = True
                            logger.info("✅ Clicked 'The businesses'")
                        except:
                            try:
                                parent.click()
                                business_found = True
                                logger.info("✅ Clicked parent of 'The businesses'")
                            except:
                                # Try clicking any arrow/icon near it
                                arrows = parent.find_elements(By.XPATH, ".//i | .//svg | .//span[contains(@class, 'arrow')]")
                                for arrow in arrows:
                                    try:
                                        arrow.click()
                                        business_found = True
                                        logger.info("✅ Clicked expansion arrow")
                                        break
                                    except:
                                        continue
                        
                        if business_found:
                            break
                
                if business_found:
                    break
                    
            except Exception as e:
                logger.debug(f"Selector failed: {e}")
                continue
        
        if not business_found:
            logger.error("❌ Could not find/click 'The businesses' in sidebar")
            return False
            
        time.sleep(3)  # Wait for expansion
        
        # Step 2: Click "Best Realtor" in the expanded section
        logger.info("🔍 Looking for 'Best Realtor' in expanded section...")
        
        realtor_found = False
        realtor_selectors = [
            "//a[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'best realtor')]",
            "//a[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'realtor')]",
            "//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'best realtor')]"
        ]
        
        for selector in realtor_selectors:
            try:
                elements = driver.find_elements(By.XPATH, selector)
                for elem in elements:
                    if elem.location['x'] < 400:  # Still in sidebar
                        logger.info(f"✅ Found 'Best Realtor' at x={elem.location['x']}")
                        driver.execute_script("arguments[0].scrollIntoView(true);", elem)
                        time.sleep(1)
                        elem.click()
                        realtor_found = True
                        logger.info("✅ Clicked 'Best Realtor'")
                        break
                
                if realtor_found:
                    break
                    
            except Exception:
                continue
        
        if not realtor_found:
            logger.error("❌ Could not find 'Best Realtor' in sidebar")
            return False
            
        time.sleep(5)  # Wait for navigation to voting section
        
        # Step 3: Now we should be in the voting section - vote for our candidates
        logger.info("🗳️ Now in voting section, looking for candidates...")
        
        votes_cast = 0
        for category, candidate in config.VOTING_RESPONSES.items():
            logger.info(f"🔍 Looking for {candidate}...")
            
            # Look for the candidate name to click
            candidate_found = False
            
            # Try partial matches since names might be formatted differently
            candidate_parts = candidate.split()
            main_identifier = candidate_parts[0] if candidate_parts else candidate  # Use first word as main identifier
            
            candidate_xpaths = [
                f"//*[contains(text(), '{candidate}')]",
                f"//*[contains(text(), '{main_identifier}')]",
                f"//label[contains(text(), '{candidate}')]",
                f"//label[contains(text(), '{main_identifier}')]",
                f"//div[@role='button' and contains(text(), '{candidate}')]",
                f"//div[@role='button' and contains(text(), '{main_identifier}')]"
            ]
            
            for xpath in candidate_xpaths:
                try:
                    elements = driver.find_elements(By.XPATH, xpath)
                    for elem in elements:
                        if elem.is_displayed() and elem.is_enabled():
                            driver.execute_script("arguments[0].scrollIntoView(true);", elem)
                            time.sleep(1)
                            elem.click()
                            logger.info(f"✅ Voted for {candidate}")
                            votes_cast += 1
                            candidate_found = True
                            break
                    
                    if candidate_found:
                        break
                        
                except Exception:
                    continue
            
            if not candidate_found:
                logger.warning(f"⚠️ Could not find {candidate}")
        
        # Step 4: Submit votes
        if votes_cast > 0:
            logger.info(f"📤 Looking for submit button...")
            
            submit_selectors = [
                "//button[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'submit')]",
                "//button[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'vote')]",
                "//input[@type='submit']",
                "//button[@type='submit']"
            ]
            
            for selector in submit_selectors:
                try:
                    submit_btn = driver.find_element(By.XPATH, selector)
                    submit_btn.click()
                    logger.info("✅ Submitted votes!")
                    break
                except:
                    continue
        
        logger.info(f"🎯 Voting complete! Cast {votes_cast} votes.")
        
        # Keep browser open for a bit to see results
        time.sleep(10)
        
        return votes_cast > 0
        
    except Exception as e:
        logger.error(f"❌ Voting failed: {e}")
        return False
        
    finally:
        if driver:
            driver.quit()
            logger.info("🔚 Browser closed")

if __name__ == "__main__":
    vote_using_sidebar()

