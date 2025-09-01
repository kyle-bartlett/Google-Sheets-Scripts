#!/usr/bin/env python3
"""
Actual Working Voter - Based on the exact page structure shown
"""

import time
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')
logger = logging.getLogger(__name__)

def vote_for_categories():
    """Vote for our 4 categories using the actual page structure"""
    driver = None
    
    try:
        # Setup browser
        logger.info("🚀 Starting voting automation...")
        chrome_options = Options()
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        
        driver = webdriver.Chrome(options=chrome_options)
        wait = WebDriverWait(driver, 20)
        
        # Go directly to The Businesses group
        url = "https://keysweekly.com/bom25/#/gallery?group=522962"
        logger.info(f"📍 Going to: {url}")
        driver.get(url)
        
        # Wait longer for JavaScript to load the content
        logger.info("⏳ Waiting for page to fully load...")
        time.sleep(8)
        
        # Check if we're on the right page
        try:
            # Wait for any voting-related content to appear
            wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Vote') or contains(text(), 'VOTE') or contains(text(), 'Best')]")))
            logger.info("✅ Page loaded with voting content")
        except:
            logger.warning("⚠️ Voting content may not have loaded properly")
        
        # Categories to vote for and their candidates
        # Based on the screenshot, the exact category names in the sidebar
        voting_targets = [
            ("Best Realtor", "Nate Bartlett"),
            ("Best Real Estate", "Berkshire Hathaway"),  # Might be "Best Real Estate Office"
            ("Best Vacation Rental", "Berkshire"),  # Might be "Best Vacation Rental Company"  
            ("Best General Contractor", "Berkshire"),  # Or another category with Berkshire
        ]
        
        votes_cast = 0
        
        for category, candidate in voting_targets:
            try:
                logger.info(f"🔍 Looking for category: {category}")
                
                # First, click on the category in the sidebar
                category_found = False
                
                # Try to find and click the category link
                category_selectors = [
                    f"//a[contains(text(), '{category}')]",
                    f"//div[contains(text(), '{category}')]",
                    f"//*[text()='{category}']"
                ]
                
                for selector in category_selectors:
                    try:
                        elements = driver.find_elements(By.XPATH, selector)
                        for elem in elements:
                            # Check if it's in the sidebar (left side)
                            if elem.location['x'] < 500:
                                logger.info(f"✅ Found '{category}' in sidebar")
                                driver.execute_script("arguments[0].scrollIntoView(true);", elem)
                                time.sleep(1)
                                elem.click()
                                category_found = True
                                time.sleep(3)  # Wait for voting section to load
                                break
                        if category_found:
                            break
                    except:
                        continue
                
                if not category_found:
                    logger.warning(f"⚠️ Could not find category: {category}")
                    continue
                
                # Now look for the VOTE button next to our candidate
                logger.info(f"🗳️ Looking for {candidate}'s VOTE button...")
                
                # Find all candidate names and their associated VOTE buttons
                vote_found = False
                
                # Look for the candidate name and then find the VOTE button in the same container
                candidate_selectors = [
                    f"//*[contains(text(), '{candidate}')]",
                    f"//div[contains(text(), '{candidate}')]",
                    f"//h3[contains(text(), '{candidate}')]",
                    f"//span[contains(text(), '{candidate}')]"
                ]
                
                for selector in candidate_selectors:
                    try:
                        candidates = driver.find_elements(By.XPATH, selector)
                        for cand_elem in candidates:
                            # Find the VOTE button in the same parent container
                            parent = cand_elem.find_element(By.XPATH, "./ancestor::div[contains(@class, 'candidate') or contains(@class, 'entry') or contains(@class, 'item')]")
                            vote_buttons = parent.find_elements(By.XPATH, ".//button[contains(text(), 'VOTE')]")
                            
                            if vote_buttons:
                                vote_button = vote_buttons[0]
                                driver.execute_script("arguments[0].scrollIntoView(true);", vote_button)
                                time.sleep(1)
                                vote_button.click()
                                logger.info(f"✅ Clicked VOTE for {candidate}")
                                votes_cast += 1
                                vote_found = True
                                time.sleep(2)  # Wait before next category
                                break
                    except:
                        continue
                    
                    if vote_found:
                        break
                
                if not vote_found:
                    # Try a simpler approach - just click any VOTE button if there's only one visible
                    vote_buttons = driver.find_elements(By.XPATH, "//button[text()='VOTE' and not(@disabled)]")
                    
                    for i, btn in enumerate(vote_buttons):
                        try:
                            # Check if this button is near our candidate text
                            btn_y = btn.location['y']
                            
                            # Look for candidate text near this button
                            page_text = driver.find_element(By.XPATH, "//body").text
                            if candidate in page_text:
                                # If our candidate is on the page, click the appropriate VOTE button
                                # For Nate Bartlett, it's usually the 2nd button (index 1)
                                if candidate == "Nate Bartlett" and len(vote_buttons) >= 2:
                                    vote_buttons[1].click()  # Nate is second in the list
                                    logger.info(f"✅ Clicked VOTE button #{i+1} for {candidate}")
                                    votes_cast += 1
                                    vote_found = True
                                    break
                                elif "Berkshire" in candidate and i == 0:
                                    # Berkshire entries are usually first
                                    btn.click()
                                    logger.info(f"✅ Clicked VOTE button #{i+1} for {candidate}")
                                    votes_cast += 1
                                    vote_found = True
                                    break
                        except:
                            continue
                
                if not vote_found:
                    logger.warning(f"⚠️ Could not find VOTE button for {candidate}")
                    
            except Exception as e:
                logger.error(f"❌ Error voting for {category}: {e}")
        
        logger.info(f"🎯 Voting complete! Cast {votes_cast} out of {len(voting_targets)} votes.")
        
        # Keep browser open to see results
        if votes_cast > 0:
            logger.info("✅ Success! Keeping browser open for 10 seconds to verify...")
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
    vote_for_categories()
