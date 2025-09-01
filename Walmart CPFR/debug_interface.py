"""
Debug script to understand the actual Marathon voting interface
"""

import time
import logging
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException
import config

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def debug_interface():
    """Debug the actual interface to understand how to access voting"""
    driver = None
    try:
        # Setup Chrome
        chrome_options = Options()
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        driver = webdriver.Chrome(options=chrome_options)
        
        # Navigate to Marathon voting
        url = config.MARATHON_URL
        logger.info(f"Navigating to: {url}")
        driver.get(url)
        time.sleep(5)
        
        # Get page info
        logger.info(f"Page title: {driver.title}")
        logger.info(f"Current URL: {driver.current_url}")
        
        # Check page source for keywords
        page_source = driver.page_source.lower()
        keywords = ['vote', 'business', 'realtor', 'dropdown', 'sidebar', 'category']
        for keyword in keywords:
            count = page_source.count(keyword)
            logger.info(f"Keyword '{keyword}': {count} occurrences")
        
        # Look for all clickable elements
        logger.info("\n=== LOOKING FOR CLICKABLE ELEMENTS ===")
        clickable_elements = driver.find_elements(By.XPATH, "//a | //button | //*[@onclick] | //*[contains(@class, 'click')]")
        logger.info(f"Found {len(clickable_elements)} potentially clickable elements")
        
        clickable_texts = []
        for elem in clickable_elements:
            try:
                if elem.is_displayed():
                    text = elem.text.strip()
                    if text and len(text) < 100:
                        location = elem.location
                        clickable_texts.append((text, elem.tag_name, location['x'], location['y']))
            except:
                continue
        
        logger.info("Clickable elements:")
        for text, tag, x, y in clickable_texts[:20]:  # Show first 20
            logger.info(f"  {tag}: '{text}' at ({x}, {y})")
        
        # Look for any dropdown or navigation elements
        logger.info("\n=== LOOKING FOR NAVIGATION ELEMENTS ===")
        nav_selectors = [
            "//nav//*",
            "//*[contains(@class, 'nav')]",
            "//*[contains(@class, 'menu')]", 
            "//*[contains(@class, 'dropdown')]",
            "//*[contains(@class, 'sidebar')]"
        ]
        
        for selector in nav_selectors:
            try:
                elements = driver.find_elements(By.XPATH, selector)
                if elements:
                    logger.info(f"Found {len(elements)} elements with selector: {selector}")
                    for elem in elements[:5]:  # Show first 5
                        try:
                            if elem.is_displayed():
                                text = elem.text.strip()
                                if text and len(text) < 200:
                                    logger.info(f"  Nav element: '{text}'")
                        except:
                            continue
            except:
                continue
        
        # Scroll and look for more content
        logger.info("\n=== SCROLLING TO FIND MORE CONTENT ===")
        original_height = driver.execute_script("return document.body.scrollHeight")
        logger.info(f"Original page height: {original_height}")
        
        for scroll in range(10):
            driver.execute_script("window.scrollBy(0, 800);")
            time.sleep(2)
            
            new_height = driver.execute_script("return document.body.scrollHeight")
            scroll_position = driver.execute_script("return window.pageYOffset")
            logger.info(f"Scroll {scroll+1}: position={scroll_position}, page_height={new_height}")
            
            # Look for vote buttons at current position
            vote_buttons = driver.find_elements(By.XPATH, "//*[contains(text(), 'Vote') or contains(text(), 'VOTE')]")
            business_elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'Business') or contains(text(), 'business')]")
            realtor_elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'Realtor') or contains(text(), 'realtor')]")
            
            logger.info(f"  Found: {len(vote_buttons)} vote, {len(business_elements)} business, {len(realtor_elements)} realtor elements")
            
            # Show any new interesting text
            current_elements = driver.find_elements(By.XPATH, "//*[string-length(text()) > 0]")
            interesting_texts = []
            for elem in current_elements:
                try:
                    if elem.is_displayed():
                        text = elem.text.strip()
                        if text and len(text) < 100:
                            text_lower = text.lower()
                            if any(keyword in text_lower for keyword in ['vote', 'business', 'realtor', 'best']):
                                location = elem.location
                                if scroll_position - 1000 < location['y'] < scroll_position + 1000:  # In current view
                                    interesting_texts.append(text)
                except:
                    continue
            
            if interesting_texts:
                logger.info(f"  Interesting texts in view: {list(set(interesting_texts))[:5]}")
        
        logger.info("\n=== KEEPING BROWSER OPEN FOR MANUAL INSPECTION ===")
        logger.info("Browser will stay open for 30 seconds so you can manually inspect...")
        time.sleep(30)
        
    except Exception as e:
        logger.error(f"Debug failed: {e}")
    finally:
        if driver:
            driver.quit()
        logger.info("Debug complete")

if __name__ == "__main__":
    debug_interface()

