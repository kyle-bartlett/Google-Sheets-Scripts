"""
Debug the SecondStreet iframe content to find the voting interface
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

def debug_iframe():
    """Debug the iframe content to understand the voting interface"""
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
        
        # Find and switch to SecondStreet iframe
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        logger.info(f"Found {len(iframes)} iframes")
        
        for i, iframe in enumerate(iframes):
            try:
                src = iframe.get_attribute('src')
                if src and 'secondstreetapp.com' in src:
                    logger.info(f"Switching to SecondStreet iframe: {src}")
                    driver.switch_to.frame(iframe)
                    time.sleep(5)  # Give iframe time to load
                    break
            except Exception as e:
                logger.error(f"Error with iframe {i+1}: {e}")
                continue
        
        # Debug iframe content
        logger.info("\n=== IFRAME CONTENT DEBUG ===")
        iframe_title = driver.title
        iframe_url = driver.current_url
        logger.info(f"iframe title: {iframe_title}")
        logger.info(f"iframe URL: {iframe_url}")
        
        # Get all text content in iframe
        logger.info("\n=== ALL TEXT CONTENT IN IFRAME ===")
        all_elements = driver.find_elements(By.XPATH, "//*[string-length(text()) > 0]")
        iframe_texts = []
        for elem in all_elements:
            try:
                if elem.is_displayed():
                    text = elem.text.strip()
                    if text and len(text) < 200:
                        iframe_texts.append(text)
            except:
                continue
        
        # Show unique texts
        unique_texts = list(set(iframe_texts))
        logger.info(f"Found {len(unique_texts)} unique text elements in iframe:")
        for text in unique_texts[:20]:  # Show first 20
            logger.info(f"  '{text}'")
        
        # Look specifically for navigation and voting elements
        logger.info("\n=== LOOKING FOR NAVIGATION ELEMENTS ===")
        nav_keywords = ['all groups', 'groups', 'food', 'community', 'business', 'businesses', 'categories']
        for keyword in nav_keywords:
            elements = driver.find_elements(By.XPATH, f"//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{keyword}')]")
            if elements:
                logger.info(f"Found {len(elements)} elements containing '{keyword}':")
                for elem in elements[:3]:
                    try:
                        if elem.is_displayed():
                            text = elem.text.strip()
                            location = elem.location
                            tag = elem.tag_name
                            logger.info(f"  {tag}: '{text}' at ({location['x']}, {location['y']})")
                    except:
                        continue
        
        # Look for voting elements
        logger.info("\n=== LOOKING FOR VOTING ELEMENTS ===")
        vote_keywords = ['vote', 'voting', 'realtor', 'best realtor']
        for keyword in vote_keywords:
            elements = driver.find_elements(By.XPATH, f"//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{keyword}')]")
            if elements:
                logger.info(f"Found {len(elements)} elements containing '{keyword}':")
                for elem in elements[:3]:
                    try:
                        if elem.is_displayed():
                            text = elem.text.strip()
                            location = elem.location
                            tag = elem.tag_name
                            logger.info(f"  {tag}: '{text}' at ({location['x']}, {location['y']})")
                    except:
                        continue
        
        # Look for clickable elements in iframe
        logger.info("\n=== CLICKABLE ELEMENTS IN IFRAME ===")
        clickable_elements = driver.find_elements(By.XPATH, "//a | //button | //*[@onclick] | //*[contains(@class, 'click')]")
        logger.info(f"Found {len(clickable_elements)} clickable elements in iframe")
        
        iframe_clickables = []
        for elem in clickable_elements:
            try:
                if elem.is_displayed():
                    text = elem.text.strip()
                    if text and len(text) < 100:
                        location = elem.location
                        tag = elem.tag_name
                        href = elem.get_attribute('href') or ''
                        iframe_clickables.append((text, tag, location['x'], location['y'], href))
            except:
                continue
        
        logger.info("Clickable elements in iframe:")
        for text, tag, x, y, href in iframe_clickables[:15]:
            logger.info(f"  {tag}: '{text}' at ({x}, {y}) -> {href}")
        
        # Try to navigate to base iframe URL (without gallery part)
        logger.info("\n=== TRYING TO NAVIGATE TO MAIN VOTING PAGE ===")
        base_iframe_url = "https://embed-1109801.secondstreetapp.com/embed/b222ccf1-e1a9-4ac5-8af8-a9f96673db22/"
        logger.info(f"Navigating to base iframe URL: {base_iframe_url}")
        driver.get(base_iframe_url)
        time.sleep(5)
        
        # Check for content after navigation
        logger.info("Checking content after navigating to base URL...")
        new_elements = driver.find_elements(By.XPATH, "//*[string-length(text()) > 0]")
        new_texts = []
        for elem in new_elements:
            try:
                if elem.is_displayed():
                    text = elem.text.strip()
                    if text and len(text) < 200:
                        new_texts.append(text)
            except:
                continue
        
        unique_new_texts = list(set(new_texts))
        logger.info(f"Found {len(unique_new_texts)} unique text elements after navigation:")
        for text in unique_new_texts[:20]:
            logger.info(f"  '{text}'")
        
        # Look for businesses dropdown again
        businesses_elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'Business') or contains(text(), 'business')]")
        if businesses_elements:
            logger.info(f"\nFound {len(businesses_elements)} business-related elements:")
            for elem in businesses_elements:
                try:
                    if elem.is_displayed():
                        text = elem.text.strip()
                        location = elem.location
                        logger.info(f"  Business element: '{text}' at ({location['x']}, {location['y']})")
                except:
                    continue
        
        logger.info("\n=== KEEPING BROWSER OPEN FOR MANUAL INSPECTION ===")
        logger.info("Browser will stay open for 60 seconds...")
        time.sleep(60)
        
    except Exception as e:
        logger.error(f"Debug failed: {e}")
    finally:
        if driver:
            driver.quit()
        logger.info("Debug complete")

if __name__ == "__main__":
    debug_iframe()

