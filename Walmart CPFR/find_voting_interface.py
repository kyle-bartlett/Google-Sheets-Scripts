"""
Script to find the correct path to the voting interface
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

def find_voting_interface():
    """Find the correct way to access the voting interface"""
    driver = None
    try:
        # Setup Chrome
        chrome_options = Options()
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        driver = webdriver.Chrome(options=chrome_options)
        
        # Start with the base Marathon URL
        base_url = "https://keysweekly.com/bom25/"
        logger.info(f"Starting at base URL: {base_url}")
        driver.get(base_url)
        time.sleep(5)
        
        logger.info(f"Page title: {driver.title}")
        logger.info(f"Current URL: {driver.current_url}")
        
        # Look for voting-related links
        voting_links = []
        all_links = driver.find_elements(By.TAG_NAME, "a")
        for link in all_links:
            try:
                if link.is_displayed():
                    text = link.text.strip()
                    href = link.get_attribute('href')
                    if text and href:
                        text_lower = text.lower()
                        href_lower = href.lower()
                        if any(keyword in text_lower or keyword in href_lower for keyword in ['vote', 'voting', 'gallery', 'poll', 'contest']):
                            voting_links.append((text, href))
            except:
                continue
        
        logger.info(f"\nFound {len(voting_links)} potential voting links:")
        for text, href in voting_links[:10]:
            logger.info(f"  '{text}' -> {href}")
        
        # Try the original URL but look for navigation to voting
        original_url = config.MARATHON_URL
        logger.info(f"\nNavigating to original URL: {original_url}")
        driver.get(original_url)
        time.sleep(5)
        
        # Look for any text that mentions groups, categories, or voting
        logger.info("\nLooking for navigation elements...")
        nav_texts = []
        all_elements = driver.find_elements(By.XPATH, "//*[string-length(text()) > 0]")
        for elem in all_elements:
            try:
                if elem.is_displayed():
                    text = elem.text.strip()
                    if text and len(text) < 100:
                        text_lower = text.lower()
                        if any(keyword in text_lower for keyword in ['group', 'all groups', 'categories', 'food', 'community', 'business', 'dropdown']):
                            location = elem.location
                            tag = elem.tag_name
                            nav_texts.append((text, tag, location['x'], location['y']))
            except:
                continue
        
        logger.info("Navigation-related elements:")
        for text, tag, x, y in nav_texts[:15]:
            logger.info(f"  {tag}: '{text}' at ({x}, {y})")
        
        # Try to click "All Groups" or similar navigation
        logger.info("\nTrying to find and click navigation elements...")
        nav_selectors = [
            "//text()[contains(., 'All Groups')]/parent::*",
            "//*[contains(text(), 'All Groups')]",
            "//text()[contains(., 'Groups')]/parent::*", 
            "//*[contains(text(), 'Groups')]",
            "//*[text()='☰']",  # Hamburger menu
            "//*[contains(@class, 'menu-toggle')]",
            "//*[contains(@class, 'nav-toggle')]"
        ]
        
        for selector in nav_selectors:
            try:
                elements = driver.find_elements(By.XPATH, selector)
                for elem in elements:
                    if elem.is_displayed():
                        logger.info(f"Found navigation element: '{elem.text}' with selector: {selector}")
                        try:
                            elem.click()
                            time.sleep(3)
                            logger.info("Clicked navigation element - checking for changes...")
                            
                            # Check if we now see "The Businesses" or voting interface
                            businesses_elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'The Businesses') or contains(text(), 'Best Realtor')]")
                            if businesses_elements:
                                logger.info(f"SUCCESS! Found {len(businesses_elements)} business/realtor elements after click!")
                                for bus_elem in businesses_elements:
                                    logger.info(f"  Business element: '{bus_elem.text}'")
                                
                                # Keep browser open longer for manual inspection
                                logger.info("Keeping browser open for 60 seconds for inspection...")
                                time.sleep(60)
                                return True
                                
                        except Exception as click_error:
                            logger.info(f"Could not click element: {click_error}")
                            continue
            except:
                continue
        
        # If we haven't found it yet, try looking for any iframe or embedded content
        logger.info("\nChecking for iframes or embedded content...")
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        logger.info(f"Found {len(iframes)} iframes")
        
        for i, iframe in enumerate(iframes):
            try:
                src = iframe.get_attribute('src')
                logger.info(f"iframe {i+1}: {src}")
                
                # Switch to iframe and check content
                driver.switch_to.frame(iframe)
                iframe_title = driver.title
                logger.info(f"iframe {i+1} title: {iframe_title}")
                
                # Look for voting content in iframe
                iframe_vote_elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'Vote') or contains(text(), 'Business') or contains(text(), 'Realtor')]")
                if iframe_vote_elements:
                    logger.info(f"Found {len(iframe_vote_elements)} voting elements in iframe {i+1}!")
                    for vote_elem in iframe_vote_elements[:5]:
                        logger.info(f"  iframe vote element: '{vote_elem.text}'")
                
                driver.switch_to.default_content()
                
            except Exception as iframe_error:
                logger.info(f"Error with iframe {i+1}: {iframe_error}")
                driver.switch_to.default_content()
                continue
        
        logger.info("\nKeeping browser open for 30 seconds for final manual inspection...")
        time.sleep(30)
        
    except Exception as e:
        logger.error(f"Search failed: {e}")
    finally:
        if driver:
            driver.quit()
        logger.info("Search complete")

if __name__ == "__main__":
    find_voting_interface()

