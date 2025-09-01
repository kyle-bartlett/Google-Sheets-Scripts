"""
Page Investigator - Let's see what's actually on the Key West page
This will help us understand why our search isn't working
"""

import time
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')
logger = logging.getLogger(__name__)

def investigate_page():
    """Investigate what's actually on the Key West page"""
    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    
    driver = webdriver.Chrome(options=chrome_options)
    driver.maximize_window()
    
    try:
        logger.info("🔍 INVESTIGATING Key West page structure...")
        driver.get("https://keysweekly.com/buk24/#/gallery?group=497175")
        time.sleep(5)
        
        logger.info(f"📄 Page title: {driver.title}")
        
        # Get all text content and look for categories
        logger.info("📜 Scanning for ALL text content...")
        
        # Scroll through the entire page and collect all text
        all_text_content = []
        
        for scroll in range(20):
            logger.info(f"📜 Scroll {scroll + 1}/20")
            driver.execute_script("window.scrollBy(0, 300);")
            time.sleep(1)
            
            # Get all text elements on current view
            all_elements = driver.find_elements(By.XPATH, "//*[text()]")
            for elem in all_elements:
                try:
                    text = elem.text.strip()
                    if text and len(text) > 2 and len(text) < 200:
                        if text not in all_text_content:
                            all_text_content.append(text)
                except:
                    pass
        
        logger.info(f"📊 Found {len(all_text_content)} unique text elements")
        
        # Look for anything with "bar" in it
        bar_related = [text for text in all_text_content if 'bar' in text.lower()]
        logger.info(f"🍺 Found {len(bar_related)} items containing 'bar':")
        for item in bar_related[:10]:  # Show first 10
            logger.info(f"   - {item}")
        
        # Look for anything with "gay" in it
        gay_related = [text for text in all_text_content if 'gay' in text.lower()]
        logger.info(f"🏳️‍🌈 Found {len(gay_related)} items containing 'gay':")
        for item in gay_related:
            logger.info(f"   - {item}")
        
        # Look for "see all" or "entrants"
        see_all_related = [text for text in all_text_content if 'see all' in text.lower() or 'entrant' in text.lower()]
        logger.info(f"👥 Found {len(see_all_related)} items with 'see all' or 'entrant':")
        for item in see_all_related:
            logger.info(f"   - {item}")
        
        # Look for common voting/contest terms
        contest_terms = ['winner', 'finalist', 'vote', 'best', 'category', 'restaurant', 'business']
        for term in contest_terms:
            matches = [text for text in all_text_content if term in text.lower()]
            if matches:
                logger.info(f"🏆 Found {len(matches)} items containing '{term}':")
                for item in matches[:5]:  # Show first 5
                    logger.info(f"   - {item}")
        
        # Show some random categories to understand the structure
        logger.info("📋 Here are some sample text elements found:")
        for i, text in enumerate(all_text_content[10:30]):  # Show elements 10-30
            logger.info(f"   {i+10}. {text}")
        
        logger.info("⏸️ Keeping browser open for 30 seconds so you can explore...")
        time.sleep(30)
        
    except Exception as e:
        logger.error(f"❌ Investigation error: {e}")
    finally:
        driver.quit()

if __name__ == "__main__":
    print("🔍 INVESTIGATING Key West page to understand structure")
    print("This will help us improve our search logic")
    input("Press Enter to start investigation...")
    investigate_page()
    print("✅ Investigation complete!")
