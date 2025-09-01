"""
Debug Bubba's Page - Let's see what's actually there
"""

import time
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')
logger = logging.getLogger(__name__)

def debug_bubba_page():
    """Debug what's actually on the Bubba's page"""
    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    
    driver = webdriver.Chrome(options=chrome_options)
    driver.maximize_window()
    
    try:
        logger.info("🌐 Navigating to Bubba's page...")
        driver.get("https://keysweekly.com/bubbas25/#/gallery/?group=516602")
        
        logger.info("⏳ Waiting 10 seconds for page to fully load...")
        time.sleep(10)  # Give it more time
        
        logger.info(f"📄 Page title: {driver.title}")
        logger.info(f"📍 Current URL: {driver.current_url}")
        
        # Check page source for any content
        page_source = driver.page_source
        logger.info(f"📊 Page source length: {len(page_source)} characters")
        
        if "gay" in page_source.lower():
            logger.info("✅ Found 'gay' somewhere in page source!")
        else:
            logger.info("❌ No 'gay' found in page source")
            
        if "bar" in page_source.lower():
            logger.info("✅ Found 'bar' somewhere in page source!")
        else:
            logger.info("❌ No 'bar' found in page source")
        
        # Get all text content
        try:
            body = driver.find_element(By.TAG_NAME, "body")
            body_text = body.text
            logger.info(f"📝 Body text length: {len(body_text)} characters")
            logger.info("📋 First 500 characters of body text:")
            logger.info(body_text[:500])
        except Exception as e:
            logger.error(f"❌ Could not get body text: {e}")
        
        # Find all links
        logger.info("🔗 Looking for ALL links on the page...")
        all_links = driver.find_elements(By.TAG_NAME, "a")
        logger.info(f"📊 Found {len(all_links)} total links")
        
        # Show first 20 links
        logger.info("📋 First 20 links found:")
        for i, link in enumerate(all_links[:20]):
            try:
                href = link.get_attribute('href')
                text = link.text.strip()
                location = link.location
                logger.info(f"   {i+1}. '{text}' -> {href} (x={location['x']}, y={location['y']})")
            except:
                logger.info(f"   {i+1}. [Could not get link info]")
        
        # Look for any elements containing gay or bar
        logger.info("🔍 Looking for ANY elements containing 'gay' or 'bar'...")
        try:
            gay_elements = driver.find_elements(By.XPATH, "//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'gay')]")
            bar_elements = driver.find_elements(By.XPATH, "//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'bar')]")
            
            logger.info(f"🏳️‍🌈 Found {len(gay_elements)} elements containing 'gay'")
            for i, elem in enumerate(gay_elements[:5]):
                try:
                    logger.info(f"   {i+1}. '{elem.text}' at {elem.location}")
                except:
                    pass
                    
            logger.info(f"🍺 Found {len(bar_elements)} elements containing 'bar'")
            for i, elem in enumerate(bar_elements[:5]):
                try:
                    logger.info(f"   {i+1}. '{elem.text}' at {elem.location}")
                except:
                    pass
        except Exception as e:
            logger.error(f"❌ Search failed: {e}")
        
        # Check if page is still loading
        logger.info("⏳ Checking if page is still loading...")
        try:
            loading_elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'loading') or contains(text(), 'Loading')]")
            if loading_elements:
                logger.info(f"⏳ Found {len(loading_elements)} loading indicators")
            else:
                logger.info("✅ No loading indicators found")
        except:
            pass
        
        logger.info("⏸️ Keeping browser open for 30 seconds for manual inspection...")
        logger.info("👀 Look at the browser window and see what's actually there!")
        time.sleep(30)
        
    except Exception as e:
        logger.error(f"❌ Debug failed: {e}")
        import traceback
        traceback.print_exc()
    finally:
        driver.quit()

if __name__ == "__main__":
    print("🔍 DEBUGGING BUBBA'S PAGE")
    print("Let's see what's actually there and why it's not working...")
    print("🚀 Starting debug automatically...")
    debug_bubba_page()
    print("✅ Debug complete!")
