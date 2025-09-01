"""
Live Interactive Demo - Watch the automation work step by step
"""

import time
import logging
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import config

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')
logger = logging.getLogger(__name__)

def setup_demo_driver():
    """Setup Chrome driver for demo viewing"""
    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    # Make sure it's NOT headless so you can see it
    
    try:
        driver = webdriver.Chrome(options=chrome_options)
        driver.maximize_window()  # Make it big so you can see everything
        logger.info("✅ Demo browser opened - you should see a Chrome window!")
        return driver
    except Exception as e:
        logger.error(f"❌ Failed to open demo browser: {e}")
        return None

def live_demo():
    """Run a live demo you can actually watch"""
    logger.info("🎬 LIVE DEMO - Keys Weekly Voting Automation")
    logger.info("="*60)
    logger.info("👀 Opening browser window - you should see Chrome open!")
    
    driver = setup_demo_driver()
    if not driver:
        logger.error("Failed to setup demo browser")
        return
    
    try:
        # Demo on Key West page (active voting)
        url = config.KEY_WEST_URL
        logger.info(f"🌐 Step 1: Navigating to Key West voting page...")
        driver.get(url)
        
        logger.info("⏸️ PAUSE: Look at the browser - you should see the Keys Weekly page!")
        time.sleep(8)
        
        logger.info(f"📄 Page title: {driver.title}")
        
        # Check page state
        logger.info("🔍 Step 2: Analyzing page to detect voting status...")
        page_text = driver.page_source.lower()
        
        if "vote" in page_text:
            logger.info("✅ Voting indicators detected!")
        if "winner" in page_text:
            logger.info("🏆 Results/winners detected!")
        if "congratulations" in page_text:
            logger.info("🎉 Congratulations text found!")
        
        time.sleep(3)
        
        # Demo scrolling
        logger.info("📜 Step 3: Scrolling through the page to find voting areas...")
        for i in range(6):
            logger.info(f"📜 Scroll {i+1}/6 - Looking for voting forms...")
            driver.execute_script("window.scrollBy(0, window.innerHeight);")
            time.sleep(2)
            
            # Look for forms and inputs
            forms = driver.find_elements(By.TAG_NAME, "form")
            inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='text'], textarea")
            buttons = driver.find_elements(By.TAG_NAME, "button")
            
            if forms or inputs:
                logger.info(f"✅ Found {len(forms)} forms, {len(inputs)} text inputs, {len(buttons)} buttons")
        
        logger.info("🔍 Step 4: Looking for our target categories...")
        time.sleep(2)
        
        # Look for our categories (even though they won't be on Key West page)
        for category in config.VOTING_RESPONSES.keys():
            logger.info(f"🔍 Searching for '{category}' category...")
            try:
                elements = driver.find_elements(By.XPATH, 
                    f"//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{category.lower()}')]")
                if elements:
                    logger.info(f"✅ Found elements containing '{category}'!")
                else:
                    logger.info(f"ℹ️ '{category}' not found (different categories on this page)")
            except Exception:
                pass
            time.sleep(1)
        
        logger.info("📊 Step 5: Demo Summary - What the automation does:")
        logger.info("   1. ✅ Navigate to voting page")
        logger.info("   2. ✅ Detect page state (voting/closed/results)")
        logger.info("   3. ✅ Scroll through page systematically")
        logger.info("   4. ✅ Find forms and input fields")
        logger.info("   5. ✅ Search for target categories")
        logger.info("   6. 🔄 Fill forms when categories match")
        logger.info("   7. 🔄 Submit votes")
        logger.info("   8. 🔄 Wait 24 hours and repeat")
        
        logger.info("⏸️ FINAL PAUSE: Take a look around the page!")
        logger.info("🎯 This is exactly what your automation does on the Marathon page!")
        logger.info("💡 When Marathon voting opens, it will find YOUR categories and vote!")
        time.sleep(15)
        
        logger.info("✅ Demo complete! Your automation is much smarter than this demo showed.")
        
    except Exception as e:
        logger.error(f"❌ Demo error: {e}")
    finally:
        logger.info("🔚 Closing browser in 5 seconds...")
        time.sleep(5)
        driver.quit()
        logger.info("👋 Demo browser closed!")

if __name__ == "__main__":
    print("🎬 LIVE DEMO - You'll see a real browser window!")
    print("👀 Watch the Chrome window that opens to see the automation work")
    print("⏳ This demo will take about 45 seconds")
    input("Press Enter to start the live demo...")
    
    live_demo()
    
    print("\n🎉 Demo complete!")
    print("💡 Your smart scheduler is running and will do exactly this on Marathon when voting opens!")
    print("📊 Check 'smart_scheduler_log.txt' anytime to see what it's doing")
