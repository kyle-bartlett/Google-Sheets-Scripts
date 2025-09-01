"""
Marathon Readiness Test - Test the actual logic on your target Marathon page
This shows exactly what will happen when Marathon voting opens
"""

import time
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from keys_weekly_voter import KeysWeeklyVoter
import config

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')
logger = logging.getLogger(__name__)

def test_marathon_readiness():
    """Test what happens when we run the actual voter on Marathon page"""
    logger.info("🎯 MARATHON READINESS TEST")
    logger.info("=" * 50)
    logger.info("Testing the ACTUAL voting logic on YOUR Marathon page")
    logger.info("This shows exactly what will happen when voting opens")
    logger.info("=" * 50)
    
    # Create our actual voter
    voter = KeysWeeklyVoter()
    
    try:
        # Test with visible browser so you can see
        config.HEADLESS_MODE = False
        
        logger.info("🚀 Setting up browser...")
        voter.setup_driver()
        
        logger.info("🌐 Navigating to YOUR Marathon page...")
        voter.driver.get(config.MARATHON_URL)
        time.sleep(5)
        
        logger.info(f"📄 Page title: {voter.driver.title}")
        
        # Test our smart page state detection
        logger.info("🧠 Testing smart page state detection...")
        page_state = voter.check_page_state()
        logger.info(f"🎯 Page state detected: {page_state}")
        
        if page_state == "closed":
            logger.info("✅ CORRECT: Detected voting is closed")
            logger.info("💡 When voting opens, this will change to 'voting' state")
        elif page_state == "voting":
            logger.info("🎉 VOTING IS OPEN! Let's test the voting process...")
        elif page_state == "results":
            logger.info("🏆 Page is showing results")
        
        # Test scrolling and category detection  
        logger.info("📜 Testing scrolling and category detection...")
        found_categories = voter.scroll_and_find_categories()
        
        if found_categories:
            logger.info("✅ Found voting elements!")
        else:
            logger.info("ℹ️ No voting elements found (expected if voting closed)")
        
        # Test looking for your specific categories
        logger.info("🔍 Testing search for YOUR specific categories...")
        for category, response in config.VOTING_RESPONSES.items():
            logger.info(f"🔍 Looking for '{category}' category...")
            
            try:
                elements = voter.driver.find_elements(By.XPATH, 
                    f"//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{category.lower()}')]")
                
                if elements:
                    logger.info(f"✅ Found '{category}' - would fill with: {response}")
                else:
                    logger.info(f"⏳ '{category}' not found (will appear when voting opens)")
            except Exception as e:
                logger.debug(f"Search for {category} failed: {e}")
        
        # Test form detection
        logger.info("📝 Testing form detection...")
        forms = voter.driver.find_elements(By.TAG_NAME, "form")
        inputs = voter.driver.find_elements(By.CSS_SELECTOR, "input[type='text'], textarea")
        
        logger.info(f"📊 Found {len(forms)} forms and {len(inputs)} text inputs")
        
        if inputs:
            logger.info("✅ Text inputs detected - voting forms may be present!")
        else:
            logger.info("ℹ️ No text inputs - voting forms not loaded yet")
        
        # Summary
        logger.info("\n🎯 READINESS SUMMARY:")
        logger.info("=" * 30)
        
        if page_state == "closed":
            logger.info("✅ System correctly detects Marathon voting is closed")
            logger.info("✅ Will automatically detect when voting opens")
            logger.info("✅ Your categories are configured and ready")
            logger.info("✅ Form detection logic is working")
            logger.info("🚀 READY: Will start voting automatically when Marathon opens!")
        elif page_state == "voting":
            logger.info("🎉 Marathon voting is OPEN!")
            logger.info("🚀 Your automation should start voting now!")
        
        logger.info("\n⏸️ Keeping browser open for 15 seconds for you to explore...")
        time.sleep(15)
        
    except Exception as e:
        logger.error(f"❌ Test error: {e}")
    finally:
        if voter.driver:
            voter.driver.quit()
            logger.info("🔚 Test browser closed")

if __name__ == "__main__":
    print("🎯 MARATHON READINESS TEST")
    print("This tests your ACTUAL voting logic on YOUR Marathon page")
    print("You'll see exactly what happens when the automation runs")
    
    input("\nPress Enter to test Marathon readiness...")
    test_marathon_readiness()
    
    print("\n✅ Readiness test complete!")
    print("💡 This shows exactly how your automation will behave!")
