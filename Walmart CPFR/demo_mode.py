"""
Demo Mode - Visual demonstration of the voting automation
Shows you exactly what the automation does when it runs
"""

import time
import logging
import sys
from keys_weekly_voter import KeysWeeklyVoter
import config

# Set up logging for demo
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger(__name__)

class DemoVoter(KeysWeeklyVoter):
    def __init__(self):
        super().__init__()
        
    def setup_driver(self):
        """Setup driver for demo - visible browser"""
        # Force visible mode for demo
        original_headless = config.HEADLESS_MODE
        config.HEADLESS_MODE = False
        
        result = super().setup_driver()
        
        # Slow down for demo viewing
        if self.driver:
            logger.info("🎬 DEMO MODE: Browser window opened for demonstration")
            logger.info("👀 You can watch the automation work in real-time!")
            time.sleep(3)
        
        return result
    
    def login(self, url):
        """Demo login with extra narration"""
        logger.info(f"🌐 DEMO: Navigating to {url}")
        result = super().login(url)
        
        if result:
            logger.info("✅ DEMO: Successfully accessed the page")
        else:
            logger.info("⚠️ DEMO: Login process had issues (this is expected for closed voting)")
        
        logger.info("⏸️ DEMO: Pausing so you can see the page...")
        time.sleep(5)
        return result
    
    def scroll_and_find_categories(self):
        """Demo scrolling with narration"""
        logger.info("📜 DEMO: Starting to scroll and look for voting categories...")
        
        for scroll_count in range(6):
            logger.info(f"📜 DEMO: Scroll {scroll_count + 1}/6 - Looking for voting forms...")
            self.driver.execute_script("window.scrollBy(0, window.innerHeight);")
            time.sleep(3)  # Slower for demo
            
            # Check for voting elements
            voting_elements = self.driver.find_elements("css selector", 
                "input[type='text'], textarea, .voting-input, .category-input")
            if voting_elements:
                logger.info(f"✅ DEMO: Found {len(voting_elements)} potential voting inputs!")
                time.sleep(2)
                return True
        
        logger.info("⚠️ DEMO: No voting forms found (expected if voting is closed)")
        return False
    
    def demo_run(self, url=None):
        """Run a complete demo of the voting process"""
        if url is None:
            url = config.MARATHON_URL
            
        logger.info("🎬 STARTING VOTING AUTOMATION DEMO")
        logger.info("=" * 50)
        logger.info("👀 Watch the browser window to see the automation in action!")
        logger.info("🎯 This demo shows you exactly what happens during voting")
        logger.info("=" * 50)
        
        try:
            # Setup browser
            if not self.setup_driver():
                return False
            
            # Navigate and login
            self.login(url)
            
            # Check page state
            page_state = self.check_page_state()
            logger.info(f"🎯 DEMO: Page state detected as '{page_state}'")
            
            if page_state == "closed":
                logger.info("⏳ DEMO: Voting is closed, but let's show you what the automation does")
                logger.info("🔍 DEMO: The real automation would wait and try again in 4 hours")
            
            # Demonstrate scrolling
            self.scroll_and_find_categories()
            
            # Show category detection attempt
            logger.info("🔍 DEMO: Attempting to find your voting categories...")
            time.sleep(2)
            
            for category in config.VOTING_RESPONSES.keys():
                logger.info(f"🔍 DEMO: Looking for '{category}' category...")
                time.sleep(1)
            
            logger.info("📊 DEMO: When voting is open, the automation would:")
            logger.info("   1. Find each category field")
            logger.info("   2. Enter your predetermined responses")
            logger.info("   3. Submit the votes")
            logger.info("   4. Wait 24 hours and repeat")
            
            logger.info("⏸️ DEMO: Keeping browser open for 15 seconds so you can explore...")
            time.sleep(15)
            
            logger.info("✅ DEMO: This is exactly what your automation does every day!")
            
        except Exception as e:
            logger.error(f"❌ DEMO error: {e}")
        finally:
            if self.driver:
                logger.info("🔚 DEMO: Closing browser")
                self.driver.quit()

def main():
    logger.info("🎬 Keys Weekly Voting Automation - DEMO MODE")
    logger.info("This will show you exactly what the automation does")
    
    choice = input("\nWhich page would you like to demo?\n1. Marathon (your target page)\n2. Key West (active voting)\nEnter 1 or 2: ").strip()
    
    demo = DemoVoter()
    
    if choice == "1":
        logger.info("🎯 Starting demo on Marathon page (your target)")
        demo.demo_run(config.MARATHON_URL)
    elif choice == "2":
        logger.info("🏝️ Starting demo on Key West page (active voting)")
        demo.demo_run(config.KEY_WEST_URL)
    else:
        logger.info("📋 Invalid choice, defaulting to Marathon page")
        demo.demo_run(config.MARATHON_URL)
    
    logger.info("\n🎉 Demo complete! This is exactly what your automation does automatically.")
    logger.info("💡 Your smart scheduler is currently monitoring and will start voting when Marathon opens!")

if __name__ == "__main__":
    main()
