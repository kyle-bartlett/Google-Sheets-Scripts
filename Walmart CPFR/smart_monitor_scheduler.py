"""
Smart Monitor and Scheduler for Keys Weekly Voting
Monitors for when voting opens and automatically starts voting
"""

import schedule
import time
import logging
import sys
import subprocess
from datetime import datetime, timedelta
from keys_weekly_voter import KeysWeeklyVoter
import config

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('smart_scheduler_log.txt'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

class SmartVotingScheduler:
    def __init__(self):
        self.voter = KeysWeeklyVoter()
        self.start_date = datetime.now()
        self.end_date = self.start_date + timedelta(days=config.TOTAL_VOTING_DAYS)
        self.vote_count = 0
        self.voting_started = False
        self.last_check_time = None
        
    def check_voting_status(self):
        """Check if voting is available and ready"""
        try:
            logger.info("🔍 Checking if Marathon voting is open...")
            self.last_check_time = datetime.now()
            
            # Use a quick headless check
            original_headless = config.HEADLESS_MODE
            config.HEADLESS_MODE = True  # Check quietly
            
            voter_temp = KeysWeeklyVoter()
            voter_temp.setup_driver()
            
            try:
                # Navigate to Marathon page
                voter_temp.driver.get(config.MARATHON_URL)
                time.sleep(5)
                
                # Check page state
                page_state = voter_temp.check_page_state()
                logger.info(f"📊 Marathon page state: {page_state}")
                
                # Look for our categories with improved detection
                categories_found = 0
                for category in config.VOTING_RESPONSES.keys():
                    try:
                        from selenium.webdriver.common.by import By
                        # Wait a bit for dynamic content
                        time.sleep(1)
                        
                        # Multiple search approaches
                        found = False
                        
                        # Approach 1: Direct text search
                        elements = voter_temp.driver.find_elements(By.XPATH, 
                            f"//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{category.lower()}')]"
                        )
                        if elements:
                            found = True
                        
                        # Approach 2: Label search
                        if not found:
                            label_elements = voter_temp.driver.find_elements(By.XPATH, 
                                f"//label[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{category.lower()}')]"
                            )
                            if label_elements:
                                found = True
                        
                        # Approach 3: Input placeholder search
                        if not found:
                            input_elements = voter_temp.driver.find_elements(By.XPATH, 
                                f"//input[contains(@placeholder, '{category}')]"
                            )
                            if input_elements:
                                found = True
                        
                        if found:
                            categories_found += 1
                            
                    except Exception:
                        continue
                
                logger.info(f"📝 Found {categories_found}/{len(config.VOTING_RESPONSES)} target categories")
                
                # Determine if voting is ready
                if page_state == "voting" and categories_found >= 2:
                    logger.info("🎉 VOTING IS OPEN AND READY!")
                    return True
                elif page_state == "closed":
                    logger.info("⏳ Voting still closed. Will check again later.")
                    return False
                elif page_state == "results":
                    logger.info("🏆 Voting appears to be in results mode.")
                    return False
                else:
                    logger.info(f"❓ Unknown state: {page_state}. Will check again later.")
                    return False
                    
            finally:
                voter_temp.driver.quit()
                config.HEADLESS_MODE = original_headless
                
        except Exception as e:
            logger.error(f"❌ Error checking voting status: {e}")
            return False
    
    def perform_vote(self):
        """Perform a voting attempt"""
        current_time = datetime.now()
        
        # Check if we're still within the 2-week period
        if current_time > self.end_date:
            logger.info("⏰ 2-week voting period completed. Stopping scheduler.")
            return schedule.CancelJob
        
        self.vote_count += 1
        logger.info(f"🗳️ VOTE ATTEMPT #{self.vote_count} at {current_time}")
        
        try:
            success = self.voter.vote(config.MARATHON_URL)
            if success:
                logger.info(f"✅ Vote #{self.vote_count} completed successfully!")
                self.voting_started = True
            else:
                logger.warning(f"⚠️ Vote #{self.vote_count} failed - will try again next cycle")
                
        except Exception as e:
            logger.error(f"❌ Error during vote #{self.vote_count}: {e}")
        
        # Calculate remaining time
        days_remaining = (self.end_date - current_time).days
        hours_remaining = ((self.end_date - current_time).seconds // 3600)
        logger.info(f"⏰ Time remaining: {days_remaining} days, {hours_remaining} hours")
        
        return None  # Continue scheduling
    
    def monitoring_check(self):
        """Periodic check to see if voting has opened"""
        if not self.voting_started:
            logger.info("👀 MONITORING CHECK - Looking for voting availability...")
            
            if self.check_voting_status():
                logger.info("🚀 VOTING DETECTED! Switching to voting mode...")
                self.voting_started = True
                
                # Cancel monitoring and start voting schedule
                schedule.clear('monitoring')
                schedule.every(config.VOTING_INTERVAL_HOURS).hours.do(self.perform_vote).tag('voting')
                
                # Perform first vote immediately
                logger.info("🎯 Performing initial vote now...")
                self.perform_vote()
            else:
                logger.info("😴 Voting not ready yet. Will check again in 4 hours.")
        
        return None
    
    def start_smart_scheduler(self):
        """Start the intelligent monitoring and voting scheduler"""
        logger.info("🧠 STARTING SMART VOTING SCHEDULER")
        logger.info("=" * 60)
        logger.info(f"📅 Monitoring period: {config.TOTAL_VOTING_DAYS} days")
        logger.info(f"⏰ Voting interval: {config.VOTING_INTERVAL_HOURS} hours")
        logger.info(f"🎯 Target URL: {config.MARATHON_URL}")
        logger.info(f"📧 Email: {config.EMAIL}")
        logger.info(f"🏁 End date: {self.end_date}")
        logger.info("=" * 60)
        
        # Initial check
        logger.info("🔍 Performing initial voting status check...")
        if self.check_voting_status():
            logger.info("🎉 Voting is already open! Starting voting schedule immediately.")
            self.voting_started = True
            schedule.every(config.VOTING_INTERVAL_HOURS).hours.do(self.perform_vote).tag('voting')
            # Perform first vote now
            self.perform_vote()
        else:
            logger.info("⏳ Voting not yet open. Starting monitoring mode.")
            logger.info("🔍 Will check every 4 hours for voting availability.")
            # Schedule monitoring checks every 4 hours
            schedule.every(4).hours.do(self.monitoring_check).tag('monitoring')
        
        logger.info("\n🎮 Scheduler is now running! Press Ctrl+C to stop.")
        logger.info("📊 Check the log file 'smart_scheduler_log.txt' for detailed activity.")
        
        try:
            while True:
                schedule.run_pending()
                time.sleep(60)  # Check every minute
        except KeyboardInterrupt:
            logger.info("\n⏹️ Scheduler stopped by user")
            logger.info("📊 Final Summary:")
            logger.info(f"   🗳️ Total votes attempted: {self.vote_count}")
            logger.info(f"   ⏰ Last check: {self.last_check_time}")
            logger.info(f"   🎯 Voting started: {self.voting_started}")
        except Exception as e:
            logger.error(f"❌ Scheduler error: {e}")
        
        return True
    
    def status_report(self):
        """Generate a status report"""
        current_time = datetime.now()
        days_remaining = (self.end_date - current_time).days
        
        logger.info("\n" + "=" * 50)
        logger.info("📊 SMART SCHEDULER STATUS REPORT")
        logger.info("=" * 50)
        logger.info(f"🕐 Current time: {current_time}")
        logger.info(f"🎯 Voting started: {self.voting_started}")
        logger.info(f"🗳️ Votes attempted: {self.vote_count}")
        logger.info(f"⏰ Days remaining: {days_remaining}")
        logger.info(f"🔍 Last check: {self.last_check_time}")
        logger.info("=" * 50)

def main():
    scheduler = SmartVotingScheduler()
    
    if len(sys.argv) > 1:
        if sys.argv[1] == "status":
            # Show current status
            scheduler.status_report()
        elif sys.argv[1] == "check":
            # Perform a one-time check
            scheduler.check_voting_status()
        elif sys.argv[1] == "vote":
            # Perform a one-time vote
            scheduler.perform_vote()
        else:
            print("Usage:")
            print("  python smart_monitor_scheduler.py          - Start smart scheduler")
            print("  python smart_monitor_scheduler.py status   - Show current status")
            print("  python smart_monitor_scheduler.py check    - Check voting availability once")
            print("  python smart_monitor_scheduler.py vote     - Attempt one vote")
    else:
        # Start the smart scheduler
        scheduler.start_smart_scheduler()

if __name__ == "__main__":
    main()
