"""
Scheduler for Keys Weekly Voting Automation
Runs the voting bot every 24 hours for 2 weeks
"""

import schedule
import time
import logging
import sys
from datetime import datetime, timedelta
from keys_weekly_voter import KeysWeeklyVoter
import config

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('scheduler_log.txt'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

class VotingScheduler:
    def __init__(self):
        self.voter = KeysWeeklyVoter()
        self.start_date = datetime.now()
        self.end_date = self.start_date + timedelta(days=config.TOTAL_VOTING_DAYS)
        self.vote_count = 0
        
    def scheduled_vote(self):
        """Function called by scheduler to perform voting"""
        current_time = datetime.now()
        
        # Check if we're still within the 2-week period
        if current_time > self.end_date:
            logger.info("2-week voting period completed. Stopping scheduler.")
            return schedule.CancelJob
        
        self.vote_count += 1
        logger.info(f"=== SCHEDULED VOTE #{self.vote_count} at {current_time} ===")
        
        try:
            success = self.voter.vote()
            if success:
                logger.info(f"Vote #{self.vote_count} completed successfully")
            else:
                logger.error(f"Vote #{self.vote_count} failed")
        except Exception as e:
            logger.error(f"Error during scheduled vote #{self.vote_count}: {e}")
        
        # Calculate days remaining
        days_remaining = (self.end_date - current_time).days
        logger.info(f"Days remaining in voting period: {days_remaining}")
        
        return None  # Continue scheduling
    
    def run_immediate_test(self):
        """Run an immediate test vote"""
        logger.info("=== RUNNING IMMEDIATE TEST VOTE ===")
        success = self.voter.test_run()
        return success
    
    def start_scheduler(self, run_test_first=True):
        """Start the 24-hour voting scheduler"""
        logger.info(f"Starting voting scheduler for {config.TOTAL_VOTING_DAYS} days")
        logger.info(f"Voting will run every {config.VOTING_INTERVAL_HOURS} hours")
        logger.info(f"Start date: {self.start_date}")
        logger.info(f"End date: {self.end_date}")
        
        # Run immediate test if requested
        if run_test_first:
            test_success = self.run_immediate_test()
            if not test_success:
                logger.error("Initial test failed. Please check configuration and try again.")
                return False
        
        # Schedule the voting job
        schedule.every(config.VOTING_INTERVAL_HOURS).hours.do(self.scheduled_vote)
        
        logger.info("Scheduler started. Press Ctrl+C to stop.")
        
        try:
            while True:
                schedule.run_pending()
                time.sleep(60)  # Check every minute
        except KeyboardInterrupt:
            logger.info("Scheduler stopped by user")
        except Exception as e:
            logger.error(f"Scheduler error: {e}")
        
        return True
    
    def run_once_now(self):
        """Run voting once immediately (for testing)"""
        return self.scheduled_vote()

def main():
    scheduler = VotingScheduler()
    
    if len(sys.argv) > 1:
        if sys.argv[1] == "test":
            # Run test only
            scheduler.run_immediate_test()
        elif sys.argv[1] == "once":
            # Run once immediately
            scheduler.run_once_now()
        elif sys.argv[1] == "schedule":
            # Start full scheduler without initial test
            scheduler.start_scheduler(run_test_first=False)
        else:
            print("Usage:")
            print("  python voting_scheduler.py test      - Run test vote only")
            print("  python voting_scheduler.py once      - Run one vote immediately")
            print("  python voting_scheduler.py schedule  - Start scheduler without test")
            print("  python voting_scheduler.py           - Run test then start scheduler")
    else:
        # Default: run test then start scheduler
        scheduler.start_scheduler(run_test_first=True)

if __name__ == "__main__":
    main()
