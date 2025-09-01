#!/usr/bin/env python3
"""
Simple Daily Scheduler
Runs the voting bot once per day. That's it.
"""

import time
import datetime
import logging
from simple_voter import vote_daily

# Simple logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')
logger = logging.getLogger(__name__)

def run_scheduler():
    """Run voting once per day"""
    logger.info("📅 Simple daily scheduler started")
    logger.info("🎯 Will vote once every 24 hours")
    
    last_vote_date = None
    
    while True:
        try:
            today = datetime.date.today()
            
            # If we haven't voted today, vote now
            if last_vote_date != today:
                logger.info(f"📅 New day ({today}) - time to vote!")
                vote_daily()
                last_vote_date = today
                logger.info("✅ Daily vote complete. Sleeping until tomorrow...")
            else:
                logger.info(f"😴 Already voted today ({today}). Sleeping...")
            
            # Sleep for 1 hour, then check again
            time.sleep(3600)  # 1 hour
            
        except KeyboardInterrupt:
            logger.info("🛑 Scheduler stopped by user")
            break
        except Exception as e:
            logger.error(f"❌ Scheduler error: {e}")
            logger.info("😴 Sleeping 1 hour before retry...")
            time.sleep(3600)

if __name__ == "__main__":
    print("🎯 SIMPLE DAILY VOTING SCHEDULER")
    print("This will vote once per day. Press Ctrl+C to stop.")
    print()
    
    input("Press Enter to start...")
    run_scheduler()
