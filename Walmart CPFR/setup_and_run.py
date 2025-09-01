"""
Setup and Run Script for Keys Weekly Voting Automation
This script helps you install dependencies and run the voting automation
"""

import subprocess
import sys
import os

def install_requirements():
    """Install required Python packages"""
    print("Installing required packages...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
        print("✅ Requirements installed successfully!")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to install requirements: {e}")
        return False

def check_chrome():
    """Check if Chrome is installed"""
    print("Checking for Google Chrome...")
    try:
        # Try to run chrome --version
        result = subprocess.run(["google-chrome", "--version"], 
                              capture_output=True, text=True, timeout=5)
        if result.returncode == 0:
            print(f"✅ Chrome found: {result.stdout.strip()}")
            return True
    except:
        pass
    
    try:
        # Try macOS Chrome path
        result = subprocess.run(["/Applications/Google Chrome.app/Contents/MacOS/Google Chrome", "--version"], 
                              capture_output=True, text=True, timeout=5)
        if result.returncode == 0:
            print(f"✅ Chrome found: {result.stdout.strip()}")
            return True
    except:
        pass
    
    print("❌ Google Chrome not found. Please install Google Chrome.")
    print("Download from: https://www.google.com/chrome/")
    return False

def run_test():
    """Run a test vote to verify everything works"""
    print("\n🧪 Running test vote...")
    try:
        from keys_weekly_voter import KeysWeeklyVoter
        voter = KeysWeeklyVoter()
        success = voter.test_run()
        if success:
            print("✅ Test vote completed successfully!")
        else:
            print("❌ Test vote failed. Check the logs for details.")
        return success
    except Exception as e:
        print(f"❌ Test failed with error: {e}")
        return False

def start_scheduler():
    """Start the voting scheduler"""
    print("\n🚀 Starting voting scheduler...")
    print("This will run votes every 24 hours for 2 weeks.")
    print("Press Ctrl+C to stop the scheduler at any time.")
    
    try:
        from voting_scheduler import VotingScheduler
        scheduler = VotingScheduler()
        scheduler.start_scheduler(run_test_first=False)
    except KeyboardInterrupt:
        print("\n⏹️ Scheduler stopped by user")
    except Exception as e:
        print(f"❌ Scheduler error: {e}")

def main():
    print("🗳️  Keys Weekly Voting Automation Setup")
    print("=" * 50)
    
    # Step 1: Install requirements
    if not install_requirements():
        print("Setup failed. Please fix the requirements installation.")
        return
    
    # Step 2: Check Chrome
    if not check_chrome():
        print("Setup failed. Please install Google Chrome.")
        return
    
    print("\n✅ Setup completed successfully!")
    
    # Ask user what they want to do
    while True:
        print("\nWhat would you like to do?")
        print("1. Run a test vote")
        print("2. Start the automated scheduler (24 hours for 2 weeks)")
        print("3. Exit")
        
        choice = input("\nEnter your choice (1-3): ").strip()
        
        if choice == "1":
            run_test()
        elif choice == "2":
            start_scheduler()
            break
        elif choice == "3":
            print("👋 Goodbye!")
            break
        else:
            print("Invalid choice. Please enter 1, 2, or 3.")

if __name__ == "__main__":
    main()
