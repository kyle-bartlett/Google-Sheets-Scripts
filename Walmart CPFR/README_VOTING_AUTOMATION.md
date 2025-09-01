# Keys Weekly Voting Automation

This automation script will automatically vote on the Keys Weekly community website every 24 hours for 2 weeks.

## Features

- **Automatic Login**: Uses your email (krbartle@gmail.com) to log in
- **Smart Category Finding**: Scrolls and searches for your 4 voting categories
- **Daily Voting**: Runs every 24 hours automatically
- **Error Handling**: Robust error handling and logging
- **Flexible URLs**: Can work with both Marathon and Key West voting pages
- **Detailed Logging**: Tracks all voting attempts and results

## Your Voting Categories

The script will vote for:
- **Realtor**: Nate Bartlett
- **Real Estate office**: Berkshire Hathaway Keys Real Estate 9141 Overseas
- **Vacation Rental Company**: Berkshire Keys Vacation Rentals
- **Best Business**: Berkshire Hathaway Keys Real Estate 9141 overseas

## Quick Start

1. **Run the setup script**:
   ```bash
   python3 setup_and_run.py
   ```

2. **Choose option 1** to run a test vote first
3. **Choose option 2** to start the automated 24-hour scheduler

## Manual Usage

### Install Dependencies
```bash
pip3 install -r requirements.txt
```

### Run a Test Vote
```bash
python3 keys_weekly_voter.py test
```

### Start the Scheduler
```bash
python3 voting_scheduler.py
```

## Configuration

Edit `config.py` to modify:
- Voting responses
- URLs (Marathon vs Key West)
- Timing settings
- Browser settings (headless mode, etc.)

## Files

- `keys_weekly_voter.py` - Main voting automation script
- `voting_scheduler.py` - 24-hour scheduler
- `config.py` - Configuration settings
- `setup_and_run.py` - Easy setup and run script
- `requirements.txt` - Python dependencies

## Logs

The script creates detailed logs:
- `voting_log.txt` - Individual voting attempts
- `scheduler_log.txt` - Scheduler activity

## Troubleshooting

### Chrome Driver Issues
The script automatically downloads the Chrome driver, but make sure Google Chrome is installed.

### Website Changes
If the website structure changes, you may need to update the category finding logic in `keys_weekly_voter.py`.

### Network Issues
The script includes retry logic and error handling for network issues.

## Current Status

- **Active URL**: https://keysweekly.com/buk24/#/gallery?group=497175 (Key West voting)
- **Future URL**: https://keysweekly.com/bom25/#/gallery/?group=522963 (Marathon - when voting opens)

The script is configured to use the Key West page since voting is currently active there. You can easily switch to the Marathon page by updating `DEFAULT_URL` in `config.py` when that voting opens.

## Safety Features

- Only votes once per 24-hour period
- Automatically stops after 2 weeks
- Detailed logging for monitoring
- Safe browser automation practices

## Running for 2 Weeks

The scheduler will automatically:
1. Run every 24 hours
2. Track voting attempts
3. Stop after 14 days
4. Log all activity

You can stop it anytime with Ctrl+C and restart later if needed.
