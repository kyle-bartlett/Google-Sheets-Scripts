"""
Configuration file for Keys Weekly voting automation
"""

# Login credentials
EMAIL = "krbartle@gmail.com"

# Voting URLs
MARATHON_URL = "https://keysweekly.com/bom25/#/gallery/?group=522963"
KEY_WEST_URL = "https://keysweekly.com/buk24/#/gallery?group=497175"

# Default to Marathon (active voting)
DEFAULT_URL = MARATHON_URL

# Voting responses for each category
VOTING_RESPONSES = {
    "Realtor": "Nate Bartlett",
    "Real Estate office": "Berkshire Hathaway Keys Real Estate 9141 Overseas",
    "Vacation Rental Company": "Berkshire Keys Vacation Rentals", 
    "Best Business": "Berkshire Hathaway Keys Real Estate 9141 overseas"
}

# Alternative category names that might appear on the site
CATEGORY_ALIASES = {
    "Realtor": ["realtor", "real estate agent", "agent"],
    "Real Estate office": ["real estate office", "realty", "real estate company"],
    "Vacation Rental Company": ["vacation rental", "rental company", "vacation rentals"],
    "Best Business": ["best business", "business", "local business"]
}

# Timing settings
VOTING_INTERVAL_HOURS = 24
TOTAL_VOTING_DAYS = 14

# Browser settings
HEADLESS_MODE = False  # Set to True to run without browser window
WAIT_TIMEOUT = 20  # seconds to wait for elements
SCROLL_PAUSE_TIME = 2  # seconds to pause between scrolls
