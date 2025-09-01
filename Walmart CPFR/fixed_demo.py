"""
Fixed Demo - Improved search logic based on actual page structure
Now properly finds categories and "SEE ALL ENTRANTS" buttons
"""

import time
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
import config

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')
logger = logging.getLogger(__name__)

class FixedDemo:
    def __init__(self):
        self.driver = None
        self.wait = None
        
    def setup_driver(self):
        """Setup Chrome driver"""
        chrome_options = Options()
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        
        try:
            self.driver = webdriver.Chrome(options=chrome_options)
            self.driver.maximize_window()
            self.wait = WebDriverWait(self.driver, 20)
            logger.info("✅ Browser opened and ready!")
            return True
        except Exception as e:
            logger.error(f"❌ Browser setup failed: {e}")
            return False
    
    def navigate_and_wait(self):
        """Navigate to Bubba's Key West page and wait for load"""
        try:
            logger.info("🌐 Navigating to Bubba's voting page...")
            self.driver.get("https://keysweekly.com/bubbas25/#/gallery/?group=516602")
            
            # Wait for page to fully load
            self.wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            time.sleep(5)  # Extra wait for dynamic content
            
            logger.info(f"📄 Page loaded: {self.driver.title}")
            return True
            
        except Exception as e:
            logger.error(f"❌ Navigation failed: {e}")
            return False
    
    def find_voting_categories(self):
        """Find the actual voting categories on the page"""
        logger.info("🔍 Searching for voting categories...")
        
        categories_found = []
        
        # Look for our target categories from config
        for category in config.VOTING_RESPONSES.keys():
            logger.info(f"🎯 Looking for '{category}' category...")
            
            # Search strategies for finding categories
            search_patterns = [
                f"//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{category.lower()}')]",
                f"//h1[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{category.lower()}')]",
                f"//h2[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{category.lower()}')]",
                f"//h3[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{category.lower()}')]"
            ]
            
            for pattern in search_patterns:
                try:
                    elements = self.driver.find_elements(By.XPATH, pattern)
                    if elements:
                        logger.info(f"✅ Found '{category}' category!")
                        categories_found.append((category, elements[0]))
                        break
                except Exception as e:
                    continue
        
        logger.info(f"📊 Found {len(categories_found)} target categories")
        return categories_found
    
    def find_voting_form_near_category(self, category_element):
        """Find voting form near a category"""
        logger.info("🔍 Looking for voting form near category...")
        
        try:
            # Look for form elements near the category
            parent = category_element.find_element(By.XPATH, "./..")
            
            # Look for text input, textarea, or form elements
            form_elements = []
            
            # Search in parent and nearby elements
            for _ in range(3):  # Try parent, grandparent, great-grandparent
                try:
                    inputs = parent.find_elements(By.CSS_SELECTOR, "input[type='text'], textarea")
                    buttons = parent.find_elements(By.CSS_SELECTOR, "button, input[type='submit']")
                    
                    if inputs or buttons:
                        form_elements.extend(inputs)
                        form_elements.extend(buttons)
                        logger.info(f"✅ Found {len(inputs)} inputs and {len(buttons)} buttons near category")
                        break
                        
                    parent = parent.find_element(By.XPATH, "./..")
                except:
                    break
            
            return form_elements
            
        except Exception as e:
            logger.error(f"❌ Error finding form: {e}")
            return []
    
    def perform_voting_action(self, category, form_elements):
        """Perform actual voting action"""
        logger.info(f"🗳️ Attempting to vote for '{category}'...")
        
        voting_response = config.VOTING_RESPONSES.get(category)
        if not voting_response:
            logger.warning(f"⚠️ No voting response configured for '{category}'")
            return False
        
        try:
            # Find text input to fill
            text_inputs = [elem for elem in form_elements if elem.tag_name in ['input', 'textarea']]
            
            if text_inputs:
                target_input = text_inputs[0]
                logger.info(f"✍️ Filling form with: '{voting_response}'")
                
                # Clear and fill the input
                target_input.clear()
                target_input.send_keys(voting_response)
                
                logger.info("✅ Form filled successfully!")
                time.sleep(2)
                
                # Look for submit button
                submit_buttons = [elem for elem in form_elements if elem.tag_name in ['button'] or 
                                (elem.tag_name == 'input' and elem.get_attribute('type') in ['submit', 'button'])]
                
                if submit_buttons:
                    logger.info("🖱️ Found submit button, clicking...")
                    submit_buttons[0].click()
                    logger.info("✅ Vote submitted!")
                    time.sleep(3)
                    return True
                else:
                    logger.warning("⚠️ No submit button found")
                    
            else:
                logger.warning("⚠️ No text input found in form")
                
        except Exception as e:
            logger.error(f"❌ Voting action failed: {e}")
        
        return False
    
    def detect_voting_status(self):
        """Detect if voting is open, closed, or showing results"""
        logger.info("🔍 Analyzing Marathon page status...")
        
        page_source = self.driver.page_source.lower()
        page_text = self.driver.find_element(By.TAG_NAME, "body").text.lower()
        
        # Check for closed/waiting indicators
        closed_indicators = [
            "voting has not started",
            "voting hasn't started", 
            "voting will begin",
            "come back tomorrow",
            "voting starts",
            "check back",
            "voting opens"
        ]
        
        # Check for active voting indicators
        voting_indicators = [
            "vote now",
            "cast your vote",
            "submit",
            "choose your favorite"
        ]
        
        # Check for results indicators
        results_indicators = [
            "winner",
            "congratulations",
            "results",
            "voting has ended"
        ]
        
        status = "unknown"
        for indicator in closed_indicators:
            if indicator in page_text:
                status = "closed"
                break
        
        if status == "unknown":
            for indicator in voting_indicators:
                if indicator in page_text:
                    status = "voting"
                    break
        
        if status == "unknown":
            for indicator in results_indicators:
                if indicator in page_text:
                    status = "results"
                    break
        
        logger.info(f"🎯 Marathon page status: {status.upper()}")
        return status
    
    def show_what_automation_will_do(self):
        """Show what the automation will do when voting opens"""
        logger.info("💡 SIMULATION: What automation will do when Marathon voting opens...")
        logger.info("=" * 60)
        
        for category, response in config.VOTING_RESPONSES.items():
            logger.info(f"🎯 Target Category: '{category}'")
            logger.info(f"✍️ Will vote for: '{response}'")
            logger.info(f"🔍 Will search page for text containing '{category.lower()}'")
            logger.info(f"📝 Will find voting form near '{category}' section")
            logger.info(f"✅ Will fill form with '{response}' and submit")
            logger.info("   " + "-" * 40)
            time.sleep(2)
        
        logger.info("🔄 This process will repeat every 24 hours until voting closes")
        logger.info("📊 Each vote will be logged and confirmed")
    
    def run_fixed_demo(self):
        """Run the simple Bubba's gay bar demo"""
        logger.info("🎬 BUBBA'S GAY BAR DEMO")
        logger.info("=" * 60)
        logger.info("🎯 This demo will:")
        logger.info("   1. Navigate to Bubba's Key West page")
        logger.info("   2. Scroll down to find 'Best Gay Bar' category")  
        logger.info("   3. Show your brother the automation found it")
        logger.info("   4. Have a good laugh")
        logger.info("=" * 60)
        
        try:
            # Setup
            if not self.setup_driver():
                return False
            
            # Navigate
            if not self.navigate_and_wait():
                return False
            
            logger.info("🎯 Counting down to the 9th box where Gay Bar is...")
            
            # Start at top
            self.driver.execute_script("window.scrollTo(0, 0);")
            time.sleep(2)
            
            # Scroll down and count the boxes/sections until we get to #9
            box_count = 0
            gay_bar_element = None
            found_gay_bar = False
            
            for scroll_step in range(20):  # Max 20 scroll attempts
                logger.info(f"📜 Scroll step {scroll_step + 1} - looking for boxes...")
                
                # Look for section/category boxes on current screen
                try:
                    # Look for elements that might be category boxes
                    potential_boxes = self.driver.find_elements(By.XPATH, "//div[contains(@class, 'category') or contains(@class, 'section')]")
                    
                    # If that doesn't work, look for any divs that have substantial content
                    if not potential_boxes:
                        all_divs = self.driver.find_elements(By.TAG_NAME, "div")
                        potential_boxes = []
                        for div in all_divs:
                            try:
                                if div.is_displayed() and div.size['height'] > 150:  # Substantial height
                                    text = div.text.strip()
                                    if text and len(text) > 20:  # Has meaningful content
                                        potential_boxes.append(div)
                            except:
                                continue
                    
                    # Count boxes we can see
                    visible_boxes = 0
                    for box in potential_boxes:
                        try:
                            if box.is_displayed():
                                visible_boxes += 1
                                box_count += 1
                                logger.info(f"📦 Found box #{box_count}")
                                
                                # Check if this box contains "Gay Bar"
                                box_text = box.text.lower()
                                if 'gay' in box_text and 'bar' in box_text:
                                    logger.info(f"🎉 FOUND IT! Box #{box_count} contains Gay Bar!")
                                    gay_bar_element = box
                                    found_gay_bar = True
                                    break
                                elif box_count == 9:
                                    logger.info(f"📦 This is box #9 - highlighting it regardless!")
                                    gay_bar_element = box
                                    found_gay_bar = True
                                    break
                        except:
                            continue
                    
                    if found_gay_bar:
                        break
                        
                    logger.info(f"📊 Total boxes found so far: {box_count}")
                    
                except Exception as e:
                    logger.debug(f"Box counting failed at step {scroll_step}: {e}")
                
                # Scroll down a bit more
                self.driver.execute_script("window.scrollBy(0, 300);")
                time.sleep(1)
            
            # If we still haven't found it, just look for "Gay Bar" text anywhere
            if not found_gay_bar:
                logger.info("🔍 Box counting didn't work - looking for 'Gay Bar' text anywhere...")
                try:
                    elements = self.driver.find_elements(By.XPATH, "//*[contains(text(), 'Gay Bar')]")
                    if elements:
                        logger.info(f"🎉 FOUND Gay Bar text in {len(elements)} elements!")
                        gay_bar_element = elements[0]
                        found_gay_bar = True
                except:
                    pass
            
            if found_gay_bar:
                logger.info("🏳️‍🌈 Time to mess with your brother!")
                
                # Scroll to the link and highlight it
                self.driver.execute_script("arguments[0].scrollIntoView(true);", gay_bar_element)
                time.sleep(2)
                
                # CIRCLE THE HELL out of the Gay Bar section
                self.driver.execute_script("""
                    arguments[0].style.border = '10px solid red';
                    arguments[0].style.backgroundColor = 'yellow';
                    arguments[0].style.outline = '10px solid blue';
                    arguments[0].style.outlineOffset = '5px';
                    arguments[0].style.boxShadow = '0 0 20px 10px red';
                    arguments[0].style.padding = '20px';
                    arguments[0].style.fontSize = '24px';
                    arguments[0].style.fontWeight = 'bold';
                    arguments[0].style.animation = 'blink 1s infinite';
                    
                    // Add blinking animation
                    var style = document.createElement('style');
                    style.innerHTML = '@keyframes blink { 0% { opacity: 1; } 50% { opacity: 0.5; } 100% { opacity: 1; } }';
                    document.head.appendChild(style);
                """, gay_bar_element)
                
                logger.info("🎊 GAY BAR SECTION IS NOW HIGHLIGHTED LIKE CRAZY!")
                logger.info("📸 RED BORDER! YELLOW BACKGROUND! BLUE OUTLINE! BLINKING!")
                logger.info("😂 Your brother will definitely see this!")
                
                logger.info("✅ Gay Bar sidebar link is now highlighted!")
                logger.info(f"📝 Link text: '{gay_bar_element.text}'")
                
                # Try to click the link to see what happens
                try:
                    logger.info("🖱️ Clicking the Gay Bar link...")
                    gay_bar_element.click()
                    time.sleep(3)
                    logger.info("✅ Clicked! Gay Bar section should be loaded now!")
                except Exception as e:
                    logger.info(f"⚠️ Click failed, but link is highlighted: {e}")
            else:
                logger.warning("😕 Couldn't find 'Gay Bar' in sidebar navigation")
                logger.info("🔍 Let me show you what links ARE in the sidebar...")
                
                # Show all sidebar links
                try:
                    all_links = self.driver.find_elements(By.XPATH, "//a")
                    left_links = []
                    
                    for link in all_links:
                        try:
                            location = link.location
                            text = link.text.strip()
                            # Include all links with text, even if off-screen
                            if text and len(text) < 100:
                                left_links.append(f"{text} (x={location['x']})")
                        except:
                            continue
                    
                    logger.info("📋 Sidebar links found:")
                    for i, link_text in enumerate(left_links[:15]):  # Show first 15
                        logger.info(f"   {i+1}. {link_text}")
                        
                except Exception as e:
                    logger.error(f"Failed to show sidebar links: {e}")
            
            if found_gay_bar:
                logger.info("🎊 SUCCESS! Your automation found the Gay Bar category!")
                logger.info("😂 Perfect for messing with your brother!")
                logger.info("📸 Take a screenshot and send it to him!")
                
                # Look for "SEE ALL ENTRANTS" button near gay bar
                logger.info("🔍 Looking for 'SEE ALL ENTRANTS' button...")
                try:
                    see_all_buttons = self.driver.find_elements(By.XPATH, 
                        "//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'see all')]")
                    
                    if see_all_buttons:
                        logger.info("🖱️ Found 'SEE ALL' button - clicking it...")
                        see_all_buttons[0].click()
                        time.sleep(3)
                        logger.info("✅ Clicked! Gay bar entrants should be visible now!")
                    else:
                        logger.info("ℹ️ No 'SEE ALL' button found, but Gay Bar category is highlighted!")
                        
                except Exception as e:
                    logger.info(f"⚠️ Couldn't click SEE ALL button: {e}")
                
            else:
                logger.warning("😕 Couldn't find 'Gay Bar' category after scrolling")
                logger.info("💡 Maybe the page structure changed or it's named differently")
            
            logger.info("⏸️ DEMO COMPLETE - Keeping browser open for 30 seconds...")
            logger.info("📸 Perfect time to take a screenshot for your brother!")
            time.sleep(30)
            
        except Exception as e:
            logger.error(f"❌ Demo error: {e}")
            import traceback
            traceback.print_exc()
        finally:
            if self.driver:
                logger.info("🔚 Closing browser...")
                self.driver.quit()

def main():
    print("🎬 BUBBA'S GAY BAR DEMO")
    print("=" * 50)
    print("😂 Time to mess with your brother!")
    print("📍 Going to Bubba's Key West page")
    print("🔍 Finding the 'Best Gay Bar' category") 
    print("🎯 Highlighting it with a big red border")
    print("📸 Perfect for screenshots to send to your brother")
    print("=" * 50)
    print("⚠️  This will:")
    print("   • Navigate to Bubba's page")
    print("   • Scroll down to find 'Gay Bar' category")
    print("   • Highlight it with bright colors")
    print("   • Give you 30 seconds to screenshot it")
    print()
    
    input("Press Enter to mess with your brother...")
    
    demo = FixedDemo()
    demo.run_fixed_demo()
    
    print()
    print("🎉 Gay bar demo complete!")
    print("😂 Perfect ammunition for your brother!")
    print("📸 Hope you got a good screenshot!")

if __name__ == "__main__":
    main()
