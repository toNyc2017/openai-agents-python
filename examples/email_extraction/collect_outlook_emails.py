#!/usr/bin/env python3
import time, re
from pathlib import Path
from playwright.async_api import async_playwright
import asyncio
import logging
import os
from datetime import datetime
import pdb

# Set up logging with timestamps
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(f'email_collection_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ───── CONFIG ────────────────────────────────────────────────────────────────
EMAIL         = "tolds@ceresinsurance.com"
PASSWORD      = "annuiTy2024!"
SEARCH_PERSON = "Lynn Gadue"
SEARCH_PERSON = "Joe Bentivoglio"
N_EMAILS      = 15
OUT_FILE      = Path("outlook_emails.txt")

# ───── HELPERS ────────────────────────────────────────────────────────────────
def clean_email_content(content: str) -> str:
    ui_bits = ["Reply", "Reply all", "Forward", "Summary by Copilot"]
    for bit in ui_bits:
        content = content.replace(bit, "")
    content = re.sub(r"\n{3,}", "\n\n", content)
    return "\n".join(line for line in content.splitlines() if line.strip())

async def scroll_message_list(page, target_count):
    """Scroll through emails in Outlook with improved reliability."""
    print("\n=== Starting scroll_message_list ===")
    print(f"Target count: {target_count}")
    
    # Track unique email IDs
    seen_email_ids = set()
    
    # Get initial email count and IDs
    email_elements = await page.locator("div[data-convid]").all()
    for element in email_elements:
        email_id = await element.get_attribute("data-convid")
        if email_id:
            seen_email_ids.add(email_id)
    
    initial_count = len(seen_email_ids)
    print(f"\nInitial unique email count: {initial_count}")
    
    # Use only the working method
    working_method = lambda: page.evaluate("""() => {
        const emails = document.querySelectorAll('div[data-convid]');
        if (emails.length === 0) return false;
        const lastEmail = emails[emails.length - 1];
        lastEmail.scrollIntoView({ behavior: 'smooth', block: 'center' });
        window.scrollBy(0, 500);
        return true;
    }""")
    
    consecutive_failures = 0
    max_consecutive_failures = 5
    
    print("\nUsing working scroll method...")
    while len(seen_email_ids) < target_count and consecutive_failures < max_consecutive_failures:
        try:
            # Execute scroll
            await working_method()
            print("✅ Scroll executed successfully")
            
            # Wait longer for content to load
            await page.wait_for_timeout(5000)
            
            # Get all email elements and their IDs
            email_elements = await page.locator("div[data-convid]").all()
            new_emails_found = False
            
            for element in email_elements:
                email_id = await element.get_attribute("data-convid")
                if email_id and email_id not in seen_email_ids:
                    seen_email_ids.add(email_id)
                    new_emails_found = True
            
            new_count = len(seen_email_ids)
            print(f"New unique email count: {new_count} (was {initial_count})")
            
            if new_emails_found:
                print(f"✅ Progress! Gained {new_count - initial_count} new unique emails")
                initial_count = new_count
                consecutive_failures = 0
                
                # Add a small delay between successful scrolls
                await page.wait_for_timeout(2000)
            else:
                consecutive_failures += 1
                print(f"❌ No new unique emails found (consecutive failures: {consecutive_failures})")
                
                # Try a more aggressive scroll on failure
                if consecutive_failures > 0:
                    await page.evaluate("""() => {
                        window.scrollBy(0, 1000);
                    }""")
                    await page.wait_for_timeout(3000)
                
                if consecutive_failures >= max_consecutive_failures:
                    print("⚠️ Reached maximum consecutive failures - assuming end of inbox")
                    return new_count
        except Exception as e:
            print(f"❌ Error during scrolling: {str(e)}")
            consecutive_failures += 1
            if consecutive_failures >= max_consecutive_failures:
                print("⚠️ Reached maximum consecutive failures - assuming end of inbox")
                return len(seen_email_ids)
            continue
    
    # Get final count
    final_count = len(seen_email_ids)
    print(f"\n=== Final Results ===")
    print(f"Initial count: {initial_count}")
    print(f"Final count: {final_count}")
    print(f"Total unique emails gained: {final_count - initial_count}")
    return final_count

# ───── MAIN ─────────────────────────────────────────────────────────────────
async def collect_outlook_emails(email, password, search_person, n_emails=10):
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        page = await browser.new_page()
        
        try:
            # Navigate to Outlook
            logger.info("Navigating to Outlook...")
            await page.goto("https://outlook.office.com")
            await page.wait_for_timeout(2000)
            
            # Login process
            logger.info("Starting login process...")
            await page.fill('input[type="email"]', email)
            await page.click('input[type="submit"]')
            await page.wait_for_timeout(2000)
            
            await page.fill('input[type="password"]', password)
            await page.click('input[type="submit"]')
            await page.wait_for_timeout(2000)
            
            # Wait for MFA
            logger.info("🔐 Please complete the MFA prompt on your phone.")
            input("When you've approved MFA, press Enter here to continue...")
            logger.info("🔐 Logged in and mailbox loaded")
            
            # Wait for inbox to load
            await page.wait_for_timeout(2000)
            
            # --- SEARCH FOR EMAILS FROM THE TARGET PERSON ---
            logger.info(f"🔍 Searching for emails from {search_person}...")
            search_query = f"from:{search_person}"

            async def apply_search_filter():
                # Try multiple selectors for the search input
                search_selectors = [
                    'input[aria-label*="Search"]',
                    'input[placeholder*="Search"]',
                    'input[aria-label*="Search"]',
                    'input[aria-label*="search"]',
                    'input[class*="search"]',
                    'input[type="search"]'
                ]

                for selector in search_selectors:
                    try:
                        logger.info(f"Trying search selector: {selector}")
                        search_box = page.locator(selector)
                        if await search_box.count() > 0:
                            logger.info(f'Found search box using selector: {selector}')
                            # Clear any existing search first
                            await search_box.fill("")
                            await page.wait_for_timeout(1000)
                            # Enter new search
                            await search_box.fill(search_query)
                            await page.wait_for_timeout(1000)
                            await search_box.press('Enter')
                            logger.info(f"Search query '{search_query}' submitted successfully")
                            # Wait longer for search results to load
                            await page.wait_for_timeout(5000)
                            
                            # Verify search was applied
                            current_search = await page.evaluate("""
                                () => {
                                    const searchSelectors = [
                                        'input[aria-label*="Search"]',
                                        'input[placeholder*="Search"]',
                                        'input[class*="search"]',
                                        'input[type="search"]'
                                    ];
                                    for (const selector of searchSelectors) {
                                        const searchBox = document.querySelector(selector);
                                        if (searchBox) {
                                            return searchBox.value;
                                        }
                                    }
                                    return '';
                                }
                            """)
                            
                            if search_query.lower() in current_search.lower():
                                logger.info("Search filter verified successfully")
                                return True
                            else:
                                logger.warning(f"Search filter verification failed. Current search: {current_search}")
                    except Exception as e:
                        logger.debug(f"Failed with selector {selector}: {str(e)}")
                        continue
                return False

            # Initial search
            if not await apply_search_filter():
                logger.error("Could not find search input field")
                await page.screenshot(path="screenshots/search_failed.png")
                return False
            
            # Wait for search results to load with better error handling
            try:
                await page.wait_for_selector("div[data-convid]", timeout=10000)
                logger.info("Search results loaded successfully")
            except Exception as e:
                logger.error(f"Timeout waiting for search results: {str(e)}")
                await page.screenshot(path="screenshots/search_timeout.png")
                return False

            # Track processed email IDs and subjects to avoid duplicates
            processed_ids = set()
            processed_subjects = set()  # Keep as secondary check
            emails_collected = 0
            consecutive_failures = 0
            max_consecutive_failures = 3

            async def verify_search_filter():
                try:
                    current_search = await page.evaluate("""
                        () => {
                            const searchSelectors = [
                                'input[aria-label*="Search"]',
                                'input[placeholder*="Search"]',
                                'input[class*="search"]',
                                'input[type="search"]'
                            ];
                            for (const selector of searchSelectors) {
                                const searchBox = document.querySelector(selector);
                                if (searchBox) {
                                    return searchBox.value;
                                }
                            }
                            return '';
                        }
                    """)
                    is_valid = f"from:{search_person}".lower() in current_search.lower()
                    if not is_valid:
                        logger.warning(f"Search filter verification failed. Current search: {current_search}")
                    return is_valid
                except Exception as e:
                    logger.error(f"Error checking search filter: {str(e)}")
                    return False

            while emails_collected < n_emails:
                # Verify search filter is still active
                if not await verify_search_filter():
                    logger.info("Search filter lost, re-applying...")
                    if not await apply_search_filter():
                        logger.error("Failed to reapply search filter")
                        return False
                    await page.wait_for_timeout(5000)

                # Wait for email list to load with increased timeout
                try:
                    await page.wait_for_selector("div[data-convid]", timeout=10000)
                except Exception as e:
                    logger.error(f"Timeout waiting for email list: {str(e)}")
                    continue

                email_elements = await page.locator("div[data-convid]").all()
                logger.info(f"Found {len(email_elements)} emails in current view")

                found_new = False
                for email_element in email_elements:
                    try:
                        # Get email ID with better error handling
                        try:
                            email_id = await email_element.get_attribute("data-convid", timeout=10000)
                            if not email_id:
                                logger.warning("Email element has no data-convid attribute")
                                continue
                                
                            if email_id in processed_ids:
                                logger.debug(f"Skipping already processed email ID: {email_id}")
                                continue
                        except Exception as e:
                            logger.error(f"Error getting email ID: {str(e)}")
                            continue

                        # Get sender information with improved selectors
                        sender = None
                        try:
                            sender = await email_element.evaluate("""
                                el => {
                                    const senderSelectors = [
                                        '[role="heading"]',
                                        'span[title]',
                                        '[aria-label*="From"]',
                                        '.ms-Persona-primaryText',
                                        'span[class*="sender"]',
                                        'div[class*="from"]',
                                        'div[class*="persona"]',
                                        'div[class*="sender"]',
                                        'div[class*="from"] span',
                                        'div[class*="persona"] span'
                                    ];
                                    for (const selector of senderSelectors) {
                                        const senderEl = el.querySelector(selector);
                                        if (senderEl) {
                                            const text = senderEl.innerText || 
                                                       senderEl.textContent || 
                                                       senderEl.getAttribute('title') || 
                                                       senderEl.getAttribute('aria-label');
                                            if (text && text.trim()) {
                                                return text.trim();
                                            }
                                        }
                                    }
                                    return '';
                                }
                            """)
                            if not sender:
                                logger.warning(f"Could not find sender for email ID: {email_id}")
                                continue
                            logger.info(f"Found sender: {sender}")
                            if search_person.lower() not in sender.lower():
                                logger.info(f"Email not from {search_person}, skipping")
                                continue
                        except Exception as e:
                            logger.error(f"Error getting sender: {str(e)}")
                            continue

                        # Get subject with improved selectors and error handling
                        subject = None
                        try:
                            subject = await email_element.evaluate("""
                                el => {
                                    const subjectSelectors = [
                                        'div[class*="subject"]',
                                        'span[class*="subject"]',
                                        'div[role="link"]',
                                        '[aria-label*="Subject"]',
                                        'div[class*="message-subject"]',
                                        'div[class*="header"] span',
                                        'div[class*="title"]',
                                        'div[class*="message"] span',
                                        'div[class*="item"] span',
                                        'div[class*="text"] span',
                                        'div[class*="content"] span',
                                        'div[class*="preview"] span',
                                        'div[class*="summary"] span',
                                        'div[class*="header"] div',
                                        'div[class*="message"] div',
                                        'div[class*="item"] div',
                                        'div[class*="text"] div',
                                        'div[class*="content"] div',
                                        'div[class*="preview"] div',
                                        'div[class*="summary"] div',
                                        'div[class*="subject"] span',
                                        'div[class*="subject"] div',
                                        'div[class*="title"] span',
                                        'div[class*="title"] div'
                                    ];
                                    
                                    // Try to find subject in the element itself first
                                    if (el.getAttribute('aria-label') && el.getAttribute('aria-label').includes('Subject')) {
                                        return el.getAttribute('aria-label').replace('Subject:', '').trim();
                                    }
                                    
                                    // Try all selectors
                                    for (const selector of subjectSelectors) {
                                        const subjectEl = el.querySelector(selector);
                                        if (subjectEl) {
                                            // Try different ways to get the text
                                            const text = subjectEl.innerText || 
                                                        subjectEl.textContent || 
                                                        subjectEl.getAttribute('title') || 
                                                        subjectEl.getAttribute('aria-label');
                                            if (text && text.trim()) {
                                                return text.trim();
                                            }
                                        }
                                    }
                                    
                                    // If no subject found, try parent elements
                                    let current = el;
                                    while (current && current.parentElement) {
                                        current = current.parentElement;
                                        if (current.getAttribute('aria-label') && current.getAttribute('aria-label').includes('Subject')) {
                                            return current.getAttribute('aria-label').replace('Subject:', '').trim();
                                        }
                                        // Try selectors on parent elements too
                                        for (const selector of subjectSelectors) {
                                            const subjectEl = current.querySelector(selector);
                                            if (subjectEl) {
                                                const text = subjectEl.innerText || 
                                                            subjectEl.textContent || 
                                                            subjectEl.getAttribute('title') || 
                                                            subjectEl.getAttribute('aria-label');
                                                if (text && text.trim()) {
                                                    return text.trim();
                                                }
                                            }
                                        }
                                    }
                                    
                                    return '';
                                }
                            """)
                            if not subject:
                                logger.warning(f"Could not find subject for email ID: {email_id}")
                                # Don't skip here - we can still process the email without a subject
                        except Exception as e:
                            logger.error(f"Error getting subject for email ID {email_id}: {str(e)}")
                            # Don't skip here - we can still process the email without a subject

                        # Click to open email with better error handling and timeout
                        try:
                            await email_element.click(timeout=10000)
                            logger.info(f"Clicked email ID {email_id}, waiting for content...")
                            # Wait longer for content to load
                            await page.wait_for_timeout(3000)
                        except Exception as e:
                            logger.error(f"Error clicking email {email_id}: {str(e)}")
                            continue

                        # Get email content with improved selectors and error handling
                        content = None
                        try:
                            # Try multiple content selectors with better error handling
                            content_selectors = [
                                'div[role="document"]',
                                'div[class*="messageBody"]',
                                'div[class*="content"]',
                                'div[class*="body"]',
                                'div[class*="message"]',
                                'div[class*="text"]',
                                'div[class*="preview"]',
                                'div[class*="summary"]',
                                'div[class*="item"]'
                            ]
                            
                            for selector in content_selectors:
                                try:
                                    content_element = page.locator(selector).first
                                    if await content_element.count() > 0:
                                        content = await content_element.inner_text(timeout=5000)
                                        if content and content.strip():
                                            logger.info(f"Found content using selector: {selector}")
                                            break
                                except Exception as e:
                                    logger.debug(f"Failed with selector {selector}: {str(e)}")
                                    continue
                            
                            if not content:
                                # Fallback to evaluate if locators fail
                                content = await page.evaluate("""
                                    () => {
                                        const contentSelectors = [
                                            'div[role="document"]',
                                            'div[class*="messageBody"]',
                                            'div[class*="content"]',
                                            'div[class*="body"]',
                                            'div[class*="message"]',
                                            'div[class*="text"]',
                                            'div[class*="preview"]',
                                            'div[class*="summary"]',
                                            'div[class*="item"]'
                                        ];
                                        
                                        for (const selector of contentSelectors) {
                                            const contentEl = document.querySelector(selector);
                                            if (contentEl) {
                                                const text = contentEl.innerText || 
                                                           contentEl.textContent || 
                                                           contentEl.getAttribute('title') || 
                                                           contentEl.getAttribute('aria-label');
                                                if (text && text.trim()) {
                                                    return text.trim();
                                                }
                                            }
                                        }
                                        return '';
                                    }
                                """)
                            
                            if not content:
                                logger.warning(f"Could not find content for email ID: {email_id}")
                                content = "No content available"
                            else:
                                logger.info("Found email content")
                                
                        except Exception as e:
                            logger.error(f"Error getting content: {str(e)}")
                            content = "Error retrieving content"

                        # Save email with better error handling
                        try:
                            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                            filename = f"collected_emails/email_{emails_collected+1}_{timestamp}.txt"
                            with open(filename, 'w', encoding='utf-8') as f:
                                f.write(f"From: {sender}\n")
                                f.write(f"Subject: {subject}\n")
                                f.write(f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                                f.write("\nContent:\n")
                                f.write(content)
                            logger.info(f"Saved email {emails_collected+1} to {filename}")
                        except Exception as e:
                            logger.error(f"Error saving email {emails_collected+1}: {str(e)}")
                            continue

                        processed_ids.add(email_id)
                        if subject:
                            processed_subjects.add(subject)
                        emails_collected += 1
                        found_new = True
                        logger.info(f"Collected {emails_collected}/{n_emails}: {subject}")

                        # Go back to email list with better error handling
                        try:
                            await page.go_back()
                            await page.wait_for_timeout(2000)
                        except Exception as e:
                            logger.error(f"Error going back to email list: {str(e)}")
                            # Try to recover by reloading the page
                            await page.reload()
                            await page.wait_for_timeout(5000)
                            continue

                        # Verify search filter is still present with improved reliability
                        try:
                            current_search = await page.evaluate("""
                                () => {
                                    const searchSelectors = [
                                        'input[aria-label*="Search"]',
                                        'input[placeholder*="Search"]',
                                        'input[class*="search"]',
                                        'input[type="search"]'
                                    ];
                                    for (const selector of searchSelectors) {
                                        const searchBox = document.querySelector(selector);
                                        if (searchBox) {
                                            return searchBox.value;
                                        }
                                    }
                                    return '';
                                }
                            """)
                            
                            if f"from:{search_person}".lower() not in current_search.lower():
                                logger.info("Search filter lost, re-applying...")
                                search_box = page.locator('input[aria-label*="Search"]').first
                                if await search_box.count() > 0:
                                    await search_box.fill(f"from:{search_person}")
                                    await search_box.press('Enter')
                                    await page.wait_for_timeout(3000)
                        except Exception as e:
                            logger.error(f"Error checking search filter: {str(e)}")
                            continue

                        if emails_collected >= n_emails:
                            break
                    except Exception as e:
                        logger.error(f"Error processing email: {str(e)}")
                        try:
                            await page.go_back()
                            await page.wait_for_timeout(2000)
                        except:
                            pass
                        continue

                if not found_new:
                    # Scroll for more emails
                    logger.info("No new matching emails found, scrolling for more…")
                    await scroll_message_list(page, n_emails)
                    await page.wait_for_timeout(2000)

            logger.info("Email collection complete")
            return True
            
        except Exception as e:
            logger.error(f"Error during email collection: {str(e)}")
            return False
        finally:
            await browser.close()

def compile_emails():
    """Compile all collected email files into a single file."""
    logger.info("Compiling all collected emails into a single file...")
    
    # Get all email files from the collected_emails directory
    email_files = sorted(Path("collected_emails").glob("email_*.txt"))
    
    if not email_files:
        logger.warning("No email files found to compile!")
        return
    
    # Create output filename with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"compiled_emails_{timestamp}.txt"
    
    try:
        with open(output_file, 'w', encoding='utf-8') as outfile:
            # Write header
            outfile.write("=" * 80 + "\n")
            outfile.write(f"COMPILED EMAILS - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            outfile.write("=" * 80 + "\n\n")
            
            # Process each email file
            for email_file in email_files:
                try:
                    with open(email_file, 'r', encoding='utf-8') as infile:
                        # Write separator between emails
                        outfile.write("-" * 80 + "\n")
                        outfile.write(f"Email from file: {email_file.name}\n")
                        outfile.write("-" * 80 + "\n\n")
                        
                        # Copy email content
                        outfile.write(infile.read())
                        outfile.write("\n\n")
                except Exception as e:
                    logger.error(f"Error processing file {email_file}: {str(e)}")
                    continue
        
        logger.info(f"Successfully compiled {len(email_files)} emails into {output_file}")
        return output_file
    except Exception as e:
        logger.error(f"Error creating compiled file: {str(e)}")
        return None

def main():
    # Your Outlook credentials
    email = "tolds@ceresinsurance.com"
    password = "annuiTy2024!"
    search_person = "Joe Bentivoglio"
    n_emails = 15
    
    logger.info("Starting email collection process...")
    success = asyncio.run(collect_outlook_emails(email, password, search_person, n_emails))
    
    if success:
        logger.info("Email collection completed successfully")
        # Compile all collected emails
        compiled_file = compile_emails()
        if compiled_file:
            logger.info(f"All emails have been compiled into: {compiled_file}")
    else:
        logger.error("Email collection failed")

if __name__ == "__main__":
    main()
