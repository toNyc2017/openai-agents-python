from __future__ import annotations as _annotations

import asyncio
import uuid
import sys
import os
import re

from pydantic import BaseModel

# Add the project root so that the examples folder is in the module search path.
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../../")))
from examples.tools.computer_use import LocalPlaywrightComputer

from agents import (
    Agent,
    MessageOutputItem,
    RunContextWrapper,
    Runner,
    TResponseInputItem,
    function_tool,
    trace,
    ModelSettings,
    enable_verbose_stdout_logging,
)
from agents.extensions.handoff_prompt import RECOMMENDED_PROMPT_PREFIX

# =============================================================================
# SUBCLASS LocalPlaywrightComputer TO USE OUTLOOK AS THE STARTING URL
# =============================================================================

class OutlookPlaywrightComputer(LocalPlaywrightComputer):
    async def _get_browser_and_page(self) -> tuple:
        width, height = self.dimensions
        launch_args = [f"--window-size={width},{height}"]
        browser = await self.playwright.chromium.launch(headless=False, args=launch_args)
        page = await browser.new_page()
        await page.set_viewport_size({"width": width, "height": height})
        # Navigate directly to Outlook.
        await page.goto("https://outlook.office.com")
        return browser, page

# Global variable for our computer instance.
computer_instance = None

async def with_timeout(coroutine, timeout_seconds=30):
    """Run a coroutine with a timeout."""
    try:
        return await asyncio.wait_for(coroutine, timeout=timeout_seconds)
    except asyncio.TimeoutError:
        return "Operation timed out after {timeout_seconds} seconds."

async def find_search_box(page: Page) -> Locator:
    """
    Attempts to locate the Outlook search box using multiple candidate selectors.
    Returns the first Locator that is found and visible.
    
    Raises:
        Exception: If no candidate selector finds a visible search box.
    """
    # List of candidate selectors to try
    candidate_selectors = [
        "input[aria-label='Search']",         # common guess
        "input[aria-label='Search mail']",      # alternative possibility
        "div[role='search'] input",             # if search input is nested in a search container
        "input[placeholder*='Search']"          # a generic search placeholder
    ]
    
    for selector in candidate_selectors:
        try:
            # Try to wait for the element to be visible (short timeout)
            locator = page.locator(selector)
            # Wait up to 5 seconds for this candidate to appear
            await locator.wait_for(state="visible", timeout=5000)
            # If found, return it.
            print(f"Found search box using selector: {selector}")
            return locator
        except Exception:
            # If not found in this candidate, continue with next.
            print(f"Selector did not match: {selector}")
            continue

    # If none of the selectors matched, raise an error.
    raise Exception("Could not locate the Outlook search box using candidate selectors.")



# =============================================================================
# Test Function: Launch Outlook, and Find the Search Box, then Pause.
# =============================================================================

async def test_find_search_box():
    enable_verbose_stdout_logging()
    async with OutlookPlaywrightComputer() as computer:
        page = computer.page
        # Optionally, wait a bit for the page to fully load and complete login/MFA.
        await asyncio.sleep(10)
        try:
            search_box = await find_search_box(page)
            print("Search box found!")
        except Exception as e:
            print(f"Error: {e}")
        
        # Pause execution to allow inspection
        input("Press Enter to close the browser and finish...")


# =============================================================================
# CONTEXT DEFINITION
# =============================================================================
class ExtractEmailContext(BaseModel):
    text: str = ""
    n_emails: int = 10
    search_person: str = "Lynn Gadue"
    email: str = "tolds@3clife.info"
    password: str = "annuiTy2024!"
    logged_in: bool = False  # Track login state
    search_complete: bool = False  # Track search state
# =============================================================================
# TOOL: LOGIN TO OUTLOOK
# =============================================================================
@function_tool(
    name_override="login_to_outlook",
    description_override="Log in to Outlook Web using the provided credentials and wait for MFA completion."
)
async def login_to_outlook(ctx: RunContextWrapper[ExtractEmailContext], email: str, password: str) -> str:
    global computer_instance
    page = computer_instance.page
    
    try:
        # Take a screenshot before starting
        await page.screenshot(path="screenshots/before_login.png")
        
        # 1. Navigate to Outlook sign-in page
        await page.goto("https://outlook.office.com")
        await page.wait_for_timeout(2000)  # Wait for page to load
        
        # 2. Wait for and fill email input
        try:
            # Try multiple selectors for email input
            email_selectors = [
                "input[type='email']",
                "input[aria-label*='email']",
                "input[placeholder*='email']"
            ]
            
            email_input = None
            for selector in email_selectors:
                try:
                    email_input = page.locator(selector)
                    await email_input.wait_for(state="visible", timeout=5000)
                    break
                except:
                    continue
            
            if not email_input:
                raise Exception("Could not find email input field")
            
            # Clear any existing text
            await email_input.click()
            await page.keyboard.down("Control")
            await page.keyboard.press("A")
            await page.keyboard.up("Control")
            await page.keyboard.press("Backspace")
            
            # Enter email
            await email_input.fill(email)
            await page.wait_for_timeout(1000)
            
            # Click next/submit
            submit_button = page.locator("input[type='submit']")
            if await submit_button.count() > 0:
                await submit_button.click()
            else:
                await page.keyboard.press("Enter")
            
            await page.wait_for_timeout(2000)
            
        except Exception as e:
            await page.screenshot(path="screenshots/email_input_error.png")
            return f"Error entering email: {str(e)}"
        
        # 3. Wait for and fill password input
        try:
            # Try multiple selectors for password input
            password_selectors = [
                "input[type='password']",
                "input[aria-label*='password']",
                "input[placeholder*='password']"
            ]
            
            password_input = None
            for selector in password_selectors:
                try:
                    password_input = page.locator(selector)
                    await password_input.wait_for(state="visible", timeout=5000)
                    break
                except:
                    continue
            
            if not password_input:
                raise Exception("Could not find password input field")
            
            # Enter password
            await password_input.fill(password)
            await page.wait_for_timeout(1000)
            
            # Click sign in
            submit_button = page.locator("input[type='submit']")
            if await submit_button.count() > 0:
                await submit_button.click()
            else:
                await page.keyboard.press("Enter")
            
            # Wait for MFA or successful login
            await page.wait_for_timeout(20000)  # 20 seconds for MFA
            
        except Exception as e:
            await page.screenshot(path="screenshots/password_input_error.png")
            return f"Error entering password: {str(e)}"
        
        # 4. Navigate to inbox
        await page.goto("https://outlook.office.com/mail/inbox")
        await page.wait_for_timeout(5000)
        
        # 5. Take a screenshot after login
        await page.screenshot(path="screenshots/after_login.png")
        
        return "Logged in successfully. Check screenshots for details."
        
    except Exception as e:
        await page.screenshot(path="screenshots/login_error.png")
        return f"Login failed: {str(e)}"



    
# =============================================================================
# TOOL: FIND SEARCH BOX
# =============================================================================
from playwright.async_api import Page, Locator
import asyncio


@function_tool(
    name_override="find_search_box",
    description_override="Try to locate the search box in Outlook with various potential selectors."
)
async def find_search_box(ctx: RunContextWrapper[ExtractEmailContext]) -> str:
    global computer_instance
    page = computer_instance.page
    
    await page.screenshot(path="screenshots/before_search.png")
    
    # List of possible selectors to try
    selectors = [
        "input[aria-label='Search']",
        "input[aria-label='Search mail']",
        "input[placeholder*='Search']",
        "div[role='search'] input",
        "#search-box",
        "input[type='search']",
        "input.searchInput",
        ".ms-SearchBox-field"
    ]
    
    for selector in selectors:
        try:
            await page.wait_for_selector(selector, timeout=5000)
            await page.screenshot(path=f"screenshots/found_selector_{selector.replace('[', '_').replace(']', '_').replace('*', '_')}.png")
            return f"Found search box with selector: {selector}"
        except:
            continue
    
    # If we get here, none of the selectors worked
    await page.screenshot(path="screenshots/no_search_box.png")
    
    # Try to get all input elements to debug
    inputs = await page.query_selector_all("input")
    input_count = len(inputs)
    
    return f"Could not find search box with any of the attempted selectors. Found {input_count} input elements on the page. Check screenshots."

# =============================================================================
# TOOL: SEARCH IN OUTLOOK
# =============================================================================
@function_tool(
    name_override="search_in_outlook",
    description_override="Search for emails using the Outlook search box with improved reliability and filter maintenance."
)
async def search_in_outlook(ctx: RunContextWrapper[ExtractEmailContext], query: str) -> str:
    global computer_instance
    page = computer_instance.page
    
    # Format the query if it doesn't have "From:" prefix and it's a person search
    if "Lynn Gadue" in query and not query.startswith("From:"):
        query = f"From:Lynn Gadue"
    elif "Gadue" in query and not query.startswith("From:"):
        query = f"From:Lynn Gadue"  # Force the correct search format
    
    # Take a screenshot before search
    await page.screenshot(path="screenshots/before_search.png")
    
    # Try multiple possible selectors with improved waiting
    selectors = [
        "input[aria-label='Search']",
        "input[aria-label='Search mail']",
        "input[placeholder*='Search']",
        "div[role='search'] input",
        "#search-box",
        "input[type='search']",
        "input.searchInput",
        ".ms-SearchBox-field"
    ]
    
    search_box = None
    for selector in selectors:
        try:
            # Wait for the element to be visible
            search_box = page.locator(selector)
            await search_box.wait_for(state="visible", timeout=5000)
            
            # Try to click and clear the search box
            await search_box.click()
            await page.wait_for_timeout(500)
            
            # Clear existing text using keyboard shortcuts
            await page.keyboard.down("Control")
            await page.keyboard.press("A")
            await page.keyboard.up("Control")
            await page.keyboard.press("Backspace")
            await page.wait_for_timeout(500)
            
            # Enter the search query
            await search_box.fill(query)
            await page.wait_for_timeout(1000)
            
            # Submit the search
            await page.keyboard.press("Enter")
            
            # Wait for results with increased timeout
            await page.wait_for_timeout(8000)
            
            # Verify search results
            try:
                # Wait for search results to load with increased timeout
                await page.wait_for_selector("div.mailListItem", timeout=15000)
                
                # Verify we're seeing emails from Lynn Gadue
                sender_count = await page.locator("text=Lynn Gadue").count()
                if sender_count > 0:
                    await page.screenshot(path="screenshots/search_results.png")
                    
                    # Try to scroll to load more results
                    await enhanced_outlook_scroll(ctx, 50)  # Scroll to load more results
                    
                    # Count again after scrolling
                    sender_count = await page.locator("text=Lynn Gadue").count()
                    return f"Search completed for '{query}'. Found {sender_count} emails from Lynn Gadue."
                else:
                    # If no Lynn Gadue emails found, try alternative result selectors
                    result_selectors = [
                        "div.mailListItem",
                        "div[data-convid]",
                        ".ZtMcN",
                        "[role='row']"
                    ]
                    
                    for result_selector in result_selectors:
                        try:
                            await page.wait_for_selector(result_selector, timeout=5000)
                            # Try searching again with a more specific query
                            await search_box.fill("From:Lynn Gadue")
                            await page.keyboard.press("Enter")
                            await page.wait_for_timeout(8000)
                            
                            # Try to scroll to load more results
                            await enhanced_outlook_scroll(ctx, 50)
                            
                            sender_count = await page.locator("text=Lynn Gadue").count()
                            if sender_count > 0:
                                return f"Search completed for '{query}'. Found {sender_count} emails from Lynn Gadue."
                        except:
                            continue
                    
                    return f"Search completed but no emails from Lynn Gadue found. Try a different search query."
                
            except Exception as e:
                await page.screenshot(path="screenshots/search_verification_error.png")
                return f"Error verifying search results: {str(e)}"
                
        except Exception as e:
            print(f"Selector {selector} failed: {str(e)}")
            continue
    
    # If we get here, none of the selectors worked
    await page.screenshot(path="screenshots/search_failed.png")
    
    # Try to diagnose the page state
    try:
        page_info = await page.evaluate("""() => {
            return {
                url: window.location.href,
                inputs: document.querySelectorAll('input').length,
                searchInputs: document.querySelectorAll('input[type="search"]').length,
                ariaSearch: document.querySelectorAll('input[aria-label*="Search"]').length,
                lynnEmails: Array.from(document.querySelectorAll('*')).filter(el => 
                    el.innerText && el.innerText.includes('Lynn Gadue')).length
            };
        }""")
        
        with open("search_failure_diagnosis.json", "w") as f:
            import json
            json.dump(page_info, f, indent=2)
    except:
        pass
    
    return f"Could not perform search with any of the selectors. Check screenshots and search_failure_diagnosis.json for details."






# =============================================================================
# TOOL: SAVE PAGE HTML
# =============================================================================
@function_tool(
    name_override="save_page_html",
    description_override="Save the current page HTML for debugging."
)
async def save_page_html(ctx: RunContextWrapper[ExtractEmailContext]) -> str:
    global computer_instance
    page = computer_instance.page
    
    html = await page.content()
    with open("page_debug.html", "w", encoding="utf-8") as f:
        f.write(html)
    
    await page.screenshot(path="screenshots/debug_screenshot.png")
    
    return "Saved current page HTML to page_debug.html and screenshot to debug_screenshot.png"




# =============================================================================
# TOOL: DIAGNOSE PAGE
# =============================================================================
@function_tool(
    name_override="diagnose_outlook_page",
    description_override="Analyze the Outlook page structure and elements to help debug extraction issues."
)
async def diagnose_outlook_page(ctx: RunContextWrapper[ExtractEmailContext]) -> str:
    global computer_instance
    page = computer_instance.page
    
    await page.screenshot(path="screenshots/diagnosis.png")
    
    # Get details about the page elements
    page_info = await page.evaluate("""() => {
        const data = {
            title: document.title,
            url: window.location.href,
            elements: {
                rows: document.querySelectorAll('[role="row"]').length,
                links: document.querySelectorAll('a').length,
                buttons: document.querySelectorAll('button').length,
                inputs: document.querySelectorAll('input').length
            },
            textContent: {
                lynnGadue: Array.from(document.querySelectorAll('*')).filter(el => 
                    el.innerText && el.innerText.includes('Lynn Gadue')).length
            },
            aria: {
                email: document.querySelectorAll('[aria-label*="email"]').length,
                inbox: document.querySelectorAll('[aria-label*="inbox"]').length,
                message: document.querySelectorAll('[aria-label*="message"]').length
            }
        };
        
        // Get some key class names
        data.classes = {};
        ['ZtMcN', 'hcptT', 'customScrollBar', 'ms-List-cell', 'messageBody'].forEach(cls => {
            data.classes[cls] = document.querySelectorAll('.' + cls).length;
        });
        
        return data;
    }""")
    
    # Save diagnostic info
    with open("outlook_diagnosis.json", "w") as f:
        import json
        json.dump(page_info, f, indent=2)
    
    # Try some direct element clicks to see if they work
    direct_click_results = []
    
    # Try clicking on Lynn Gadue text
    try:
        await page.click("text=Lynn Gadue", {timeout: 3000})
        await page.wait_for_timeout(1000)
        await page.screenshot(path="screenshots/after_lynn_click.png")
        direct_click_results.push("Successfully clicked 'Lynn Gadue' text")
    except Exception as e:
        direct_click_results.push(f"Failed to click 'Lynn Gadue' text: {str(e)}")
    
    return f"Diagnosis complete. Found {page_info['textContent']['lynnGadue']} elements with 'Lynn Gadue' text. Details saved to outlook_diagnosis.json."



# =============================================================================
# TOOL: ENHANCED OUTLOOK SCROLL
# =============================================================================


@function_tool(
    name_override="enhanced_outlook_scroll",
    description_override="Scroll through emails in Outlook with improved reliability."
)
async def enhanced_outlook_scroll(ctx: RunContextWrapper[ExtractEmailContext], n_emails: int) -> str:
    global computer_instance
    page = computer_instance.page
    
    await page.screenshot(path="screenshots/before_scroll.png")
    
    # Try multiple scroll container selectors
    scroll_container_selectors = [
        "div.customScrollBar",
        "[role='region']",
        ".ms-ScrollablePane--contentContainer",
        ".ScrollRegion",
        "div[class*='scroll']",
        "div[class*='list']",
        "div[class*='mailList']"
    ]
    
    scroll_container = None
    for selector in scroll_container_selectors:
        try:
            container = page.locator(selector)
            if await container.count() > 0:
                scroll_container = container
                break
        except:
            continue
    
    if not scroll_container:
        return "Could not find scrollable container. Check screenshots for diagnosis."
    
    # Get initial email count
    initial_count = await page.locator("div[data-convid]").count()
    
    # Try multiple scrolling methods
    scroll_methods = [
        # Method 1: Direct JavaScript scrolling
        lambda: page.evaluate("""() => {
            const container = document.querySelector('div[class*="scroll"]') || 
                            document.querySelector('[role="region"]') ||
                            document.querySelector('.ms-ScrollablePane--contentContainer');
            if (!container) return false;
            container.scrollTop += 800;
            return true;
        }"""),
        
        # Method 2: Scroll into view
        lambda: page.evaluate("""() => {
            const emails = document.querySelectorAll('div[data-convid]');
            if (emails.length === 0) return false;
            emails[emails.length - 1].scrollIntoView({ behavior: 'smooth', block: 'center' });
            return true;
        }"""),
        
        # Method 3: Keyboard scrolling
        lambda: page.keyboard.press("PageDown"),
        
        # Method 4: Scroll to top and back
        lambda: page.evaluate("""() => {
            window.scrollTo(0, 0);
            const elements = document.querySelectorAll('div[data-convid]');
            if (elements.length > 0) {
                elements[elements.length - 1].scrollIntoView({ behavior: 'smooth', block: 'center' });
                return true;
            }
            return false;
        }"""),
        
        # Method 5: Infinite scroll simulation
        lambda: page.evaluate("""() => {
            const container = document.querySelector('div[class*="scroll"]') || 
                            document.querySelector('[role="region"]') ||
                            document.querySelector('.ms-ScrollablePane--contentContainer');
            if (!container) return false;
            
            // Scroll to bottom
            container.scrollTop = container.scrollHeight;
            
            // Wait for content to load
            return new Promise(resolve => {
                setTimeout(() => {
                    resolve(true);
                }, 1000);
            });
        }""")
    ]
    
    # Try each scroll method multiple times
    for _ in range(5):  # Try 5 times
        for scroll_method in scroll_methods:
            try:
                await scroll_method()
                await page.wait_for_timeout(2000)  # Increased wait time for content to load
                
                # Check if we got more emails
                new_count = await page.locator("div[data-convid]").count()
                if new_count > initial_count:
                    await page.screenshot(path="screenshots/after_successful_scroll.png")
                    return f"Successfully scrolled. Initial emails: {initial_count}, Current emails: {new_count}"
            except:
                continue
    
    # If we get here, try keyboard scrolling as a last resort
    for _ in range(10):  # Try 10 times
        await page.keyboard.press("PageDown")
        await page.wait_for_timeout(1000)  # Increased wait time
    
    # Get final count
    final_count = await page.locator("div[data-convid]").count()
    await page.screenshot(path="screenshots/after_scroll.png")
    
    return f"Scroll completed. Initial emails: {initial_count}, Final emails: {final_count}"



# =============================================================================
# TOOL: MINIMAL OUTLOOK SCROLL
# =============================================================================

@function_tool(
    name_override="minimal_outlook_scroll",
    description_override="Scroll down emails in Outlook using multiple methods (direct JS, scrollIntoView, keyboard)."
)
async def minimal_outlook_scroll(ctx: RunContextWrapper[ExtractEmailContext], n_emails: int, search_person: str) -> str:
    global computer_instance
    page = computer_instance.page
    selector = "div.customScrollBar.jEpCF"
    try:
        initial_result = await page.evaluate(f"""() => {{
            const el = document.querySelector("{selector}");
            if (!el) return {{ success: false, reason: "Element not found" }};
            const initial = el.scrollTop;
            el.scrollTop += 500;
            return {{
                success: el.scrollTop > initial,
                initial: initial,
                current: el.scrollTop
            }};
        }}""")
        await page.wait_for_timeout(300)
        if (initial_result or {}).get("success", False):
            return f"Scrolled from {initial_result['initial']}px to {initial_result['current']}px."
        else:
            await page.evaluate(f"""() => {{
                const el = document.querySelector("{selector}");
                if (el) el.scrollIntoView({{ behavior: 'smooth', block: 'center' }});
            }}""")
            await page.wait_for_timeout(300)
            await page.keyboard.press("PageDown")
            return "Attempted scrolling using scrollIntoView and PageDown."
    except Exception as e:
        return f"Error during scrolling: {e}"

# =============================================================================
# TOOL: COLLECT EMAILS
# =============================================================================
@function_tool(
    name_override="collect_emails",
    description_override="Iterate over emails from search results with improved reliability and filter maintenance."
)
async def collect_emails(ctx: RunContextWrapper[ExtractEmailContext], n_emails: int) -> str:
    global computer_instance
    page = computer_instance.page
    
    # Take a screenshot before starting
    await page.screenshot(path="screenshots/before_collection.png")
    
    # First verify we're still in the search results
    try:
        # Check if we see Lynn Gadue emails
        sender_count = await page.locator("text=Lynn Gadue").count()
        if sender_count == 0:
            # If no Lynn Gadue emails found, try to reapply the search
            search_box = page.locator("input[aria-label*='Search']")
            await search_box.fill("From:Lynn Gadue")
            await page.keyboard.press("Enter")
            await page.wait_for_timeout(5000)
            sender_count = await page.locator("text=Lynn Gadue").count()
    except:
        # If we can't verify the search, try to reapply it
        await page.goto("https://outlook.office.com/mail/")
        await page.wait_for_timeout(1000)
        search_box = page.locator("input[aria-label*='Search']")
        await search_box.fill("From:Lynn Gadue")
        await page.keyboard.press("Enter")
        await page.wait_for_timeout(5000)
    
    # Create output file
    output_path = "outlook_emails.txt"
    processed_count = 0
    
    with open(output_path, "w", encoding="utf-8") as f:
        f.write("Starting email collection...\n\n")
        
        # Process first email
        print("Processing first email...")
        try:
            # Wait for email list to load
            await page.wait_for_selector("div[data-convid]", timeout=10000)
            
            # Get and click first email
            first_email = page.locator("div[data-convid]").first
            await first_email.click()
            await page.wait_for_timeout(3000)
            
            # Extract first email content
            content1 = await page.evaluate("""() => {
                const mainContent = document.querySelector('[role="main"]');
                if (mainContent) return mainContent.innerText;
                const messageBody = document.querySelector('.messageBody, .emailBody, .messageContent, .emailContent');
                if (messageBody) return messageBody.innerText;
                const readingPane = document.querySelector('.readingPane, .reading-pane');
                if (readingPane) return readingPane.innerText;
                return document.body.innerText;
            }""")
            
            f.write("=== FIRST EMAIL ===\n")
            f.write(content1)
            f.write("\n\n")
            
            # Go back to inbox
            await page.go_back()
            await page.wait_for_timeout(2000)
            
            # Process second email
            print("Processing second email...")
            second_email = page.locator("div[data-convid]").nth(1)
            await second_email.click()
            await page.wait_for_timeout(3000)
            
            # Extract second email content
            content2 = await page.evaluate("""() => {
                const mainContent = document.querySelector('[role="main"]');
                if (mainContent) return mainContent.innerText;
                const messageBody = document.querySelector('.messageBody, .emailBody, .messageContent, .emailContent');
                if (messageBody) return messageBody.innerText;
                const readingPane = document.querySelector('.readingPane, .reading-pane');
                if (readingPane) return readingPane.innerText;
                return document.body.innerText;
            }""")
            
            f.write("=== SECOND EMAIL ===\n")
            f.write(content2)
            f.write("\n\n")
            
            processed_count = 2
            print(f"Successfully processed {processed_count} emails")
            
        except Exception as e:
            print(f"Error processing emails: {str(e)}")
            return f"Error processing emails: {str(e)}"
    
    return f"Processed {processed_count} emails. Results saved to {output_path}"

@function_tool(
    name_override="diagnose_search_and_emails",
    description_override="Diagnose the search box and email list state to help debug extraction issues."
)
async def diagnose_search_and_emails(ctx: RunContextWrapper[ExtractEmailContext]) -> str:
    global computer_instance
    page = computer_instance.page
    
    # Take a screenshot of the current state
    await page.screenshot(path="screenshots/diagnose_search.png")
    
    # Get page information
    page_info = await page.evaluate("""() => {
        const data = {
            url: window.location.href,
            searchBoxes: {
                total: document.querySelectorAll('input').length,
                searchInputs: document.querySelectorAll('input[type="search"]').length,
                ariaSearch: document.querySelectorAll('input[aria-label*="Search"]').length,
                searchPlaceholder: document.querySelectorAll('input[placeholder*="Search"]').length
            },
            emailList: {
                rows: document.querySelectorAll('[role="row"]').length,
                listItems: document.querySelectorAll('.ms-List-cell').length,
                emailItems: document.querySelectorAll('div[data-convid]').length,
                visibleEmails: Array.from(document.querySelectorAll('div[data-convid]')).filter(el => {
                    const rect = el.getBoundingClientRect();
                    return rect.top >= 0 && rect.bottom <= window.innerHeight;
                }).length
            },
            searchResults: {
                visible: document.querySelectorAll('.ZtMcN').length,
                total: document.querySelectorAll('.ZtMcN').length
            }
        };
        
        // Get the current search box value if any
        const searchBox = document.querySelector('input[aria-label*="Search"]') || 
                         document.querySelector('input[type="search"]') ||
                         document.querySelector('input[placeholder*="Search"]');
        if (searchBox) {
            data.currentSearchValue = searchBox.value;
        }
        
        return data;
    }""")
    
    # Save diagnostic info
    with open("search_diagnosis.json", "w") as f:
        import json
        json.dump(page_info, f, indent=2)
    
    return f"""Diagnostic Results:
1. Search Box Status:
   - Total input fields: {page_info['searchBoxes']['total']}
   - Search inputs: {page_info['searchBoxes']['searchInputs']}
   - Aria search fields: {page_info['searchBoxes']['ariaSearch']}
   - Search placeholders: {page_info['searchBoxes']['searchPlaceholder']}
   - Current search value: {page_info.get('currentSearchValue', 'None')}

2. Email List Status:
   - Total rows: {page_info['emailList']['rows']}
   - List items: {page_info['emailList']['listItems']}
   - Email items: {page_info['emailList']['emailItems']}
   - Visible emails: {page_info['emailList']['visibleEmails']}

3. Search Results:
   - Visible results: {page_info['searchResults']['visible']}
   - Total results: {page_info['searchResults']['total']}

Screenshots and detailed JSON saved for further analysis."""

@function_tool(
    name_override="analyze_emails",
    description_override="Analyze the collected emails and create a summary of key points and open questions."
)
async def analyze_emails(ctx: RunContextWrapper[ExtractEmailContext]) -> str:
    try:
        # Read the collected emails
        with open("outlook_emails.txt", "r", encoding="utf-8") as f:
            email_content = f.read()
        
        print("\nDEBUG: Reading email content for analysis...")
        print(f"DEBUG: Email content length: {len(email_content)}")
        print(f"DEBUG: First 200 chars of email content:\n{email_content[:200]}\n")
        
        # Prepare the prompt for GPT-4
        prompt = f"""Please analyze these emails and create a chronological summary of the discussions. Focus on:
1. Key topics discussed
2. Important decisions or conclusions reached
3. Open questions or unresolved issues
4. Any action items or next steps mentioned

Here are the emails to analyze:
{email_content}"""
        
        try:
            # Call OpenAI API
            response = await ctx.client.chat.completions.create(
                model="gpt-4",
                messages=[
                    {"role": "system", "content": "You are an expert at analyzing email threads and extracting key information."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.7,
                max_tokens=2000
            )
            
            # Save the analysis
            analysis = response.choices[0].message.content
            print("\nDEBUG: Saving analysis to email_takeaways.txt...")
            with open("email_takeaways.txt", "w", encoding="utf-8") as f:
                f.write(analysis)
            
            return "Email analysis complete. Results saved to email_takeaways.txt"
        
        except Exception as e:
            print(f"DEBUG: Error during OpenAI API call: {str(e)}")
            return f"Error analyzing emails: {str(e)}"
    
    except Exception as e:
        print(f"DEBUG: Error reading email file: {str(e)}")
        return f"Error reading email file: {str(e)}"

# =============================================================================
# DYNAMIC AGENT INSTRUCTIONS
# =============================================================================

default_context = ExtractEmailContext()
INSTRUCTIONS = f"""{RECOMMENDED_PROMPT_PREFIX}
You are an email extraction agent that will help collect and analyze emails from Lynn Gadue in Outlook.
Follow these steps in order:

1. First, log into Outlook using the login_to_outlook function
2. Then, search for emails from Lynn Gadue using the search_in_outlook function
3. Next, collect the emails using the collect_emails function, specifying how many emails to collect
4. Finally, analyze the collected emails using the analyze_emails function to generate a summary of key points and open questions

After each step, verify that it completed successfully before moving to the next step.
If any step fails, try to diagnose the issue and retry the step.
Make sure to maintain the search filter for Lynn Gadue throughout the process.

Credentials:
Username: {default_context.email}
Password: {default_context.password}
"""
# =============================================================================
# AGENT DEFINITION
# =============================================================================
email_agent = Agent[ExtractEmailContext](
    name="Outlook Email Extraction Agent",
    instructions=INSTRUCTIONS,
    tools=[
        login_to_outlook, 
        find_search_box, 
        search_in_outlook, 
        enhanced_outlook_scroll,
        collect_emails,
        diagnose_outlook_page,
        diagnose_search_and_emails,
        save_page_html,
        analyze_emails
    ],
    model_settings=ModelSettings(tool_choice="required"),
)
# =============================================================================
# MAIN RUNNER
# =============================================================================

async def main():
    global computer_instance
    enable_verbose_stdout_logging()
    
    async with OutlookPlaywrightComputer() as computer:
        computer_instance = computer
        
        # Provide an initial input to start the process.
        input_items: list[TResponseInputItem] = [{"role": "user", "content": "start email extraction"}]
        context = ExtractEmailContext(n_emails=5)  # For testing, extract 5 emails.

         # Ensure "screenshots" directory exists
        os.makedirs("screenshots", exist_ok=True)
        
        with trace("Outlook Email Extraction Agent"):
            result = await Runner.run(email_agent, input_items, context=context, max_turns=30)
            
            for new_item in result.new_items:
                agent_name = new_item.agent.name
                if isinstance(new_item, MessageOutputItem):
                    print(f"{agent_name}: {new_item.content}")
                else:
                    print(f"{agent_name}: Received item of type {new_item.__class__.__name__}")
        
        print("Email extraction task complete.")

if __name__ == "__main__":
    asyncio.run(main())


#if __name__ == "__main__":
#    asyncio.run(test_find_search_box())

