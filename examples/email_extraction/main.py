from __future__ import annotations as _annotations

import asyncio
import uuid
import sys
import os
import re
import pdb
import logging
import json
import datetime
from typing import Any

from pydantic import BaseModel

# Add detailed trace logging
class DetailedTraceLogger:
    def __init__(self, filename=None):
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        self.log_file = filename or f"agent_trace_log_{timestamp}.jsonl"
        
        # Store original stdout/stderr
        self.original_stdout = sys.stdout
        self.original_stderr = sys.stderr
        
        # Create a root logger
        self.logger = logging.getLogger()
        self.logger.setLevel(logging.INFO)  # Changed from DEBUG to INFO
        
        # Remove any existing handlers
        self.logger.handlers.clear()
        
        # File handler for JSONL trace log with custom formatter
        class JSONFormatter(logging.Formatter):
            def format(self, record):
                # Skip certain noisy messages
                msg = record.getMessage()
                if "Creating trace" in msg or "Using selector" in msg:
                    return None
                
                # Create the base log entry
                log_entry = {
                    "timestamp": datetime.datetime.now().isoformat(),
                    "level": record.levelname,
                    "message": msg
                }
                
                # Add extra fields if they exist
                if hasattr(record, 'details'):
                    log_entry.update(record.details)
                
                return json.dumps(log_entry)
        
        # Create handlers
        self.trace_handler = logging.FileHandler(self.log_file)
        self.trace_handler.setFormatter(JSONFormatter())
        self.trace_handler.setLevel(logging.INFO)  # Changed from DEBUG to INFO
        
        # Console handler with minimal formatting
        self.console_handler = logging.StreamHandler(self.original_stdout)
        self.console_handler.setFormatter(logging.Formatter('%(message)s'))
        self.console_handler.setLevel(logging.INFO)
        
        # Add handlers to logger
        self.logger.addHandler(self.trace_handler)
        self.logger.addHandler(self.console_handler)
        
        # Set up stdout/stderr capture
        self.setup_stdout_capture()
    
    def setup_stdout_capture(self):
        """Capture stdout and stderr to also log to the trace file"""
        class StreamToLogger:
            def __init__(self, logger, level, original_stream):
                self.logger = logger
                self.level = level
                self.original_stream = original_stream
            
            def write(self, buf):
                if not buf or not buf.strip():
                    return
                self.original_stream.write(buf)
                self.original_stream.flush()
                self.logger.log(self.level, buf.rstrip())
            
            def flush(self):
                self.original_stream.flush()
            
            def isatty(self):
                return hasattr(self.original_stream, 'isatty') and self.original_stream.isatty()
            
            def fileno(self):
                return self.original_stream.fileno()
        
        # Redirect stdout and stderr while preserving originals
        sys.stdout = StreamToLogger(self.logger, logging.INFO, self.original_stdout)
        sys.stderr = StreamToLogger(self.logger, logging.ERROR, self.original_stderr)
    
    def restore_streams(self):
        """Restore original stdout/stderr"""
        sys.stdout = self.original_stdout
        sys.stderr = self.original_stderr

# Initialize basic logging
logger = logging.getLogger()
logger.setLevel(logging.INFO)

# Create a formatter that includes timestamp and log level
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

# Create file handler for all logs
file_handler = logging.FileHandler('email_extraction.log')
file_handler.setFormatter(formatter)
file_handler.setLevel(logging.DEBUG)
logger.addHandler(file_handler)

# Create console handler that only shows important messages
class ConsoleFilter(logging.Filter):
    def filter(self, record):
        # Only show messages that are INFO level or higher
        # and contain certain keywords
        keywords = ['Starting', 'Successfully', 'Error', 'Failed', 'Summary']
        return (record.levelno >= logging.INFO and 
                any(keyword in record.getMessage() for keyword in keywords))

console_handler = logging.StreamHandler(sys.stdout)
console_handler.setFormatter(logging.Formatter('%(message)s'))
console_handler.addFilter(ConsoleFilter())
logger.addHandler(console_handler)

# Verify logging is working
logger.info("=== Starting Email Extraction Process ===")
logger.info("Logging system initialized successfully")
logger.info("Detailed logs will be saved to email_extraction.log")

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

def log_success(operation: str, details: dict) -> None:
    """
    Log a successful operation with its details.
    
    Args:
        operation: The name of the operation that succeeded
        details: A dictionary containing details about the operation
    """
    try:
        logger.info(f"Operation '{operation}' completed successfully")
        logger.debug(f"Success details: {json.dumps(details, indent=2)}")
    except Exception as e:
        logger.error(f"Failed to log success for operation '{operation}': {e}")

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
    email: str = "tolds@ceresinsurance.com"
    password: str = "annuiTy2024!"
    logged_in: bool = False  # Track login state
    search_complete: bool = False  # Track search state
    openai_client: Any = None  # Add OpenAI client to context

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
        
        # Navigate to the specific OAuth2 endpoint
        #await page.goto("https://login.microsoftonline.com/common/oauth2/authorize?client_id=00000002-0000-0ff1-ce00-000000000000&redirect_uri=https%3a%2f%2foutlook.office.com%2fowa%2f&resource=00000002-0000-0ff1-ce00-000000000000&response_mode=form_post&response_type=code+id_token&scope=openid")
        await page.goto("https://outlook.office.com")
        # Wait for and fill email input using role-based selector
        email_input = page.get_by_role("textbox", name="Enter your email, phone, or")
        await email_input.wait_for(state="visible", timeout=5000)
        await email_input.click()
        await email_input.fill(email)
        await page.get_by_role("button", name="Next").click()
        
        # Wait for and fill password input
        password_input = page.locator("#i0118")
        await password_input.wait_for(state="visible", timeout=5000)
        await password_input.fill(password)
        await page.get_by_role("button", name="Sign in").click()
        
        # Wait for MFA or successful login
        await page.wait_for_timeout(20000)  # 20 seconds for MFA
        
        # Navigate to inbox
        await page.goto("https://outlook.office.com/mail/")
        await page.wait_for_timeout(5000)
        
        # Take a screenshot after login
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
        #"input[aria-label='Search']",
        #"input[aria-label='Search mail']",
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
            #pdb.set_trace()
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
# Add the search helper function before the tool definition
async def _search_in_outlook_impl(page, query, logger):
    logger.info(f"Starting search with query: {query}")
    # Format the query if it doesn't have "From:" prefix and it's a person search
    if "Lynn Gadue" in query and not query.startswith("From:"):
        query = f"From:Lynn Gadue"
    elif "Gadue" in query and not query.startswith("From:"):
        query = f"From:Lynn Gadue"  # Force the correct search format
    # Take a screenshot before search
    await page.screenshot(path="screenshots/before_search.png")
    # Try multiple possible selectors with improved waiting
    selectors = [
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
            logger.info(f"Found search box with selector: {selector}")
            break
        except Exception as e:
            logger.debug(f"Selector {selector} failed: {str(e)}")
            continue
    if not search_box:
        logger.error("Could not find search box with any selector")
        return "Could not find search box with any selector"
    try:
        # Clear existing text
        await search_box.click()
        await page.wait_for_timeout(500)
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
        await page.wait_for_timeout(8000)
        # Try multiple selectors for search results
        result_selectors = [
            "div.mailListItem",
            "div[data-convid]",
            ".ZtMcN",
            "[role='row']",
            "div[class*='mailListItem']",
            "div[class*='messageListItem']"
        ]
        for result_selector in result_selectors:
            try:
                logger.debug(f"Trying to find results with selector: {result_selector}")
                await page.wait_for_selector(result_selector, timeout=5000)
                logger.info(f"Found search results with selector: {result_selector}")
                # Count results
                result_count = await page.locator(result_selector).count()
                logger.info(f"Found {result_count} search results")
                # Take screenshot of results
                await page.screenshot(path="screenshots/search_results.png")
                return f"Search completed for '{query}'. Found {result_count} results."
            except Exception as e:
                logger.debug(f"Result selector {result_selector} failed: {str(e)}")
                continue
        logger.warning("Could not find any search results with the tried selectors")
        return "Search completed but could not verify results. Check screenshots for details."
    except Exception as e:
        logger.error(f"Error during search: {str(e)}")
        await page.screenshot(path="screenshots/search_error.png")
        return f"Error during search: {str(e)}"

# Update the tool to use the helper
@function_tool(
    name_override="search_in_outlook",
    description_override="Search for emails using the Outlook search box with improved reliability and filter maintenance."
)
async def search_in_outlook(ctx: RunContextWrapper[ExtractEmailContext], query: str) -> str:
    global computer_instance
    page = computer_instance.page
    return await _search_in_outlook_impl(page, query, logger)




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
    description_override="Diagnose the Outlook page structure and elements to help debug extraction issues."
)
async def diagnose_outlook_page(ctx: RunContextWrapper[ExtractEmailContext]) -> str:
    global computer_instance
    page = computer_instance.page
    
    # Take a screenshot of the current page
    await page.screenshot(path="screenshots/diagnose_search.png")
    
    # Gather basic page info
    page_info = await page.evaluate("""() => {
        const data = {
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
        
        // Get search box value
        const searchBox = document.querySelector('input[aria-label*="Search"]') || 
                         document.querySelector('input[type="search"]') ||
                         document.querySelector('input[placeholder*="Search"]');
        if (searchBox) {
            data.currentSearchValue = searchBox.value;
        }
        
        return data;
    }""")
    
    # Save high-level page info as JSON
    with open("search_diagnosis.json", "w") as f:
        import json
        json.dump(page_info, f, indent=2)
    
    # Capture innerHTML and innerText of the first email element
    email_html = await page.evaluate("""() => {
        const el = document.querySelector('div[data-convid]');
        return el ? el.innerHTML : "No email element found";
    }""")
    with open("email_sample.html", "w") as f:
        f.write(email_html)
    
    email_text = await page.evaluate("""() => {
        const el = document.querySelector('div[data-convid]');
        return el ? el.innerText : "No email element found";
    }""")
    with open("email_sample.txt", "w") as f:
        f.write(email_text)
    
    # Try direct element clicks to see if they work
    direct_click_results = []
    
    # Try clicking on Lynn Gadue text
    try:
        await page.click("text=Lynn Gadue", timeout=3000)
        await page.wait_for_timeout(1000)
        await page.screenshot(path="screenshots/after_lynn_click.png")
        direct_click_results.append("Successfully clicked 'Lynn Gadue' text")
    except Exception as e:
        direct_click_results.append(f"Failed to click 'Lynn Gadue' text: {str(e)}")
    
    return f"""Diagnostic Results:
1. Search Box Status:
   - Total input fields: {page_info['elements']['inputs']}
   - Current search value: {page_info.get('currentSearchValue', 'None')}

2. Email List Status:
   - Total rows: {page_info['elements']['rows']}
   - Lynn Gadue mentions: {page_info['textContent']['lynnGadue']}

3. Click Test Results:
   - {direct_click_results[0]}

4. Email Sample Files:
   - Saved HTML: email_sample.html
   - Saved Text: email_sample.txt

5. Screenshot: screenshots/diagnose_search.png
"""



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
def clean_email_content(content: str) -> str:
    """Clean up email content by removing UI elements and normalizing formatting."""
    # Remove UI elements
    ui_elements = [
        "Reply", "Reply all", "Forward",
        "Summary by Copilot",
        "To:​", "Cc:​",
        "​​"  # Remove zero-width spaces
    ]
    
    cleaned = content
    for element in ui_elements:
        cleaned = cleaned.replace(element, "")
    
    # Remove excessive blank lines (more than 2 consecutive)
    cleaned = re.sub(r'\n{3,}', '\n\n', cleaned)
    
    # Remove lines that are just whitespace
    cleaned = '\n'.join(line for line in cleaned.split('\n') if line.strip())
    
    return cleaned.strip()

@function_tool(
    name_override="collect_emails",
    description_override="Collect emails from Lynn Gadue and save their content to outlook_emails.txt"
)
async def collect_emails(ctx: RunContextWrapper[ExtractEmailContext]) -> str:
    global computer_instance
    page = computer_instance.page
    
    # Get number of emails to collect from context
    n_emails = ctx.context.n_emails
    search_person = ctx.context.search_person
    logger.info(f"Starting email collection for {n_emails} emails from {search_person}")
    
    # Take a screenshot before starting
    await page.screenshot(path="screenshots/before_collect.png")
    
    # Track processed email IDs and subjects to avoid duplicates
    processed_ids = set()
    processed_subjects = set()
    emails_collected = 0
    
    try:
        logger.info("Waiting for email list to load...")
        # Wait for email list to load with multiple selectors
        selectors = [
            "div[data-convid]",
            "div[role='row']",
            "div[class*='mailListItem']",
            "div[class*='messageListItem']"
        ]
        
        email_elements = None
        for selector in selectors:
            try:
                logger.info(f"Trying to find emails with selector: {selector}")
                await page.wait_for_selector(selector, timeout=5000)
                email_elements = await page.locator(selector).all()
                if email_elements:
                    logger.info(f"Found {len(email_elements)} emails with selector: {selector}")
                    break
            except Exception as e:
                logger.debug(f"Selector {selector} failed: {str(e)}")
                continue
        
        if not email_elements:
            logger.error("Could not find any email elements with any selector")
            await page.screenshot(path="screenshots/no_emails_found.png")
            return "No email elements found. Please check the search results and try again."
        
        # Clear the output file before starting
        with open("outlook_emails.txt", "w", encoding="utf-8") as f:
            f.write("=== Email Collection Started ===\n\n")
        
        while emails_collected < n_emails:
            try:
                logger.info(f"Attempting to collect email {emails_collected + 1} of {n_emails}")
                
                # Find an unprocessed email from the search person
                current_email = None
                for i, email in enumerate(email_elements):
                    try:
                        logger.info(f"Checking email {i + 1} of {len(email_elements)}")
                        
                        # Verify the element is visible
                        if not await email.is_visible():
                            logger.debug(f"Email {i + 1} is not visible, skipping")
                            continue
                        
                        # Get email ID and check for duplicates
                        email_id = await email.evaluate("el => el.getAttribute('data-convid')")
                        if email_id in processed_ids:
                            logger.debug(f"Email {i + 1} already processed, skipping")
                            continue
                        
                        # Get sender information
                        sender_info = await email.evaluate("""
                            el => {
                                const senderSelectors = [
                                    '[role=\"heading\"]',
                                    'span[title]',
                                    '[aria-label*="From"]',
                                    '.ms-Persona-primaryText',
                                    'span[class*="sender"]',
                                    'div[class*="from"]'
                                ];
                                
                                for (const selector of senderSelectors) {
                                    const senderEl = el.querySelector(selector);
                                    if (senderEl && senderEl.innerText.trim()) {
                                        return senderEl.innerText.trim();
                                    }
                                }
                                return '';
                            }
                        """)
                        logger.info(f"Found sender info: {sender_info}")
                        
                        # Verify this is from the search person
                        if not sender_info or search_person.lower() not in sender_info.lower():
                            logger.debug(f"Email {i + 1} not from {search_person}, skipping")
                            continue
                        
                        # Get subject to check for duplicates
                        subject = await email.evaluate(f"""
                            (el) => {{
                                // Try header selectors first
                                const headerSelectors = [
                                    'div[class*="message-subject"]',
                                    'div[class*="subject"]',
                                    'span[class*="subject"]',
                                    'div[role="link"] > span',
                                    '[aria-label*="Subject"]'
                                ];
                                for (const selector of headerSelectors) {{
                                    const subjectEl = el.querySelector(selector);
                                    if (subjectEl) {{
                                        const text = subjectEl.innerText || subjectEl.textContent || subjectEl.getAttribute('title') || subjectEl.getAttribute('aria-label');
                                        if (text && text.trim()) {{
                                            return text.trim();
                                        }}
                                    }}
                                }}
                                // Fallback: parse innerText lines
                                const lines = (el.innerText || '').split('\\n').map(l => l.trim()).filter(Boolean);
                                // Find sender and date line indices
                                let senderIdx = -1, dateIdx = -1;
                                for (let i = 0; i < lines.length; i++) {{
                                    if (lines[i].toLowerCase().includes('{search_person.lower()}')) senderIdx = i;
                                    // Heuristic: date line often contains a day and a number, e.g., "Fri 4/11"
                                    if (/\\b(?:mon|tue|wed|thu|fri|sat|sun)\\b/i.test(lines[i]) && /\\d/.test(lines[i])) dateIdx = i;
                                }}
                                // If both found and at least one line between, return the line(s) between
                                if (senderIdx !== -1 && dateIdx !== -1 && dateIdx > senderIdx + 1) {{
                                    const subjectLines = lines.slice(senderIdx + 1, dateIdx);
                                    return subjectLines.join(' ').trim();
                                }}
                                // Fallbacks: as before
                                for (let i = 0; i < lines.length; i++) {{
                                    if (/^(Re:|Fwd:)/i.test(lines[i])) {{
                                        return lines[i];
                                    }}
                                }}
                                if (lines.length >= 5) {{
                                    return lines[4];
                                }}
                                return '';
                            }}
                        """)
                        logger.info(f"Found subject: {subject}")
                        
                        # Diagnostic: If subject is empty, save innerHTML and innerText for first 3 emails
                        if (not subject) and (i < 3):
                            try:
                                html = await email.evaluate("el => el.innerHTML")
                                text = await email.evaluate("el => el.innerText")
                                with open(f"email_element_{i+1}.html", "w", encoding="utf-8") as f_html:
                                    f_html.write(html)
                                with open(f"email_element_{i+1}.txt", "w", encoding="utf-8") as f_txt:
                                    f_txt.write(text)
                                logger.info(f"Saved diagnostic files for email {i+1}: email_element_{i+1}.html and email_element_{i+1}.txt")
                            except Exception as diag_e:
                                logger.error(f"Failed to save diagnostic files for email {i+1}: {diag_e}")
                        
                        if subject and subject not in processed_subjects:
                            current_email = email
                            processed_ids.add(email_id)
                            processed_subjects.add(subject)
                            logger.info(f"Found unprocessed email: {subject}")
                            break
                            
                    except Exception as e:
                        logger.error(f"Error checking email {i + 1}: {str(e)}")
                        continue
                
                if not current_email:
                    logger.error("No more unprocessed emails found from the search person")
                    await page.screenshot(path="screenshots/no_more_emails.png")
                    return "No more unprocessed emails found from the search person. Please check the search results and try again."
                
                # Process the found email
                try:
                    logger.info("Clicking on email to view content...")
                    # Click and wait for content
                    await current_email.click()
                    await page.wait_for_timeout(3000)
                    
                    # Try multiple selectors for the email content
                    content_selectors = [
                        "div[role='document']",
                        "div[class*='messageBody']",
                        "div[class*='MessageBody']",
                        "div[class*='mail-message-content']",
                        "div[class*='message-content']"
                    ]
                    
                    content = None
                    for selector in content_selectors:
                        try:
                            logger.info(f"Trying content selector: {selector}")
                            content_element = page.locator(selector)
                            if await content_element.count() > 0:
                                content = await content_element.inner_text()
                                if content.strip():
                                    logger.info("Found content using selector")
                                    break
                        except Exception as e:
                            logger.debug(f"Selector {selector} failed: {str(e)}")
                            continue
                    
                    if not content:
                        logger.warning("Could not find email content")
                        content = "No content available"
                    
                    # Clean and write the content
                    cleaned_content = clean_email_content(content)
                    with open("outlook_emails.txt", "a", encoding="utf-8") as f:
                        f.write(f"=== EMAIL {emails_collected + 1} ===\n")
                        f.write(f"Subject: {subject}\n")
                        f.write(f"From: {sender_info}\n")
                        f.write(cleaned_content)
                        f.write("\n\n")
                    
                    emails_collected += 1
                    logger.info(f"Successfully collected email {emails_collected}: {subject}")
                    
                    # Navigate back and wait for list to reload
                    logger.info("Navigating back to email list...")
                    await page.go_back()
                    await page.wait_for_timeout(2000)

                    # After navigating back, check if the search filter is still present
                    try:
                        search_box = page.get_by_placeholder("Search")
                        current_search_value = ""
                        if await search_box.count() > 0:
                            try:
                                current_search_value = await search_box.input_value()
                            except Exception as e:
                                logger.debug(f"Could not get search box value: {str(e)}")
                        search_should_be = f"from:{search_person}"
                        if search_should_be.lower() not in current_search_value.lower():
                            logger.info(f"Search filter lost (current value: '{current_search_value}'). Re-applying search filter...")
                            await _search_in_outlook_impl(page, search_should_be, logger)
                            await page.wait_for_timeout(5000)
                            # Refresh email_elements after re-search
                            email_elements = await page.locator("div[data-convid]").all()
                        else:
                            logger.info("Search filter still present after navigating back.")
                    except Exception as e:
                        logger.error(f"Error checking or re-applying search filter after navigating back: {str(e)}")
                    
                except Exception as e:
                    logger.error(f"Error processing email: {str(e)}")
                    try:
                        await page.go_back()
                        await page.wait_for_timeout(2000)
                    except:
                        pass
                    continue
                
            except Exception as e:
                logger.error(f"Error in collection loop: {str(e)}")
                # Try to recover by going back to the inbox
                try:
                    logger.info("Attempting to recover by going back to inbox...")
                    await page.goto("https://outlook.office.com/mail/")
                    await page.wait_for_timeout(5000)
                except:
                    pass
                continue
        
        if emails_collected == 0:
            logger.error("No emails were collected. Please check the search results and try again.")
            return "No emails were collected. Please check the search results and try again."
        
        summary = f"""
=== Email Collection Summary ===
Total emails collected: {emails_collected}
Search person: {search_person}
Processed subjects: {', '.join(processed_subjects)}
Check email_extraction.log for detailed process information
Check outlook_emails.txt for the collected email content
"""
        logger.info(summary)
        return f"Successfully collected {emails_collected} emails from {search_person}. Check outlook_emails.txt for results."
        
    except Exception as e:
        logger.error(f"Failed to process emails: {str(e)}")
        await page.screenshot(path="screenshots/collection_error.png")
        return f"Error collecting emails: {str(e)}"

@function_tool(
    name_override="diagnose_search_and_emails",
    description_override="Diagnose the search box and email list state to help debug extraction issues."
)
async def diagnose_search_and_emails(ctx: RunContextWrapper[ExtractEmailContext]) -> str:
    global computer_instance
    page = computer_instance.page

    # Take a screenshot of the current page
    await page.screenshot(path="screenshots/diagnose_search.png")

    # Gather basic page info
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

        // Get search box value
        const searchBox = document.querySelector('input[aria-label*="Search"]') || 
                          document.querySelector('input[type="search"]') ||
                          document.querySelector('input[placeholder*="Search"]');
        if (searchBox) {
            data.currentSearchValue = searchBox.value;
        }

        return data;
    }""")

    # Save high-level page info as JSON
    with open("search_diagnosis.json", "w") as f:
        import json
        json.dump(page_info, f, indent=2)

    # Capture innerHTML and innerText of the first email element
    email_html = await page.evaluate("""() => {
        const el = document.querySelector('div[data-convid]');
        return el ? el.innerHTML : "No email element found";
    }""")
    with open("email_sample.html", "w") as f:
        f.write(email_html)

    email_text = await page.evaluate("""() => {
        const el = document.querySelector('div[data-convid]');
        return el ? el.innerText : "No email element found";
    }""")
    with open("email_sample.txt", "w") as f:
        f.write(email_text)

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

4. Email Sample Files:
   - Saved HTML: email_sample.html
   - Saved Text: email_sample.txt

5. Screenshot: screenshots/diagnose_search.png
"""


@function_tool(
    name_override="analyze_emails",
    description_override="Analyze the collected emails using OpenAI to create a detailed summary of key points and open questions."
)
async def analyze_emails(ctx: RunContextWrapper[ExtractEmailContext]) -> str:
    try:
        # Check if we've already analyzed these emails
        if os.path.exists("email_takeaways.txt"):
            with open("email_takeaways.txt", "r", encoding="utf-8") as f:
                existing_analysis = f.read()
                if existing_analysis:
                    logger.info("Email analysis already exists, skipping re-analysis")
                    return "Email analysis already exists. Check email_takeaways.txt for results."
        
        # Read the collected emails
        with open("outlook_emails.txt", "r", encoding="utf-8") as f:
            email_content = f.read()
        
        logger.debug("Reading email content for analysis...")
        logger.debug(f"Email content length: {len(email_content)}")
        
        if not ctx.context.openai_client:
            logger.error("OpenAI client not available in context")
            return "Error: OpenAI client not configured. Please ensure the client is set in the context."
        
        # Prepare the prompt for GPT-4
        prompt = f"""Please analyze these emails and create a detailed summary. Focus on:
1. Key topics discussed
2. Important decisions or conclusions reached
3. Open questions or unresolved issues
4. Any action items or next steps mentioned
5. Timeline of events
6. Key stakeholders involved

Here are the emails to analyze:
{email_content}"""
        
        try:
            # Call OpenAI API
            response = await ctx.context.openai_client.chat.completions.create(
                model="gpt-4",
                messages=[
                    {"role": "system", "content": "You are an expert at analyzing email threads and extracting key information. Provide a clear, structured analysis."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.7,
                max_tokens=2000
            )
            
            # Save the analysis
            analysis = response.choices[0].message.content
            with open("email_takeaways.txt", "w", encoding="utf-8") as f:
                f.write(analysis)
            
            logger.info("Email analysis complete using OpenAI")
            return "Email analysis complete. Results saved to email_takeaways.txt"
        
        except Exception as e:
            logger.error(f"Error during OpenAI API call: {str(e)}")
            return f"Error analyzing emails with OpenAI: {str(e)}"
    
    except Exception as e:
        logger.error(f"Error reading email file: {str(e)}")
        return f"Error reading email file: {str(e)}"

# =============================================================================
# DYNAMIC AGENT INSTRUCTIONS
# =============================================================================

default_context = ExtractEmailContext()
INSTRUCTIONS = f"""{RECOMMENDED_PROMPT_PREFIX}
You are an email extraction agent that will help collect and analyze emails from Lynn Gadue in Outlook.
Follow these steps in order:

1. First, log into Outlook using the login_to_outlook function
2. Then, find the search box of the Outlook UI using the find_search_box function
3. Then, search for emails from Lynn Gadue using the search_in_outlook function
4. Next, collect the content of the emails using the collect_emails function. The number of emails to collect is specified in the context (n_emails).
   - Note: You may need to scroll down the sorted list to ensure enough emails are visible. but don't scroll more than 5 times for purposse of this testing
   - It may be necessary to click on each email to see and extract its contents. you can determine this by looking at the page to see if there are still emails below
5. Note that it might be necessary to re-sort or re-search the inbox once you've clicked on a particular email and want to collect content from the next email. We only want to collect content of emails from the specific sender designated for this run of the agent.
6. Finally, analyze the collected emails using the analyze_emails function to generate a summary of key points and open questions

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
    model_settings=ModelSettings(
        tool_choice="required",
        temperature=0.7,
        max_tokens=2000
    ),
)
# =============================================================================
# MAIN RUNNER
# =============================================================================

async def main():
    global computer_instance
    
    # Initialize the detailed logger first
    logger = DetailedTraceLogger()
    
    try:
        # Then enable verbose stdout logging
        #enable_verbose_stdout_logging()
        
        # Initialize OpenAI client
        from openai import AsyncOpenAI
        openai_client = AsyncOpenAI()
        
        async with OutlookPlaywrightComputer() as computer:
            computer_instance = computer
            
            # Provide an initial input to start the process.
            input_items: list[TResponseInputItem] = [{"role": "user", "content": "start email extraction"}]
            context = ExtractEmailContext(n_emails=5, openai_client=openai_client)
            
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
    except Exception as e:
        print(f"Error during execution: {str(e)}")
        import traceback
        traceback.print_exc()
    finally:
        # Restore original streams before exiting
        logger.restore_streams()




#async def debug_main():
#    global computer_instance
#    #enable_verbose_stdout_logging()
    
#    async with OutlookPlaywrightComputer() as computer:
#        computer_instance = computer
        
#        # Create context instance.
#        context = ExtractEmailContext(n_emails=5)
#        ctx = RunContextWrapper(context)
#        os.makedirs("screenshots", exist_ok=True)
        
#        # Step 1: Login
#        login_result = await login_to_outlook(ctx, email=context.email, password=context.password)
#        print("Login result:", login_result)
#        pdb.set_trace()  # Pause here to inspect after login
        
#        # Step 2: Search for emails
#        search_result = await search_in_outlook(ctx, query="from:Lynn Gadue")
#        print("Search result:", search_result)
#        pdb.set_trace()  # Pause here to inspect after search
        
#        # Step 3: Scroll to load more emails
#        scroll_result = await enhanced_outlook_scroll(ctx, n_emails=context.n_emails)
#        print("Scroll result:", scroll_result)
#        pdb.set_trace()  # Pause here to inspect after scrolling
        
#        # Step 4: Collect emails
#        collect_result = await collect_emails(ctx)
#        print("Collect emails result:", collect_result)
#        pdb.set_trace()  # Pause here to inspect after collection
        
#        # Additional steps for diagnosing and analyzing can also be added.


if __name__ == "__main__":
    asyncio.run(main())
    #asyncio.run(debug_main())


#if __name__ == "__main__":
#    asyncio.run(test_find_search_box())

