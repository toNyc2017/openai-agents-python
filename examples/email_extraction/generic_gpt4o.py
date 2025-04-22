import logging
import time
import os
import argparse  # Added for command line arguments
import socket  # Added for debugging port check
import json
from datetime import datetime
from playwright.async_api import async_playwright, TimeoutError
from agents import Agent, Runner, function_tool, enable_verbose_stdout_logging
from dotenv import load_dotenv
import urllib.request

# Load API key from .env if available (for OpenAI Agents usage)
load_dotenv()

# Enable verbose logging for the agent
enable_verbose_stdout_logging()

# Configuration constants for extensibility
REMOTE_DEBUGGING_URL = os.environ.get("CHROME_DEBUG_URL", "http://localhost:9222")
MAX_LOAD_MORE_CLICKS = 20       # Safety cap on number of "Load More" clicks
MAX_SCROLL_ITERATIONS = 30     # Safety cap on scroll loops for infinite scroll
SCROLL_PAUSE_SECONDS = 1.0     # Delay between scroll attempts (seconds)
OUTPUT_TEXT_FILE = "page_content.txt"
OUTPUT_SCREENSHOT_FILE = "page_screenshot.png"

def check_chrome_debugging():
    """Check if Chrome is running with remote debugging enabled"""
    max_retries = 5
    retry_delay = 2  # seconds
    
    for attempt in range(max_retries):
        try:
            # Try to connect to the debugging port
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            result = sock.connect_ex(('localhost', 9222))
            sock.close()
            
            if result == 0:
                # Also verify we can get the JSON response
                try:
                    with urllib.request.urlopen('http://localhost:9222/json') as response:
                        if response.status == 200:
                            return True
                except:
                    pass
                
            if attempt < max_retries - 1:
                print(f"Waiting for Chrome debugging server to be ready (attempt {attempt + 1}/{max_retries})...")
                time.sleep(retry_delay)
                
        except Exception as e:
            print(f"Debug check attempt {attempt + 1} failed: {str(e)}")
            if attempt < max_retries - 1:
                time.sleep(retry_delay)
    
    return False

# Set up detailed logging
log_filename = f"generic_gpt4o_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_filename),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)

# Define tools (functions) for agent to use, using the OpenAI Agents SDK
@function_tool
async def click_load_more() -> str:
    """
    Checks for a visible 'Load More' button on the page and clicks it if found.
    Returns a message indicating whether the button was clicked or not found.
    """
    global page
    
    # Take a screenshot before looking for the button
    await page.screenshot(path="before_load_more.png", full_page=True)
    
    # Log all buttons on the page for debugging
    buttons = await page.evaluate("""() => {
        const buttons = Array.from(document.querySelectorAll('button, [role="button"], .load-more, [class*="load-more"]'));
        return buttons.map(b => ({
            text: b.textContent,
            className: b.className,
            id: b.id,
            role: b.getAttribute('role'),
            isVisible: b.offsetParent !== null,
            rect: b.getBoundingClientRect()
        }));
    }""")
    
    logger.info(f"Found {len(buttons)} potential button elements:")
    for btn in buttons:
        logger.info(f"Button: {json.dumps(btn, indent=2)}")
    
    # Try multiple selectors to find the Load More button
    selectors = [
        # Exact class name from the screenshot
        ".DirectoryEntries_LoadMoreEntriesButton",
        "[class*='DirectoryEntries_LoadMoreEntriesButton']",
        # Custom element with class
        "vx-button.DirectoryEntries_LoadMoreEntriesButton",
        # More general selectors as fallback
        "[class*='LoadMoreEntries']",
        "vx-button:has-text('Load More')",
        "vx-button:has-text(/Load.*More/i)",
        # Original selectors as final fallback
        "button:has-text('Load More')",
        "[role='button']:has-text(/Load.*More/i)",
        "[aria-label*='load more' i]"
    ]
    
    for selector in selectors:
        try:
            logger.info(f"Trying selector: {selector}")
            # Increase timeout to give more time to find the element
            element = await page.wait_for_selector(selector, timeout=2000, state="visible")
            if element:
                # Check if element is visible
                is_visible = await element.is_visible()
                logger.info(f"Found element with selector {selector}, visible: {is_visible}")
                
                if is_visible:
                    # Get element properties for logging
                    text = await element.text_content()
                    logger.info(f"Found visible 'Load More' button with text: {text}")
                    
                    # Take screenshot before clicking
                    await page.screenshot(path="before_click.png")
                    
                    # Try to scroll the button into view
                    await element.scroll_into_view_if_needed()
                    await page.wait_for_timeout(1000)  # Longer pause after scrolling
                    
                    try:
                        # First try normal click
                        await element.click()
                    except Exception as click_error:
                        logger.info(f"Normal click failed: {click_error}, trying JavaScript click")
                        # If normal click fails, try JavaScript click
                        await page.evaluate("(element) => element.click()", element)
                    
                    logger.info("Successfully clicked the button")
                    
                    # Take screenshot after clicking
                    await page.wait_for_timeout(2000)  # Longer wait for any updates
                    await page.screenshot(path="after_click.png")
                    
                    return f"Clicked 'Load More' button with text: {text}"
        except Exception as e:
            logger.info(f"Selector {selector} failed: {str(e)}")
    
    # If we get here, we didn't find a clickable Load More button
    logger.info("No clickable 'Load More' button found")
    return "No 'Load More' button found or button not clickable"

@function_tool
async def scroll_page() -> str:
    """
    Scrolls the page down by one viewport height or to the bottom.
    Returns a message indicating the action.
    """
    global page
    try:
        # Get current scroll position
        prev_pos = await page.evaluate("window.pageYOffset")
        
        # Scroll down one viewport height
        await page.evaluate("window.scrollBy(0, window.innerHeight)")
        
        # Wait for any dynamic content to load
        await page.wait_for_timeout(SCROLL_PAUSE_SECONDS * 1000)
        
        # Get new scroll position
        new_pos = await page.evaluate("window.pageYOffset")
        
        # Log scroll progress
        logger.info(f"Scrolled from {prev_pos} to {new_pos}")
        
        # Take a screenshot after scrolling
        await page.screenshot(path=f"scroll_{new_pos}.png")
        
        return f"Scrolled page from position {prev_pos} to {new_pos}"
    except Exception as e:
        logger.error(f"Scrolling error: {e}")
        return "Scrolling failed"

@function_tool
async def get_visible_text() -> str:
    """
    Retrieves all visible text on the page.
    """
    global page
    try:
        # Get all visible text using JavaScript
        text = await page.evaluate("""() => {
            const walker = document.createTreeWalker(
                document.body,
                NodeFilter.SHOW_TEXT,
                {
                    acceptNode: function(node) {
                        if (!node.textContent.trim()) return NodeFilter.FILTER_REJECT;
                        const style = window.getComputedStyle(node.parentElement);
                        return (style.display !== 'none' && 
                                style.visibility !== 'hidden' && 
                                style.opacity !== '0') 
                                ? NodeFilter.FILTER_ACCEPT 
                                : NodeFilter.FILTER_REJECT;
                    }
                }
            );
            const textNodes = [];
            while (walker.nextNode()) textNodes.push(walker.currentNode.textContent.trim());
            return textNodes.join('\n');
        }""")
        
        # Save text to file
        with open(OUTPUT_TEXT_FILE, "w", encoding="utf-8") as f:
            f.write(text)
        
        logger.info(f"Captured {len(text)} characters of text")
        return text
    except Exception as e:
        logger.error(f"Error getting text: {e}")
        return ""

# Main procedure
async def main():
    global page  # declare the page as global so that tool functions can use it
    
    # Check if Chrome is running in debug mode
    if not check_chrome_debugging():
        print("\nERROR: Chrome is not running with remote debugging enabled or not ready.")
        print("\nPlease ensure Chrome is launched with:")
        print("\n/Applications/Google\\ Chrome.app/Contents/MacOS/Google\\ Chrome --remote-debugging-port=9222")
        print("\nAnd wait a few seconds for it to initialize.")
        print("\nYou can verify it's working by visiting: http://localhost:9222")
        print("\nThen run this script again.")
        return

    # Set up argument parser
    parser = argparse.ArgumentParser(description='Automate browser actions and capture page content.')
    parser.add_argument('--url', type=str, help='URL to navigate to (optional)')
    args = parser.parse_args()
    
    logging.info("Connecting to Chrome at %s ..." % REMOTE_DEBUGGING_URL)
    async with async_playwright() as p:
        try:
            # Connect to an existing Chrome instance via Chrome DevTools Protocol
            browser = await p.chromium.connect_over_cdp(REMOTE_DEBUGGING_URL)
        except Exception as e:
            logging.error("Failed to connect to Chrome. Is Chrome running with --remote-debugging-port? Error: %s", e)
            return

        logging.info("Connected to Chrome DevTools at %s" % REMOTE_DEBUGGING_URL)
        contexts = browser.contexts
        if len(contexts) == 0:
            logging.error("No browser contexts found. Ensure Chrome is running with a debugging port.")
            return
        context = contexts[0]  # use the first context (default profile)
        pages = context.pages
        
        # Find the first non-DevTools page
        target_page = None
        for p in pages:
            current_url = p.url
            if not current_url.startswith('devtools://'):
                target_page = p
                break
        
        if target_page:
            page = target_page
            logging.info(f"Connected to page with URL: {page.url}")
            
            # Wait for the page to be stable
            try:
                await page.wait_for_load_state("domcontentloaded", timeout=5000)
                await page.wait_for_load_state("networkidle", timeout=5000)
            except TimeoutError:
                logging.warning("Page load wait timed out, proceeding anyway...")
            
            # Try to get the title, but don't fail if we can't
            try:
                title = await page.title()
                logging.info(f"Working with page - Title: \"{title}\" URL: {page.url}")
            except Exception as e:
                logging.warning(f"Could not get page title: {e}")
                logging.info(f"Working with page - URL: {page.url}")
        else:
            logging.warning("No suitable pages found. Please ensure your target page is open in Chrome.")
            return
            
        # If URL is provided, navigate to it
        if args.url:
            logging.info(f"Navigating to URL: {args.url}")
            try:
                await page.goto(args.url, wait_until="networkidle", timeout=30000)
                logging.info(f"Successfully navigated to {args.url}")
            except Exception as e:
                logging.error(f"Failed to navigate to URL: {e}")
                return
        
        # Ensure the page is fully loaded
        try:
            await page.wait_for_load_state("domcontentloaded", timeout=5000)
            await page.wait_for_load_state("networkidle", timeout=5000)
        except TimeoutError:
            logging.info("Page load wait timed out, proceeding with current content...")

        # Set up the agent with tools
        agent = Agent(
            name="BrowserAutomationAgent",
            instructions="Use the click_load_more, scroll_page, and get_visible_text tools to gather content from the page.",
            tools=[click_load_more, scroll_page, get_visible_text],
            model="gpt-4"
        )

        # Iteratively click 'Load More' buttons if present, and scroll the page
        load_more_clicks = 0
        scroll_iter = 0
        while True:
            # Run the agent to perform the task of clicking 'Load More'
            result = await Runner.run(agent, input="Click the Load More button.")
            if "Clicked" in result.final_output:
                load_more_clicks += 1
                if load_more_clicks >= MAX_LOAD_MORE_CLICKS:
                    logging.info("Reached maximum number of Load More clicks (%d). Stopping clicks." % MAX_LOAD_MORE_CLICKS)
                    break
                # If clicked, continue loop to look for another button or to scroll after content loads
                continue

            # No 'Load More' found (or no longer visible), so attempt infinite scrolling
            scroll_iter += 1
            # Run the agent to perform the task of scrolling the page
            await Runner.run(agent, input="Scroll the page.")
            # Scroll to bottom and check if page height increases
            prev_height = await page.evaluate("document.body.scrollHeight")
            try:
                await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            except Exception as e:
                logging.warning(f"Error scrolling to bottom: {e}")
            await asyncio.sleep(SCROLL_PAUSE_SECONDS)  # wait for new content to load (if any)
            new_height = await page.evaluate("document.body.scrollHeight")
            if new_height == prev_height:
                logging.info("Reached end of page or no more content to load.")
                break  # exit loop when no more new content appears
            if scroll_iter >= MAX_SCROLL_ITERATIONS:
                logging.info("Reached maximum scroll iterations (%d). Stopping scroll." % MAX_SCROLL_ITERATIONS)
                break

        # At this point, we assume all content is loaded.
        # Capture all visible text directly
        logging.info("Collecting visible text from the page...")
        try:
            # Get the raw text directly from the page
            raw_text = await page.evaluate("document.body.innerText")
            
            if raw_text:
                # Save raw text to file
                with open(OUTPUT_TEXT_FILE, "w", encoding="utf-8") as f:
                    f.write(raw_text)
                logging.info(f"Raw text saved to {OUTPUT_TEXT_FILE} (length: {len(raw_text)} characters).")
                print("\nRaw text content:")
                print("=" * 80)
                print(raw_text)
                print("=" * 80)
            else:
                logging.warning("No text found on the page (or page is empty).")
        except Exception as e:
            logging.error(f"Error getting page text: {e}")

        # Save a full-page screenshot
        try:
            await page.screenshot(path=OUTPUT_SCREENSHOT_FILE, full_page=True)
            logging.info(f"Screenshot of the page saved to {OUTPUT_SCREENSHOT_FILE}.")
        except Exception as e:
            logging.error(f"Failed to capture screenshot: {e}")

        # Close out the browser context and connection
        try:
            await context.close()
            await browser.close()
        except Exception as e:
            logging.debug(f"Error during browser close: {e}")

        logging.info("Browser automation task completed.")

if __name__ == "__main__":
    import asyncio
    asyncio.run(main())
