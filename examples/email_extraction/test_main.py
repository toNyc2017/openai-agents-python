from __future__ import annotations as _annotations

import asyncio
import sys
import os

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

# Import types from Playwright.
from playwright.async_api import Page, Locator

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

# =============================================================================
# HELPER FUNCTION: FIND THE SEARCH BOX
# =============================================================================

async def find_search_box(page: Page) -> Locator:
    """
    Attempts to locate the Outlook search box using multiple candidate selectors.
    Returns the first Locator that is found and visible.
    
    Raises:
        Exception: If no candidate selector finds a visible search box.
    """
    candidate_selectors = [
        "input[aria-label='Search']",
        "input[aria-label='Search mail']",
        "div[role='search'] input",
        "input[placeholder*='Search']"
    ]
    
    for selector in candidate_selectors:
        try:
            locator = page.locator(selector)
            await locator.wait_for(state="visible", timeout=5000)
            print(f"Found search box using selector: {selector}")
            return locator
        except Exception:
            print(f"Selector did not match: {selector}")
            continue

    raise Exception("Could not locate the Outlook search box using candidate selectors.")

# =============================================================================
# CONTEXT DEFINITION
# =============================================================================

class ExtractEmailContext(BaseModel):
    text: str = ""         # Additional instructions if needed.
    n_emails: int = 5      # Number of emails to extract (for testing).
    search_person: str = "Lynn Gadue"
    email: str = "tolds@3clife.info"
    password: str = "annuiTy2024!"

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

    # 1. Navigate to Outlook sign-in page.
    await page.goto("https://outlook.office.com")
    await page.wait_for_selector("input[type='email']", timeout=15000)
    await page.fill("input[type='email']", email)
    await page.click("input[type='submit']")  # Click Next.
    await page.wait_for_timeout(2000)

    # 2. Fill password.
    await page.wait_for_selector("input[type='password']", timeout=15000)
    await page.fill("input[type='password']", password)
    await page.click("input[type='submit']")

    # 3. Wait for an element indicating the user is signed in.
    await page.wait_for_selector("div[role='navigation']", timeout=30000)

    # 4. Navigate explicitly to the inbox.
    await page.goto("https://outlook.office.com/mail/inbox")

    # 5. Wait for the search box to appear.
    await page.wait_for_selector("input[aria-label='Search']", timeout=30000)

    # 6. Ensure the screenshots directory exists.
    os.makedirs("screenshots", exist_ok=True)
    
    # 7. Take a screenshot after arriving in the inbox.
    await page.screenshot(path="screenshots/outlook_login.png")
    return "Logged in successfully; screenshot saved as outlook_login.png."

# =============================================================================
# TOOL: SEARCH IN OUTLOOK
# =============================================================================

@function_tool(
    name_override="search_in_outlook",
    description_override="Search for emails using the Outlook search box by dynamically determining its selector."
)
async def search_in_outlook(ctx: RunContextWrapper[ExtractEmailContext], query: str) -> str:
    global computer_instance
    page = computer_instance.page
    
    try:
        search_box = await find_search_box(page)
    except Exception as e:
        return f"Error locating search box: {e}"
    
    try:
        await search_box.click()
        await search_box.fill(query)
        await page.keyboard.press("Enter")
    except Exception as e:
        return f"Error interacting with search box: {e}"
    
    try:
        await page.wait_for_selector("div.mailListItem", timeout=15000)
    except Exception as e:
        return f"Search results did not load in time: {e}"
    
    await page.screenshot(path="screenshots/search_results.png")
    return f"Searched for '{query}'; screenshot saved as search_results.png."

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
    description_override="Iterate over the first n emails from search results, open each one, extract content, and save to a file."
)
async def collect_emails(ctx: RunContextWrapper[ExtractEmailContext], n_emails: int) -> str:
    global computer_instance
    page = computer_instance.page
    emails = page.locator("div.mailListItem")
    total = await emails.count()
    count = min(n_emails, total)
    if count == 0:
        return "No emails found."
    output_path = "outlook_emails.txt"
    with open(output_path, "w", encoding="utf-8") as f:
        for i in range(count):
            await emails.nth(i).click()
            await page.wait_for_selector("div.messageBody", timeout=15000)
            sender = await page.locator("span.senderName").inner_text() or "Unknown Sender"
            subject = await page.locator("h1.subjectLine").inner_text() or "No Subject"
            date = await page.locator("span.emailDate").inner_text() or "Unknown Date"
            body = await page.locator("div.messageBody").inner_text() or ""
            screenshot_path = f"screenshots/email_{i+1}.png"
            await page.screenshot(path=screenshot_path)
            f.write(f"Email {i+1} - {date}\nSubject: {subject}\nFrom: {sender}\n")
            f.write(body)
            f.write("\n\n" + "-"*40 + "\n\n")
        f.flush()
    return f"Saved {count} emails to {output_path} and screenshots in 'screenshots/' folder."

# =============================================================================
# DYNAMIC AGENT INSTRUCTIONS
# =============================================================================

default_context = ExtractEmailContext()
INSTRUCTIONS = f"""{RECOMMENDED_PROMPT_PREFIX}
Your task is to automate Outlook webmail using browser automation. Follow these steps:
1. Log in to https://outlook.office.com using the provided credentials.
2. Wait for MFA to complete.
3. Use the search box to search for emails from "From:{default_context.search_person}".
4. Wait for the search results to load and take a screenshot.
5. Optionally, perform a minimal scroll action to load additional emails.
6. Iterate over the first {default_context.n_emails} emails in the results:
   - Open each email,
   - Extract its full text (including subject, sender, date, and body),
   - Take a screenshot of the open email.
7. Save all extracted email texts into a text file, with headers for each email.
8. Return a final summary of all saved file paths.
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
    tools=[login_to_outlook, search_in_outlook, minimal_outlook_scroll, collect_emails],
    model_settings=ModelSettings(tool_choice="required"),
)

# =============================================================================
# TEST FUNCTION: LOGIN AND FIND SEARCH BOX, THEN PAUSE
# =============================================================================

async def test_find_search_box():
    enable_verbose_stdout_logging()
    async with OutlookPlaywrightComputer() as computer:
        global computer_instance
        computer_instance = computer
        os.makedirs("screenshots", exist_ok=True)
        
        page = computer_instance.page
        
        # Log in programmatically using the login tool.
        # Since automatic login might be disrupted by MFA, prompt the user to log in manually if needed.
        #login_result = await login_to_outlook(None, default_context.email, default_context.password)
        # Log in programmatically using the raw login function
        login_result = await login_to_outlook.raw_function(None, default_context.email, default_context.password)

        print(f"Login result: {login_result}")
        
        # If login wasn't successful or if you prefer manual intervention, pause here.
        input("If not already logged in, please complete login manually (enter credentials, complete MFA) and then press Enter...")
        
        # Now try to find the search box.
        try:
            search_box = await find_search_box(page)
            print("Search box found!")
        except Exception as e:
            print(f"Error locating search box: {e}")
        
        # Pause execution so you can inspect the browser.
        input("Press Enter to close the browser and finish...")

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

# =============================================================================
# ENTRY POINT
# =============================================================================

if __name__ == "__main__":
    # To test login and search box detection, run:
    asyncio.run(test_find_search_box())
    # To run the full agent workflow, comment out the above line and uncomment the line below:
    # asyncio.run(main())
