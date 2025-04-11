from __future__ import annotations

import asyncio
import sys
import os
import re
import pdb  # For interactive debugging

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
        return f"Operation timed out after {timeout_seconds} seconds."

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
        await page.wait_for_timeout(2000)
        
        # 2. Wait for and fill email input
        try:
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
                except Exception:
                    continue
            
            if not email_input:
                raise Exception("Could not find email input field")
            
            await email_input.click()
            await page.keyboard.down("Control")
            await page.keyboard.press("A")
            await page.keyboard.up("Control")
            await page.keyboard.press("Backspace")
            await email_input.fill(email)
            await page.wait_for_timeout(1000)
            
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
                except Exception:
                    continue
            
            if not password_input:
                raise Exception("Could not find password input field")
            
            await password_input.fill(password)
            await page.wait_for_timeout(1000)
            
            submit_button = page.locator("input[type='submit']")
            if await submit_button.count() > 0:
                await submit_button.click()
            else:
                await page.keyboard.press("Enter")
            
            await page.wait_for_timeout(20000)
            
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
# TOOL: SEARCH IN OUTLOOK (Modified)
# =============================================================================
@function_tool(
    name_override="search_in_outlook",
    description_override="Search for emails using the Outlook search box with improved reliability and filter maintenance."
)
async def search_in_outlook(ctx: RunContextWrapper[ExtractEmailContext], query: str) -> str:
    """
    This function uses the 'search_person' from the context to build a query that forces
    the search to 'From:{search_person}'.
    """
    global computer_instance
    page = computer_instance.page

    # Use the search person value from context.
    search_person = ctx.context.search_person or "Lynn Gadue"
    query = f"From:{search_person}"

    await page.screenshot(path="screenshots/before_search.png")
    
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
            search_box = page.locator(selector)
            await search_box.wait_for(state="visible", timeout=5000)
            await search_box.click()
            await page.wait_for_timeout(500)
            
            # Clear existing text
            await page.keyboard.down("Control")
            await page.keyboard.press("A")
            await page.keyboard.up("Control")
            await page.keyboard.press("Backspace")
            await page.wait_for_timeout(500)
            
            await search_box.fill(query)
            await page.wait_for_timeout(1000)
            await page.keyboard.press("Enter")
            await page.wait_for_timeout(8000)
            
            try:
                await page.wait_for_selector("div.mailListItem", timeout=15000)
                sender_count = await page.locator("text:" + search_person).count()
                if sender_count > 0:
                    await page.screenshot(path="screenshots/search_results.png")
                    # Recheck count (in case scrolling loads more)
                    sender_count = await page.locator("text:" + search_person).count()
                    return f"Search completed for '{query}'. Found {sender_count} emails from {search_person}."
                else:
                    return f"Search completed but no emails from {search_person} found."
            except Exception as e:
                await page.screenshot(path="screenshots/search_verification_error.png")
                return f"Error verifying search results: {str(e)}"
                
        except Exception as e:
            print(f"Selector {selector} failed: {str(e)}")
            continue
    
    await page.screenshot(path="screenshots/search_failed.png")
    return "Could not perform search with any of the selectors. Check screenshots for details."

# =============================================================================
# TOOL: SAVE PAGE HTML (Unchanged)
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
# TOOL: DIAGNOSE OUTLOOK PAGE (Unchanged)
# =============================================================================
@function_tool(
    name_override="diagnose_outlook_page",
    description_override="Analyze the Outlook page structure and elements to help debug extraction issues."
)
async def diagnose_outlook_page(ctx: RunContextWrapper[ExtractEmailContext]) -> str:
    global computer_instance
    page = computer_instance.page
    await page.screenshot(path="screenshots/diagnosis.png")
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
        data.classes = {};
        ['ZtMcN', 'hcptT', 'customScrollBar', 'ms-List-cell', 'messageBody'].forEach(cls => {
            data.classes[cls] = document.querySelectorAll('.' + cls).length;
        });
        return data;
    }""")
    with open("outlook_diagnosis.json", "w") as f:
        import json
        json.dump(page_info, f, indent=2)
    try:
        await page.click("text:Lynn Gadue", timeout=3000)
        await page.wait_for_timeout(1000)
        await page.screenshot(path="screenshots/after_lynn_click.png")
    except Exception as e:
        print(f"Error clicking on 'Lynn Gadue': {str(e)}")
    return (f"Diagnosis complete. Found {page_info['textContent']['lynnGadue']} elements with 'Lynn Gadue' text. "
            "Details saved to outlook_diagnosis.json.")

# =============================================================================
# TOOL: ENHANCED OUTLOOK SCROLL (Unchanged)
# =============================================================================
@function_tool(
    name_override="enhanced_outlook_scroll",
    description_override="Scroll through emails in Outlook with improved reliability."
)
async def enhanced_outlook_scroll(ctx: RunContextWrapper[ExtractEmailContext], n_emails: int) -> str:
    global computer_instance
    page = computer_instance.page
    await page.screenshot(path="screenshots/before_scroll.png")
    
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
        except Exception:
            continue
    
    if not scroll_container:
        return "Could not find scrollable container. Check screenshots for diagnosis."
    
    initial_count = await page.locator("div[data-convid]").count()
    scroll_methods = [
        lambda: page.evaluate("""() => {
            const container = document.querySelector('div[class*="scroll"]') ||
                            document.querySelector('[role="region"]') ||
                            document.querySelector('.ms-ScrollablePane--contentContainer');
            if (!container) return false;
            container.scrollTop += 800;
            return true;
        }"""),
        lambda: page.evaluate("""() => {
            const emails = document.querySelectorAll('div[data-convid]');
            if (emails.length === 0) return false;
            emails[emails.length - 1].scrollIntoView({ behavior: 'smooth', block: 'center' });
            return true;
        }"""),
        lambda: page.keyboard.press("PageDown"),
        lambda: page.evaluate("""() => {
            window.scrollTo(0, 0);
            const elements = document.querySelectorAll('div[data-convid]');
            if (elements.length > 0) {
                elements[elements.length - 1].scrollIntoView({ behavior: 'smooth', block: 'center' });
                return true;
            }
            return false;
        }"""),
        lambda: page.evaluate("""() => {
            const container = document.querySelector('div[class*="scroll"]') ||
                            document.querySelector('[role="region"]') ||
                            document.querySelector('.ms-ScrollablePane--contentContainer');
            if (!container) return false;
            container.scrollTop = container.scrollHeight;
            return new Promise(resolve => {
                setTimeout(() => {
                    resolve(true);
                }, 1000);
            });
        }""")
    ]
    
    for _ in range(5):
        for scroll_method in scroll_methods:
            try:
                await scroll_method()
                await page.wait_for_timeout(2000)
                new_count = await page.locator("div[data-convid]").count()
                if new_count > initial_count:
                    await page.screenshot(path="screenshots/after_successful_scroll.png")
                    return f"Successfully scrolled. Initial emails: {initial_count}, Current emails: {new_count}"
            except Exception:
                continue
    
    for _ in range(10):
        await page.keyboard.press("PageDown")
        await page.wait_for_timeout(1000)
    
    final_count = await page.locator("div[data-convid]").count()
    await page.screenshot(path="screenshots/after_scroll.png")
    return f"Scroll completed. Initial emails: {initial_count}, Final emails: {final_count}"

# =============================================================================
# TOOL: COLLECT EMAILS (Unchanged)
# =============================================================================
@function_tool(
    name_override="collect_emails",
    description_override="Collect the first 3 emails from Lynn Gadue and save their content to outlook_emails.txt"
)
async def collect_emails(ctx: RunContextWrapper[ExtractEmailContext]) -> str:
    global computer_instance
    page = computer_instance.page
    await page.screenshot(path="screenshots/before_collect.png")
    sender_count = await page.locator("text:Lynn Gadue").count()
    print(f"Found {sender_count} emails from Lynn Gadue")
    if sender_count == 0:
        return "No emails from Lynn Gadue found. Please run search_in_outlook first."
    
    with open("outlook_emails.txt", "w", encoding="utf-8") as f:
        f.write("=== Email Collection Started ===\n\n")
    
    try:
        await page.wait_for_selector("div[data-convid]", timeout=10000)
        for i in range(3):
            print(f"\nProcessing email {i+1} of 3...")
            await page.evaluate(f"""() => {{
                const emails = document.querySelectorAll('div[data-convid]');
                if (emails[{i}]) {{
                    emails[{i}].scrollIntoView({{behavior: 'smooth', block: 'center'}});
                }}
            }}""")
            await page.wait_for_timeout(2000)
            
            email = page.locator("div[data-convid]").nth(i)
            await email.click()
            await page.wait_for_timeout(3000)
            await page.wait_for_selector("div[role='document']", timeout=5000)
            
            content = await page.evaluate("""() => {
                const mainContent = document.querySelector('[role="main"]');
                if (mainContent) return mainContent.innerText;
                const messageBody = document.querySelector('.messageBody, .emailBody, .messageContent, .emailContent');
                if (messageBody) return messageBody.innerText;
                const readingPane = document.querySelector('.readingPane, .reading-pane');
                if (readingPane) return readingPane.innerText;
                return document.body.innerText;
            }""")
            
            with open("outlook_emails.txt", "a", encoding="utf-8") as f:
                f.write(f"=== EMAIL {i+1} ===\n")
                f.write(content)
                f.write("\n\n")
            
            await page.screenshot(path=f"screenshots/email_{i+1}.png")
            
            for attempt in range(3):
                try:
                    await page.go_back()
                    await page.wait_for_timeout(2000)
                    await page.wait_for_selector("div[data-convid]", timeout=5000)
                    break
                except Exception as e:
                    if attempt == 2:
                        print(f"Warning: Could not return to inbox after {attempt+1} attempts")
                        await page.goto("https://outlook.office.com/mail/inbox")
                        await page.wait_for_timeout(3000)
                    else:
                        print(f"Retrying inbox navigation (attempt {attempt+1})")
                        await page.wait_for_timeout(1000)
            print(f"Successfully processed email {i+1}")
        
        print("\nCompleted processing 3 emails")
        log_success("collect_emails", {
            "total_emails_collected": 3,
            "screenshots": ["email_1.png", "email_2.png", "email_3.png"]
        })
        return "Email collection completed. Successfully collected 3 emails. Check outlook_emails.txt for the results."
        
    except Exception as e:
        print(f"Failed to process emails: {str(e)}")
        await page.screenshot(path="screenshots/collection_error.png")
        return f"Error collecting emails: {str(e)}"

# =============================================================================
# TOOL: DIAGNOSE SEARCH AND EMAILS (Unchanged)
# =============================================================================
@function_tool(
    name_override="diagnose_search_and_emails",
    description_override="Diagnose the search box and email list state to help debug extraction issues."
)
async def diagnose_search_and_emails(ctx: RunContextWrapper[ExtractEmailContext]) -> str:
    global computer_instance
    page = computer_instance.page
    await page.screenshot(path="screenshots/diagnose_search.png")
    
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
        const searchBox = document.querySelector('input[aria-label*="Search"]') || 
                         document.querySelector('input[type="search"]') ||
                         document.querySelector('input[placeholder*="Search"]');
        if (searchBox) {
            data.currentSearchValue = searchBox.value;
        }
        return data;
    }""")
    
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

# =============================================================================
# TOOL: ANALYZE EMAILS (Unchanged)
# =============================================================================
@function_tool(
    name_override="analyze_emails",
    description_override="Analyze the collected emails and create a summary of key points and open questions."
)
async def analyze_emails(ctx: RunContextWrapper[ExtractEmailContext]) -> str:
    try:
        with open("outlook_emails.txt", "r", encoding="utf-8") as f:
            email_content = f.read()
        
        print("\nDEBUG: Reading email content for analysis...")
        print(f"DEBUG: Email content length: {len(email_content)}")
        print(f"DEBUG: First 200 chars of email content:\n{email_content[:200]}\n")
        
        prompt = f"""Please analyze these emails and create a chronological summary of the discussions. Focus on:
1. Key topics discussed
2. Important decisions or conclusions reached
3. Open questions or unresolved issues
4. Any action items or next steps mentioned

Here are the emails to analyze:
{email_content}"""
        
        try:
            response = await ctx.client.chat.completions.create(
                model="gpt-4",
                messages=[
                    {"role": "system", "content": "You are an expert at analyzing email threads and extracting key information."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.7,
                max_tokens=2000
            )
            
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
# MAIN RUNNER WITH PDB BREAKPOINTS
# =============================================================================

async def main():
    global computer_instance
    enable_verbose_stdout_logging()
    
    async with OutlookPlaywrightComputer() as computer:
        computer_instance = computer
        
        # Inspect computer_instance and page before proceeding.
        pdb.set_trace()  # <-- BREAKPOINT: inspect computer_instance and computer_instance.page
        
        # Provide an initial input to start the process.
        input_items: list[TResponseInputItem] = [{"role": "user", "content": "start email extraction"}]
        context = ExtractEmailContext(n_emails=5)  # For testing, extract 5 emails.
        
        # Ensure "screenshots" directory exists
        os.makedirs("screenshots", exist_ok=True)
        
        # Check context and input_items before running Runner.
        pdb.set_trace()  # <-- BREAKPOINT: inspect context and input_items
        
        with trace("Outlook Email Extraction Agent"):
            result = await Runner.run(email_agent, input_items, context=context, max_turns=30)
            
            for new_item in result.new_items:
                agent_name = new_item.agent.name
                if isinstance(new_item, MessageOutputItem):
                    print(f"{agent_name}: {new_item.content}")
                else:
                    print(f"{agent_name}: Received item of type {new_item.__class__.__name__}")
        
        # Final breakpoint after processing.
        pdb.set_trace()  # <-- BREAKPOINT: inspect final state before finishing
        print("Email extraction task complete.")

if __name__ == "__main__":
    asyncio.run(main())
