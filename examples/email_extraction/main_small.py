from __future__ import annotations as _annotations

import asyncio
import os
import re
from pydantic import BaseModel

# Add the project root so that the examples folder is in the module search path.
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../../")))
from examples.tools.computer_use import LocalPlaywrightComputer

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

async def main():
    async with OutlookPlaywrightComputer() as computer:
        page = computer.page
        
        # 1. Login to Outlook
        print("Logging in to Outlook...")
        try:
            # Wait for and fill email
            await page.wait_for_selector("input[type='email']", timeout=10000)
            await page.fill("input[type='email']", "tolds@3clife.info")
            await page.click("input[type='submit']")
            
            # Wait for and fill password
            await page.wait_for_selector("input[type='password']", timeout=10000)
            await page.fill("input[type='password']", "annuiTy2024!")
            await page.click("input[type='submit']")
            
            # Wait for MFA if needed
            await page.wait_for_timeout(20000)  # 20 seconds for MFA
            
            print("Login successful")
        except Exception as e:
            print(f"Login failed: {str(e)}")
            return
        
        # 2. Wait for inbox to load
        print("Waiting for inbox to load...")
        await page.wait_for_timeout(5000)
        
        # 3. Find and process the first 2 emails
        print("Looking for emails...")
        try:
            # Wait for email list to load
            await page.wait_for_selector("div[data-convid]", timeout=10000)
            
            # Open file for writing
            with open("outlook_email_small.txt", "w", encoding="utf-8") as f:
                # Process first email
                print("Processing first email...")
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
            
            print("Both emails saved to outlook_email_small.txt")
            
        except Exception as e:
            print(f"Failed to process emails: {str(e)}")
            return

if __name__ == "__main__":
    asyncio.run(main()) 