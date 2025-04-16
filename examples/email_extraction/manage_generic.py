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
from agents.items import ItemHelpers

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
        self.logger.setLevel(logging.INFO)
        
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
        self.trace_handler.setLevel(logging.INFO)
        
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
file_handler = logging.FileHandler('generic_management.log')
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
logger.info("=== Starting Generic Page Management ===")
logger.info("Logging system initialized successfully")
logger.info("Detailed logs will be saved to generic_management.log")

class GenericPageContext(BaseModel):
    """Context for managing a generic page"""
    current_url: str = ""
    page_title: str = ""
    logged_in: bool = False
    page_info: dict = {}
    openai_client: Any = None  # Add OpenAI client to context

class GenericPageComputer(LocalPlaywrightComputer):
    """Computer class that can connect to an existing browser instance"""
    def __init__(self):
        super().__init__()  # Don't pass dimensions to parent

    async def _get_browser_and_page(self) -> tuple:
        # Try to connect to an existing browser first
        try:
            # The default port for Playwright's debugging protocol
            browser = await self.playwright.chromium.connect_over_cdp("http://localhost:9222")
            logger.info("Connected to existing browser instance")
        except Exception as e:
            logger.warning(f"Could not connect to existing browser: {e}")
            logger.info("Launching new browser instance")
            # Fall back to launching a new browser if connection fails
            browser = await self.playwright.chromium.launch(headless=False)
        
        # Get the first available page or create a new one
        pages = browser.contexts[0].pages if browser.contexts else []
        if pages:
            page = pages[0]
            logger.info(f"Using existing page with URL: {page.url}")
        else:
            page = await browser.new_page()
            logger.info("Created new page")
        
        # Set viewport size
        await page.set_viewport_size({"width": 1920, "height": 1080})
        return browser, page

@function_tool(
    name_override="examine_page",
    description_override="Examine the current page and gather information about its structure and interactive elements."
)
async def examine_page(ctx: RunContextWrapper[GenericPageContext]) -> str:
    """Examine the current page and gather information about its structure"""
    global computer_instance
    page = computer_instance.page
    
    try:
        # Take a screenshot
        await page.screenshot(path="screenshots/page_examination.png")
        
        # Gather page information
        page_info = await page.evaluate("""() => {
            const data = {
                url: window.location.href,
                title: document.title,
                interactiveElements: {
                    buttons: Array.from(document.querySelectorAll('button')).map(btn => ({
                        text: btn.textContent.trim(),
                        ariaLabel: btn.getAttribute('aria-label'),
                        role: btn.getAttribute('role'),
                        disabled: btn.disabled
                    })),
                    inputs: Array.from(document.querySelectorAll('input')).map(input => ({
                        type: input.type,
                        placeholder: input.placeholder,
                        ariaLabel: input.getAttribute('aria-label'),
                        value: input.value
                    })),
                    links: Array.from(document.querySelectorAll('a')).map(link => ({
                        text: link.textContent.trim(),
                        href: link.href,
                        ariaLabel: link.getAttribute('aria-label')
                    })),
                    clickableElements: Array.from(document.querySelectorAll('[role="button"], [role="link"]')).map(el => ({
                        text: el.textContent.trim(),
                        role: el.getAttribute('role'),
                        ariaLabel: el.getAttribute('aria-label')
                    }))
                },
                mainContent: {
                    headings: Array.from(document.querySelectorAll('h1, h2, h3, h4, h5, h6')).map(h => h.textContent.trim()),
                    lists: document.querySelectorAll('ul, ol').length,
                    tables: document.querySelectorAll('table').length
                },
                navigation: {
                    menus: document.querySelectorAll('[role="menu"]').length,
                    tabs: document.querySelectorAll('[role="tab"]').length
                }
            };
            return data;
        }""")
        
        # Save the information to a JSON file
        with open("page_examination.json", "w") as f:
            json.dump(page_info, f, indent=2)
        
        # Update context
        ctx.context.current_url = page_info["url"]
        ctx.context.page_title = page_info["title"]
        ctx.context.page_info = page_info
        
        # Create a summary
        summary = f"""Page Examination Results:
URL: {page_info["url"]}
Title: {page_info["title"]}

Interactive Elements:
- Buttons: {len(page_info["interactiveElements"]["buttons"])}
- Inputs: {len(page_info["interactiveElements"]["inputs"])}
- Links: {len(page_info["interactiveElements"]["links"])}
- Clickable Elements: {len(page_info["interactiveElements"]["clickableElements"])}

Main Content:
- Headings: {len(page_info["mainContent"]["headings"])}
- Lists: {page_info["mainContent"]["lists"]}
- Tables: {page_info["mainContent"]["tables"]}

Navigation:
- Menus: {page_info["navigation"]["menus"]}
- Tabs: {page_info["navigation"]["tabs"]}

Detailed information has been saved to page_examination.json
Screenshot saved to screenshots/page_examination.png
"""
        
        return summary
    except Exception as e:
        logger.error(f"Error examining page: {str(e)}")
        await page.screenshot(path="screenshots/examination_error.png")
        return f"Error examining page: {str(e)}"

# Agent instructions
INSTRUCTIONS = f"""{RECOMMENDED_PROMPT_PREFIX}
You are a generic page management agent. Your task is to:
1. Examine the current page and understand its structure
2. Identify available actions and interactive elements
3. Provide a clear summary of what can be done on this page

After completing these steps, your task is done.

Note: The page examination will be saved to page_examination.json and a screenshot will be taken."""

# Agent definition
page_agent = Agent[GenericPageContext](
    name="Generic Page Management Agent",
    instructions=INSTRUCTIONS,
    tools=[
        examine_page
    ],
    model="gpt-4",
    model_settings=ModelSettings(
        tool_choice="required",
        temperature=0.7,
        max_tokens=2000
    ),
)

async def main():
    global computer_instance
    
    # Initialize the detailed logger first
    logger = DetailedTraceLogger()
    
    try:
        # Initialize OpenAI client
        from openai import AsyncOpenAI
        openai_client = AsyncOpenAI()
        
        async with GenericPageComputer() as computer:
            computer_instance = computer
            
            # Provide an initial input to start the process
            input_items: list[TResponseInputItem] = [{"role": "user", "content": "examine the current page"}]
            context = GenericPageContext(openai_client=openai_client)
            
            # Ensure "screenshots" directory exists
            os.makedirs("screenshots", exist_ok=True)
            
            with trace("Generic Page Management Agent"):
                result = await Runner.run(page_agent, input_items, context=context, max_turns=30)
                
                for new_item in result.new_items:
                    agent_name = new_item.agent.name
                    if isinstance(new_item, MessageOutputItem):
                        print(f"{agent_name}: {ItemHelpers.text_message_output(new_item)}")
                    else:
                        print(f"{agent_name}: Received item of type {new_item.__class__.__name__}")
            
            print("Page examination complete.")
    except Exception as e:
        print(f"Error during execution: {str(e)}")
        import traceback
        traceback.print_exc()
    finally:
        # Restore original streams before exiting
        logger.restore_streams()

if __name__ == "__main__":
    asyncio.run(main()) 