from playwright.async_api import Page
import asyncio
from datetime import datetime
import os

async def main(page: Page):
    """
    Download the program plan from Smartsheet as an Excel file.
    Uses an existing logged-in browser session.
    
    Args:
        page: The Playwright Page object from an active browser session
    """
    
    smartsheet_url = os.getenv("SMARTSHEET_PROJECT_URL")
    smartsheet_name = os.getenv("SMARTSHEET_PROJECT_NAME")
    
    try:
        # Navigate to sheet (already logged in)
        #print("Navigating to sheet...")
        await page.goto(smartsheet_url, wait_until="domcontentloaded")
        await asyncio.sleep(4)
        
        print("✓ On sheet, proceeding with export...")
        
        # Export using keyboard navigation
        #print("Opening File menu...")
        await page.click('text=File')
        await asyncio.sleep(1)
        
        # Navigate to Export (9 down arrows)
        #print("Navigating to Export...")
        for i in range(9):
            await page.keyboard.press('ArrowDown')
            await asyncio.sleep(0.2)
        
        # Open Export submenu
        #print("Opening Export submenu...")
        await page.keyboard.press('ArrowRight')
        await asyncio.sleep(1)
        
        # Navigate to "Export to Microsoft Excel" (5 down arrows)
        #print("Navigating to Export to Microsoft Excel...")
        for i in range(5):
            await page.keyboard.press('ArrowDown')
            await asyncio.sleep(0.2)
        
        # Wait for download - SET UP BEFORE PRESSING ENTER!
        #print("Waiting for download...")
        async with page.expect_download(timeout=30000) as download_info:
            # Select Export to Microsoft Excel
            #print("Selecting Export to Microsoft Excel...")
            await page.keyboard.press('Enter')
        
        download = await download_info.value
        
        # Create folder and save file
        today = datetime.now()
        folder_name = f"{smartsheet_name}_program_plan"
        filename = f"{smartsheet_name}_program_plan_{today.year}_{today.month:02d}_{today.day:02d}.xlsx"
        
        # CORRECT PATH - NOT INSIDE VENV!
        base_path = "/mnt/c/Users/krpop/Amway Corp/Global Account Management Community - Workspace Core Team - Workspace Core Team/Program Status"
        folder_path = f"{base_path}/{folder_name}"
        
        # Create the folder if it doesn't exist
        #print(f"Creating folder if needed: {folder_name}")
        os.makedirs(folder_path, exist_ok=True)
        
        # Full path with folder and filename
        downloads_path = f"{folder_path}/{filename}"
        
        # Save the file
        #print(f"Saving file to: {downloads_path}")
        await download.save_as(downloads_path)
        
        print(f"\n✓ Success! Program plan saved to: {folder_name}")
        
        await asyncio.sleep(2)
        
    except Exception as e:
        print(f"\n!!! ERROR creating program plan: {e}")
        raise


if __name__ == '__main__':
    print("This script must be called with an active browser session.")
    print("Run smartsheet_login.py first, then pass the page object to this function.")