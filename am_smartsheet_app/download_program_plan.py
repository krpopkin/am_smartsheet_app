from playwright.sync_api import sync_playwright
import time
from datetime import datetime
import os
import json

def main():
    """
    Download the program plan from Smartsheet as an Excel file.
    Requires a valid session to be saved by smartsheet_login.py first.
    """
    
    sheet_url = "https://app.smartsheet.com/sheets/HCh3Jrfcx25f8mvJP8pGCVxg834CfR6W5xqWV781?view=grid"
    session_file = "/mnt/c/Users/krpop/.smartsheet_session.json"
    
    # Check if session exists
    if not os.path.exists(session_file):
        print("⚠ ERROR: No saved session found!")
        print("Please run smartsheet_login.py first to login and save session.")
        return
    
    with sync_playwright() as p:
        browser = None
        try:
            # Launch browser
            browser = p.chromium.launch(
                headless=False,
                slow_mo=500,
                args=['--disable-blink-features=AutomationControlled']
            )
            
            context = browser.new_context(
                viewport={'width': 1920, 'height': 1080},
                user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36',
            )
            
            page = context.new_page()
            page.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined});")
            
            # Load the saved session cookies
            #print("Loading saved session...")
            with open(session_file, 'r') as f:
                cookies = json.load(f)
            context.add_cookies(cookies)
            
            # Navigate to sheet
            #print("Navigating to sheet...")
            page.goto(sheet_url, wait_until="domcontentloaded")
            time.sleep(4)
            
            # Check if session is still valid (not redirected to login)
            if "login" in page.url.lower() or page.locator("#loginEmail").count() > 0:
                print("⚠ ERROR: Session expired!")
                print("Please run smartsheet_login.py first to refresh your login session.")
                return
            
            #print("✓ Session valid, proceeding with export...")
            
            # Export using keyboard navigation
            #print("Opening File menu...")
            page.click('text=File')
            time.sleep(1)
            
            # Navigate to Export (9 down arrows)
            #print("Navigating to Export...")
            for i in range(9):
                page.keyboard.press('ArrowDown')
                time.sleep(0.2)
            
            # Open Export submenu
            #print("Opening Export submenu...")
            page.keyboard.press('ArrowRight')
            time.sleep(1)
            
            # Navigate to "Export to Microsoft Excel" (5 down arrows)
            #print("Navigating to Export to Microsoft Excel...")
            for i in range(5):
                page.keyboard.press('ArrowDown')
                time.sleep(0.2)
            
            # Select Export to Microsoft Excel
            #print("Selecting Export to Microsoft Excel...")
            page.keyboard.press('Enter')
            time.sleep(1)
            
            # Wait for download
            #print("Waiting for download...")
            with page.expect_download(timeout=30000) as download_info:
                pass
            
            download = download_info.value
            
            # Create folder and save file
            today = datetime.now()
            folder_name = "am_program_plan"
            filename = f"am_program_plan_{today.year}_{today.month:02d}_{today.day:02d}.xlsx"
            
            base_path = "/mnt/c/Users/krpop/Amway Corp/Global Account Management Community - Workspace Core Team - Workspace Core Team/Program Status"
            folder_path = f"{base_path}/{folder_name}"
            
            # Create the folder if it doesn't exist
            #print(f"Creating folder if needed: {folder_name}")
            os.makedirs(folder_path, exist_ok=True)
            
            # Full path with folder and filename
            downloads_path = f"{folder_path}/{filename}"
            
            # Save the file
            #print(f"Saving file: {filename}")
            download.save_as(downloads_path)
            
            print(f"\n✓ Success! File saved to: {filename}")
            
            time.sleep(2)
            
        except Exception as e:
            print(f"\n!!! ERROR: {e}")
        finally:
            if browser:
                browser.close()
    
    #print("Download complete!")


if __name__ == '__main__':
    main()