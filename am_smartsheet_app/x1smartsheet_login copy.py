from playwright.sync_api import sync_playwright
import time
from datetime import datetime
import os

def main():
    # This version uses keyboard navigation which is often more reliable than hover

    email = "ken.popkin@amway.com"
    password = "Smartsheet1!"
    sheet_url = "https://app.smartsheet.com/sheets/HCh3Jrfcx25f8mvJP8pGCVxg834CfR6W5xqWV781?view=grid"

    with sync_playwright() as p:
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
        
        try:
            # Login
            #print("Logging in...")
            page.goto("https://app.smartsheet.com/b/login", wait_until="domcontentloaded")
            page.fill('#loginEmail', email)
            page.click('#formControl')
            time.sleep(2)
            
            if page.locator("text='Sign in with email and password'").count() > 0:
                page.click("text='Sign in with email and password'")
                time.sleep(2)
            
            page.fill('#loginPassword', password)
            page.click('#formControl')
            time.sleep(5)
            
            # Go to sheet
            #print("Going to sheet...")
            page.goto(sheet_url, wait_until="domcontentloaded")
            time.sleep(4)
            
            #print("\nExporting using keyboard navigation...")
            
            # Click File menu
            #print("  Opening File menu...")
            page.click('text=File')
            time.sleep(1)
            
            # Navigate to Export using Down arrow (9 times based on manual testing)
            #print("  Navigating to Export...")
            for i in range(9):
                page.keyboard.press('ArrowDown')
                time.sleep(0.2)
            
            # Press Right arrow to open Export submenu
            #print("  Opening Export submenu...")
            page.keyboard.press('ArrowRight')
            time.sleep(1)
            
            # Navigate to "Export to Microsoft Excel" using Down arrow (5 times)
            #print("  Navigating to Export to Microsoft Excel...")
            for i in range(5):
                page.keyboard.press('ArrowDown')
                time.sleep(0.2)
            
            # Press Enter to select
            #print("  Selecting Export to Microsoft Excel...")
            page.keyboard.press('Enter')
            time.sleep(1)
            
            # Wait for download
            #print("  Waiting for download...")
            with page.expect_download(timeout=30000) as download_info:
                pass
            
            download = download_info.value
            
            # Create folder and save file
            today = datetime.now()
            #folder_name = f"am_program_status_{today.year}_{today.month:02d}_{today.day:02d}"
            folder_name = f"am_program_plan"
            filename = f"am_program_plan_{today.year}_{today.month:02d}_{today.day:02d}.xlsx"
            
            base_path = "/mnt/c/Users/krpop/Amway Corp/Global Account Management Community - Workspace Core Team - Workspace Core Team/Program Status"
            folder_path = f"{base_path}/{folder_name}"
            
            # Create the folder if it doesn't exist
            #print(f"  Creating folder: {folder_name}")
            os.makedirs(folder_path, exist_ok=True)
            
            # Full path with folder and filename
            downloads_path = f"{folder_path}/{filename}"
            
            #print(f"  Saving as: {filename}")
            download.save_as(downloads_path)
            
            print(f"\n✓ Success! File saved to: {downloads_path}")
            
            time.sleep(10)
            #print("\nClosing browser...")
            
        except Exception as e:
            print(f"\n!!! ERROR: {e}")
            time.sleep(30)
        
        browser.close()
        print("Done!")
        
if __name__ == '__main__':
    main()