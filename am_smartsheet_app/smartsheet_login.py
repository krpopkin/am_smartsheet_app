from playwright.sync_api import sync_playwright
import time
import os
import json

def main():
    """
    Log into Smartsheet and save the browser session state.
    If a valid session already exists, it will verify and reuse it.
    """
    
    email = "ken.popkin@amway.com"
    password = "Smartsheet1!"
    sheet_url = "https://app.smartsheet.com/sheets/HCh3Jrfcx25f8mvJP8pGCVxg834CfR6W5xqWV781?view=grid"
    
    # Path to store browser session state
    session_file = "/mnt/c/Users/krpop/.smartsheet_session.json"
    
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
        
        # Check if we have a saved session
        if os.path.exists(session_file):
            print("Found existing session, checking if still valid...")
            
            try:
                # Load saved cookies
                with open(session_file, 'r') as f:
                    cookies = json.load(f)
                
                # Add cookies to context
                context.add_cookies(cookies)
                
                # Test if session is still valid
                page.goto(sheet_url, wait_until="domcontentloaded")
                time.sleep(3)
                
                # Check if we're still logged in or got redirected to login
                if "login" in page.url.lower() or page.locator("#loginEmail").count() > 0:
                    print("Session expired, logging in fresh...")
                    perform_login(page, email, password, sheet_url)
                    # Save new session
                    cookies = context.cookies()
                    with open(session_file, 'w') as f:
                        json.dump(cookies, f)
                    print("✓ New session saved!")
                else:
                    print("✓ Session is still valid! No need to login again.")
                    
            except Exception as e:
                print(f"Error checking session: {e}")
                print("Logging in fresh...")
                perform_login(page, email, password, sheet_url)
                # Save new session
                cookies = context.cookies()
                with open(session_file, 'w') as f:
                    json.dump(cookies, f)
                print("✓ New session saved!")
        else:
            print("No existing session found, logging in...")
            perform_login(page, email, password, sheet_url)
            # Save session
            cookies = context.cookies()
            with open(session_file, 'w') as f:
                json.dump(cookies, f)
            print("✓ Session saved!")
        
        time.sleep(2)
        browser.close()
    
    print("Login process complete!")


def perform_login(page, email, password, sheet_url):
    """
    Perform fresh login.
    """
    
    try:
        # Login
        #print("  Navigating to login page...")
        page.goto("https://app.smartsheet.com/b/login", wait_until="domcontentloaded")
        page.fill('#loginEmail', email)
        page.click('#formControl')
        time.sleep(2)
        
        if page.locator("text='Sign in with email and password'").count() > 0:
            #print("  Selecting email/password login...")
            page.click("text='Sign in with email and password'")
            time.sleep(2)
        
        print("  Entering password...")
        page.fill('#loginPassword', password)
        page.click('#formControl')
        time.sleep(5)
        
        # Navigate to the sheet to verify login worked
        #print("  Navigating to sheet...")
        page.goto(sheet_url, wait_until="domcontentloaded")
        time.sleep(3)
        
        # Verify we're actually on the sheet
        if "login" in page.url.lower():
            raise Exception("Login failed - still on login page")
        
        print("✓ Successfully logged in!")
        
    except Exception as e:
        print(f"!!! Login ERROR: {e}")
        raise


if __name__ == '__main__':
    main()