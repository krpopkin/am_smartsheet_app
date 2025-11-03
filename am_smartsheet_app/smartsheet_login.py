from playwright.async_api import async_playwright
import asyncio
import os
from dotenv import load_dotenv

async def main():
    """
    Log into Smartsheet and return the browser session objects.
    Returns: tuple of (playwright, browser, context, page)
    """
    load_dotenv()
    
    email = "ken.popkin@amway.com"
    password = "Smartsheet1!"
    sheet_url = os.getenv("SMARTSHEET_PROJECT_URL")
    
    # Start Playwright async
    p = await async_playwright().start()
    
    browser = await p.chromium.launch(
        headless=False,
        slow_mo=500,
        args=['--disable-blink-features=AutomationControlled']
    )
    
    context = await browser.new_context(
        viewport={'width': 1920, 'height': 1080},
        user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36',
    )
    
    page = await context.new_page()
    await page.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined});")
    
    print("Logging in...")
    await perform_login(page, email, password, sheet_url)
    
    print("Login complete!")
    
    # Return all objects so they can be used by subsequent scripts
    return p, browser, context, page


async def perform_login(page, email, password, sheet_url):
    """
    Perform fresh login.
    """
    
    try:
        await page.goto("https://app.smartsheet.com/b/login", wait_until="domcontentloaded")
        await asyncio.sleep(2)
        
        await page.fill('#loginEmail', email)
        await page.click('#formControl')
        await asyncio.sleep(3)
        
        if await page.locator("text='Sign in with email and password'").count() > 0:
            await page.click("text='Sign in with email and password'")
            await asyncio.sleep(2)
        
        await page.fill('#loginPassword', password)
        await page.click('#formControl')
        await asyncio.sleep(5)
        
        await page.goto(sheet_url, wait_until="domcontentloaded")
        await asyncio.sleep(3)
        
        if "login" in page.url.lower():
            raise Exception("Login failed - still on login page")
        
        #print("✓ Login successful!")
        
    except Exception as e:
        print(f"!!! Login ERROR: {e}")
        raise


if __name__ == '__main__':
    async def run():
        p, browser, context, page = await main()
        input("Press Enter to close browser...")
        await browser.close()
        await p.stop()
    
    asyncio.run(run())