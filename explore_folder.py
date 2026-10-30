import os
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

load_dotenv()

SAP_URL = os.getenv("SAP_URL", "http://report.maycha.com.vn/reports/browse")
SAP_USERNAME = os.getenv("SAP_USERNAME", "")
SAP_PASSWORD = os.getenv("SAP_PASSWORD", "")

def explore_folder(folder_name="01 THỊNH THẾ VINH HOA"):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context_args = {
            "viewport": {"width": 1400, "height": 900},
            "accept_downloads": True
        }
        if SAP_USERNAME and SAP_PASSWORD:
            context_args["http_credentials"] = {
                "username": SAP_USERNAME,
                "password": SAP_PASSWORD
            }

        context = browser.new_context(**context_args)
        page = context.new_page()

        print(f"[*] Navigating to: {SAP_URL}")
        page.goto(SAP_URL, wait_until="domcontentloaded", timeout=60000)

        # Wait for folder list
        tile_selector = f'a.tile:has-text("{folder_name}")'
        page.wait_for_selector(tile_selector, timeout=30000)
        print(f"[+] Found folder: {folder_name}")
        
        # Click into folder
        print(f"[*] Clicking into: {folder_name}...")
        page.click(tile_selector)
        page.wait_for_load_state("domcontentloaded")
        page.wait_for_timeout(3000)

        # List all items inside folder
        print(f"\n===== ITEMS IN FOLDER [{folder_name}] =====")
        items = page.locator('a.tile span.name').all_text_contents()
        for idx, item in enumerate(items, 1):
            print(f" {idx}. {item}")
        print("============================================\n")

        print("Browser is open. Check the report names and export buttons.")
        input("Press Enter here when finished: ")
        browser.close()

if __name__ == "__main__":
    explore_folder("01 THỊNH THẾ VINH HOA")
