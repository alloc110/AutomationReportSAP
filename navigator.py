import os
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

load_dotenv()

SAP_URL = os.getenv("SAP_URL", "http://report.maycha.com.vn/reports/browse")
SAP_USERNAME = os.getenv("SAP_USERNAME", "")
SAP_PASSWORD = os.getenv("SAP_PASSWORD", "")

def interactive_navigator():
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

        while True:
            page.wait_for_timeout(2000)
            current_url = page.url
            print(f"\n==========================================")
            print(f"CURRENT URL: {current_url}")
            print(f"==========================================")

            # Lấy danh sách items hiện tại
            tiles = page.locator('a.tile')
            count = tiles.count()
            print(f"Found {count} items in current view:")
            items = []
            for i in range(count):
                name = tiles.nth(i).locator('span.name').inner_text().strip()
                href = tiles.nth(i).get_attribute('href')
                items.append((name, href))
                print(f"  [{i + 1}] {name} (href: {href})")

            print("\nOptions:")
            print("  - Enter item number (1, 2, ...) to click and enter folder/report")
            print("  - Type 'export' to detect and test download/export buttons on current page")
            print("  - Type 'url' to print current URL")
            print("  - Type 'q' to quit")
            
            choice = input("\nYour choice: ").strip()

            if choice.lower() == 'q':
                break
            elif choice.isdigit():
                idx = int(choice) - 1
                if 0 <= idx < count:
                    name, href = items[idx]
                    print(f"[*] Clicking into: {name}...")
                    tiles.nth(idx).click()
                    page.wait_for_load_state("domcontentloaded")
                else:
                    print("[!] Invalid number.")
            elif choice.lower() == 'export':
                print("[*] Detecting potential Export/Download buttons...")
                export_buttons = page.locator('button, a, [role="button"]').filter(has_text="Export").all()
                if not export_buttons:
                    export_buttons = page.locator('[title*="Export" i], [aria-label*="Export" i]').all()

                print(f"[+] Found {len(export_buttons)} export-related elements.")
                for b_idx, btn in enumerate(export_buttons, 1):
                    txt = btn.inner_text().strip() or btn.get_attribute("title") or btn.get_attribute("aria-label")
                    print(f"   Export button [{b_idx}]: {txt}")
            else:
                print("[*] Continuing...")

        browser.close()

if __name__ == "__main__":
    interactive_navigator()
