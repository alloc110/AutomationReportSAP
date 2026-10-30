import os
import sys
import time
import logging

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

from datetime import datetime, timedelta
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright
from telegram_bot import send_telegram_message, send_telegram_document

load_dotenv()

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger("P01_Report")

SAP_USERNAME = os.getenv("SAP_USERNAME", "")
SAP_PASSWORD = os.getenv("SAP_PASSWORD", "")
HEADLESS = os.getenv("HEADLESS", "true").lower() in ("true", "1", "yes")
DOWNLOAD_DIR = os.getenv("DOWNLOAD_DIR", "downloads")

# Direct URL to Report P01
REPORT_PAGE_URL = "http://report.maycha.com.vn/reports/report/02%20KIM%20B%E1%BA%A2O%20M%C3%83N%20%C4%90%C6%AF%E1%BB%9CNG/02.%20Mua%20h%C3%A0ng/P01.%20B%C3%A1o%20c%C3%A1o%20t%E1%BB%95ng%20h%E1%BB%A3p%20mua%20h%C3%A0ng%20theo%20PR"

def run_p01_report():
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    # Date range: yesterday
    yesterday = datetime.now() - timedelta(days=1)
    # SSRS expects MM/DD/YYYY format for LCID 1033 (en-US)
    date_str_mdy = yesterday.strftime("%m/%d/%Y")
    date_str_dmy = yesterday.strftime("%d/%m/%Y")
    logger.info(f"Target date for report: {date_str_dmy} (input format: {date_str_mdy})")

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=HEADLESS,
            args=["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"]
        )

        context_args = {
            "accept_downloads": True,
            "viewport": {"width": 1600, "height": 1000}
        }
        if SAP_USERNAME and SAP_PASSWORD:
            context_args["http_credentials"] = {
                "username": SAP_USERNAME,
                "password": SAP_PASSWORD
            }

        context = browser.new_context(**context_args)
        page = context.new_page()

        try:
            logger.info(f"Navigating to report page: {REPORT_PAGE_URL}")
            page.goto(REPORT_PAGE_URL, wait_until="networkidle", timeout=60000)

            # Wait for Report Viewer iframe
            logger.info("Waiting for ReportViewer iframe...")
            iframe_element = page.wait_for_selector('iframe.viewer', timeout=45000)
            frame = iframe_element.content_frame()
            if not frame:
                raise Exception("Could not attach to ReportViewer iframe!")

            logger.info("Attached to ReportViewer iframe successfully.")

            # Wait for parameters toolbar to be ready inside frame
            frame.wait_for_selector("#ReportViewerControl_ctl04_ctl03_txtValue", timeout=30000)
            time.sleep(1)

            # 1. Fill Date Inputs (Từ ngày & Đến ngày)
            logger.info(f"Setting FromDate and ToDate: {date_str_mdy}")
            frame.evaluate(f"""function() {{
                $get('ReportViewerControl_ctl04_ctl03_txtValue').value = '{date_str_mdy}';
                $get('ReportViewerControl_ctl04_ctl05_txtValue').value = '{date_str_mdy}';
            }}""")

            # 2. Trigger PostBack to load dependent parameters (item, store)
            logger.info("Triggering ASP.NET postback for ToDate...")
            frame.evaluate("""function() {
                window.setTimeout(function() {
                    __doPostBack('ReportViewerControl$ctl04$ctl05$txtValue', '');
                }, 0);
            }""")

            # Wait for dependent parameters (ctl13: Mặt hàng) to become enabled
            logger.info("Waiting for dependent parameters to become active...")
            frame.wait_for_function("""function() {
                var inp = $get('ReportViewerControl_ctl04_ctl13_txtValue');
                return inp && !inp.disabled;
            }""", timeout=30000)
            logger.info("Mặt hàng and Cửa hàng controls are now enabled.")
            time.sleep(1)

            # 3. Select All for 'Mặt hàng' (ctl13)
            logger.info("Opening 'Mặt hàng' dropdown...")
            frame.click("#ReportViewerControl_ctl04_ctl13_ctl01")
            frame.wait_for_selector("#ReportViewerControl_ctl04_ctl13_divDropDown_ctl00", timeout=15000)
            select_all_item = frame.locator("#ReportViewerControl_ctl04_ctl13_divDropDown_ctl00")
            if not select_all_item.is_checked():
                select_all_item.check()
                logger.info("Checked (Select All) for Mặt hàng.")
            time.sleep(1)
            frame.click("#ReportViewerControl_ctl04_ctl13_ctl01")  # Close dropdown
            time.sleep(1)

            # 4. Select All for 'Cửa hàng' (ctl17)
            logger.info("Opening 'Cửa hàng' dropdown...")
            frame.click("#ReportViewerControl_ctl04_ctl17_ctl01")
            frame.wait_for_selector("#ReportViewerControl_ctl04_ctl17_divDropDown_ctl00", timeout=15000)
            select_all_store = frame.locator("#ReportViewerControl_ctl04_ctl17_divDropDown_ctl00")
            if not select_all_store.is_checked():
                select_all_store.check()
                logger.info("Checked (Select All) for Cửa hàng.")
            time.sleep(1)
            frame.click("#ReportViewerControl_ctl04_ctl17_ctl01")  # Close dropdown
            time.sleep(1)

            # 5. Click 'Xem báo cáo' button
            logger.info("Clicking 'Xem báo cáo' button...")
            frame.click("#ReportViewerControl_ctl04_ctl00")

            # 6. Wait for report rendering to finish
            logger.info("Waiting for report rendering to complete...")
            start_wait = time.time()
            export_ready = False
            while time.time() - start_wait < 120:
                time.sleep(2)
                res = frame.evaluate("""function() {
                    var exportBtn = $get('ReportViewerControl_ctl05_ctl04_ctl00_ButtonLink');
                    return exportBtn ? (exportBtn.offsetWidth > 0 && exportBtn.offsetHeight > 0) : false;
                }""")
                if res:
                    export_ready = True
                    break

            if not export_ready:
                raise Exception("Report did not finish rendering within timeout (120s)!")

            logger.info(f"Report rendered successfully in {int(time.time() - start_wait)} seconds.")

            # 7. Click Export dropdown button
            export_btn = frame.locator("#ReportViewerControl_ctl05_ctl04_ctl00_ButtonLink")
            logger.info("Opening Export dropdown menu...")
            export_btn.click()
            time.sleep(1)

            # 8. Click 'Excel' option and capture download
            excel_link = frame.locator('#ReportViewerControl_ctl05_ctl04_ctl00_Menu a[title="Excel"]').first
            if excel_link.count() == 0:
                excel_link = frame.locator('a:has-text("Excel")').first

            target_filename = f"P01_BaoCaoTongHopMuaHangTheoPR_{yesterday.strftime('%Y%m%d')}.xlsx"
            file_saved_path = os.path.join(DOWNLOAD_DIR, target_filename)
            logger.info(f"Triggering Excel download to {file_saved_path}...")
            with page.expect_download(timeout=120000) as download_info:
                excel_link.click()

            download = download_info.value
            download.save_as(file_saved_path)
            file_size_kb = round(os.path.getsize(file_saved_path) / 1024, 2)
            logger.info(f"Excel file successfully saved: {file_saved_path} ({file_size_kb} KB)")

            # 9. Send file to Telegram
            caption = (
                f"📊 <b>P01. Báo cáo tổng hợp mua hàng theo PR</b>\n"
                f"📅 <b>Kỳ báo cáo:</b> {date_str_dmy}\n"
                f"🕒 <b>Xuất lúc:</b> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"📁 <b>Tệp:</b> <code>{target_filename}</code> ({file_size_kb} KB)"
            )
            logger.info("Sending report document to Telegram...")
            success = send_telegram_document(file_saved_path, caption=caption)

            # 10. Clean up local file after sending
            if success:
                logger.info(f"Report sent to Telegram successfully. Cleaning up: {file_saved_path}")
                try:
                    os.remove(file_saved_path)
                except Exception as de:
                    logger.warning(f"Could not delete local file: {de}")
            else:
                logger.warning("Telegram sending failed, keeping file for manual check.")

            logger.info("P01 report workflow finished successfully!")

        except Exception as e:
            logger.error(f"Error during report execution: {e}")
            send_telegram_message(f"❌ <b>[P01 Report Error]:</b> {str(e)}")
            raise e
        finally:
            context.close()
            browser.close()

if __name__ == "__main__":
    run_p01_report()
