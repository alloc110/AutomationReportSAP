import os
import sys
import time
import shutil
import logging
from datetime import datetime
from urllib.parse import urljoin, quote
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright
from telegram_bot import send_telegram_message, send_telegram_document

load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger("CrawlerBot")

BASE_URL = os.getenv("SAP_URL", "http://report.maycha.com.vn/reports/browse")
SAP_USERNAME = os.getenv("SAP_USERNAME", "")
SAP_PASSWORD = os.getenv("SAP_PASSWORD", "")
HEADLESS = os.getenv("HEADLESS", "false").lower() in ("true", "1", "yes")
DOWNLOAD_DIR = os.getenv("DOWNLOAD_DIR", "downloads")

# SSRS / Power BI Report Server URL formats for direct export:
# Excel:  http://host/ReportServer?/Folder/ReportName&rs:Format=EXCELOPENXML
# PDF:    http://host/ReportServer?/Folder/ReportName&rs:Format=PDF
# CSV:    http://host/ReportServer?/Folder/ReportName&rs:Format=CSV
EXPORT_FORMAT = os.getenv("EXPORT_FORMAT", "EXCELOPENXML")  # or PDF, CSV

def export_ssrs_report(page, report_path: str, report_name: str) -> str | None:
    """Download SSRS report directly or via UI"""
    # Xử lý đường dẫn sang ReportServer endpoint
    # Example: /reports/browse/01%20THINH -> /ReportServer?/01 THINH/report_name&rs:Format=EXCELOPENXML
    server_base = BASE_URL.split("/reports")[0]
    clean_path = report_path.replace("browse/", "").replace("/reports/browse/", "")
    export_url = f"{server_base}/ReportServer?/{clean_path}&rs:Format={EXPORT_FORMAT}"

    logger.info(f"Triggering export URL: {export_url}")
    try:
        with page.expect_download(timeout=90000) as download_info:
            page.goto(export_url)
        download = download_info.value
        filename = f"{report_name}.xlsx" if "EXCEL" in EXPORT_FORMAT else download.suggested_filename
        save_path = os.path.join(DOWNLOAD_DIR, filename)
        download.save_as(save_path)
        logger.info(f"Downloaded report to: {save_path}")
        return save_path
    except Exception as e:
        logger.warning(f"Direct export URL failed ({e}), attempting UI export...")
        return None

def crawl_and_download(page, current_url: str, visited_folders: set):
    """Recursive crawler to visit all folders and download all reports"""
    if current_url in visited_folders:
        return
    visited_folders.add(current_url)

    logger.info(f"Scanning folder: {current_url}")
    page.goto(current_url, wait_until="domcontentloaded", timeout=60000)
    page.wait_for_timeout(2000)

    # 1. Thu thập danh sách items trong folder hiện tại
    tiles = page.locator('a.tile')
    count = tiles.count()
    sub_folders = []
    reports = []

    for i in range(count):
        tile = tiles.nth(i)
        name = tile.locator('span.name').inner_text().strip()
        href = tile.get_attribute('href')
        is_folder = tile.locator('app-item-type-icon.glyphui-folder').count() > 0

        if is_folder:
            full_folder_url = urljoin(current_url + "/", href)
            sub_folders.append((name, full_folder_url))
        else:
            reports.append((name, href))

    logger.info(f"Found {len(sub_folders)} sub-folders and {len(reports)} reports.")

    # 2. Xử lý tải toàn bộ báo cáo trong folder này
    for report_name, report_href in reports:
        logger.info(f"--> Processing report: {report_name}")
        file_path = export_ssrs_report(page, report_href, report_name)
        
        if file_path and os.path.isfile(file_path):
            caption = f"📊 <b>Report:</b> <code>{report_name}</code>\n🕒 <b>Time:</b> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
            success = send_telegram_document(file_path, caption=caption)
            if success:
                logger.info(f"Sent {report_name} to Telegram. Deleting local file...")
                try:
                    os.remove(file_path)
                    logger.info(f"Deleted local file: {file_path}")
                except Exception as de:
                    logger.warning(f"Failed to delete file {file_path}: {de}")
            else:
                logger.error(f"Failed to send {report_name} to Telegram.")

    # 3. Đệ quy đi vào các folder con
    for sub_name, sub_url in sub_folders:
        logger.info(f"Entering sub-folder: [{sub_name}] -> {sub_url}")
        crawl_and_download(page, sub_url, visited_folders)

def main():
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    today_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    send_telegram_message(f"🚀 <b>[AutoBot]</b> Starting full report crawl & download at {today_str}...")

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=HEADLESS,
            args=["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"]
        )
        context_args = {
            "accept_downloads": True,
            "viewport": {"width": 1400, "height": 900}
        }
        if SAP_USERNAME and SAP_PASSWORD:
            context_args["http_credentials"] = {
                "username": SAP_USERNAME,
                "password": SAP_PASSWORD
            }

        context = browser.new_context(**context_args)
        page = context.new_page()

        visited_folders = set()
        try:
            crawl_and_download(page, BASE_URL, visited_folders)
            send_telegram_message(f"✅ <b>[AutoBot]</b> All folders crawled and reports sent successfully at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}!")
        except Exception as e:
            logger.error(f"Crawler error: {e}")
            send_telegram_message(f"❌ <b>[AutoBot]</b> Job failed: {str(e)}")
        finally:
            context.close()
            browser.close()

if __name__ == "__main__":
    main()
