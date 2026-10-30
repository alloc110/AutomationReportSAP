import os
import sys
import logging
import urllib.parse
from datetime import datetime
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright
from telegram_bot import send_telegram_message, send_telegram_document

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

load_dotenv()

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger("Costing_Report")

SAP_USERNAME = os.getenv("SAP_USERNAME", "")
SAP_PASSWORD = os.getenv("SAP_PASSWORD", "")
HEADLESS = os.getenv("HEADLESS", "true").lower() in ("true", "1", "yes")
DOWNLOAD_DIR = os.getenv("DOWNLOAD_DIR", "downloads")

# Đường dẫn chính thức của Báo cáo Xuất Nhập Tồn trong thư mục 03. COSTING
REPORT_PATH = "/02 KIM BẢO MÃN ĐƯỜNG/03. COSTING/Costing. Báo Cáo Xuất Nhập Tồn"
REPORT_PAGE_URL = "http://report.maycha.com.vn/reports/report/02%20KIM%20B%E1%BA%A2O%20M%C3%83N%20%C4%90%C6%AF%E1%BB%9CNG/03.%20COSTING/Costing.%20B%C3%A1o%20C%C3%A1o%20Xu%E1%BA%A5t%20Nh%E1%BA%ADp%20T%E1%BB%93n"

# Danh mục mã các nhóm hàng:
# 100: 10-Thành phẩm
# 101: 11-Hàng Hóa
# 102: 20-Bán Thành Phẩm (Semi Product) - CK
# 103: 30-Nguyên vật liệu
# 104: 32-Bao Bì
# 110: 12-Dịch vụ vận chuyển
ALL_ITEM_GROUPS = ["100", "101", "102", "103", "104", "110"]

def run_costing_report(month: int = None, year: int = None, item_groups: list = None, use_ui: bool = False):
    """
    Xuất báo cáo Costing. Báo Cáo Xuất Nhập Tồn
    :param month: Tháng (1-12), mặc định là tháng hiện tại
    :param year: Năm (ví dụ: 2026), mặc định là năm hiện tại
    :param item_groups: Danh sách mã nhóm hàng cần chọn (mặc định chọn tất cả)
    :param use_ui: True nếu muốn tương tác qua giao diện web, False để tải trực tiếp qua ReportServer (nhanh hơn)
    """
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    now = datetime.now()
    if month is None:
        month = now.month
    if year is None:
        year = now.year
    if item_groups is None:
        item_groups = ALL_ITEM_GROUPS

    logger.info(f"Target Period for Costing Report: Tháng {month}/{year} | Nhóm hàng: {item_groups}")

    target_filename = f"Costing_BaoCaoXuatNhapTon_{year}_{month:02d}.xlsx"
    file_saved_path = os.path.join(DOWNLOAD_DIR, target_filename)

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=HEADLESS,
            args=["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"]
        )
        context = browser.new_context(
            http_credentials={"username": SAP_USERNAME, "password": SAP_PASSWORD} if SAP_USERNAME else None,
            accept_downloads=True,
            viewport={"width": 1600, "height": 1000}
        )
        page = context.new_page()

        try:
            if not use_ui:
                # Cách 1: Xuất trực tiếp qua endpoint ReportServer (tốc độ cao ~3 giây)
                logger.info("Exporting report directly via ReportServer endpoint...")
                encoded_path = urllib.parse.quote(REPORT_PATH)
                query_parts = [
                    f"month={month}",
                    f"year={year}",
                ]
                for grp in item_groups:
                    query_parts.append(f"Itemgroup={grp}")
                query_parts.append("rs:Format=EXCELOPENXML")

                export_url = f"http://report.maycha.com.vn/ReportServer?{encoded_path}&" + "&".join(query_parts)
                logger.info(f"Triggering direct download URL...")

                with page.expect_download(timeout=90000) as download_info:
                    try:
                        page.goto(export_url)
                    except Exception:
                        pass
                download = download_info.value
                download.save_as(file_saved_path)

            else:
                # Cách 2: Tương tác qua giao diện Web ReportViewer
                logger.info(f"Navigating to report UI: {REPORT_PAGE_URL}")
                page.goto(REPORT_PAGE_URL, wait_until="networkidle", timeout=60000)

                iframe_element = page.wait_for_selector('iframe.viewer', timeout=45000)
                frame = iframe_element.content_frame()
                frame.wait_for_selector("#ReportViewerControl_ctl04", timeout=30000)

                # 1. Chọn Tháng
                logger.info(f"Selecting Month: {month}")
                frame.select_option("#ReportViewerControl_ctl04_ctl03_ddValue", str(month))

                # 2. Chọn Năm
                logger.info(f"Selecting Year: {year}")
                # Year option in SSRS might be index or string value
                frame.select_option("#ReportViewerControl_ctl04_ctl05_ddValue", label=str(year))

                # 3. Mở dropdown Nhóm hàng và chọn
                logger.info("Opening 'Nhóm hàng' dropdown...")
                frame.click("#ReportViewerControl_ctl04_ctl07_ctl01")
                frame.wait_for_selector("#ReportViewerControl_ctl04_ctl07_divDropDown_ctl00", timeout=15000)

                if len(item_groups) == len(ALL_ITEM_GROUPS):
                    # Chọn tất cả
                    select_all_cb = frame.locator("#ReportViewerControl_ctl04_ctl07_divDropDown_ctl00")
                    if not select_all_cb.is_checked():
                        select_all_cb.check()
                    logger.info("Checked (Select All) for Nhóm hàng.")
                else:
                    # Chọn cụ thể từng nhóm
                    for val in item_groups:
                        cb = frame.locator(f'#ReportViewerControl_ctl04_ctl07_divDropDown input[value="{val}"]')
                        if cb.count() > 0 and not cb.is_checked():
                            cb.check()
                            logger.info(f"Checked item group: {val}")

                # Đóng dropdown
                frame.click("#ReportViewerControl_ctl04_ctl07_ctl01")

                # 4. Bấm Xem báo cáo
                logger.info("Clicking 'Xem báo cáo' button...")
                frame.click("#ReportViewerControl_ctl04_ctl00")

                # 5. Chờ kết xuất xong
                logger.info("Waiting for report rendering...")
                frame.wait_for_selector("#ReportViewerControl_ctl05_ctl04_ctl00_ButtonLink", timeout=120000)

                # 6. Xuất Excel
                logger.info("Triggering Excel export...")
                frame.click("#ReportViewerControl_ctl05_ctl04_ctl00_ButtonLink")
                excel_link = frame.locator('#ReportViewerControl_ctl05_ctl04_ctl00_Menu a[title="Excel"]').first
                if excel_link.count() == 0:
                    excel_link = frame.locator('a:has-text("Excel")').first

                with page.expect_download(timeout=120000) as download_info:
                    excel_link.click()
                download = download_info.value
                download.save_as(file_saved_path)

            file_size_kb = round(os.path.getsize(file_saved_path) / 1024, 2)
            logger.info(f"Report Excel saved successfully: {file_saved_path} ({file_size_kb} KB)")

            # Gửi tài liệu qua Telegram
            caption = (
                f"📊 <b>Costing. Báo Cáo Xuất Nhập Tồn</b>\n"
                f"📅 <b>Kỳ báo cáo:</b> Tháng {month}/{year}\n"
                f"📦 <b>Nhóm hàng:</b> {len(item_groups)} nhóm\n"
                f"🕒 <b>Xuất lúc:</b> {now.strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"📁 <b>Tệp:</b> <code>{target_filename}</code> ({file_size_kb} KB)"
            )
            logger.info("Sending document to Telegram...")
            success = send_telegram_document(file_saved_path, caption=caption)

            if success:
                logger.info(f"Report sent to Telegram successfully. Cleaning up: {file_saved_path}")
                try:
                    os.remove(file_saved_path)
                except Exception as de:
                    logger.warning(f"Could not delete local file: {de}")
            else:
                logger.warning("Telegram sending failed, keeping file locally.")

            logger.info("Costing report workflow finished successfully!")
            return file_saved_path

        except Exception as e:
            logger.error(f"Error during Costing report export: {e}")
            send_telegram_message(f"❌ <b>[Costing Report Error]:</b> {str(e)}")
            raise e
        finally:
            context.close()
            browser.close()

if __name__ == "__main__":
    # Chạy thử nghiệm xuất tháng 10/2026 với tất cả nhóm hàng
    run_costing_report(month=10, year=2026)
