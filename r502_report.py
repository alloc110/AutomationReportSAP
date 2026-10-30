import os
import sys
import time
import logging
import urllib.parse
from datetime import datetime, timedelta
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright
from telegram_bot import send_telegram_message, send_telegram_document

# Đảm bảo in tiếng Việt chuẩn trên Windows console
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
logger = logging.getLogger("R502_Report")

SAP_USERNAME = os.getenv("SAP_USERNAME", "")
SAP_PASSWORD = os.getenv("SAP_PASSWORD", "")
HEADLESS = os.getenv("HEADLESS", "true").lower() in ("true", "1", "yes")
DOWNLOAD_DIR = os.getenv("DOWNLOAD_DIR", "downloads")

# Đường dẫn báo cáo R502 trong thư mục 03. COSTING
REPORT_PATH = "/02 KIM BẢO MÃN ĐƯỜNG/03. COSTING/R502. Báo cáo chi tiết lệnh sản xuất"
REPORT_PAGE_URL = f"http://report.maycha.com.vn/reports/report/{urllib.parse.quote('02 KIM BẢO MÃN ĐƯỜNG/03. COSTING/R502. Báo cáo chi tiết lệnh sản xuất')}"
REPORT_SERVER_URL = f"http://report.maycha.com.vn/ReportServer/Pages/ReportViewer.aspx?{urllib.parse.quote(REPORT_PATH)}&rc:showbackbutton=true"


def run_r502_report(target_date: datetime = None) -> bool:
    """
    Xuất báo cáo R502. Báo cáo chi tiết lệnh sản xuất (KIM BẢO MÃN ĐƯỜNG)
    Mặc định lấy ngày hôm qua cho cả Từ ngày và Đến ngày.
    """
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    if target_date is None:
        target_date = datetime.now() - timedelta(days=1)

    date_str_vn = target_date.strftime("%d/%m/%Y")
    date_str_mdy = f"{target_date.month}/{target_date.day}/{target_date.year}"
    date_suffix = target_date.strftime("%Y%m%d")

    target_filename = f"R502_ChiTietLenhSanXuat_KBMD_{date_suffix}.xlsx"
    file_saved_path = os.path.join(DOWNLOAD_DIR, target_filename)

    logger.info(f"Target Date for R502: {date_str_vn} (MDY: {date_str_mdy})")

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=HEADLESS,
            args=["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"]
        )

        context = browser.new_context(
            http_credentials={
                "username": SAP_USERNAME,
                "password": SAP_PASSWORD
            },
            viewport={"width": 1600, "height": 900},
            accept_downloads=True
        )

        page = context.new_page()

        try:
            logger.info("Mở trang báo cáo R502...")
            page.goto(REPORT_SERVER_URL, timeout=60000, wait_until="load")

            # Đợi bảng tham số hiển thị
            page.wait_for_selector("#ReportViewerControl_ctl04_ctl03_txtValue", timeout=30000)
            time.sleep(1)

            # Điền Từ ngày và Đến ngày
            logger.info(f"Điền tham số Từ ngày & Đến ngày: {date_str_mdy}")
            page.fill("#ReportViewerControl_ctl04_ctl03_txtValue", date_str_mdy)
            page.fill("#ReportViewerControl_ctl04_ctl05_txtValue", date_str_mdy)
            time.sleep(1)

            # Bấm 'Xem báo cáo' (View Report)
            logger.info("Nhấn 'Xem báo cáo'...")
            page.click("#ReportViewerControl_ctl04_ctl00")

            # Chờ báo cáo kết xuất xong (AsyncWait biến mất)
            logger.info("Đang chờ báo cáo kết xuất dữ liệu...")
            for i in range(30):
                time.sleep(2)
                is_loading = page.evaluate("""() => {
                    const el = document.querySelector('#ReportViewerControl_AsyncWait');
                    return el && el.style.display !== 'none';
                }""")
                if not is_loading:
                    break

            time.sleep(2)
            logger.info("Báo cáo đã hiển thị! Bắt đầu tải file Excel...")

            # Mở menu Export
            page.click("#ReportViewerControl_ctl05_ctl04_ctl00_Button")
            time.sleep(1)

            # Chọn xuất định dạng Excel
            with page.expect_download(timeout=60000) as download_info:
                page.click("a[title='Excel']")

            download = download_info.value
            download.save_as(file_saved_path)

            file_size_kb = round(os.path.getsize(file_saved_path) / 1024, 2)
            logger.info(f"Xuất file thành công: {file_saved_path} ({file_size_kb} KB)")

            # Gửi file qua Telegram
            caption = (
                f"📊 <b>BÁO CÁO CHI TIẾT LỆNH SẢN XUẤT (R502)</b>\n"
                f"🏢 <b>Đơn vị:</b> 02 KIM BẢO MÃN ĐƯỜNG\n"
                f"📅 <b>Kỳ báo cáo:</b> {date_str_vn}\n"
                f"💾 <b>Kích thước:</b> {file_size_kb} KB\n"
                f"⏰ <b>Thời gian xuất:</b> {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"
            )
            send_telegram_document(file_saved_path, caption=caption)
            return True

        except Exception as e:
            logger.error(f"Lỗi trong quá trình xuất báo cáo R502: {e}")
            send_telegram_message(
                f"⚠️ <b>[CẢNH BÁO LỖI]</b> Xuất Báo cáo Chi Tiết Lệnh Sản Xuất (R502) thất bại: {e}"
            )
            return False

        finally:
            browser.close()


if __name__ == "__main__":
    run_r502_report()

