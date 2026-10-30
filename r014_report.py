import os
import sys
import time
import logging
from datetime import datetime, timedelta
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
logger = logging.getLogger("R014_Report")

SAP_USERNAME = os.getenv("SAP_USERNAME", "")
SAP_PASSWORD = os.getenv("SAP_PASSWORD", "")
HEADLESS = os.getenv("HEADLESS", "true").lower() in ("true", "1", "yes")
DOWNLOAD_DIR = os.getenv("DOWNLOAD_DIR", "downloads")

REPORT_PAGE_URL = "http://report.maycha.com.vn/reports/report/02%20KIM%20B%E1%BA%A2O%20M%C3%83N%20%C4%90%C6%AF%E1%BB%9CNG/04.%20INVENTORY/R014-03.%20S%E1%BB%95%20Chi%20Ti%E1%BA%BFt%20V%E1%BA%ADt%20T%C6%B0"

def run_r014_report(target_date: datetime = None):
    """
    Xuất báo cáo 'R014-03. Sổ Chi Tiết Vật Tư'
    :param target_date: Ngày cần xuất (mặc định là ngày hôm qua)
    """
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    if target_date is None:
        target_date = datetime.now() - timedelta(days=1)

    date_str_mdy = target_date.strftime("%m/%d/%Y")
    date_str_dmy = target_date.strftime("%d/%m/%Y")
    logger.info(f"Target Date for R014-03: {date_str_dmy} (input format: {date_str_mdy})")

    target_filename = f"R014_03_SoChiTietVatTu_{target_date.strftime('%Y%m%d')}.xlsx"
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

        def wait_ajax(frm):
            time.sleep(1)
            frm.wait_for_function("""() => {
                const prm = window.Sys && window.Sys.WebForms && window.Sys.WebForms.PageRequestManager && window.Sys.WebForms.PageRequestManager.getInstance();
                return !prm || !prm.get_isInAsyncPostBack();
            }""", timeout=45000)
            time.sleep(1)

        try:
            logger.info(f"Navigating to R014-03 report: {REPORT_PAGE_URL}")
            page.goto(REPORT_PAGE_URL, wait_until="networkidle", timeout=60000)

            iframe_elem = page.wait_for_selector("iframe.viewer", timeout=45000)
            frame = iframe_elem.content_frame()
            frame.wait_for_selector("#ReportViewerControl_ctl04", timeout=30000)
            time.sleep(2)

            # 1. Điền Từ ngày & Đến ngày
            logger.info(f"Setting Từ ngày and Đến ngày: {date_str_mdy}")
            frame.evaluate(f"""() => {{
                $get('ReportViewerControl_ctl04_ctl03_txtValue').value = '{date_str_mdy}';
                $get('ReportViewerControl_ctl04_ctl05_txtValue').value = '{date_str_mdy}';
            }}""")

            # 2. Chọn tất cả Kho (ctl07)
            logger.info("Opening 'Kho' dropdown and selecting (Select All)...")
            frame.click("#ReportViewerControl_ctl04_ctl07_ctl01")
            frame.wait_for_selector("#ReportViewerControl_ctl04_ctl07_divDropDown_ctl00", timeout=15000)
            kho_all = frame.locator("#ReportViewerControl_ctl04_ctl07_divDropDown_ctl00")
            if not kho_all.is_checked():
                kho_all.check()
            time.sleep(1)
            frame.click("#ReportViewerControl_ctl04_ctl07_ctl01")  # Đóng dropdown Kho
            time.sleep(1)

            # 3. Chọn tất cả Nhóm hàng (ctl09)
            logger.info("Opening 'Nhóm hàng' dropdown and selecting (Select All)...")
            frame.click("#ReportViewerControl_ctl04_ctl09_ctl01")
            frame.wait_for_selector("#ReportViewerControl_ctl04_ctl09_divDropDown_ctl00", timeout=15000)
            nhom_all = frame.locator("#ReportViewerControl_ctl04_ctl09_divDropDown_ctl00")
            if not nhom_all.is_checked():
                nhom_all.check()
            time.sleep(1)
            frame.click("#ReportViewerControl_ctl04_ctl09_ctl01")  # Đóng dropdown Nhóm hàng
            time.sleep(1)

            # 4. Chờ nạp danh mục 'Từ mã hàng' (ctl15)
            logger.info("Waiting for 'Từ mã hàng' to be populated...")
            frame.wait_for_function("""() => {
                const el = $get('ReportViewerControl_ctl04_ctl15_ddValue');
                return el && !el.disabled && el.options.length > 1;
            }""", timeout=45000)

            # Chọn mã đầu tiên trong danh sách (index=1, vì index=0 là <Select a Value>)
            logger.info("Selecting first item in 'Từ mã hàng'...")
            frame.select_option("#ReportViewerControl_ctl04_ctl15_ddValue", index=1)
            wait_ajax(frame)

            # 5. Chờ nạp danh mục 'Tới mã hàng' (ctl19)
            logger.info("Waiting for 'Tới mã hàng' to be populated...")
            frame.wait_for_function("""() => {
                const el = $get('ReportViewerControl_ctl04_ctl19_ddValue');
                return el && !el.disabled && el.options.length > 1;
            }""", timeout=45000)

            # Chọn mã cuối cùng để bao quát toàn bộ dải vật tư
            ctl19_count = frame.evaluate("""() => $get('ReportViewerControl_ctl04_ctl19_ddValue').options.length""")
            logger.info(f"Selecting last item (index {ctl19_count - 1}) in 'Tới mã hàng'...")
            frame.select_option("#ReportViewerControl_ctl04_ctl19_ddValue", index=ctl19_count - 1)
            time.sleep(1)

            # 6. Bấm nút 'Xem báo cáo'
            logger.info("Clicking 'Xem báo cáo' button...")
            frame.click("#ReportViewerControl_ctl04_ctl00")

            # 7. Chờ báo cáo kết xuất hoàn tất (báo cáo lớn, có thể mất 30-90 giây)
            logger.info("Waiting for report rendering to complete (this may take up to 2 minutes)...")
            start_wait = time.time()
            export_ready = False
            while time.time() - start_wait < 180:
                time.sleep(3)
                status = frame.evaluate("""() => {
                    const exp = $get('ReportViewerControl_ctl05_ctl04_ctl00_ButtonLink');
                    const asyncWait = $get('ReportViewerControl_AsyncWait');
                    const isWait = asyncWait ? (asyncWait.style.display !== 'none') : false;
                    const isExpDisabled = exp ? (exp.getAttribute('disabled') === 'disabled' || exp.disabled || exp.classList.contains('Disabled')) : true;
                    return !isWait && !isExpDisabled;
                }""")
                if status:
                    export_ready = True
                    break

            if not export_ready:
                raise Exception("Report did not finish rendering within timeout (180s)!")

            logger.info(f"Report rendered successfully in {int(time.time() - start_wait)} seconds.")

            # 8. Mở menu Export và tải Excel
            logger.info("Triggering Excel download...")
            frame.click("#ReportViewerControl_ctl05_ctl04_ctl00_ButtonLink")
            time.sleep(1)

            excel_link = frame.locator('#ReportViewerControl_ctl05_ctl04_ctl00_Menu a[title="Excel"]').first
            if excel_link.count() == 0:
                excel_link = frame.locator('a:has-text("Excel")').first

            with page.expect_download(timeout=120000) as download_info:
                excel_link.click()

            download = download_info.value
            download.save_as(file_saved_path)
            file_size_kb = round(os.path.getsize(file_saved_path) / 1024, 2)
            logger.info(f"Excel file successfully saved: {file_saved_path} ({file_size_kb} KB)")

            # 9. Gửi Telegram
            caption = (
                f"📊 <b>R014-03. Sổ Chi Tiết Vật Tư</b>\n"
                f"🏢 <b>Đơn vị:</b> 02 KIM BẢO MÃN ĐƯỜNG\n"
                f"📅 <b>Kỳ báo cáo:</b> {date_str_dmy}\n"
                f"🕒 <b>Xuất lúc:</b> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"📁 <b>Tệp:</b> <code>{target_filename}</code> ({file_size_kb} KB)"
            )
            logger.info("Sending document to Telegram...")
            success = send_telegram_document(file_saved_path, caption=caption)

            # 10. Dọn dẹp tệp local
            if success:
                logger.info(f"Report sent to Telegram successfully. Cleaning up: {file_saved_path}")
                try:
                    os.remove(file_saved_path)
                except Exception as de:
                    logger.warning(f"Could not delete local file: {de}")
            else:
                logger.warning("Telegram sending failed, keeping file locally.")

            logger.info("R014-03 report workflow finished successfully!")
            return file_saved_path

        except Exception as e:
            logger.error(f"Error during R014-03 report export: {e}")
            send_telegram_message(f"❌ <b>[R014-03 Report Error]:</b> {str(e)}")
            raise e
        finally:
            context.close()
            browser.close()

if __name__ == "__main__":
    run_r014_report()

