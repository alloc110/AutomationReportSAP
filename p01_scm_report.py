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
logger = logging.getLogger("P01_SCM_Report")

SAP_USERNAME = os.getenv("SAP_USERNAME", "")
SAP_PASSWORD = os.getenv("SAP_PASSWORD", "")
HEADLESS = os.getenv("HEADLESS", "true").lower() in ("true", "1", "yes")
DOWNLOAD_DIR = os.getenv("DOWNLOAD_DIR", "downloads")

REPORT_PAGE_URL = "http://report.maycha.com.vn/reports/report/01%20TH%E1%BB%8ANH%20TH%E1%BA%BE%20VINH%20HOA/02.%20Mua%20h%C3%A0ng%20(SCM)/P01.%20B%C3%A1o%20c%C3%A1o%20t%E1%BB%95ng%20h%E1%BB%A3p%20y%C3%AAu%20c%E1%BA%A7u%20h%C3%A0ng%20h%C3%B3a%20theo%20Purchase%20Request"


def run_p01_scm_report(target_date: datetime = None):
    """
    Xuất báo cáo 'P01. Báo cáo tổng hợp yêu cầu hàng hóa theo Purchase Request'
    Đơn vị: 01 THỊNH THẾ VINH HOA / 02. Mua hàng (SCM)
    :param target_date: Ngày cần xuất (mặc định: hôm qua)
    """
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    if target_date is None:
        target_date = datetime.now() - timedelta(days=1)

    date_str_mdy = target_date.strftime("%m/%d/%Y")
    date_str_dmy = target_date.strftime("%d/%m/%Y")
    logger.info(f"Target Date for P01 SCM: {date_str_dmy} (input format: {date_str_mdy})")

    target_filename = f"P01_TongHopYeuCauHangHoa_SCM_{target_date.strftime('%Y%m%d')}.xlsx"
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
            logger.info(f"Navigating to P01 SCM report: {REPORT_PAGE_URL}")
            page.goto(REPORT_PAGE_URL, wait_until="networkidle", timeout=60000)

            iframe_elem = page.wait_for_selector("iframe.viewer", timeout=45000)
            frame = iframe_elem.content_frame()
            frame.wait_for_selector("#ReportViewerControl_ctl04_ctl03_txtValue", timeout=30000)
            time.sleep(2)

            # 1. Điền ngày và đánh dấu NULL cho Ngày yêu cầu
            logger.info(f"Setting Từ ngày and Đến ngày: {date_str_mdy}, setting Ngày yêu cầu NULL...")
            frame.evaluate(f"""() => {{
                $get('ReportViewerControl_ctl04_ctl03_txtValue').value = '{date_str_mdy}';
                $get('ReportViewerControl_ctl04_ctl05_txtValue').value = '{date_str_mdy}';
                $get('ReportViewerControl_ctl04_ctl07_cbNull').checked = true;
                $get('ReportViewerControl_ctl04_ctl09_cbNull').checked = true;
                window.setTimeout(function() {{
                    __doPostBack('ReportViewerControl$ctl04$ctl05$txtValue', '');
                }}, 0);
            }}""")

            # 2. Chờ dropdowns phụ thuộc được nạp xong
            logger.info("Waiting for dependent dropdowns to populate...")
            time.sleep(5)
            frame.wait_for_function("""() => {
                const div13 = $get('ReportViewerControl_ctl04_ctl13_divDropDown');
                const div17 = $get('ReportViewerControl_ctl04_ctl17_divDropDown');
                return div13 && div17;
            }""", timeout=30000)

            # 3. Chọn tất cả Mặt hàng (ctl13)
            logger.info("Opening 'Mặt hàng' dropdown and selecting (Select All)...")
            frame.click("#ReportViewerControl_ctl04_ctl13_ctl01")
            frame.wait_for_selector("#ReportViewerControl_ctl04_ctl13_divDropDown_ctl00", timeout=15000)
            cb_item = frame.locator("#ReportViewerControl_ctl04_ctl13_divDropDown_ctl00")
            if not cb_item.is_checked():
                cb_item.check()
            time.sleep(1)
            frame.click("#ReportViewerControl_ctl04_ctl13_ctl01")
            time.sleep(1)

            # 4. Chọn tất cả Cửa hàng (ctl17)
            logger.info("Opening 'Cửa hàng' dropdown and selecting (Select All)...")
            frame.click("#ReportViewerControl_ctl04_ctl17_ctl01")
            frame.wait_for_selector("#ReportViewerControl_ctl04_ctl17_divDropDown_ctl00", timeout=15000)
            cb_store = frame.locator("#ReportViewerControl_ctl04_ctl17_divDropDown_ctl00")
            if not cb_store.is_checked():
                cb_store.check()
            time.sleep(1)
            frame.click("#ReportViewerControl_ctl04_ctl17_ctl01")
            time.sleep(1)

            # 5. Bấm nút 'Xem báo cáo'
            logger.info("Clicking 'Xem báo cáo' button...")
            frame.click("#ReportViewerControl_ctl04_ctl00")

            # 6. Chờ render hoàn tất
            logger.info("Waiting for report rendering to complete...")
            start_wait = time.time()
            export_ready = False
            while time.time() - start_wait < 120:
                time.sleep(2)
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
                raise Exception("Report did not finish rendering within timeout (120s)!")

            logger.info(f"Report rendered successfully in {int(time.time() - start_wait)} seconds.")

            # 7. Tải file Excel
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

            # 8. Gửi Telegram
            caption = (
                f"📊 <b>P01. Báo cáo tổng hợp yêu cầu hàng hóa theo Purchase Request</b>\n"
                f"🏢 <b>Đơn vị:</b> 01 THỊNH THẾ VINH HOA (SCM)\n"
                f"📅 <b>Kỳ báo cáo:</b> {date_str_dmy}\n"
                f"🕒 <b>Xuất lúc:</b> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"📁 <b>Tệp:</b> <code>{target_filename}</code> ({file_size_kb} KB)"
            )
            logger.info("Sending document to Telegram...")
            success = send_telegram_document(file_saved_path, caption=caption)

            # 9. Dọn dẹp tệp local
            if success:
                logger.info(f"Report sent to Telegram successfully. Cleaning up: {file_saved_path}")
                try:
                    os.remove(file_saved_path)
                except Exception as de:
                    logger.warning(f"Could not delete local file: {de}")
            else:
                logger.warning("Telegram sending failed, keeping file locally.")

            logger.info("P01 SCM report workflow finished successfully!")
            return file_saved_path

        except Exception as e:
            logger.error(f"Error during P01 SCM report export: {e}")
            send_telegram_message(f"❌ <b>[P01 SCM Report Error]:</b> {str(e)}")
            raise e
        finally:
            context.close()
            browser.close()


if __name__ == "__main__":
    run_p01_scm_report()

