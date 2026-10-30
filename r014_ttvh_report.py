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

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger("R014_TTVH_Report")

SAP_USERNAME = os.getenv("SAP_USERNAME", "")
SAP_PASSWORD = os.getenv("SAP_PASSWORD", "")
HEADLESS = os.getenv("HEADLESS", "true").lower() in ("true", "1", "yes")
DOWNLOAD_DIR = os.getenv("DOWNLOAD_DIR", "downloads")

REPORT_PAGE_URL = "http://report.maycha.com.vn/reports/report/01%20TH%E1%BB%8ANH%20TH%E1%BA%BE%20VINH%20HOA/04.%20Inventory/R014-03.%20S%E1%BB%95%20Chi%20Ti%E1%BA%BFt%20V%E1%BA%ADt%20T%C6%B0"


def run_r014_ttvh_report(target_date: datetime = None) -> bool:
    """
    Xuất báo cáo 'R014-03. Sổ Chi Tiết Vật Tư' của 01 THỊNH THẾ VINH HOA
    :param target_date: Ngày cần xuất (mặc định là ngày hôm qua)
    """
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    if target_date is None:
        target_date = datetime.now() - timedelta(days=1)

    date_str_mdy = target_date.strftime("%m/%d/%Y")
    date_str_dmy = target_date.strftime("%d/%m/%Y")
    logger.info(f"Target Date for R014-03 (TTVH): {date_str_dmy} (input format: {date_str_mdy})")

    target_filename = f"R014_03_SoChiTietVatTu_TTVH_{target_date.strftime('%Y%m%d')}.xlsx"
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
            logger.info(f"Navigating to R014-03 (TTVH): {REPORT_PAGE_URL}")
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
            frame.click("#ReportViewerControl_ctl04_ctl07_ctl01")
            time.sleep(1)

            # 3. Chọn tất cả Nhóm hàng (ctl09)
            logger.info("Opening 'Nhóm hàng' dropdown and selecting (Select All)...")
            frame.click("#ReportViewerControl_ctl04_ctl09_ctl01")
            frame.wait_for_selector("#ReportViewerControl_ctl04_ctl09_divDropDown_ctl00", timeout=15000)
            nhom_all = frame.locator("#ReportViewerControl_ctl04_ctl09_divDropDown_ctl00")
            if not nhom_all.is_checked():
                nhom_all.check()
            time.sleep(1)
            frame.click("#ReportViewerControl_ctl04_ctl09_ctl01")
            time.sleep(1)

            # 4. Kích hoạt PostBack để nạp danh mục vật tư
            logger.info("Kích hoạt nạp danh mục vật tư...")
            frame.evaluate("""() => {
                window.setTimeout(() => {
                    __doPostBack('ReportViewerControl$ctl04$ctl09$txtValue', '');
                }, 0);
            }""")

            # Đợi ctl15 mở
            for _ in range(45):
                time.sleep(1)
                ctl15_ready = frame.evaluate("""() => {
                    const dd = $get('ReportViewerControl_ctl04_ctl15_ddValue');
                    return dd && !dd.disabled && dd.options.length > 1;
                }""")
                if ctl15_ready:
                    break

            logger.info("Chọn mã hàng đầu tiên trong 'Từ mã hàng' (ctl15)...")
            frame.evaluate("""() => {
                const dd = $get('ReportViewerControl_ctl04_ctl15_ddValue');
                dd.selectedIndex = 1;
                if (dd.onchange) {
                    dd.onchange();
                } else {
                    dd.dispatchEvent(new Event('change', { bubbles: true }));
                }
            }""")

            # Đợi ctl19 nạp xong
            logger.info("Chờ nạp dropdown 'Tới mã hàng' (ctl19)...")
            for _ in range(45):
                time.sleep(1)
                ctl19_ready = frame.evaluate("""() => {
                    const dd = $get('ReportViewerControl_ctl04_ctl19_ddValue');
                    return dd && !dd.disabled && dd.options.length > 1;
                }""")
                if ctl19_ready:
                    break

            # Chọn mã cuối cùng
            frame.evaluate("""() => {
                const dd = $get('ReportViewerControl_ctl04_ctl19_ddValue');
                dd.selectedIndex = dd.options.length - 1;
            }""")
            time.sleep(1)

            # 5. Bấm 'Xem báo cáo'
            logger.info("Clicking 'Xem báo cáo' button...")
            frame.click("#ReportViewerControl_ctl04_ctl00")

            # 6. Chờ SSRS render
            logger.info("Chờ SSRS kết xuất báo cáo R014-03 TTVH...")
            start_wait = time.time()
            ready = False
            for _ in range(90):
                time.sleep(2)
                ok = frame.evaluate("""() => {
                    const exp = $get('ReportViewerControl_ctl05_ctl04_ctl00_ButtonLink') || $get('ReportViewerControl_ctl05_ctl04_ctl00_Button');
                    const asyncWait = $get('ReportViewerControl_AsyncWait');
                    const isWait = asyncWait ? (asyncWait.style.display !== 'none') : false;
                    const isDis = exp ? (exp.getAttribute('disabled') === 'disabled' || exp.disabled || exp.classList.contains('Disabled')) : true;
                    return !isWait && !isDis;
                }""")
                if ok:
                    ready = True
                    break

            if not ready:
                raise Exception("Báo cáo R014-03 TTVH không kết xuất xong trong thời gian cho phép (180s)!")

            logger.info(f"Kết xuất hoàn tất trong {int(time.time() - start_wait)} giây. Bắt đầu tải Excel...")

            # 7. Tải file Excel
            btn_sel = "#ReportViewerControl_ctl05_ctl04_ctl00_ButtonLink" if frame.query_selector("#ReportViewerControl_ctl05_ctl04_ctl00_ButtonLink") else "#ReportViewerControl_ctl05_ctl04_ctl00_Button"
            frame.click(btn_sel)
            time.sleep(1)
            excel_link = frame.locator('#ReportViewerControl_ctl05_ctl04_ctl00_Menu a[title="Excel"], a[title="Excel"], a:has-text("Excel")').first
            with page.expect_download(timeout=120000) as dl_info:
                excel_link.click()

            download = dl_info.value
            download.save_as(file_saved_path)
            size_kb = round(os.path.getsize(file_saved_path) / 1024, 2)
            logger.info(f"Đã tải về file Excel: {file_saved_path} ({size_kb} KB)")

            # 8. Gửi file qua Telegram
            caption = (
                f"📊 <b>R014-03. Sổ Chi Tiết Vật Tư</b>\n"
                f"🏢 <b>Đơn vị:</b> 01 THỊNH THẾ VINH HOA\n"
                f"📅 <b>Kỳ báo cáo:</b> {date_str_dmy}\n"
                f"🕒 <b>Xuất lúc:</b> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"📁 <b>Tệp:</b> <code>{target_filename}</code> ({size_kb} KB)"
            )
            send_telegram_document(file_saved_path, caption=caption)
            return True

        except Exception as e:
            logger.error(f"Lỗi xuất báo cáo R014-03 TTVH: {e}")
            send_telegram_message(f"❌ <b>[R014-03 TTVH Report Error]:</b> {str(e)}")
            return False

        finally:
            browser.close()


if __name__ == "__main__":
    run_r014_ttvh_report()

