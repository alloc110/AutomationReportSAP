import os
import sys
import time
import logging
import argparse
import requests
import urllib.parse
from datetime import datetime, timedelta
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

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
logger = logging.getLogger("AutoReportSAP")

# Cấu hình từ .env
SAP_USERNAME = os.getenv("SAP_USERNAME", "")
SAP_PASSWORD = os.getenv("SAP_PASSWORD", "")
HEADLESS = os.getenv("HEADLESS", "true").lower() in ("true", "1", "yes")
DOWNLOAD_DIR = os.getenv("DOWNLOAD_DIR", "downloads")
SCHEDULE_TIME = os.getenv("RUN_AT_TIME", "08:00")
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")

os.makedirs(DOWNLOAD_DIR, exist_ok=True)


# ==============================================================================
# 1. MODULE TELEGRAM BOT
# ==============================================================================
def send_telegram_message(message: str) -> bool:
    """Gửi tin nhắn thông báo lên nhóm Telegram"""
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        logger.warning("Chưa cấu hình TELEGRAM_BOT_TOKEN hoặc TELEGRAM_CHAT_ID trong .env")
        return False

    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    payload = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": message,
        "parse_mode": "HTML"
    }

    try:
        response = requests.post(url, json=payload, timeout=30)
        response.raise_for_status()
        logger.info("Đã gửi thông báo Telegram thành công.")
        return True
    except Exception as e:
        logger.error(f"Lỗi gửi thông báo Telegram: {e}")
        return False


def send_telegram_document(file_path: str, caption: str = "") -> bool:
    """Gửi tệp tài liệu (Excel) lên nhóm Telegram"""
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        logger.warning("Chưa cấu hình TELEGRAM_BOT_TOKEN hoặc TELEGRAM_CHAT_ID trong .env")
        return False

    if not os.path.isfile(file_path):
        logger.error(f"Không tìm thấy tệp để gửi: {file_path}")
        return False

    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument"
    data = {
        "chat_id": TELEGRAM_CHAT_ID,
        "caption": caption,
        "parse_mode": "HTML"
    }

    try:
        with open(file_path, "rb") as doc:
            files = {"document": doc}
            response = requests.post(url, data=data, files=files, timeout=180)
            response.raise_for_status()
        logger.info(f"Đã gửi tệp {file_path} qua Telegram thành công.")
        return True
    except Exception as e:
        logger.error(f"Lỗi gửi tệp qua Telegram: {e}")
        return False


def wait_ajax_ready(frm, timeout_ms=45000):
    """Hàm phụ trợ chờ AJAX PostBack của SSRS hoàn tất an toàn"""
    time.sleep(1)
    frm.wait_for_function("""() => {
        const prm = window.Sys && window.Sys.WebForms && window.Sys.WebForms.PageRequestManager && window.Sys.WebForms.PageRequestManager.getInstance();
        return !prm || !prm.get_isInAsyncPostBack();
    }""", timeout=timeout_ms)
    time.sleep(1)


# ==============================================================================
# 2. BÁO CÁO 1: P01. BÁO CÁO TỔNG HỢP MUA HÀNG THEO PR
# ==============================================================================
def run_p01_report(target_date: datetime = None):
    """
    Xuất báo cáo 'P01. Báo cáo tổng hợp mua hàng theo PR'
    Đường dẫn: 02 KIM BẢO MÃN ĐƯỜNG / 02. Mua hàng
    """
    if target_date is None:
        target_date = datetime.now() - timedelta(days=1)

    url = "http://report.maycha.com.vn/reports/report/02%20KIM%20B%E1%BA%A2O%20M%C3%83N%20%C4%90%C6%AF%E1%BB%9CNG/02.%20Mua%20h%C3%A0ng/P01.%20B%C3%A1o%20c%C3%A1o%20t%E1%BB%95ng%20h%E1%BB%A3p%20mua%20h%C3%A0ng%20theo%20PR"
    date_str_mdy = target_date.strftime("%m/%d/%Y")
    date_str_dmy = target_date.strftime("%d/%m/%Y")
    target_filename = f"P01_BaoCaoTongHopMuaHangTheoPR_{target_date.strftime('%Y%m%d')}.xlsx"
    file_saved_path = os.path.join(DOWNLOAD_DIR, target_filename)

    logger.info(f"[P01] Đang xử lý cho ngày: {date_str_dmy}")

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
            page.goto(url, wait_until="networkidle", timeout=60000)
            frame = page.wait_for_selector('iframe.viewer', timeout=45000).content_frame()
            frame.wait_for_selector("#ReportViewerControl_ctl04_ctl03_txtValue", timeout=30000)
            time.sleep(1)

            # Điền Từ ngày và Đến ngày
            frame.evaluate(f"""function() {{
                $get('ReportViewerControl_ctl04_ctl03_txtValue').value = '{date_str_mdy}';
                $get('ReportViewerControl_ctl04_ctl05_txtValue').value = '{date_str_mdy}';
            }}""")

            # Kích hoạt postback để nạp Mặt hàng & Cửa hàng
            frame.evaluate("""function() {
                window.setTimeout(function() {
                    __doPostBack('ReportViewerControl$ctl04$ctl05$txtValue', '');
                }, 0);
            }""")

            # Chờ dropdown Mặt hàng sẵn sàng
            frame.wait_for_function("""function() {
                var inp = $get('ReportViewerControl_ctl04_ctl13_txtValue');
                return inp && !inp.disabled;
            }""", timeout=30000)
            time.sleep(1)

            # Chọn (Select All) Mặt hàng
            frame.click("#ReportViewerControl_ctl04_ctl13_ctl01")
            frame.wait_for_selector("#ReportViewerControl_ctl04_ctl13_divDropDown_ctl00", timeout=15000)
            cb_item = frame.locator("#ReportViewerControl_ctl04_ctl13_divDropDown_ctl00")
            if not cb_item.is_checked():
                cb_item.check()
            time.sleep(1)
            frame.click("#ReportViewerControl_ctl04_ctl13_ctl01")
            time.sleep(1)

            # Chọn (Select All) Cửa hàng
            frame.click("#ReportViewerControl_ctl04_ctl17_ctl01")
            frame.wait_for_selector("#ReportViewerControl_ctl04_ctl17_divDropDown_ctl00", timeout=15000)
            cb_store = frame.locator("#ReportViewerControl_ctl04_ctl17_divDropDown_ctl00")
            if not cb_store.is_checked():
                cb_store.check()
            time.sleep(1)
            frame.click("#ReportViewerControl_ctl04_ctl17_ctl01")
            time.sleep(1)

            # Bấm Xem báo cáo
            frame.click("#ReportViewerControl_ctl04_ctl00")

            # Chờ kết xuất
            start_w = time.time()
            while time.time() - start_w < 120:
                time.sleep(2)
                res = frame.evaluate("""function() {
                    var btn = $get('ReportViewerControl_ctl05_ctl04_ctl00_ButtonLink');
                    return btn ? (btn.offsetWidth > 0 && btn.offsetHeight > 0) : false;
                }""")
                if res:
                    break

            # Tải Excel
            frame.click("#ReportViewerControl_ctl05_ctl04_ctl00_ButtonLink")
            time.sleep(1)
            excel_link = frame.locator('#ReportViewerControl_ctl05_ctl04_ctl00_Menu a[title="Excel"], a:has-text("Excel")').first
            with page.expect_download(timeout=120000) as dl_info:
                excel_link.click()

            download = dl_info.value
            download.save_as(file_saved_path)
            size_kb = round(os.path.getsize(file_saved_path) / 1024, 2)
            logger.info(f"[P01] Đã tải về file Excel: {file_saved_path} ({size_kb} KB)")

            # Gửi Telegram
            caption = (
                f"📊 <b>P01. Báo cáo tổng hợp mua hàng theo PR</b>\n"
                f"🏢 <b>Đơn vị:</b> 02 KIM BẢO MÃN ĐƯỜNG\n"
                f"📅 <b>Kỳ báo cáo:</b> {date_str_dmy}\n"
                f"🕒 <b>Xuất lúc:</b> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"📁 <b>Tệp:</b> <code>{target_filename}</code> ({size_kb} KB)"
            )
            success = send_telegram_document(file_saved_path, caption=caption)
            if success:
                try:
                    os.remove(file_saved_path)
                except Exception:
                    pass
            return True

        except Exception as e:
            logger.error(f"[P01] Lỗi xuất báo cáo: {e}")
            send_telegram_message(f"❌ <b>[P01 Report Error]:</b> {str(e)}")
            raise e
        finally:
            context.close()
            browser.close()


# ==============================================================================
# 3. BÁO CÁO 2: COSTING. BÁO CÁO XUẤT NHẬP TỒN
# ==============================================================================
def run_costing_report(month: int = None, year: int = None):
    """
    Xuất báo cáo 'Costing. Báo Cáo Xuất Nhập Tồn'
    Đường dẫn: 02 KIM BẢO MÃN ĐƯỜNG / 03. COSTING
    """
    now = datetime.now()
    if month is None:
        month = now.month
    if year is None:
        year = now.year

    url = "http://report.maycha.com.vn/reports/report/02%20KIM%20B%E1%BA%A2O%20M%C3%83N%20%C4%90%C6%AF%E1%BB%9CNG/03.%20COSTING/Costing.%20B%C3%A1o%20C%C3%A1o%20Xu%E1%BA%A5t%20Nh%E1%BA%ADp%20T%E1%BB%93n"
    target_filename = f"Costing_BaoCaoXuatNhapTon_{year}_{month:02d}.xlsx"
    file_saved_path = os.path.join(DOWNLOAD_DIR, target_filename)

    logger.info(f"[Costing] Đang xử lý cho: Tháng {month}/{year}")

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
            page.goto(url, wait_until="networkidle", timeout=60000)
            frame = page.wait_for_selector('iframe.viewer', timeout=45000).content_frame()
            frame.wait_for_selector("#ReportViewerControl_ctl04", timeout=30000)

            # Chọn Tháng và Năm
            frame.select_option("#ReportViewerControl_ctl04_ctl03_ddValue", str(month))
            frame.select_option("#ReportViewerControl_ctl04_ctl05_ddValue", label=str(year))

            # Mở dropdown Nhóm hàng và chọn tất cả
            frame.click("#ReportViewerControl_ctl04_ctl07_ctl01")
            frame.wait_for_selector("#ReportViewerControl_ctl04_ctl07_divDropDown_ctl00", timeout=15000)
            cb_all = frame.locator("#ReportViewerControl_ctl04_ctl07_divDropDown_ctl00")
            if not cb_all.is_checked():
                cb_all.check()
            frame.click("#ReportViewerControl_ctl04_ctl07_ctl01")

            # Bấm Xem báo cáo
            frame.click("#ReportViewerControl_ctl04_ctl00")

            # Chờ kết xuất
            frame.wait_for_selector("#ReportViewerControl_ctl05_ctl04_ctl00_ButtonLink", timeout=120000)

            # Xuất Excel
            frame.click("#ReportViewerControl_ctl05_ctl04_ctl00_ButtonLink")
            excel_link = frame.locator('#ReportViewerControl_ctl05_ctl04_ctl00_Menu a[title="Excel"], a:has-text("Excel")').first
            with page.expect_download(timeout=120000) as dl_info:
                excel_link.click()

            download = dl_info.value
            download.save_as(file_saved_path)
            size_kb = round(os.path.getsize(file_saved_path) / 1024, 2)
            logger.info(f"[Costing] Đã tải về file Excel: {file_saved_path} ({size_kb} KB)")

            # Gửi Telegram
            caption = (
                f"📊 <b>Costing. Báo Cáo Xuất Nhập Tồn</b>\n"
                f"🏢 <b>Đơn vị:</b> 02 KIM BẢO MÃN ĐƯỜNG\n"
                f"📅 <b>Kỳ báo cáo:</b> Tháng {month}/{year}\n"
                f"🕒 <b>Xuất lúc:</b> {now.strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"📁 <b>Tệp:</b> <code>{target_filename}</code> ({size_kb} KB)"
            )
            success = send_telegram_document(file_saved_path, caption=caption)
            if success:
                try:
                    os.remove(file_saved_path)
                except Exception:
                    pass
            return True

        except Exception as e:
            logger.error(f"[Costing] Lỗi xuất báo cáo: {e}")
            send_telegram_message(f"❌ <b>[Costing Report Error]:</b> {str(e)}")
            raise e
        finally:
            context.close()
            browser.close()


# ==============================================================================
# 4. BÁO CÁO 3: R014-03. SỔ CHI TIẾT VẬT TƯ
# ==============================================================================
def run_r014_report(target_date: datetime = None):
    """
    Xuất báo cáo 'R014-03. Sổ Chi Tiết Vật Tư'
    Đường dẫn: 02 KIM BẢO MÃN ĐƯỜNG / 04. INVENTORY
    """
    if target_date is None:
        target_date = datetime.now() - timedelta(days=1)

    url = "http://report.maycha.com.vn/reports/report/02%20KIM%20B%E1%BA%A2O%20M%C3%83N%20%C4%90%C6%AF%E1%BB%9CNG/04.%20INVENTORY/R014-03.%20S%E1%BB%95%20Chi%20Ti%E1%BA%BFt%20V%E1%BA%ADt%20T%C6%B0"
    date_str_mdy = target_date.strftime("%m/%d/%Y")
    date_str_dmy = target_date.strftime("%d/%m/%Y")
    target_filename = f"R014_03_SoChiTietVatTu_{target_date.strftime('%Y%m%d')}.xlsx"
    file_saved_path = os.path.join(DOWNLOAD_DIR, target_filename)

    logger.info(f"[R014-03] Đang xử lý cho ngày: {date_str_dmy}")

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
            page.goto(url, wait_until="networkidle", timeout=60000)
            frame = page.wait_for_selector("iframe.viewer", timeout=45000).content_frame()
            frame.wait_for_selector("#ReportViewerControl_ctl04", timeout=30000)
            time.sleep(2)

            # Điền Từ ngày và Đến ngày
            frame.evaluate(f"""() => {{
                $get('ReportViewerControl_ctl04_ctl03_txtValue').value = '{date_str_mdy}';
                $get('ReportViewerControl_ctl04_ctl05_txtValue').value = '{date_str_mdy}';
            }}""")

            # Chọn tất cả Kho
            frame.click("#ReportViewerControl_ctl04_ctl07_ctl01")
            frame.wait_for_selector("#ReportViewerControl_ctl04_ctl07_divDropDown_ctl00", timeout=15000)
            cb_kho = frame.locator("#ReportViewerControl_ctl04_ctl07_divDropDown_ctl00")
            if not cb_kho.is_checked():
                cb_kho.check()
            time.sleep(1)
            frame.click("#ReportViewerControl_ctl04_ctl07_ctl01")
            time.sleep(1)

            # Chọn tất cả Nhóm hàng
            frame.click("#ReportViewerControl_ctl04_ctl09_ctl01")
            frame.wait_for_selector("#ReportViewerControl_ctl04_ctl09_divDropDown_ctl00", timeout=15000)
            cb_nhom = frame.locator("#ReportViewerControl_ctl04_ctl09_divDropDown_ctl00")
            if not cb_nhom.is_checked():
                cb_nhom.check()
            time.sleep(1)
            frame.click("#ReportViewerControl_ctl04_ctl09_ctl01")
            time.sleep(1)

            # Chờ nạp Từ mã hàng và chọn mã đầu tiên
            frame.wait_for_function("""() => {
                const el = $get('ReportViewerControl_ctl04_ctl15_ddValue');
                return el && !el.disabled && el.options.length > 1;
            }""", timeout=45000)
            frame.select_option("#ReportViewerControl_ctl04_ctl15_ddValue", index=1)
            wait_ajax_ready(frame)

            # Chờ nạp Tới mã hàng và chọn mã cuối cùng
            frame.wait_for_function("""() => {
                const el = $get('ReportViewerControl_ctl04_ctl19_ddValue');
                return el && !el.disabled && el.options.length > 1;
            }""", timeout=45000)
            count19 = frame.evaluate("""() => $get('ReportViewerControl_ctl04_ctl19_ddValue').options.length""")
            frame.select_option("#ReportViewerControl_ctl04_ctl19_ddValue", index=count19 - 1)
            time.sleep(1)

            # Bấm Xem báo cáo
            frame.click("#ReportViewerControl_ctl04_ctl00")

            # Chờ kết xuất hoàn tất (báo cáo lớn, có thể mất ~50-90 giây)
            logger.info("[R014-03] Đang chờ máy chủ kết xuất (khoảng ~50 giây)...")
            start_wait = time.time()
            ready = False
            while time.time() - start_wait < 180:
                time.sleep(3)
                ok = frame.evaluate("""() => {
                    const exp = $get('ReportViewerControl_ctl05_ctl04_ctl00_ButtonLink');
                    const asyncWait = $get('ReportViewerControl_AsyncWait');
                    const isWait = asyncWait ? (asyncWait.style.display !== 'none') : false;
                    const isDis = exp ? (exp.getAttribute('disabled') === 'disabled' || exp.disabled || exp.classList.contains('Disabled')) : true;
                    return !isWait && !isDis;
                }""")
                if ok:
                    ready = True
                    break

            if not ready:
                raise Exception("Báo cáo R014-03 không kết xuất xong trong thời gian cho phép (180s)!")

            logger.info(f"[R014-03] Kết xuất hoàn tất trong {int(time.time() - start_wait)} giây.")

            # Tải Excel
            frame.click("#ReportViewerControl_ctl05_ctl04_ctl00_ButtonLink")
            time.sleep(1)
            excel_link = frame.locator('#ReportViewerControl_ctl05_ctl04_ctl00_Menu a[title="Excel"], a:has-text("Excel")').first
            with page.expect_download(timeout=120000) as dl_info:
                excel_link.click()

            download = dl_info.value
            download.save_as(file_saved_path)
            size_kb = round(os.path.getsize(file_saved_path) / 1024, 2)
            logger.info(f"[R014-03] Đã tải về file Excel: {file_saved_path} ({size_kb} KB)")

            # Gửi Telegram
            caption = (
                f"📊 <b>R014-03. Sổ Chi Tiết Vật Tư</b>\n"
                f"🏢 <b>Đơn vị:</b> 02 KIM BẢO MÃN ĐƯỜNG\n"
                f"📅 <b>Kỳ báo cáo:</b> {date_str_dmy}\n"
                f"🕒 <b>Xuất lúc:</b> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"📁 <b>Tệp:</b> <code>{target_filename}</code> ({size_kb} KB)"
            )
            success = send_telegram_document(file_saved_path, caption=caption)
            if success:
                try:
                    os.remove(file_saved_path)
                except Exception:
                    pass
            return True

        except Exception as e:
            logger.error(f"[R014-03] Lỗi xuất báo cáo: {e}")
            send_telegram_message(f"❌ <b>[R014-03 Report Error]:</b> {str(e)}")
            raise e
        finally:
            context.close()
            browser.close()


# ==============================================================================
# 5. BÁO CÁO 4: P01. BÁO CÁO TỔNG HỢP YÊU CẦU HÀNG HÓA THEO PR (SCM)
# ==============================================================================
def run_p01_scm_report(target_date: datetime = None):
    """
    Xuất báo cáo 'P01. Báo cáo tổng hợp yêu cầu hàng hóa theo Purchase Request'
    Đường dẫn: 01 THỊNH THẾ VINH HOA / 02. Mua hàng (SCM)
    """
    if target_date is None:
        target_date = datetime.now() - timedelta(days=1)

    url = "http://report.maycha.com.vn/reports/report/01%20TH%E1%BB%8ANH%20TH%E1%BA%BE%20VINH%20HOA/02.%20Mua%20h%C3%A0ng%20(SCM)/P01.%20B%C3%A1o%20c%C3%A1o%20t%E1%BB%95ng%20h%E1%BB%A3p%20y%C3%AAu%20c%E1%BA%A7u%20h%C3%A0ng%20h%C3%B3a%20theo%20Purchase%20Request"
    date_str_mdy = target_date.strftime("%m/%d/%Y")
    date_str_dmy = target_date.strftime("%d/%m/%Y")
    target_filename = f"P01_TongHopYeuCauHangHoa_SCM_{target_date.strftime('%Y%m%d')}.xlsx"
    file_saved_path = os.path.join(DOWNLOAD_DIR, target_filename)

    logger.info(f"[P01 SCM] Đang xử lý cho ngày: {date_str_dmy}")

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
            page.goto(url, wait_until="networkidle", timeout=60000)
            frame = page.wait_for_selector("iframe.viewer", timeout=45000).content_frame()
            frame.wait_for_selector("#ReportViewerControl_ctl04_ctl03_txtValue", timeout=30000)
            time.sleep(2)

            # Điền ngày và gán NULL cho Ngày yêu cầu
            frame.evaluate(f"""() => {{
                $get('ReportViewerControl_ctl04_ctl03_txtValue').value = '{date_str_mdy}';
                $get('ReportViewerControl_ctl04_ctl05_txtValue').value = '{date_str_mdy}';
                $get('ReportViewerControl_ctl04_ctl07_cbNull').checked = true;
                $get('ReportViewerControl_ctl04_ctl09_cbNull').checked = true;
                window.setTimeout(function() {{
                    __doPostBack('ReportViewerControl$ctl04$ctl05$txtValue', '');
                }}, 0);
            }}""")

            # Chờ dropdowns phụ thuộc được nạp xong
            time.sleep(5)
            frame.wait_for_function("""() => {
                const div13 = $get('ReportViewerControl_ctl04_ctl13_divDropDown');
                const div17 = $get('ReportViewerControl_ctl04_ctl17_divDropDown');
                return div13 && div17;
            }""", timeout=30000)

            # Chọn tất cả Mặt hàng (ctl13)
            frame.click("#ReportViewerControl_ctl04_ctl13_ctl01")
            frame.wait_for_selector("#ReportViewerControl_ctl04_ctl13_divDropDown_ctl00", timeout=15000)
            cb_item = frame.locator("#ReportViewerControl_ctl04_ctl13_divDropDown_ctl00")
            if not cb_item.is_checked():
                cb_item.check()
            time.sleep(1)
            frame.click("#ReportViewerControl_ctl04_ctl13_ctl01")
            time.sleep(1)

            # Chọn tất cả Cửa hàng (ctl17)
            frame.click("#ReportViewerControl_ctl04_ctl17_ctl01")
            frame.wait_for_selector("#ReportViewerControl_ctl04_ctl17_divDropDown_ctl00", timeout=15000)
            cb_store = frame.locator("#ReportViewerControl_ctl04_ctl17_divDropDown_ctl00")
            if not cb_store.is_checked():
                cb_store.check()
            time.sleep(1)
            frame.click("#ReportViewerControl_ctl04_ctl17_ctl01")
            time.sleep(1)

            # Bấm Xem báo cáo
            frame.click("#ReportViewerControl_ctl04_ctl00")

            # Chờ render
            start_wait = time.time()
            ready = False
            while time.time() - start_wait < 120:
                time.sleep(2)
                ok = frame.evaluate("""() => {
                    const exp = $get('ReportViewerControl_ctl05_ctl04_ctl00_ButtonLink');
                    const asyncWait = $get('ReportViewerControl_AsyncWait');
                    const isWait = asyncWait ? (asyncWait.style.display !== 'none') : false;
                    const isDis = exp ? (exp.getAttribute('disabled') === 'disabled' || exp.disabled || exp.classList.contains('Disabled')) : true;
                    return !isWait && !isDis;
                }""")
                if ok:
                    ready = True
                    break

            if not ready:
                raise Exception("Báo cáo P01 SCM không kết xuất xong trong thời gian cho phép (120s)!")

            # Tải Excel
            frame.click("#ReportViewerControl_ctl05_ctl04_ctl00_ButtonLink")
            time.sleep(1)
            excel_link = frame.locator('#ReportViewerControl_ctl05_ctl04_ctl00_Menu a[title="Excel"], a:has-text("Excel")').first
            with page.expect_download(timeout=120000) as dl_info:
                excel_link.click()

            download = dl_info.value
            download.save_as(file_saved_path)
            size_kb = round(os.path.getsize(file_saved_path) / 1024, 2)
            logger.info(f"[P01 SCM] Đã tải về file Excel: {file_saved_path} ({size_kb} KB)")

            # Gửi Telegram
            caption = (
                f"📊 <b>P01. Báo cáo tổng hợp yêu cầu hàng hóa theo Purchase Request</b>\n"
                f"🏢 <b>Đơn vị:</b> 01 THỊNH THẾ VINH HOA (SCM)\n"
                f"📅 <b>Kỳ báo cáo:</b> {date_str_dmy}\n"
                f"🕒 <b>Xuất lúc:</b> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"📁 <b>Tệp:</b> <code>{target_filename}</code> ({size_kb} KB)"
            )
            success = send_telegram_document(file_saved_path, caption=caption)
            if success:
                try:
                    os.remove(file_saved_path)
                except Exception:
                    pass
            return True

        except Exception as e:
            logger.error(f"[P01 SCM] Lỗi xuất báo cáo: {e}")
            send_telegram_message(f"❌ <b>[P01 SCM Report Error]:</b> {str(e)}")
            raise e
        finally:
            context.close()
            browser.close()


# ==============================================================================
# 6. BÁO CÁO 5: R502. BÁO CÁO CHI TIẾT LỆNH SẢN XUẤT (KBMD)
# ==============================================================================
def run_r502_report(target_date: datetime = None):
    """
    Xuất báo cáo 'R502. Báo cáo chi tiết lệnh sản xuất'
    Đường dẫn: 02 KIM BẢO MÃN ĐƯỜNG / 03. COSTING / R502. Báo cáo chi tiết lệnh sản xuất
    """
    if target_date is None:
        target_date = datetime.now() - timedelta(days=1)

    url = "http://report.maycha.com.vn/reports/report/02%20KIM%20B%E1%BA%A2O%20M%C3%83N%20%C4%90%C6%AF%E1%BB%9CNG/03.%20COSTING/R502.%20B%C3%A1o%20c%C3%A1o%20chi%20ti%E1%BA%BFt%20l%E1%BB%87nh%20s%E1%BA%A3n%20xu%E1%BA%A5t"
    date_str_mdy = f"{target_date.month}/{target_date.day}/{target_date.year}"
    date_str_dmy = target_date.strftime("%d/%m/%Y")
    target_filename = f"R502_ChiTietLenhSanXuat_KBMD_{target_date.strftime('%Y%m%d')}.xlsx"
    file_saved_path = os.path.join(DOWNLOAD_DIR, target_filename)

    logger.info(f"\n[R502] Bắt đầu xử lý cho ngày: {date_str_dmy} (MDY: {date_str_mdy})")

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=HEADLESS,
            args=["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"]
        )
        context = browser.new_context(
            http_credentials={"username": SAP_USERNAME, "password": SAP_PASSWORD},
            viewport={"width": 1600, "height": 900},
            accept_downloads=True
        )
        page = context.new_page()

        try:
            logger.info(f"[R502] Đang truy cập URL: {url}")
            page.goto(url, wait_until="load", timeout=60000)

            frame_element = page.wait_for_selector("iframe.viewer", timeout=60000)
            frame = frame_element.content_frame()
            if not frame:
                raise Exception("Không tìm thấy iframe báo cáo!")

            # Chờ bảng tham số
            frame.wait_for_selector("#ReportViewerControl_ctl04_ctl03_txtValue", timeout=30000)
            time.sleep(1)

            # Điền Từ ngày và Đến ngày
            logger.info(f"[R502] Điền Từ ngày & Đến ngày: {date_str_mdy}")
            frame.fill("#ReportViewerControl_ctl04_ctl03_txtValue", date_str_mdy)
            frame.fill("#ReportViewerControl_ctl04_ctl05_txtValue", date_str_mdy)
            time.sleep(1)

            # Bấm 'Xem báo cáo'
            logger.info("[R502] Nhấn nút 'Xem báo cáo'...")
            frame.click("#ReportViewerControl_ctl04_ctl00")

            # Chờ kết xuất
            logger.info("[R502] Đang chờ SSRS xử lý dữ liệu báo cáo...")
            start_wait = time.time()
            ready = False
            for _ in range(60):
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
                raise Exception("Báo cáo R502 không kết xuất xong trong thời gian cho phép (120s)!")

            logger.info(f"[R502] Kết xuất hoàn tất trong {int(time.time() - start_wait)} giây.")

            # Tải Excel
            btn_selector = "#ReportViewerControl_ctl05_ctl04_ctl00_ButtonLink" if frame.query_selector("#ReportViewerControl_ctl05_ctl04_ctl00_ButtonLink") else "#ReportViewerControl_ctl05_ctl04_ctl00_Button"
            frame.click(btn_selector)
            time.sleep(1)
            excel_link = frame.locator('#ReportViewerControl_ctl05_ctl04_ctl00_Menu a[title="Excel"], a[title="Excel"], a:has-text("Excel")').first
            with page.expect_download(timeout=120000) as dl_info:
                excel_link.click()

            download = dl_info.value
            download.save_as(file_saved_path)
            size_kb = round(os.path.getsize(file_saved_path) / 1024, 2)
            logger.info(f"[R502] Đã tải về file Excel: {file_saved_path} ({size_kb} KB)")

            # Gửi Telegram
            caption = (
                f"📊 <b>R502. Báo cáo chi tiết lệnh sản xuất</b>\n"
                f"🏢 <b>Đơn vị:</b> 02 KIM BẢO MÃN ĐƯỜNG\n"
                f"📅 <b>Kỳ báo cáo:</b> {date_str_dmy}\n"
                f"🕒 <b>Xuất lúc:</b> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"📁 <b>Tệp:</b> <code>{target_filename}</code> ({size_kb} KB)"
            )
            success = send_telegram_document(file_saved_path, caption=caption)
            if success:
                try:
                    os.remove(file_saved_path)
                except Exception:
                    pass
            return True

        except Exception as e:
            logger.error(f"[R502] Lỗi xuất báo cáo: {e}")
            send_telegram_message(f"❌ <b>[R502 Report Error]:</b> {str(e)}")
            raise e
        finally:
            context.close()
            browser.close()


# ==============================================================================
# 7. BÁO CÁO 6: R014-03. SỔ CHI TIẾT VẬT TƯ (THỊNH THẾ VINH HOA)
# ==============================================================================
def run_r014_ttvh_report(target_date: datetime = None):
    """
    Xuất báo cáo 'R014-03. Sổ Chi Tiết Vật Tư' của 01 THỊNH THẾ VINH HOA
    Đường dẫn: 01 THỊNH THẾ VINH HOA / 04. Inventory / R014-03. Sổ Chi Tiết Vật Tư
    """
    if target_date is None:
        target_date = datetime.now() - timedelta(days=1)

    url = "http://report.maycha.com.vn/reports/report/01%20TH%E1%BB%8ANH%20TH%E1%BA%BE%20VINH%20HOA/04.%20Inventory/R014-03.%20S%E1%BB%95%20Chi%20Ti%E1%BA%BFt%20V%E1%BA%ADt%20T%C6%B0"
    date_str_mdy = target_date.strftime("%m/%d/%Y")
    date_str_dmy = target_date.strftime("%d/%m/%Y")
    target_filename = f"R014_03_SoChiTietVatTu_TTVH_{target_date.strftime('%Y%m%d')}.xlsx"
    file_saved_path = os.path.join(DOWNLOAD_DIR, target_filename)

    logger.info(f"\n[R014-03 TTVH] Bắt đầu xử lý cho ngày: {date_str_dmy} (MDY: {date_str_mdy})")

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
            logger.info(f"[R014-03 TTVH] Đang truy cập URL: {url}")
            page.goto(url, wait_until="networkidle", timeout=60000)

            iframe_elem = page.wait_for_selector("iframe.viewer", timeout=45000)
            frame = iframe_elem.content_frame()
            frame.wait_for_selector("#ReportViewerControl_ctl04", timeout=30000)
            time.sleep(2)

            # 1. Điền Từ ngày & Đến ngày
            logger.info(f"[R014-03 TTVH] Điền Từ ngày & Đến ngày: {date_str_mdy}")
            frame.evaluate(f"""() => {{
                $get('ReportViewerControl_ctl04_ctl03_txtValue').value = '{date_str_mdy}';
                $get('ReportViewerControl_ctl04_ctl05_txtValue').value = '{date_str_mdy}';
            }}""")

            # 2. Chọn tất cả Kho (ctl07)
            logger.info("[R014-03 TTVH] Mở dropdown Kho và chọn (Select All)...")
            frame.click("#ReportViewerControl_ctl04_ctl07_ctl01")
            frame.wait_for_selector("#ReportViewerControl_ctl04_ctl07_divDropDown_ctl00", timeout=15000)
            kho_all = frame.locator("#ReportViewerControl_ctl04_ctl07_divDropDown_ctl00")
            if not kho_all.is_checked():
                kho_all.check()
            time.sleep(1)
            frame.click("#ReportViewerControl_ctl04_ctl07_ctl01")
            time.sleep(1)

            # 3. Chọn tất cả Nhóm hàng (ctl09)
            logger.info("[R014-03 TTVH] Mở dropdown Nhóm hàng và chọn (Select All)...")
            frame.click("#ReportViewerControl_ctl04_ctl09_ctl01")
            frame.wait_for_selector("#ReportViewerControl_ctl04_ctl09_divDropDown_ctl00", timeout=15000)
            nhom_all = frame.locator("#ReportViewerControl_ctl04_ctl09_divDropDown_ctl00")
            if not nhom_all.is_checked():
                nhom_all.check()
            time.sleep(1)
            frame.click("#ReportViewerControl_ctl04_ctl09_ctl01")
            time.sleep(1)

            # 4. Kích hoạt PostBack để nạp danh mục vật tư
            logger.info("[R014-03 TTVH] Kích hoạt nạp danh mục vật tư...")
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

            logger.info("[R014-03 TTVH] Chọn mã hàng đầu tiên trong 'Từ mã hàng' (ctl15)...")
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
            logger.info("[R014-03 TTVH] Chờ nạp dropdown 'Tới mã hàng' (ctl19)...")
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
            logger.info("[R014-03 TTVH] Nhấn nút 'Xem báo cáo'...")
            frame.click("#ReportViewerControl_ctl04_ctl00")

            # 6. Chờ SSRS render
            logger.info("[R014-03 TTVH] Chờ SSRS kết xuất báo cáo...")
            start_wait = time.time()
            ready = False
            for _ in range(120):
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
                raise Exception("Báo cáo R014-03 TTVH không kết xuất xong trong thời gian cho phép (240s)!")

            logger.info(f"[R014-03 TTVH] Kết xuất hoàn tất trong {int(time.time() - start_wait)} giây. Bắt đầu tải Excel...")

            # 7. Tải file Excel
            btn_sel = "#ReportViewerControl_ctl05_ctl04_ctl00_ButtonLink" if frame.query_selector("#ReportViewerControl_ctl05_ctl04_ctl00_ButtonLink") else "#ReportViewerControl_ctl05_ctl04_ctl00_Button"
            frame.click(btn_sel)
            time.sleep(1)
            excel_link = frame.locator('#ReportViewerControl_ctl05_ctl04_ctl00_Menu a[title="Excel"], a[title="Excel"], a:has-text("Excel")').first
            with page.expect_download(timeout=180000) as dl_info:
                excel_link.click()

            download = dl_info.value
            download.save_as(file_saved_path)
            size_kb = round(os.path.getsize(file_saved_path) / 1024, 2)
            logger.info(f"[R014-03 TTVH] Đã tải về file Excel: {file_saved_path} ({size_kb} KB)")

            # 8. Gửi file qua Telegram
            caption = (
                f"📊 <b>R014-03. Sổ Chi Tiết Vật Tư</b>\n"
                f"🏢 <b>Đơn vị:</b> 01 THỊNH THẾ VINH HOA\n"
                f"📅 <b>Kỳ báo cáo:</b> {date_str_dmy}\n"
                f"🕒 <b>Xuất lúc:</b> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"📁 <b>Tệp:</b> <code>{target_filename}</code> ({size_kb} KB)"
            )
            success = send_telegram_document(file_saved_path, caption=caption)
            if success:
                try:
                    os.remove(file_saved_path)
                except Exception:
                    pass
            return True

        except Exception as e:
            logger.error(f"[R014-03 TTVH] Lỗi xuất báo cáo: {e}")
            send_telegram_message(f"❌ <b>[R014-03 TTVH Report Error]:</b> {str(e)}")
            raise e
        finally:
            context.close()
            browser.close()


# ==============================================================================
# 8. ĐIỀU PHỐI CHẠY TẤT CẢ CÁC BÁO CÁO
# ==============================================================================
def run_all_reports():
    """Chạy tuần tự toàn bộ các báo cáo và gửi tổng kết qua Telegram"""
    start_time = datetime.now()
    logger.info("=" * 65)
    logger.info(f"BẮT ĐẦU TIẾN TRÌNH XUẤT BÁO CÁO SAP - {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info("=" * 65)

    results = []

    # 1. P01 Mua hàng theo PR (02 Kim Bảo Mãn Đường)
    logger.info("\n>>> [1/6] Đang xử lý: P01. Báo cáo tổng hợp mua hàng theo PR (KBMD)...")
    try:
        run_p01_report()
        results.append(("P01 Mua hàng theo PR (KBMD)", "Thành công ✅"))
    except Exception as e:
        results.append(("P01 Mua hàng theo PR (KBMD)", f"Thất bại ❌ ({e})"))

    time.sleep(3)

    # 2. P01 SCM (01 Thịnh Thế Vinh Hoa)
    logger.info("\n>>> [2/6] Đang xử lý: P01. Yêu cầu hàng hóa theo PR (TTVH SCM)...")
    try:
        run_p01_scm_report()
        results.append(("P01 Yêu cầu hàng hóa (TTVH SCM)", "Thành công ✅"))
    except Exception as e:
        results.append(("P01 Yêu cầu hàng hóa (TTVH SCM)", f"Thất bại ❌ ({e})"))

    time.sleep(3)

    # 3. Costing Xuất Nhập Tồn (02 Kim Bảo Mãn Đường)
    logger.info("\n>>> [3/6] Đang xử lý: Costing. Báo Cáo Xuất Nhập Tồn (KBMD)...")
    try:
        run_costing_report()
        results.append(("Costing. Báo Cáo Xuất Nhập Tồn (KBMD)", "Thành công ✅"))
    except Exception as e:
        results.append(("Costing. Báo Cáo Xuất Nhập Tồn (KBMD)", f"Thất bại ❌ ({e})"))

    time.sleep(3)

    # 4. R014-03 Sổ Chi Tiết Vật Tư (02 Kim Bảo Mãn Đường)
    logger.info("\n>>> [4/6] Đang xử lý: R014-03. Sổ Chi Tiết Vật Tư (KBMD)...")
    try:
        run_r014_report()
        results.append(("R014-03. Sổ Chi Tiết Vật Tư (KBMD)", "Thành công ✅"))
    except Exception as e:
        results.append(("R014-03. Sổ Chi Tiết Vật Tư (KBMD)", f"Thất bại ❌ ({e})"))

    time.sleep(3)

    # 5. R502 Báo cáo chi tiết lệnh sản xuất (02 Kim Bảo Mãn Đường)
    logger.info("\n>>> [5/6] Đang xử lý: R502. Báo cáo chi tiết lệnh sản xuất (KBMD)...")
    try:
        run_r502_report()
        results.append(("R502. Báo cáo chi tiết lệnh sản xuất (KBMD)", "Thành công ✅"))
    except Exception as e:
        results.append(("R502. Báo cáo chi tiết lệnh sản xuất (KBMD)", f"Thất bại ❌ ({e})"))

    time.sleep(3)

    # 6. R014-03 Sổ Chi Tiết Vật Tư (01 Thịnh Thế Vinh Hoa)
    logger.info("\n>>> [6/6] Đang xử lý: R014-03. Sổ Chi Tiết Vật Tư (TTVH)...")
    try:
        run_r014_ttvh_report()
        results.append(("R014-03. Sổ Chi Tiết Vật Tư (TTVH)", "Thành công ✅"))
    except Exception as e:
        results.append(("R014-03. Sổ Chi Tiết Vật Tư (TTVH)", f"Thất bại ❌ ({e})"))

    # Tổng kết
    elapsed = round((datetime.now() - start_time).total_seconds(), 1)
    logger.info("\n" + "=" * 65)
    logger.info(f"KẾT QUẢ TIẾN TRÌNH XUẤT BÁO CÁO (Tổng thời gian: {elapsed}s):")
    for name, status in results:
        logger.info(f" - {name}: {status}")
    logger.info("=" * 65)

    # Gửi thông báo tổng kết lên Telegram
    summary_msg = (
        f"📋 <b>TỔNG KẾT TIẾN TRÌNH XUẤT CÁC BÁO CÁO SAP MAYCHA</b>\n"
        f"🕒 <b>Tổng thời gian:</b> {elapsed} giây\n"
        f"📅 <b>Thời điểm xuất:</b> {start_time.strftime('%d/%m/%Y %H:%M:%S')}\n\n"
    )
    for name, status in results:
        summary_msg += f"• <b>{name}:</b> {status}\n"

    send_telegram_message(summary_msg)


def start_scheduler():
    """Chạy dịch vụ lên lịch hàng ngày trong cùng 1 file"""
    import schedule

    logger.info(f"Khởi động dịch vụ Lên lịch tự động. Chạy định kỳ lúc: {SCHEDULE_TIME} hàng ngày.")
    send_telegram_message(f"⏰ <b>[Scheduler]</b> Dịch vụ tự động xuất các báo cáo SAP đã khởi động. Lịch chạy: <b>{SCHEDULE_TIME}</b> mỗi ngày.")

    schedule.every().day.at(SCHEDULE_TIME).do(run_all_reports)

    try:
        while True:
            schedule.run_pending()
            time.sleep(30)
    except (KeyboardInterrupt, SystemExit):
        logger.info("Dịch vụ lên lịch đã dừng.")


def main():
    parser = argparse.ArgumentParser(description="Chương trình tự động xuất các báo cáo SAP Maycha và gửi qua Telegram")
    parser.add_argument("--schedule", "-s", action="store_true", help="Chạy chế độ lập lịch định kỳ (Scheduler) theo giờ cấu hình trong .env")
    parser.add_argument("--p01", action="store_true", help="Chỉ chạy riêng báo cáo P01 (Kim Bảo Mãn Đường)")
    parser.add_argument("--p01-scm", action="store_true", help="Chỉ chạy riêng báo cáo P01 SCM (Thịnh Thế Vinh Hoa)")
    parser.add_argument("--costing", action="store_true", help="Chỉ chạy riêng báo cáo Costing")
    parser.add_argument("--r014", action="store_true", help="Chỉ chạy riêng báo cáo R014-03 (Kim Bảo Mãn Đường)")
    parser.add_argument("--r014-ttvh", action="store_true", help="Chỉ chạy riêng báo cáo R014-03 (Thịnh Thế Vinh Hoa)")
    parser.add_argument("--r502", action="store_true", help="Chỉ chạy riêng báo cáo R502 (Chi tiết lệnh sản xuất)")

    args = parser.parse_args()

    if args.schedule:
        start_scheduler()
    elif args.p01:
        run_p01_report()
    elif getattr(args, 'p01_scm', False):
        run_p01_scm_report()
    elif args.costing:
        run_costing_report()
    elif args.r014:
        run_r014_report()
    elif getattr(args, 'r014_ttvh', False):
        run_r014_ttvh_report()
    elif args.r502:
        run_r502_report()
    else:
        # Mặc định: Chạy toàn bộ các báo cáo tích hợp
        run_all_reports()


if __name__ == "__main__":
    main()
