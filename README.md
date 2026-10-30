# SAP SSRS Report Automation System (Maycha)

Hệ thống tự động hóa đăng nhập cổng báo cáo SAP SSRS của Maycha, kết xuất các báo cáo tài chính/kho/mua hàng sang định dạng Excel và gửi tự động về nhóm Telegram.

---

## 📌 Các báo cáo đang được tự động hóa

1. **P01. Báo cáo tổng hợp mua hàng theo PR**
   - Vị trí: `02 KIM BẢO MÃN ĐƯỜNG / 02. Mua hàng`
   - Kỳ báo cáo: Ngày hôm qua (tự động tính).
   - Tự động nạp danh sách và chọn tất cả các mặt hàng, cửa hàng.

2. **P01. Báo cáo tổng hợp yêu cầu hàng hóa theo Purchase Request (SCM)**
   - Vị trí: `01 THỊNH THẾ VINH HOA / 02. Mua hàng (SCM)`
   - Kỳ báo cáo: Ngày hôm qua (tự động tính).
   - Tự động xử lý checkbox NULL cho Ngày yêu cầu, nạp và chọn tất cả Mặt hàng, Cửa hàng.

3. **Costing. Báo Cáo Xuất Nhập Tồn**
   - Vị trí: `02 KIM BẢO MÃN ĐƯỜNG / 03. COSTING`
   - Kỳ báo cáo: Tháng & Năm hiện tại (tự động tính).
   - Tự động nạp và chọn tất cả thành phần: Thành phẩm, Bán thành phẩm, Nguyên vật liệu, Bao bì, Hàng hóa, Dịch vụ...

4. **R014-03. Sổ Chi Tiết Vật Tư**
   - Vị trí: `02 KIM BẢO MÃN ĐƯỜNG / 04. INVENTORY`
   - Kỳ báo cáo: Ngày hôm qua (tự động tính).
   - Tự động nạp và chọn: Tất cả Kho, Tất cả Nhóm hàng, Dải mã hàng từ mã đầu tiên đến mã cuối cùng (hơn 920 mã).

5. **R502. Báo cáo chi tiết lệnh sản xuất**
   - Vị trí: `02 KIM BẢO MÃN ĐƯỜNG / 03. COSTING`
   - Kỳ báo cáo: Ngày hôm qua (tự động tính).
   - Tự động điền Từ ngày, Đến ngày và kết xuất file Excel chi tiết LSX.

6. **R014-03. Sổ Chi Tiết Vật Tư (Thịnh Thế Vinh Hoa)**
   - Vị trí: `01 THỊNH THẾ VINH HOA / 04. Inventory`
   - Kỳ báo cáo: Ngày hôm qua (tự động tính).
   - Tự động nạp và chọn: Tất cả Kho, Tất cả Nhóm hàng, Dải hơn 2,340 mã hàng.

---

## 🚀 Hướng dẫn sử dụng (Chỉ cần chạy 1 file duy nhất)

Kích hoạt môi trường conda có Playwright:
```powershell
conda activate playwright
```

### 1. Xuất và gửi toàn bộ các báo cáo ngay lập tức
```powershell
python main.py
```
> *(Chạy xong sẽ tự động gửi từng file báo cáo kèm tin nhắn tổng kết lên Telegram)*

### 2. Chạy tự động theo lịch trình hàng ngày (Scheduler)
```powershell
python main.py --schedule
```
> *(Hệ thống sẽ giữ tiến trình chạy ngầm và tự động kích hoạt xuất toàn bộ báo cáo vào giờ đã cấu hình trong `.env`, mặc định `08:00` hàng ngày)*

### 3. Tùy chọn chạy riêng lẻ từng báo cáo (nếu cần)
- Chỉ chạy báo cáo P01 (Kim Bảo Mãn Đường):
  ```powershell
  python main.py --p01
  ```
- Chỉ chạy báo cáo P01 SCM (Thịnh Thế Vinh Hoa):
  ```powershell
  python main.py --p01-scm
  # Hoặc chạy trực tiếp file module:
  python p01_scm_report.py
  ```
- Chỉ chạy báo cáo Costing (Xuất Nhập Tồn):
  ```powershell
  python main.py --costing
  # Hoặc: python costing_report.py
  ```
- Chỉ chạy báo cáo R014-03 (Sổ Chi Tiết Vật Tư - Kim Bảo Mãn Đường):
  ```powershell
  python main.py --r014
  # Hoặc: python r014_report.py
  ```
- Chỉ chạy báo cáo R014-03 (Sổ Chi Tiết Vật Tư - Thịnh Thế Vinh Hoa):
  ```powershell
  python main.py --r014-ttvh
  # Hoặc: python r014_ttvh_report.py
  ```
- Chỉ chạy báo cáo R502 (Chi Tiết Lệnh Sản Xuất):
  ```powershell
  python main.py --r502
  # Hoặc: python r502_report.py
  ```

---

## ⚙️ Cấu hình file `.env`

```ini
SAP_USERNAME=your_username
SAP_PASSWORD=your_password
HEADLESS=true
TELEGRAM_BOT_TOKEN=your_bot_token
TELEGRAM_CHAT_ID=-5362864260
RUN_AT_TIME=08:00
DOWNLOAD_DIR=downloads
```
