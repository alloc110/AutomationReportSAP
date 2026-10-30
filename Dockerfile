# Sử dụng image chính thức của Microsoft Playwright Python
FROM mcr.microsoft.com/playwright/python:v1.48.0-jammy

# Thiết lập thư mục làm việc
WORKDIR /app

# Đặt biến môi trường ngăn Python tạo tệp .pyc và bật unbuffered log
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV TZ=Asia/Ho_Chi_Minh

# Cài đặt múi giờ Việt Nam
RUN ln -snf /usr/share/zoneinfo/$TZ /etc/localtime && echo $TZ > /etc/timezone

# Cài đặt dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy toàn bộ mã nguồn vào container
COPY . .

# Tạo thư mục cho downloads và screenshots
RUN mkdir -p downloads screenshots

# Mặc định khởi chạy scheduler để chạy theo lịch trình
CMD ["python", "scheduler.py"]
