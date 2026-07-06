# Sử dụng môi trường Python 3.11 siêu nhẹ
FROM python:3.11-slim

# Thiết lập thư mục làm việc trong container
WORKDIR /app

# Copy file thư viện và cài đặt
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy toàn bộ mã nguồn vào container
COPY . .

# Khởi chạy server Gunicorn (Tối ưu cho môi trường Production)
# Cloud Run tự động cấp cổng qua biến môi trường $PORT (thường là 8080)
RUN chmod -R 755 /app/Data
CMD exec gunicorn --bind :$PORT --workers 1 --threads 8 --timeout 60 app:app
