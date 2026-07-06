# Dùng image Python bản nhẹ

FROM python:3.10-slim

# Ngăn Python tạo ra các file .pyc và ép in log trực tiếp ra console (Tốt cho Cloud Run)

ENV PYTHONDONTWRITEBYTECODE 1
ENV PYTHONUNBUFFERED 1

# Tạo thư mục làm việc trong container

WORKDIR /app

# Ưu tiên copy file requirements trước để tận dụng Docker cache

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

Copy TẤT CẢ mã nguồn VÀ các file Excel cấu hình vào thư mục /app trong container

# Docker sẽ tự động copy cả nội dung bên trong thư mục static và templates

COPY . .

# Mở port (Cloud Run sẽ tự động truyền biến môi trường PORT vào đây)

EXPOSE 8080

# Chạy ứng dụng bằng Gunicorn thay vì Flask thuần.

# - workers 1, threads 8: Tối ưu cho quá trình I/O bound (xử lý file Excel) trên Cloud Run

# - timeout 0: Giao việc quản lý timeout lại cho Cloud Run

CMD exec gunicorn --bind :$PORT --workers 1 --threads 8 --timeout 0 app:app