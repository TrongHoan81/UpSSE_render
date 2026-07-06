<<<<<<< HEAD
# Nâng cấp lên python 3.11 để tương thích với các thư viện mới
FROM python:3.11-slim

# Ngăn Python tạo file .pyc và ép in log trực tiếp ra console
ENV PYTHONDONTWRITEBYTECODE 1
ENV PYTHONUNBUFFERED 1

# Thiết lập thư mục làm việc
WORKDIR /app

# Copy requirements và cài đặt
COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# Copy toàn bộ mã nguồn
COPY . .

# Mở port cho Cloud Run
EXPOSE 8080

# Chạy ứng dụng
CMD exec gunicorn --bind :$PORT --workers 1 --threads 8 --timeout 0 app:app
=======
# Nâng cấp lên python 3.11 để tương thích với các thư viện mới
FROM python:3.11-slim

# Ngăn Python tạo file .pyc và ép in log trực tiếp ra console
ENV PYTHONDONTWRITEBYTECODE 1
ENV PYTHONUNBUFFERED 1

# Thiết lập thư mục làm việc
WORKDIR /app

# Copy requirements và cài đặt
COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# Copy toàn bộ mã nguồn
COPY . .

# Mở port cho Cloud Run
EXPOSE 8080

# Chạy ứng dụng
CMD exec gunicorn --bind :$PORT --workers 1 --threads 8 --timeout 0 app:app
>>>>>>> 3ef276fb99249cd2c8b68fa99a9d36212337caed
