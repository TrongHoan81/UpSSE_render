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
```eof

### Giải thích tại sao làm vậy:
#*   **Chuyển sang `python:3.11-slim`:** Đây là phiên bản Python ổn định, hỗ trợ hoàn hảo cho `MarkupSafe 3.x` và các thư viện hiện đại.
#*   **Thêm `pip install --upgrade pip`:** Trong môi trường Docker cũ, pip thường bị lỗi thời, gây ra việc không tìm thấy các bản cập nhật mới nhất của thư viện. Dòng này giúp cập nhật trình cài đặt lên phiên bản mới nhất trước khi cài thư viện.
#*   **Đã loại bỏ lệnh `chmod` lỗi:** Dockerfile mới này đã sạch sẽ và không còn lỗi thừa thư mục nữa.

#Sau khi sửa 2 file này, bạn hãy thử Build lại (Deploy lại) lên Google Cloud Run. Nó sẽ vượt qua bước cài đặt thư viện một cách suôn sẻ!
