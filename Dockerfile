# ใช้ Python base image
FROM python:3.10-slim

# ตั้ง working directory ใน container
WORKDIR /app

# คัดลอกไฟล์ทั้งหมดจากโฟลเดอร์ปัจจุบันเข้า container
COPY . .

# ติดตั้ง dependencies
RUN pip install --no-cache-dir -r requirements.txt

# เปิดพอร์ต 5000 (Flask ใช้พอร์ตนี้โดย default)
EXPOSE 5000

# สั่งรัน Flask app
CMD ["python", "app.py"]
