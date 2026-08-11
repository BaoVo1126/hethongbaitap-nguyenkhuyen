# 🚀 SmartQuiz AI — Dynamic Exam & Practice Platform (Hệ thống trung tâm luyện thi CLC Nguyễn Khuyến)

> **Nền tảng quản lý & tạo đề thi tự động bằng Flask, hỗ trợ bóc tách đề từ nhiều định dạng file (PDF/DOCX/OCR) và xáo trộn đề theo mã số học sinh.**

[![Python](https://img.shields.io/badge/Python-3.9+-3776AB?style=flat&logo=python&logoColor=white)](https://www.python.org/)
[![Flask](https://img.shields.io/badge/Framework-Flask-000000?style=flat&logo=flask&logoColor=white)](https://flask.palletsprojects.com/)
[![Demo](https://img.shields.io/badge/Live-Demo-brightgreen?style=flat&logo=render)](https://he-thong-bai-tap-smartquizai.onrender.com/)

- **🌐 Live Demo**: https://he-thong-bai-tap-smartquizai.onrender.com/
---

## 📌 Mục đích dự án (Summary)
Dự án giải quyết vấn đề **chuyển đổi đề thi truyền thống (file PDF, Word, ảnh scan) thành bài trắc nghiệm tương tác trực tuyến** mà không cần nhập liệu thủ công. Hệ thống tự động bóc tách câu hỏi, công thức, hình ảnh đi kèm, đồng thời quản lý luồng làm bài và chấm điểm tự động.

---

## ✨ Điểm sáng kỹ thuật (Key Highlights)

* **Multi-Format OCR & Parsing**: Trích xuất tự động văn bản, bảng biểu, công thức và **ảnh minh họa** từ file PDF (`PyMuPDF`), Word (`python-docx`), Excel và ảnh scan (`PaddleOCR` / `Tesseract`).
* **Xáo trộn đề thông minh (Anti-Cheat)**: Sử dụng Seeded Random từ `SHA-256(ExamID + StudentCode)` để sinh thứ tự câu hỏi và đáp án riêng cho từng học sinh mà vẫn giữ nguyên trạng thái khi reload.
* **Kiểm soát phòng thi**: Đếm ngược thời gian từ server, mở tab riêng biệt và tự động nộp bài khi hết giờ hoặc chuyển tab/thoát trang (`visibilitychange`).
* **Chấm điểm & Lưu trữ**: Chấm điểm tự động server-side, đo thời gian làm bài chính xác và ghi nhận lịch sử/điểm danh.

---

## 🛠️ Tech Stack

* **Backend**: Python, Flask, SQLite3 (Auto-migration).
* **Processing/OCR**: PyMuPDF, python-docx, PaddleOCR, Tesseract, Pillow.
* **Deployment**: Render (Production).

---

## 🚀 Hướng dẫn chạy nhanh (Quick Start)

### 1. Cài đặt môi trường
```bash
git clone [https://github.com/your-username/smartquiz-ai.git](https://github.com/your-username/smartquiz-ai.git)
cd smartquiz-ai
python -m venv venv
source venv/bin/activate  # Trên Windows: venv\Scripts\activate
```
### 2. Cài thư viện & Chạy ứng dụng
```bash
pip install flask pillow pytesseract pymupdf python-docx python-pptx openpyxl requests paddleocr
python app.py
```
- **🎓 Học sinh**: http://127.0.0.1:5000/

- **🛠️ Quản trị / Giáo viên** : http://127.0.0.1:5000/admin


