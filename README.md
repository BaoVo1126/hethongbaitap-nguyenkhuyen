# Hệ thống Trung tâm — Demo

Demo Flask chạy trên localhost: upload đề (PDF/DOCX/PPTX/TXT/ảnh) → hệ thống tự trích xuất câu hỏi
(PyMuPDF cho PDF chữ, PaddleOCR/Tesseract cho bản scan/ảnh) → giáo viên xem lại & đăng → học sinh làm
bài (mở ở tab riêng, có thể giới hạn thời gian, hết giờ tự nộp hoặc nộp sớm tuỳ ý) → chấm điểm tự
động, hiện điểm + thời gian làm bài trước rồi mới tới phần xem lại đúng/sai → nộp bài = điểm danh có
mặt.

## Cài đặt

```bash
pip install flask pillow pytesseract pymupdf python-docx python-pptx requests
# PaddleOCR (tuỳ chọn, cho OCR ảnh/PDF scan chất lượng cao hơn Tesseract):
pip install paddleocr paddlepaddle pdf2image
```

Cần cài thêm engine Tesseract nếu muốn dùng làm dự phòng:
```bash
sudo apt-get install tesseract-ocr tesseract-ocr-vie   # Ubuntu
brew install tesseract tesseract-lang                   # macOS
```

## Chạy demo

```bash
python app.py
```
- Học sinh: http://127.0.0.1:5000/
- Giáo viên: http://127.0.0.1:5000/admin

Database SQLite (`center.db`) tự tạo/migrate khi chạy lần đầu, không mất dữ liệu cũ khi nâng cấp.

## Làm bài ở tab riêng

Từ trang nhập họ tên + mã số, bấm "Bắt đầu làm bài" sẽ **mở bài làm ở một tab trình duyệt mới**
(form dùng `target="_blank"`), tab nhập thông tin vẫn giữ nguyên. Trong lúc làm bài (chưa nộp), nếu
học sinh lỡ đóng/tải lại tab đó, trình duyệt sẽ hỏi xác nhận trước khi rời trang — tránh mất bài làm
dở do bấm nhầm. Lưu ý: đây không phải chế độ kiosk khoá cứng trình duyệt (không có công nghệ web
chuẩn nào làm được điều đó vì lý do bảo mật) — học sinh vẫn có thể đóng tab nếu cố ý.

## Giới hạn thời gian làm bài

Khi tạo đề (hoặc sửa sau ở trang "Xem/Sửa"), giáo viên có thể đặt **số phút giới hạn**. Nếu có đặt:

- Học sinh thấy đồng hồ đếm ngược ở đầu trang làm bài.
- Đồng hồ tính theo thời điểm học sinh **bắt đầu** làm bài (lưu ở server) + số phút giới hạn — tải
  lại trang không "được" thêm giờ.
- Hết giờ, trình duyệt **tự động nộp bài** với những câu đã trả lời tại thời điểm đó.
- Học sinh cũng có thể bấm **"Nộp bài" để nộp sớm** bất cứ lúc nào, không cần đợi hết giờ.
- Để trống ô giới hạn thời gian = không giới hạn, không hiện đồng hồ.

## Trang kết quả

Sau khi nộp, trang kết quả hiện **khung điểm số + thời gian làm bài + thời điểm nộp** ở trên cùng,
rồi mới tới phần xem lại từng câu đúng/sai bên dưới.

## Giới hạn của bản demo

- Đồng hồ đếm ngược và cảnh báo rời trang chạy ở trình duyệt (client-side) — không phải cơ chế chống
  gian lận cấp production (học sinh có thể sửa giờ máy, dùng devtools...). Phù hợp quy mô lớp học nhỏ,
  tin tưởng lẫn nhau.
- Chưa có xác thực đăng nhập cho khu vực giáo viên (`/admin`).
- SQLite phù hợp demo/localhost; lên production nên đổi sang Postgres.
