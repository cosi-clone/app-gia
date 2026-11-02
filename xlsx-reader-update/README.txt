EXCEL READER — BẢN VERCEL

Deploy:
1. Giải nén. Đưa NỘI DUNG thư mục này vào repository GitHub.
2. Trên Vercel chọn Add New Project, import repository.
3. Framework Preset: Other.
4. Root Directory: thư mục chứa index.html, api và vercel.json.
5. Build Command: để trống. Output Directory: .
6. Deploy. Không cần database, biến môi trường hoặc pip package.

Bộ này đã có đủ thư viện JSZip và PDF.js, không cần tải lại.
Không sử dụng server.py và không chạy python3 server.py trên Vercel.
Nếu muốn chạy thử cả API ở local với Vercel CLI: npx vercel dev

Sử dụng:
- Upload PDF/TXT/EPUB hoặc dán link PDF rồi bấm Tải PDF từ link.
- Upload XLSX, chọn sheet/cột/dòng, bấm Chèn truyện.
- Khôi phục template để trả lại nội dung gốc.
- Đổi sheet bỏ phần chèn hiện tại. File trên máy không bị sửa.

Giới hạn:
- API PDF dạng phản hồi thường được giới hạn 4.000.000 byte (~4 MB)
  để ở dưới mức phản hồi 4,5 MB của Vercel Functions.
- PDF online tải trực tiếp từ nguồn có CORS vẫn có giới hạn 100 MB.
- PDF upload được xử lý ngay trong trình duyệt, tối đa 100 MB.
- XLSX tối đa 50 MB, 500 cột, 100.000 dòng.
- Chưa có xuất XLSX hoặc lưu phiên đọc qua reload.
- Chỉ giữ định dạng Excel cơ bản. Công thức hiện giá trị cache, không tính lại.
- Ô gộp/dòng ẩn được bỏ qua khi chèn. Thiếu dòng thì tự thêm dòng.
- PDF scan cần OCR. Link yêu cầu đăng nhập/captcha không được hỗ trợ.
- Link PTIT mẫu chưa xác nhận truy cập thành công từ môi trường kiểm tra.
- Chưa deploy kiểm thử trên tài khoản Vercel thực.

Kiểm tra đã thực hiện: cú pháp JS/Python, cấu hình JSON, tải PDF qua API
bằng phản hồi mô phỏng, giới hạn dung lượng và phản hồi HTTP local.
Các luồng đọc/chèn XLSX, TXT, PDF đã được kiểm tra ở bản trước.
