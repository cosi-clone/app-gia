EXCEL READER — UPLOAD PDF, KHÔNG CẦN PYTHON

CHẠY LOCAL:
1. Giải nén vào thư mục mới (để không sót file Python/API từ bản cũ).
2. Mở Terminal trong thư mục chứa index.html.
3. Chạy: npx serve . --listen 3000
4. Mở http://localhost:3000
Lần đầu npx có thể hỏi tải gói serve: chọn y. Cần Node.js/npm.
Đây là web tĩnh HTML/CSS/JavaScript, không cần Python, API hoặc database.

SỬ DỤNG:
1. Chọn file PDF từ máy. App tự trích xuất nội dung ngay trong trình duyệt.
2. Chọn XLSX làm template, chọn sheet/cột/hàng bắt đầu.
3. Bấm Mở tab Excel để đọc.
4. Bấm □ góc trên phải trong tab mới để bật fullscreen (Esc để thoát).
Không còn phần nhập link PDF hoặc tải truyện online.
PDF/XLSX không được upload lên server; chỉ đọc trong trình duyệt.

CÁCH LẶP HÀNG:
- Giữ các hàng trước dòng bắt đầu.
- Khối mẫu là toàn bộ các hàng từ dòng bắt đầu đến cuối sheet gốc.
- Nếu không đủ ô cho truyện, lặp nguyên khối này đến khi đủ chỗ.
- Các cột khác, giá trị và định dạng của hàng mẫu cũng được lặp theo.
- Ô gộp được lặp nếu nằm hoàn toàn trong khối; dòng ẩn vẫn ẩn.
- Không chèn truyện vào dòng ẩn hoặc ô gộp của cột được chọn.
- Không cho chọn dòng bắt đầu cắt ngang vùng ô gộp; chọn dòng đầu vùng gộp
  hoặc dòng sau vùng đó. Nếu vùng không có ô chèn được thì báo lỗi.
- Mỗi đoạn truyện vào một ô. Tối đa 100.000 hàng sau khi lặp.
- Khối lặp cuối có thể dư hàng; các hàng dư giữ nội dung mẫu.
- Không sửa file XLSX gốc, không tạo công thức Excel mới.

TAB ĐỌC:
- Thanh tiêu đề/ribbon/thanh công thức dùng HTML bạn gửi và CSS gốc kèm điều chỉnh.
- Không có form upload/link trong tab này. Chỉ hiện bảng và các sheet.
- Bấm ô để xem nội dung ở thanh công thức; phím mũi tên di chuyển khi bảng có focus.
- Có nút chỉnh font/cỡ chữ/đậm/nghiêng cho cột truyện.
- Các nút ribbon khác chủ yếu phục vụ giao diện, không phải bộ Excel đầy đủ.
- Cuộn liên tục với cơ chế chỉ dựng hàng gần vùng đang xem.
- Phiên đọc được lưu trong IndexedDB cùng trình duyệt/cùng origin; tab có thể reload.
- Link tab chứa mã phiên cục bộ, không dùng để chia sẻ cho người/máy khác.
- Không thay domain/cổng giữa trang chuẩn bị và tab đọc. Xóa dữ liệu trang sẽ mất phiên.
- Trình duyệt cần cho phép mở tab mới và lưu dữ liệu. Chặn popup/bộ nhớ sẽ có thông báo.

VERCEL:
- Đẩy nội dung gói lên GitHub và import vào Vercel.
- Framework: Other; Build Command: trống; Output Directory: .
- Root Directory là thư mục chứa index.html và vercel.json.
- Không cần Python, Vercel Function, biến môi trường hoặc database.
- Nếu sửa đè repository cũ, xóa api/pdf.py và dev.py cũ, thay vercel.json.
- Không commit file PDF/XLSX cá nhân.

GIỚI HẠN:
- PDF tối đa 100 MB, cần có lớp văn bản; PDF scan cần OCR trước.
- XLSX tối đa 50 MB, 500 cột, 100.000 hàng sau khi lặp.
- Chưa có xuất XLSX hoặc tính lại/dịch công thức khi lặp; công thức hiện cache.
- Định dạng Excel phức tạp (ảnh/biểu đồ/conditional format) chưa hỗ trợ đầy đủ.
- Các tính năng và giao diện ở tab Excel giữ nguyên bản trước.

KIỂM TRA:
- Chromium: upload PDF local, XLSX mẫu 11 sheet, lặp nguyên khối hàng,
  màu ô và ô gộp, mở tab riêng, thanh công thức, reload và fullscreen.
- Truyện thử 1.000 đoạn: cuộn đến đoạn cuối, số hàng DOM được giới hạn.
