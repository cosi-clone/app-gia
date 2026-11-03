# Excel Reader — Node.js + Firebase

Toàn bộ trình đọc PDF/XLSX cũ được giữ trong `public/`, thêm đăng nhập,
mã kích hoạt một lần và dashboard quản trị.

## Bắt đầu

1. Cài Node.js **20.19 trở lên**.
2. Đọc **[SETUP_FIREBASE.md](SETUP_FIREBASE.md)** để tạo Firebase project,
   bật Authentication, Firestore, cấu hình service account và UID admin.
3. Trong thư mục chứa `package.json`:

```bash
npm ci
cp .env.example .env
```

4. Điền Firebase config và service account trong `.env`, deploy Firestore rules/indexes.
5. Chạy:

```bash
npm run dev
```

Mở **http://localhost:3000**. Không cần Python. Không chạy bản này bằng `npx serve`
vì phần đăng nhập/kích hoạt cần backend Node.js.

## Tính năng

- Đăng nhập/đăng ký email và mật khẩu; Google login tùy chọn.
- Xác minh email, quên mật khẩu, đăng xuất.
- Tài khoản mới/chưa kích hoạt/hết hạn được chuyển sang nhập mã.
- Admin tạo mã số tháng tùy chọn (1–1200) hoặc vĩnh viễn, tối đa 100 mã/lần.
- Mã chỉ dùng một lần trên toàn hệ thống, kể cả hai request đồng thời.
- Gia hạn cộng tiếp nếu còn hạn; đã hết hạn thì tính từ thời điểm kích hoạt mới.
- Dashboard hiển thị mã, trạng thái, email/UID sử dụng, ngày tạo/ngày dùng.
- Tìm mã/email, lọc trạng thái, phân trang, copy mã và thu hồi mã chưa dùng.
- Admin có thể tạo tài khoản và khóa/mở khóa quyền dùng app.
- PDF upload từ máy, XLSX template, lặp hàng, tab Excel riêng, fullscreen như trước.
- Phiên đọc cục bộ tách theo tài khoản.

## Cấu trúc

```text
public/                 Giao diện cũ + đăng nhập/kích hoạt/admin
client/firebase.js      Firebase SDK client source
server/                 Node.js API, kiểm tra token, Firestore transaction
api/service.js          Vercel Function entry
scripts/                Build và tạo admin
tests/                  Test dịch vụ/API/cấu hình và integration emulator
firestore.rules         Chặn truy cập Firestore trực tiếp từ trình duyệt
firestore.indexes.json  Indexes phục vụ dashboard
.env.example            Mẫu cấu hình; không chứa key thật
vercel.json             Cấu hình deploy
```

## Kiểm tra

```bash
npm test
```

Đã kiểm tra domain/API, cạnh tranh mã dùng một lần bằng Firestore Emulator,
và Chromium với Firebase Authentication + Firestore Emulator thật: đăng nhập,
xác minh email, tạo/đổi mã, upload PDF/XLSX, tab đọc/reload và từ chối dùng lại mã.
Dashboard cũng được kiểm tra tạo mã tháng/vĩnh viễn, thống kê, chủ mã và thu hồi.
Chưa cấu hình hoặc deploy lên Firebase/Vercel production của bạn.

## Giới hạn

PDF tối đa 100 MB, XLSX tối đa 50 MB; PDF cần lớp văn bản. Trình đọc giữ các định dạng
Excel cơ bản, chưa phải Excel đầy đủ và chưa tính lại công thức hoặc xuất XLSX.
Mã/thuê bao là quản lý quyền dùng dịch vụ, không phải cổng thanh toán hoặc DRM chống
sao chép frontend. Các tài khoản cũ cần upload lại PDF/XLSX để tạo phiên có chủ sở hữu.

Service account chỉ dùng ở backend; không commit `.env`, private key hay tài liệu cá nhân.
