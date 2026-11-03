EXCEL READER — NODE.JS + FIREBASE

Bản này đã thêm đăng nhập và mã kích hoạt. Đọc README.md và SETUP_FIREBASE.md
trước khi chạy. Cần cấu hình Firebase Authentication, Firestore và UID admin.

Chạy:
  npm ci
  cp .env.example .env
  [Điền cấu hình Firebase thật trong .env]
  npm run dev

Mở http://localhost:3000
Không dùng Python. Không chạy bằng npx serve vì cần API Node.js.

Trang admin: /admin.html
Tài khoản mới cần xác minh email và nhập mã do admin cấp.
Mã dùng một lần; có gói theo tháng hoặc vĩnh viễn.
PDF/XLSX vẫn được đọc cục bộ; giao diện Excel và lặp hàng như bản trước.

Hướng dẫn Firebase, quyền admin, Firestore rules/indexes và Vercel nằm trong
SETUP_FIREBASE.md. Không gửi hoặc commit service account/.env chứa key thật.
