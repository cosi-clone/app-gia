# Excel Reader — Firebase và Node.js

Project này bao gồm toàn bộ phần đọc PDF/XLSX cũ và phần đăng nhập/kích hoạt mới.
Không cần Python. Cần Node.js 20.19 trở lên, tài khoản Firebase và cấu hình riêng của bạn.

## 1. Tạo Firebase project

1. Mở Firebase Console: https://console.firebase.google.com/ và tạo project.
2. Project settings → General → Your apps → thêm Web app.
3. Copy object `firebaseConfig` vào biến `FIREBASE_WEB_CONFIG` trong `.env`.
   Dùng JSON hợp lệ, có dấu ngoặc kép; xem `.env.example`.
4. Build → Authentication → Get started → Sign-in method → bật Email/Password.
5. Nếu muốn Google login, bật Google và đặt `ENABLE_GOOGLE_LOGIN=true`.
6. Authentication → Settings → Authorized domains: thêm `localhost`, `127.0.0.1`
   nếu cần cho máy local, và domain Vercel/domain riêng dùng thực tế.
7. Build → Firestore Database → Create database → chọn Standard edition,
   database `(default)` và khu vực phù hợp. Không chọn chế độ rules mở cho mọi người.

## 2. Cấu hình server

Trong thư mục chứa `package.json`:

```bash
npm ci
cp .env.example .env
```

Firebase Console → Project settings → Service accounts → Generate new private key.
Copy toàn bộ JSON vào `FIREBASE_SERVICE_ACCOUNT` trong `.env` (JSON một dòng được
bọc bằng dấu nháy đơn). `private_key` giữ các ký tự `\n` theo JSON gốc.
`project_id` của service account phải trùng `projectId` của Web app.

Không đưa key vào `public/`, không gửi key lên chat/GitHub và không dùng tiền tố
`VITE_`/`NEXT_PUBLIC_` cho service account. Biến web config là cấu hình công khai;
service-account private key chỉ được đọc bởi backend Node.js.

`.env` cần có:

```dotenv
PORT=3000
FIREBASE_WEB_CONFIG='{"apiKey":"…","authDomain":"….firebaseapp.com","projectId":"…","appId":"…"}'
FIREBASE_SERVICE_ACCOUNT='{"type":"service_account","project_id":"…","private_key":"…","client_email":"…"}'
ADMIN_UIDS=UID_CUA_ADMIN
ENABLE_GOOGLE_LOGIN=false
USE_FIREBASE_EMULATORS=false
```

## 3. Tạo admin

Cách A — Firebase Console:

1. Authentication → Users → Add user: tạo email và mật khẩu admin.
2. Copy UID tài khoản đó vào `ADMIN_UIDS` trong `.env`.
3. Admin đăng nhập và xác minh email theo giao diện trước khi sử dụng.

Cách B — script quản trị, cần service account đã cấu hình:

1. Thêm `BOOTSTRAP_ADMIN_EMAIL` và `BOOTSTRAP_ADMIN_PASSWORD` vào `.env`.
   Mật khẩu tối thiểu 12 ký tự.
2. Chạy:

```bash
npm run admin:create
```

3. Script tạo/tìm tài khoản, đánh dấu email đã xác minh rồi in UID.
4. Copy UID vào `ADMIN_UIDS`. Xóa `BOOTSTRAP_ADMIN_PASSWORD` sau đó.

Nhiều admin: `ADMIN_UIDS=uid1,uid2`. Quyền admin lấy từ biến server, không lấy
từ email hay dữ liệu do trình duyệt gửi lên. Thay biến xong phải restart/redeploy.

## 4. Deploy Firestore rules và indexes

Cài Firebase CLI nếu chưa có:

```bash
npm install -g firebase-tools
firebase login
firebase deploy --only firestore:rules,firestore:indexes --project YOUR_PROJECT_ID
```

Rules của project chặn toàn bộ thao tác Firestore từ client. Mọi dữ liệu mã,
quyền sử dụng và dashboard đi qua Node.js + Firebase Admin SDK có kiểm tra token.
Không thay rules bằng `allow read, write: if true`.

Sau khi deploy indexes, đợi các index chuyển sang Enabled trước khi lọc mã theo
trạng thái hoặc email. Nếu API báo thiếu index, deploy lại file indexes và đợi.

## 5. Chạy local

```bash
npm run dev
```

Mở http://localhost:3000. Lệnh này build Firebase client rồi chạy Node.js server.
Không dùng `npx serve` cho bản có đăng nhập vì lệnh đó không chạy backend.

Sau khi đã build có thể dùng `npm start`. Khi thay `client/firebase.js`, chạy
`npm run build` lại. Thay HTML/CSS/JS trong `public/` chỉ cần reload trang.

## 6. Luồng sử dụng

### Người dùng

1. Đăng ký email/mật khẩu hoặc đăng nhập bằng Google nếu đã bật.
2. Xác minh email (Google thường đã có email xác minh).
3. Tài khoản chưa có gói được chuyển sang màn hình nhập mã.
4. Nhập mã hợp lệ → mở ứng dụng upload PDF/XLSX như trước.
5. Tài khoản hết hạn được chuyển về màn hình gia hạn. Người còn hạn có thể
   mở Tài khoản / Gia hạn để nhập mã mới và cộng tiếp thời gian.

### Admin

- Mở `/admin.html` hoặc bấm Quản trị trên trang chuẩn bị Excel.
- Tạo 1–100 mã/lần. Chọn số tháng 1–1200 hoặc vĩnh viễn.
- Xem toàn bộ mã, trạng thái, ngày tạo/ngày dùng, email và UID đã sử dụng.
- Lọc theo trạng thái; tìm mã đầy đủ hoặc email chính xác; Tải thêm để phân trang.
- Copy mã cấp cho người dùng qua kênh riêng. Hệ thống không tự bán hàng/gửi mã.
- Thu hồi mã chưa dùng. Mã đã dùng vẫn giữ lịch sử và không tái sử dụng được.
- Cấp tài khoản email/mật khẩu (mật khẩu tối thiểu 12 ký tự), xem tài khoản,
  khóa/mở khóa quyền dùng app. Tài khoản được cấp vẫn cần xác minh email và mã.
- Không có chức năng xóa tài khoản hoặc thay đổi quyền admin trong dashboard.

## 7. Quy tắc thời hạn và mã

- Mã ngẫu nhiên 144 bit, dạng `EX-…`, mỗi mã dùng một lần trên toàn hệ thống.
- Đổi mã cập nhật đồng thời mã, thành viên và lịch sử bằng Firestore transaction.
  Hai request đồng thời không thể cùng nhận quyền từ một mã.
- Hết hạn dựa trên giờ server, không dựa vào ngày giờ máy người dùng.
- Gói tháng dùng tháng lịch theo giờ Việt Nam (UTC+7). Ngày 31 chuyển sang
  tháng ít ngày sẽ chốt ở ngày cuối tháng đó, giữ giờ địa phương.
- Còn hạn: cộng vào ngày hết hạn hiện tại. Hết hạn/chưa có gói: tính từ lúc đổi mã.
- Gói vĩnh viễn không có ngày hết hạn. Không tiêu thụ mã tháng cho tài khoản
  đã có gói vĩnh viễn. Admin có quyền dùng app độc lập với gói.
- Thử mã tối đa 20 lần/10 phút/tài khoản. Mã bị thu hồi/đã dùng không cấp quyền.
- Khóa tài khoản không trả lại mã, không hoàn thời gian và không xóa dữ liệu.
- App kiểm tra quyền khi mở trang, upload tệp, mở tab đọc; kiểm tra lại mỗi phút
  và khi quay lại tab. Vì vậy bản có đăng nhập cần kết nối mạng để xác thực quyền.

## 8. Vercel

1. Đẩy nội dung project lên GitHub, giữ `api/`, `server/`, `public/`, `client/`,
   `scripts/`, `package.json`, `package-lock.json`, `vercel.json`.
2. Vercel → Import Git repository → Framework Other.
3. Root Directory: thư mục chứa `package.json`.
4. Build Command: `npm run build`. Output Directory: `public`.
5. Thêm Environment Variables: `FIREBASE_WEB_CONFIG`, `FIREBASE_SERVICE_ACCOUNT`,
   `ADMIN_UIDS`, `ENABLE_GOOGLE_LOGIN`; đặt `USE_FIREBASE_EMULATORS=false` hoặc bỏ.
   Trong ô giá trị Vercel, nhập JSON trực tiếp, không thêm dấu nháy đơn ngoài cùng.
6. Chọn Node.js runtime đáp ứng Node 20.19 trở lên và deploy.
7. Thêm domain deploy vào Firebase Auth Authorized domains rồi test đăng nhập.
8. Khi sửa biến môi trường, redeploy để Vercel dùng giá trị mới.

Không cần Firebase Cloud Functions; backend chạy Node.js local hoặc Vercel Functions.
Firestore và Firebase Authentication là dịch vụ Firebase cần được cấu hình riêng.

## 9. Kiểm thử

```bash
npm test
```

Các test kiểm tra giới hạn, phân quyền API, thời hạn, mã một lần, cạnh tranh
đổi mã, gia hạn, vĩnh viễn, thu hồi, khóa tài khoản và thống kê.

Để kiểm tra transaction trên Firestore Emulator (CLI/Java chỉ cần cho test này):

```bash
firebase emulators:exec --project demo-excel-reader --only firestore "npm run test:integration"
```

Test chỉ chạy với `FIRESTORE_EMULATOR_HOST`, không tác động Firestore production.
Firebase CLI mới có thể cần Java 21 trở lên; Java không cần khi chạy app bình thường.

## 10. Dữ liệu cũ và giới hạn còn lại

- PDF/XLSX vẫn xử lý cục bộ trong trình duyệt, không upload lên Firebase/server.
- Giao diện Excel, lặp hàng từ vị trí chọn, tab riêng, fullscreen được giữ lại.
- Phiên đọc mới gắn với UID. Phiên cục bộ cũ trước khi có đăng nhập chưa có chủ
  sở hữu sẽ không mở được sau nâng cấp; upload lại hai tệp để tạo phiên mới.
- Dữ liệu tài khoản/mã mới nằm trong `members`, `activationCodes`, `redemptions`,
  `rateLimits`. Đây không phải cơ chế đồng bộ PDF/XLSX giữa các máy.
- Chưa có thanh toán, gửi email cấp mã tự động hoặc xuất XLSX.
- Hạn chế hiển thị Excel/PDF của bản cũ vẫn áp dụng; xem README.txt.
- Vì trình đọc PDF/XLSX chạy ở client, người nhận mã nguồn có thể tự sửa code
  thành một bản offline riêng. Cơ chế này bảo vệ tài khoản/mã/thời hạn của dịch vụ,
  không phải DRM để chống sao chép mã nguồn frontend.

## Tài liệu chính thức

- Firebase Admin: https://firebase.google.com/docs/admin/setup
- Xác thực token: https://firebase.google.com/docs/auth/admin/verify-id-tokens
- Transaction: https://firebase.google.com/docs/firestore/manage-data/transactions
