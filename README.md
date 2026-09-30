# HRM · Sổ tay Quản lý nghỉ phép

Hệ thống quản lý nghỉ phép năm, Đại học Y Dược TP. Hồ Chí Minh.

```
Trình duyệt ──► Cloudflare Pages (public/)  ──fetch──►  Google Apps Script Web App  ──►  Google Sheet
                giao diện tĩnh                          API (apps-script/Code.gs)          dữ liệu
```

- **Giao diện** (`public/`) được lưu trên GitHub và Cloudflare Pages tự động triển khai.
- **API và dữ liệu** vẫn là Google Apps Script cùng Google Sheet như hiện tại. Không cần di chuyển dữ liệu.

## Cấu trúc thư mục

| Đường dẫn | Mô tả |
|---|---|
| `public/index.html` | Toàn bộ giao diện (HTML/CSS/JS) |
| `public/config.js` | **Nơi duy nhất cần sửa**: `API_URL` của Web App |
| `public/_headers` | Header bảo mật và cache cho Cloudflare Pages |
| `apps-script/Code.gs` | Mã backend, dán vào dự án Apps Script |
| `apps-script/appsscript.json` | Manifest (chỉ cần khi dùng `clasp`) |

---

## Bước 1: Cập nhật Apps Script (backend)

1. Mở dự án Apps Script đang chạy.
2. Thay nội dung `Code.gs` bằng file `apps-script/Code.gs`.
3. Sửa dòng `APP_URL` thành địa chỉ Cloudflare Pages (xem bước 3). Link này được dùng trong email thông báo.
4. File `index.html` trong Apps Script **có thể xóa**. Nếu giữ lại cũng không ảnh hưởng.
5. Chọn **Triển khai → Quản lý các lần triển khai → biểu tượng bút chì → Phiên bản: Phiên bản mới → Triển khai**.
   - Thực thi dưới dạng: **Tôi** (tài khoản chủ Sheet)
   - Ai có quyền truy cập: **Bất kỳ ai**
   - Nên **sửa lần triển khai cũ** thay vì tạo mới, để URL `/exec` giữ nguyên.
6. Sao chép **URL ứng dụng web** (kết thúc bằng `/exec`).

> Kiểm tra nhanh: mở `URL/exec?action=ping` trên trình duyệt. Nếu thấy `{"success":true,"version":"HRM_V310",...}` là backend đã sẵn sàng.

## Bước 2: Đưa mã lên GitHub

1. Tạo repository mới trên GitHub, ví dụ `hrm-nghiphep`. Nên để ở chế độ **Private**.
2. Mở `public/config.js` và dán URL `/exec` vào `API_URL`.
3. Đẩy mã lên:

```bash
git init
git add .
git commit -m "HRM nghỉ phép: giao diện Cloudflare Pages"
git branch -M main
git remote add origin https://github.com/<tai-khoan>/hrm-nghiphep.git
git push -u origin main
```

Nếu không dùng dòng lệnh: trên GitHub chọn **Add file → Upload files** rồi kéo toàn bộ thư mục vào.

## Bước 3: Kết nối Cloudflare Pages

1. Đăng nhập [dash.cloudflare.com](https://dash.cloudflare.com), chọn **Workers & Pages → Create → Pages → Connect to Git**.
2. Chọn repository `hrm-nghiphep`.
3. Cấu hình build:
   - Framework preset: **None**
   - Build command: *(để trống)*
   - Build output directory: **`public`**
4. Bấm **Save and Deploy**. Sau khoảng 1 phút sẽ có địa chỉ dạng `https://hrm-nghiphep.pages.dev`.
5. (Tùy chọn) Vào **Custom domains** để gắn tên miền riêng, ví dụ `nghiphep.ump.edu.vn`.
6. Quay lại **Bước 1, mục 3** để cập nhật `APP_URL` bằng địa chỉ thật, rồi triển khai lại phiên bản mới.

Từ nay, mỗi lần `git push` lên nhánh `main`, Cloudflare sẽ tự động cập nhật giao diện.

---

## Cập nhật về sau

| Thay đổi | Việc cần làm |
|---|---|
| Giao diện (`public/`) | `git push`, Cloudflare tự triển khai |
| Backend (`Code.gs`) | Dán lại vào Apps Script, rồi **Quản lý các lần triển khai → Phiên bản mới** |
| URL Web App thay đổi | Sửa `public/config.js`, rồi `git push` |

## Xử lý sự cố

| Hiện tượng | Nguyên nhân thường gặp |
|---|---|
| "Chưa cấu hình API_URL trong config.js" | Chưa dán URL `/exec` vào `config.js` |
| "Máy chủ trả về dữ liệu không hợp lệ" | Web App chưa để **Bất kỳ ai**, hoặc đang dùng URL `/dev` thay cho `/exec` |
| Sửa `Code.gs` nhưng không có tác dụng | Chưa tạo **phiên bản mới** trong Quản lý các lần triển khai |
| Phản hồi chậm 1–3 giây | Đặc điểm bình thường của Apps Script |
| Email thông báo dẫn link cũ | Chưa sửa `APP_URL` trong `Code.gs` |

## Lưu ý bảo mật

- Web App ở chế độ "Bất kỳ ai" nên ai có URL `/exec` đều gọi được API. Mọi thao tác vẫn yêu cầu token đăng nhập, nhưng việc đăng nhập hiện **chỉ dựa trên số CCCD**. Nên bổ sung mật khẩu hoặc mã OTP qua email.
- Không đưa ID Google Sheet hay thông tin nhạy cảm nào khác vào thư mục `public/`.
