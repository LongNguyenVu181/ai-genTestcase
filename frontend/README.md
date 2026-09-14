# TestPilot AI Web

Source code React/Vite bám theo bộ mockup và design spec đã chốt.

## Chạy local

```bash
npm install
npm run dev
```

Mở: `http://localhost:5173`

## Build production

```bash
npm run build
npm run preview
```

Thư mục build: `dist/`.

## Flow đã implement

- `/login` — Đăng nhập
- `/dashboard` — Bảng điều khiển / tạo dự án
- `/project/:projectId` — Chi tiết dự án (Web App + API)
- `/project/:projectId/folder/:folderId/upload` — Tải tài liệu
- `/project/:projectId/folder/:folderId/testcases` — Kho testcase (accordion)
- `/api-keys` — Quản lý khóa API

## Lưu ý

- Đây là frontend demo với mock data.
- Nút “Phân tích tài liệu” mô phỏng processing rồi chuyển sang cùng màn Kho testcase; không có màn phân tích riêng.
- “Tải Excel” hiện export CSV UTF-8 để Excel mở trực tiếp, tránh thêm dependency xlsx.
- API key trong demo chỉ giữ bản masked sau khi bấm Lưu. Production phải gửi secret về backend và mã hóa/lưu trữ an toàn, không lưu full key ở frontend.
- Sidebar chỉ giữ đúng menu đã chốt: Bảng điều khiển, Quản lý khóa API, Dự án mới, tìm kiếm/danh sách dự án.

## V3 layout update
- Nội dung các màn bên trong được giới hạn khoảng 1240px để không bị kéo giãn trên màn 1920px.
- Sidebar thu gọn còn khoảng 248px.
- Card, tiêu đề và khoảng cách được nén nhẹ để giao diện chắc và cân hơn.


## V4 layout update
- Login không còn trải full 1920px: 2 cột được gom vào khung trung tâm.
- Card đăng nhập lớn và gần phần nội dung hơn.
- Màn hình nội bộ giới hạn vùng nội dung khoảng 1180px.
- Sidebar thu còn 232px để UI chắc hơn.


## V5 - 24 inch target
V5 không còn để layout tự giãn theo màn 27 inch/2K. Toàn app dùng canvas desktop tối đa 1440px.
Chi tiết xem `VIEWPORT_24INCH.md`.


## V6 - Nền kiểu Steam profile
- Phần trống ngoài canvas được đổ nền gradient xanh/xám nhẹ.
- App canvas nổi như một workspace riêng, bo góc và có shadow.
- Bên trong app có dải nền phía trên; các panel thao tác nổi trên nền đó.
- Không thêm menu mới, chỉ thay treatment của background/panel.


## V7 - Dense layout
- Canvas desktop giảm còn 1360px để đỡ giãn trên màn lớn.
- Login được gom lại, tăng density, giảm khoảng trống phía dưới.
- App bên trong có nền chìm phần dưới thay vì trắng phẳng.
- Các panel thao tác nổi rõ hơn trên nền.


## V8 - Balanced wide desktop
- Canvas app tăng lên tối đa 1680px trên màn 1920px.
- Gutter hai bên giảm mạnh.
- Content nghiệp vụ vẫn giới hạn khoảng 1320px để không bị kéo giãn.
- Dashboard giữ cột phải 330px và phần chính linh hoạt.


## V9 - Wide workspace
- Canvas app tối đa 1820px, gần full màn 1920.
- Nội dung chung tối đa 1480px.
- Riêng Kho testcase tối đa 1580px.
- Kho testcase có toolbar rộng/sticky và card accordion sử dụng nhiều chiều ngang hơn.
- Modal thêm/sửa testcase tăng lên 900px.
- Hai cạnh màn có chi tiết nền nhẹ dạng rail/dots, không phải menu.


## V11 - Pastel background on all pages
- Áp dụng nền pastel cho toàn bộ các trang, không chỉ Login.
- Kết hợp:
  - abstract gradient blobs
  - line-art cành lá rất nhạt ở hai cạnh
  - dotted detail ở phần đáy nội dung
- Giữ nguyên wide workspace của màn Kho testcase.
