# Prompt cho AI code tiếp dự án

Bạn đang làm việc với source React/Vite của **TestPilot AI**.

## Source of truth
1. Đọc `docs/design-spec.json` để hiểu flow, route, component, field, action và design token.
2. Dùng `docs/design-spec.pdf` để đối chiếu visual.
3. Giữ nguyên nguyên tắc: UI tiếng Việt, nền sáng tối giản, không tự thêm menu/banner/chức năng ngoài spec.

## Flow
`/login` -> `/dashboard` -> `/project/:projectId` -> `/project/:projectId/folder/:folderId/upload` -> `/project/:projectId/folder/:folderId/testcases`

`/api-keys` là cấu hình toàn cục.

## Quy tắc quan trọng
- Sidebar chỉ có: Bảng điều khiển, Quản lý khóa API, Dự án mới, tìm kiếm dự án, danh sách dự án.
- Màn phân tích testcase KHÔNG phải màn riêng. Khi bấm "Phân tích tài liệu", hiển thị trạng thái processing rồi chuyển sang chính màn Kho testcase.
- Kho testcase dùng accordion: header chỉ ID + tên + loại + chevron; mở ra mới thấy Điều kiện tiên quyết, Dữ liệu kiểm thử, Bước thực hiện, Kết quả mong đợi, Loại testcase.
- Loại testcase: Giao diện, Kiểm tra dữ liệu, Chức năng, Ngoại lệ, Popup, Luồng.
- Step là mảng, render dạng 1., 2., 3....
- Sửa màu xanh, Xóa màu đỏ.
- API key sau khi lưu chỉ hiển thị masked, không render lại full secret.
