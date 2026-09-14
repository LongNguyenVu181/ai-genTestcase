# V5 target viewport

Thiết kế V5 được khóa theo cảm giác hiển thị của màn hình 24 inch Full HD.

CSS không thể nhận biết kích thước vật lý 24 inch hay 27 inch.
Vì vậy source sử dụng **canvas tối đa 1440px**:
- Màn 24 inch / laptop: UI dùng gần hết viewport.
- Màn 27 inch / 2K / ultrawide: UI không bị kéo giãn; phần dư trở thành outer gutter.
- Sidebar: 228px.
- Khu vực nội dung: phần còn lại trong canvas 1440px.
- Login: 2 cột 760px + 520px với gap 48px.

Nếu muốn rộng hơn về sau chỉ cần đổi:
`--desktop-canvas: 1440px;`
