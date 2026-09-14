# TestPilot AI — Integrated Multi-Agent + New UI

Project này ghép **UI React/Vite mới** với **pipeline AI từ `multiA.txt`** để test lại prompt trên giao diện mới.

## Kiến trúc

```text
Browser (React/Vite :5173)
        |
        | /api proxy
        v
FastAPI backend :8000
        |
        +-- Web Agent 1 -> Rule Matrix -> Review UI
        +-- Web Agent 2 -> Testcase tree -> Accordion / XMind / Excel
        +-- API Agent 1 -> API Rule Matrix
        +-- API Agent 2 -> API Testcase tree
        +-- Qwen DashScope OpenAI-compatible API
```

## 1. Cài đặt lần đầu trên Windows

Mở PowerShell tại thư mục project:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\setup.ps1
```

Hoặc chạy thủ công:

```powershell
cd backend
py -m pip install -r requirements.txt

cd ..\frontend
npm install
```

## 2. Chạy dev

Cách nhanh:

```powershell
.\start-dev.ps1
```

Hoặc mở 2 terminal.

Terminal backend:

```powershell
cd backend
py -m uvicorn app:app --reload --host 127.0.0.1 --port 8000
```

Terminal frontend:

```powershell
cd frontend
npm run dev
```

Mở:
- UI: http://localhost:5173
- Backend Swagger: http://127.0.0.1:8000/docs

## 3. Flow test Web prompt

1. Login demo -> Dashboard.
2. Mở `Quản lý khóa API`.
3. Paste Qwen API Key.
4. Kiểm tra Base URL, chọn model: `qwen-max`, `qwen3.8-max` hoặc `qwen-plus`.
5. `Kiểm tra kết nối` -> `Lưu cấu hình`.
6. Vào Project -> Web App -> tạo/chọn folder -> `Tải tài liệu`.
7. Upload PDF/MD/TXT/DOCX.
8. `Phân tích tài liệu`: chạy **Agent 1 thật**.
9. UI hiện Rule Matrix JSON + metrics để QA Lead review/chỉnh.
10. `Chốt Rule Matrix & Tạo testcase`: chạy **Agent 2 thật**.
11. Điều hướng vào `Kho testcase`, hiển thị accordion từ kết quả AI.
12. Có thể tải JSON, Excel, XMind.

## 4. Flow API

Khi tạo/chọn folder API, màn Upload yêu cầu 2 file:
- API Design / Spec
- BA / Business

Sau đó chạy API Agent 1 -> review matrix -> API Agent 2 tương tự.

## 5. Bảo mật API key

Bản này dành cho **local prompt testing**:
- full API key không được commit vào source;
- full key chỉ nằm trong `sessionStorage` của tab hiện tại;
- backend nhận key theo request và không lưu vào `RUN_STORE`.

Khi production nên chuyển key sang backend secret manager/vault và session server-side.

## 6. Build production

```powershell
.\build-prod.ps1
cd backend
py -m uvicorn app:app --host 0.0.0.0 --port 8000
```

Sau khi `frontend/dist` tồn tại, FastAPI sẽ serve luôn React SPA và API trên port 8000.

## Ghi chú
- Backend giữ `RUN_STORE`/cache in-memory để test nhanh. Restart backend sẽ mất kết quả run.
- `.doc` cũ phụ thuộc `antiword`; nên ưu tiên `.docx`, PDF, MD, TXT.
- Source AI gốc được giữ tại `docs/multiA_original.txt` để đối chiếu.


## V1.1 — Fix kiểm tra kết nối Qwen

- Connection test tăng output budget từ 16 lên 1024 tokens, có fallback 4096 cho reasoning model.
- Tránh false-negative với Qwen 3.8 Max khi reasoning_content tiêu thụ token trước content.
- Storage key frontend được version hóa (`.v2`) để không nạp nhầm Base URL/API config cũ từ các bản localhost trước.
- Thêm nút **Khôi phục mặc định** ở màn Quản lý khóa API.
- Base URL mặc định: `https://token-plan.ap-southeast-1.maas.aliyuncs.com/compatible-mode/v1`.

## V1.2 — Project persistence & redirect
- Dự án mới được lưu trong `localStorage` và xuất hiện ngay ở sidebar.
- Sau khi bấm **Tạo dự án**, UI chuyển thẳng vào màn **Chi tiết dự án** của project vừa tạo.
- Tên/mô tả/loại project và danh sách folder được dùng thật ở Project Detail.
- Folder tạo mới trong Project Detail cũng được lưu vào project trước khi chuyển sang Upload.
- Breadcrumb ở Upload/Kho testcase hiển thị đúng tên project hiện tại.


## V1.3 — Single AI Stage + SQLite testcase store

### Thay đổi kiến trúc

AI chỉ còn **một nhiệm vụ**:

`Tài liệu -> Agent 1 / QA Brain -> Rule Matrix JSON`

Rule Matrix JSON vẫn giữ nguyên schema của pipeline cũ nhưng **không hiển thị trên UI**.

Sau đó backend làm hoàn toàn bằng code:

`Rule Matrix -> validate -> map testcase -> SQLite -> UI review -> Excel/XMind`

Không còn AI Agent 2 trong flow UI.

### DB

SQLite tự tạo tại:

`backend/data/testpilot.db`

Testcase được lưu theo:
- project_id
- folder_id
- scope web/api
- run_id

Thêm/Sửa/Xóa testcase trên UI cập nhật trực tiếp DB.

### Excel

Excel không còn phụ thuộc output AI/XMind. Backend dựng file `.xlsx` trực tiếp từ testcase trong DB, theo đúng thứ tự cột:

`ID -> Name -> PreConditions -> Importance -> Step -> Data test -> Expected Result -> Actual Result -> ... review/automation`

### UI review

- Testcase được nhóm theo `feature_group + feature_name`.
- Trong cùng một field/feature, case `VALIDATION` được ưu tiên trước.
- Expanded card hiển thị các field theo thứ tự Excel.
- Backend validate các field bắt buộc: ID, Name, Step, Expected Result, Type.


### Kiểm tra build trong gói này

- Backend compile: OK
- SQLite CRUD smoke test: OK
- Excel generation smoke test: OK


## V1.4 — Folder routing + workspace endpoint

Base URL mặc định:

`https://ws-2vuxxf5tta2cjplh.ap-southeast-1.maas.aliyuncs.com/compatible-mode/v1`

Rule mở folder:

- testcase `null` / `0` → **Tải tài liệu**
- testcase `> 0` → **Kho testcase**

Count được lấy từ SQLite/backend thay vì tin vào `folder.count` trong localStorage.

Nếu người dùng mở trực tiếp URL `/testcases` của folder rỗng thì UI cũng tự chuyển về Upload.


## V1.5 — Qwen output budget

`QWEN_MAX_OUTPUT_TOKENS = 32768`

Áp dụng cho pipeline AI phân tích tài liệu dùng biến output budget chung.
Health-check kết nối vẫn giữ budget nhỏ riêng, không dùng 32768.


## V1.6 — Senior hierarchy / sequential IDs / source Excel name

- ID hiển thị: `TC_001 -> TC_002 -> ... -> TC_NNN`, đánh lại tự động sau thêm/xóa.
- `sourceRuleId` của AI vẫn được giữ ngầm để traceability.
- Excel download dùng basename của tài liệu nguồn, ví dụ `Trai_phieu_SRS.docx -> Trai_phieu_SRS.xlsx`.
- Thêm màn **Quản lý testcase**: `Project -> Web App/API -> Màn hình -> Testcase`.
- Kho testcase nhận query `screen=` để review riêng một màn hình.
- Prompt Agent 1 giữ nguyên logic Senior; Python dedup được làm bảo thủ hơn để tránh mất rule khác feature/rule_name.
- Backend invariant: số testcase mapping phải bằng `merge.final_rules`.
- UI hiển thị Raw rules / Dedup / Final rules để so với baseline Senior 203 case.
- Giữ nguyên: SQLite, hidden Rule Matrix JSON, Excel/XMind bằng code, workspace endpoint, max_tokens=32768, empty-folder -> Upload.


## V1.7 — Explorer UI + clean generate flow + faster Web reading

Luồng UI chuẩn:

`Project -> Web App/API -> Folder màn hình -> Testcase`

- Không còn tạo screen folder ngang cấp với Web App/API.
- `screen_name` từ Rule Matrix trở thành folder con tự động.
- Click Web App/API luôn điều hướng được; nếu scope chưa có testcase thì tự chuyển sang Tải tài liệu.
- Tải tài liệu và Kho testcase bỏ breadcrumb + text kỹ thuật dư thừa.
- Button/icon được căn giữa và tăng khoảng cách.
- Testcase theo screen được lấy trực tiếp từ SQLite toàn project/scope.
- Excel theo screen vẫn đặt tên theo tài liệu source gần nhất.
- `QWEN_MAX_OUTPUT_TOKENS = 32768`.
- Base URL mặc định giữ workspace endpoint đã xác nhận.
- Web Agent1 chạy tối đa 2 initial semantic chunks song song để giảm thời gian đọc tài liệu.
  Có thể đặt `TESTPILOT_WEB_PARALLEL_WORKERS=1` nếu provider báo rate limit/429.
- Prompt Senior, Rule Matrix schema, recursive split, validation mapping và bảo toàn số rule/testcase được giữ nguyên.
