# TestPilot AI — Combined Patch 1.11.4 + 1.11.5

Patch này đã gộp hai nhánh thay đổi:

## Phần từ 1.11.4 — API source-role / multi-document
- API Spec/Design là nguồn sinh AUTH/PERMISSION/VALIDATION/HAPPY_PATH.
- BA là nguồn sinh BUSINESS_RULE.
- BA error code/message có thể enrich testcase kỹ thuật đã có từ API Spec thay vì sinh case trùng.
- BA không tự tạo API workspace/card mới chỉ vì heading hoặc wording nghiệp vụ khác nhau.
- Giữ strict API target scope và cross-document reconciliation/dedup.

## Phần từ 1.11.5 — Web Discovery -> Inventory -> QA Rules
- Web BA được đọc theo 2 stage: Discovery trước, QA Design sau.
- Discovery list screen -> field/control/grid/popup -> logic tương ứng, chưa sinh testcase.
- Giữ screen ownership qua semantic chunking.
- DOCX được extract đúng thứ tự paragraph/table; giữ heading và strikethrough marker.
- XLSX Web được extract theo sheet/row structure.
- QA testcase chỉ được derive sau khi có canonical inventory.
- Mapper giữ testcase theo từng screen liên tục thay vì sort global theo category.

## File thay thế
Copy đè 3 file sau vào project hiện tại:

- `backend/app.py`
- `backend/pipeline_core.py`
- `backend/testcase_mapper.py`

Sau đó restart backend và chạy `Phân tích tài liệu mới` để tạo run mới.

## Xác nhận merge
- Toàn bộ 34 function API trong `pipeline_core.py` có source hash giống bản 1.11.4.
- `_process_api_analysis_job()` trong `app.py` có source hash giống bản 1.11.4.
- Web code lấy từ bản 1.11.5.
- `python -m py_compile` pass cho cả 3 file trong package này.

Version runtime vẫn là `1.11.5`; đây là package hợp nhất, không phải một nhánh logic API mới.
