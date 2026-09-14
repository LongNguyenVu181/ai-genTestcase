# Prompt audit V1.6

## Kết luận

Prompt Agent 1 không bị viết lại theo UI. Cấu trúc Senior-QA gốc vẫn được giữ: `category` là metadata QA nội bộ, còn `feature_group + feature_name` là trục tổ chức output.

Trong project tích hợp, khác biệt chủ đích duy nhất ở prompt là câu mô tả downstream: Agent 2 đã được thay bằng Python deterministic mapping. Các rule về scope, validation atomicity, Data Grid, Business Flow, Derived rules, duplicate control, schema JSON vẫn giữ nguyên.

## V1.6 xử lý chỗ có thể làm giảm case ngoài prompt

1. Python dedup được làm bảo thủ hơn: signature có thêm `feature_group`, `feature_name`, `rule_name`.
2. Có invariant bắt buộc `final_rules == mapped_testcases`; nếu lệch backend báo lỗi thay vì âm thầm thiếu case.
3. UI hiển thị `raw_rules`, `exact_duplicates_removed`, `final_rules` để đối soát lần chạy tiếp theo với bộ Senior 203 case.
4. Rule Matrix JSON vẫn xử lý ngầm, không show lên UI.
