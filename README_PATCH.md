# TestPilot AI V1.11.1 — Web Human QA Patch

Apply over V1.11.0.

Changes:
- Web presentation template is exactly six sections: UI, VALIDATE, FUNCTION, POPUP, DATA_GRID, EXCEPTION.
- Permission and general screen UI are both under UI; Permission is ordered first.
- Legacy groups PRECONDITION_PERMISSION / GENERAL_UI / FILTER are normalized automatically.
- Popup is a first-class tester section while internal QA category remains intact.
- Web testcase steps use concrete tester verbs and source-grounded actions; generic phrases such as "Thiết lập dữ liệu cho..." and "Quan sát kết quả" are removed.
- Search/Reset/Create/Approve/Cancel/Popup/Grid Mapping/Sort/Pagination/Validation have deterministic human-style step rendering.
- Exact-length N-1/N/N+1 remains three independent testcase records.
- Tester-facing wording continues to remove AI/meta phrases and keeps `Mapping` instead of `Ánh xạ`.
- Web Rule Matrix schema version bumped to 3.5 and cache namespace bumped to avoid stale V1.11.0 output.

Backend compile and deterministic mapper smoke tests passed.
