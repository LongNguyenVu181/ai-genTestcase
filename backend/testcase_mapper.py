# -*- coding: utf-8 -*-
from __future__ import annotations

import re
from typing import Any

import pipeline_core as core

ALLOWED_TYPES = {
    "Giao diện", "Kiểm tra dữ liệu", "Chức năng", "Ngoại lệ", "Popup", "Luồng",
    "Auth", "Permission", "Validation", "Happy Path", "Business Rule",
}

WEB_REVIEW_CATEGORY_ORDER = {
    "UI": 1,
    "VALIDATION": 2,
    "ACTION": 3,
    "BUSINESS_FLOW": 4,
    "DATA_GRID": 5,
    "EXCEPTION": 6,
}

API_REVIEW_CATEGORY_ORDER = {
    "AUTH": 1,
    "PERMISSION": 2,
    "VALIDATION": 3,
    "HAPPY_PATH": 4,
    "BUSINESS_RULE": 5,
}

API_CATEGORY_TITLES = {
    "AUTH": "1. Kiểm tra xác thực và token",
    "PERMISSION": "2. Kiểm tra phân quyền",
    "VALIDATION": "3. Kiểm tra Validate Input",
    "HAPPY_PATH": "4. Kiểm tra Luồng chính (Happy Path)",
    "BUSINESS_RULE": "5. Kiểm tra Luồng nghiệp vụ (Business Rule)",
}

EXTRA_BLANK_FIELDS = {
    "importance": "",
    "actualResult": "",
    "run1": "",
    "run2": "",
    "run3": "",
    "currentResult": "",
    "note": "",
    "errorCode": "",
    "qcWriter": "",
    "sprintWriter": "",
    "qcExecutor": "",
    "sprintExecutor": "",
    "reviewer": "",
    "reviewDate": "",
    "reviewContent": "",
    "needAuto": "",
    "automated": "",
    "smoke": "",
    "regression": "",
    "outdated": "",
    "outdatedDate": "",
}


def _clean(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def _display_mapping_term(value: Any) -> str:
    return re.sub(r"(?i)\bánh\s+xạ\b", "Mapping", str(value or "")).strip()


def _humanize_generated_text(value: Any) -> str:
    """Remove generator/meta wording from tester-facing content without changing business meaning."""
    text = str(value or "")
    if not text:
        return ""
    text = _display_mapping_term(text)
    patterns = (
        r"(?i)\bMục\s+tiêu\s+kiểm\s+thử\s*[:：-]?\s*",
        r"(?i)\bTest\s+objective\s*[:：-]?\s*",
        r"(?i)\btheo\s+mục\s+tiêu\s+kiểm\s+thử\b",
        r"(?i)\b(?:do|được)\s+AI\s+(?:sinh|tạo|đề\s+xuất)\b",
        r"(?i)\btheo\s+(?:phân\s+tích\s+)?(?:của\s+)?AI\s*[,;:]?\s*",
        r"(?i)\bAI\s+(?:đề\s+xuất|suy\s+luận|xác\s+định|phân\s+tích|sinh|tạo)\b",
        r"(?i)\bAI[- ]generated\b",
        r"(?i)\bgenerated\s+by\s+AI\b",
        r"(?i)\bthe\s+model\s+(?:suggests|generated|inferred)\b",
    )
    for pattern in patterns:
        text = re.sub(pattern, "", text)
    text = re.sub(r"\(\s*AI\s*\)|\[\s*AI\s*\]", "", text, flags=re.I)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\s+([,.;:])", r"\1", text)
    text = re.sub(r"([:;,-])\s*\1+", r"\1", text)
    return text.strip(" \t-–—:;,.")


def _web_feature_name(rule: dict) -> str:
    group = _clean(rule.get("feature_group")).upper()
    name = _clean(rule.get("feature_name"))
    if group == "UI":
        if core._normalize_text(name) in {"permission", "quyen truy cap", "điều kiện truy cập", "dieu kien truy cap"}:
            return "Permission"
        return name or "Giao diện chung"
    if group == "DATA_GRID":
        norm = core._normalize_text(name)
        if not norm or norm.startswith(("cot ", "cột ")) or "mapping" in norm or ("anh xa" in norm or "ánh xạ" in norm):
            return "Data Grid"
    return name or {
        "VALIDATE": "Validation",
        "FUNCTION": "Chức năng",
        "POPUP": "Popup",
        "EXCEPTION": "Ngoại lệ",
    }.get(group, "Khác")


def _web_type(rule: dict) -> str:
    group = _clean(rule.get("feature_group")).upper()
    if group == "POPUP":
        return "Popup"
    if group == "EXCEPTION":
        return "Ngoại lệ"
    if group == "VALIDATE":
        return "Kiểm tra dữ liệu"
    if group == "UI":
        return "Giao diện"
    category = _clean(rule.get("category")).upper()
    return {
        "UI": "Giao diện",
        "VALIDATION": "Kiểm tra dữ liệu",
        "ACTION": "Chức năng",
        "DATA_GRID": "Chức năng",
        "BUSINESS_FLOW": "Luồng",
        "EXCEPTION": "Ngoại lệ",
    }.get(category, "Chức năng")


def _api_type(rule: dict) -> str:
    category = _clean(rule.get("category")).upper()
    return {
        "AUTH": "Auth",
        "PERMISSION": "Permission",
        "VALIDATION": "Validation",
        "HAPPY_PATH": "Happy Path",
        "BUSINESS_RULE": "Business Rule",
    }.get(category, "Business Rule")


def _precondition(rule: dict, *, api: bool = False) -> str:
    if api and _clean(rule.get("precondition")):
        return str(rule.get("precondition") or "").strip()
    condition = _clean(rule.get("test_condition"))
    if not condition:
        return ""
    category = _clean(rule.get("category")).upper()
    feature_group = _clean(rule.get("feature_group")).upper()
    markers = (
        "tiền điều kiện", "pre-condition", "precondition", "đã đăng nhập", "có quyền",
        "vai trò", "role", "trạng thái hiện tại", "đang ở trạng thái", "đã tồn tại",
        "dependency", "phụ thuộc",
    )
    if feature_group in {"PRECONDITION_PERMISSION", "UI"} and _web_feature_name(rule) == "Permission":
        return condition
    if category in {"AUTHENTICATION", "AUTHORIZATION"} and api:
        return condition
    lowered = condition.casefold()
    return condition if any(m in lowered for m in markers) else ""


def _test_data(rule: dict) -> str:
    if _clean(rule.get("test_data")):
        return str(rule.get("test_data") or "").strip()
    condition = _clean(rule.get("test_condition"))
    if not condition:
        return ""
    category = _clean(rule.get("category")).upper()
    if category in {"VALIDATION", "REQUEST_VALIDATION", "RESPONSE_VALIDATION"}:
        return condition
    lowered = condition.casefold()
    data_markers = (
        "nhập ", "chọn ", "giá trị", "dữ liệu", "boundary", "biên", "ký tự", "độ dài",
        "min", "max", "null", "rỗng", "enum", "format", "định dạng", "ngày ", "số ",
        "tham số", "parameter", "header", "body", "request",
    )
    return condition if any(m in lowered for m in data_markers) else ""


def _ensure_sentence(value: str) -> str:
    text = _humanize_generated_text(value).strip()
    if not text:
        return ""
    return text if text.endswith((".", "!", "?")) else text + "."


def _looks_like_action(text: str) -> bool:
    norm = core._normalize_search_text(text)
    verbs = (
        "nhap ", "chon ", "bo chon", "nhan ", "click ", "xoa ", "tai len", "tim kiem",
        "mo ", "dong ", "xac nhan", "chuyen trang", "sap xep", "keo ", "chon file",
    )
    return any(norm.startswith(v) for v in verbs)


def _rule_text(rule: dict) -> str:
    return core._normalize_search_text(" ".join(
        _clean(rule.get(k)) for k in
        ("target", "rule_name", "test_objective", "test_condition", "expected_result", "feature_name")
    ))


def _function_action_name(rule: dict) -> str:
    text = _rule_text(rule)
    candidates = (
        ("dat lai", "Đặt lại"), ("reset", "Đặt lại"),
        ("tim kiem", "Tìm kiếm"), ("search", "Tìm kiếm"),
        ("them moi", "Thêm mới"), ("tao moi", "Thêm mới"), ("create", "Thêm mới"),
        ("phe duyet", "Phê duyệt"), ("approve", "Phê duyệt"),
        ("tu choi", "Từ chối"), ("reject", "Từ chối"),
        ("huy", "Hủy"), ("cancel", "Hủy"),
        ("xac nhan", "Xác nhận"), ("confirm", "Xác nhận"),
        ("hold", "Hold"),
        ("tai len", "Tải lên"), ("upload", "Tải lên"),
        ("tai xuong", "Tải xuống"), ("download", "Tải xuống"),
        ("xem chi tiet", "Xem chi tiết"), ("chi tiet", "Xem chi tiết"),
        ("chinh sua", "Chỉnh sửa"), ("edit", "Chỉnh sửa"),
        ("tao ban sao", "Tạo bản sao"), ("copy", "Tạo bản sao"),
    )
    for key, label in candidates:
        if key in text:
            return label
    target = _humanize_generated_text(_clean(rule.get("target")))
    feature = _humanize_generated_text(_clean(rule.get("feature_name")))
    return target or feature


def _condition_mentions_action(condition: str, action: str) -> bool:
    c = core._normalize_search_text(condition)
    a = core._normalize_search_text(action)
    if not c or not a:
        return False
    return a in c and any(v in c for v in ("nhan", "click", "chon", "thuc hien", "bam", "mo", "tai"))


def _web_feature_sort_rank(tc: dict) -> int:
    group = _clean(tc.get("featureGroup")).upper()
    feature = core._normalize_search_text(tc.get("featureName", ""))
    if group == "UI" and feature == "permission":
        return 0
    if group == "UI" and feature == "giao dien chung":
        return 1
    return 2


def _steps(rule: dict, screen_name: str, *, api: bool = False) -> list[str]:
    target = _humanize_generated_text(_clean(rule.get("target")))
    condition = _humanize_generated_text(_clean(rule.get("test_condition")))
    pre = _precondition(rule, api=api)
    data = _humanize_generated_text(_test_data(rule))
    category = _clean(rule.get("category")).upper()
    group = _clean(rule.get("feature_group")).upper()
    feature = _web_feature_name(rule) if not api else ""
    text = _rule_text(rule)

    steps: list[str] = []
    if api:
        if target:
            steps.append(f"Chuẩn bị request và dữ liệu cho {target}.")
        if condition and condition not in {pre, data}:
            steps.append(_ensure_sentence(condition))
        steps.append("Gửi request đến endpoint tương ứng.")
        steps.append("Kiểm tra response và kết quả nghiệp vụ.")
    else:
        if screen_name:
            steps.append(f"Mở màn hình {screen_name}.")

        # UI / permission: observe or access the concrete element instead of generic 'check UI'.
        if group == "UI":
            if feature == "Permission":
                if condition and condition != pre and _looks_like_action(condition):
                    steps.append(_ensure_sentence(condition))
                steps.append(f"Truy cập màn hình {screen_name}." if screen_name else "Truy cập màn hình.")
                steps.append("Kiểm tra quyền truy cập và nội dung được phép hiển thị.")
            elif "placeholder" in text:
                if target:
                    steps.append(f"Không nhập dữ liệu vào {target}.")
                    steps.append(f"Quan sát nội dung gợi ý tại {target}.")
            elif any(k in text for k in ("mac dinh", "default", "gia tri ban dau", "ban dau")):
                if target:
                    steps.append(f"Quan sát {target} ngay khi màn hình được tải.")
                    steps.append(f"Kiểm tra giá trị mặc định của {target}.")
            elif target:
                steps.append(f"Quan sát {target} trên màn hình.")
                steps.append(f"Đối chiếu nội dung hiển thị của {target} với tài liệu.")

        # Field/control validation: use the concrete condition as the tester action whenever possible.
        elif group == "VALIDATE" or category == "VALIDATION":
            if "placeholder" in text and target:
                steps.append(f"Không nhập dữ liệu vào {target}.")
                steps.append(f"Quan sát nội dung gợi ý tại {target}.")
            elif any(k in text for k in ("mac dinh", "default", "gia tri ban dau", "ban dau")) and target:
                steps.append(f"Quan sát {target} ngay khi màn hình được tải.")
                steps.append(f"Kiểm tra giá trị mặc định của {target}.")
            elif condition and condition != pre:
                steps.append(_ensure_sentence(condition))
                if target:
                    steps.append(f"Kiểm tra giá trị/trạng thái hiển thị tại {target}.")
            elif data and target:
                steps.append(_ensure_sentence(data if _looks_like_action(data) else f"Nhập {data} vào {target}"))
                steps.append(f"Kiểm tra giá trị/trạng thái hiển thị tại {target}.")
            elif target:
                steps.append(f"Thao tác trực tiếp trên {target} theo điều kiện của testcase.")
                steps.append(f"Kiểm tra trạng thái của {target}.")

        # Popup is a container: keep the exact source-grounded trigger/interaction if present.
        elif group == "POPUP":
            popup_name = feature if feature and feature != "Popup" else (target or "popup")
            if condition and condition != pre and _looks_like_action(condition):
                steps.append(_ensure_sentence(condition))
            else:
                steps.append(f"Thực hiện thao tác mở {popup_name}.")
            if category == "VALIDATION" and condition and condition != pre and not _looks_like_action(condition):
                steps.append(_ensure_sentence(condition))
            elif category in {"ACTION", "BUSINESS_FLOW"}:
                action = _function_action_name(rule)
                if action and action.casefold() not in popup_name.casefold():
                    steps.append(f"Nhấn {action} trên {popup_name}.")
            else:
                steps.append(f"Kiểm tra nội dung và trạng thái hiển thị của {popup_name}.")

        # Data Grid: distinguish semantic Mapping, sort and pagination.
        elif group == "DATA_GRID" or category == "DATA_GRID":
            if "mapping" in text:
                col = target or _humanize_generated_text(_clean(rule.get("rule_name"))).replace("Mapping ", "")
                steps.append(f"Quan sát {col} trong danh sách.")
                steps.append(f"Đối chiếu giá trị hiển thị của {col} với dữ liệu bản ghi.")
            elif "sap xep" in text or "sort" in text:
                col = target or feature
                steps.append(f"Nhấn chức năng sắp xếp tại {col}.")
                steps.append(f"Kiểm tra thứ tự dữ liệu của {col} sau khi sắp xếp.")
            elif "phan trang" in text or "pagination" in text or "page size" in text:
                steps.append("Thực hiện thao tác phân trang được mô tả trong tài liệu.")
                steps.append("Kiểm tra trang hiện tại và dữ liệu danh sách sau khi chuyển trang.")
            elif target:
                steps.append(f"Quan sát {target} trong danh sách.")
                steps.append(f"Kiểm tra cách hiển thị của {target}.")

        # Technical exception: execute only the source-described condition.
        elif group == "EXCEPTION" or category == "EXCEPTION":
            if condition and condition != pre:
                steps.append(_ensure_sentence(condition))
            elif target:
                steps.append(f"Thực hiện {target} trong điều kiện lỗi được mô tả trong tài liệu.")
            steps.append("Kiểm tra thông báo/trạng thái lỗi theo tài liệu.")

        # Business function/action: convert the feature into concrete tester verbs.
        else:
            action = _function_action_name(rule)
            condition_is_action = bool(condition and condition != pre and _looks_like_action(condition))
            if condition_is_action:
                steps.append(_ensure_sentence(condition))
            elif data and target and action == "Tìm kiếm" and _looks_like_action(data):
                steps.append(_ensure_sentence(data))

            already_triggered = _condition_mentions_action(condition, action) if action else False
            if action == "Tìm kiếm":
                if not already_triggered:
                    steps.append("Nhấn nút Tìm kiếm.")
                steps.append("Kiểm tra danh sách kết quả trả về.")
            elif action == "Đặt lại":
                if not already_triggered:
                    steps.append("Nhấn nút Đặt lại.")
                steps.append("Kiểm tra các điều kiện tìm kiếm trở về giá trị mặc định.")
            elif action == "Thêm mới":
                if not already_triggered:
                    steps.append("Nhấn nút Thêm mới.")
                steps.append("Kiểm tra màn hình thêm mới được mở đúng theo tài liệu.")
            elif action in {"Phê duyệt", "Từ chối", "Hủy", "Xác nhận", "Hold", "Tải lên", "Tải xuống", "Xem chi tiết", "Chỉnh sửa", "Tạo bản sao"}:
                if not already_triggered:
                    if action == "Tải lên":
                        steps.append("Thực hiện tải tệp lên theo dữ liệu của testcase.")
                    elif action == "Tải xuống":
                        steps.append("Nhấn chức năng Tải xuống.")
                    else:
                        suffix = " trên bản ghi đã chọn" if action in {"Phê duyệt","Từ chối","Hủy","Hold","Xem chi tiết","Chỉnh sửa","Tạo bản sao"} else ""
                        steps.append(f"Nhấn {action}{suffix}.")
                steps.append(f"Kiểm tra kết quả của chức năng {action} theo tài liệu.")
            elif action:
                if not already_triggered:
                    steps.append(f"Thực hiện {action}.")
                steps.append(f"Kiểm tra kết quả của {action} theo tài liệu.")

    # De-duplicate and remove vague machine-like filler.
    banned = {
        "quan sát kết quả", "thiet lap du lieu kiem thu", "thiết lập dữ liệu kiểm thử",
        "ap dung dieu kien kiem thu", "áp dụng điều kiện kiểm thử", "thuc hien thao tac tuong ung",
        "thực hiện thao tác tương ứng", "kiem tra theo yeu cau", "kiểm tra theo yêu cầu",
    }
    result: list[str] = []
    seen = set()
    for step in steps:
        step = _humanize_generated_text(step)
        norm = core._normalize_text(step).strip(" .")
        key = step.casefold().strip()
        if key and norm not in banned and key not in seen:
            seen.add(key)
            result.append(step)
    return result[:5]



def _api_expected_result(rule: dict) -> str:
    """Render one common Expected Result template for every API testcase."""
    rows = [
        ("HTTP Code", rule.get("expected_http_code", "")),
        ("Status", rule.get("expected_status", "")),
        ("Code", rule.get("expected_code", "")),
        ("Message", rule.get("expected_message", "")),
        ("TraceId", rule.get("expected_trace_id", "")),
        ("Data/Body", rule.get("expected_data_body", "")),
        ("Business Result", rule.get("business_result", "")),
    ]
    return "\n".join(f"{label}: {str(value or '').strip()}" for label, value in rows)


def _api_steps(rule: dict, method: str, endpoint_path: str, endpoint_summary: str = "") -> list[str]:
    target = _clean(rule.get("target"))
    raw_data = str(rule.get("test_data") or "")
    effective_method = method
    effective_path = endpoint_path
    if _clean(effective_method).upper() == "UNMAPPED":
        effective_method = ""
    if _clean(effective_path).upper() == "UNMAPPED":
        effective_path = ""

    # Senior-style wrong Method/URL cases keep the mutation in Data test. Reflect it
    # in the actual call step so the generated testcase can be executed directly.
    mm = re.search(r"(?im)^\s*Method\s*:\s*(GET|POST|PUT|PATCH|DELETE|HEAD|OPTIONS)\b", raw_data)
    if mm:
        effective_method = mm.group(1).upper()
    um = re.search(r"(?im)^\s*(?:URL|Endpoint)\s*:\s*(\S+)", raw_data)
    if um:
        effective_path = um.group(1).strip()

    steps = ["Chuẩn bị request và dữ liệu cần thiết."]
    if target:
        steps.append(f"Thiết lập {target} theo dữ liệu đã chuẩn bị.")
    call_target = " ".join(x for x in (effective_method, effective_path) if x).strip()
    if not call_target:
        call_target = _clean(endpoint_summary) or "API tương ứng"
    steps.append(f"Gửi request {call_target}.")
    steps.append("Kiểm tra response và các trường kết quả.")
    if _clean(rule.get("business_result")):
        steps.append("Kiểm tra kết quả nghiệp vụ và side effect liên quan.")
    return steps[:5]

def validate_testcase(tc: dict) -> dict:
    errors: dict[str, str] = {}
    if not _clean(tc.get("id")):
        errors["id"] = "ID không được để trống."
    if not _clean(tc.get("name")):
        errors["name"] = "Tên testcase không được để trống."
    steps = tc.get("steps")
    if not isinstance(steps, list) or not any(_clean(x) for x in steps):
        errors["steps"] = "Testcase phải có ít nhất 1 bước thực hiện."
    if not _clean(tc.get("expectedResult")):
        errors["expectedResult"] = "Kết quả mong đợi không được để trống."
    if _clean(tc.get("type")) not in ALLOWED_TYPES:
        errors["type"] = "Loại testcase không hợp lệ."
    return {"valid": not errors, "errors": errors}


def validate_and_attach(tc: dict) -> dict:
    tc = dict(tc)
    tc["validation"] = validate_testcase(tc)
    return tc


def _base_case(
    *,
    tc_id: str,
    name: str,
    pre_condition: str,
    steps: list[str],
    test_data: str,
    expected_result: str,
    tc_type: str,
    source_rule_id: str,
    source_requirement: str,
    feature_group: str,
    feature_name: str,
    category: str,
    screen: str,
) -> dict:
    tc = {
        "id": tc_id,
        "name": _humanize_generated_text(name),
        "preCondition": _humanize_generated_text(pre_condition),
        "steps": [_humanize_generated_text(step) for step in (steps or []) if _humanize_generated_text(step)],
        "testData": _humanize_generated_text(test_data),
        "expectedResult": _humanize_generated_text(expected_result),
        "type": tc_type,
        "sourceRuleId": source_rule_id,
        "sourceRequirement": source_requirement,
        "featureGroup": feature_group,
        "featureName": feature_name,
        "category": category,
        "screen": screen,
        **EXTRA_BLANK_FIELDS,
    }
    return validate_and_attach(tc)



def order_and_renumber_testcases(testcases: list[dict], scope: str | None = None) -> list[dict]:
    """Return tester-facing order and assign TC_001..TC_N only after ordering is final."""
    items = [dict(tc) for tc in (testcases or [])]
    if not items:
        return items

    resolved_scope = (scope or "").strip().lower()
    if not resolved_scope:
        resolved_scope = "api" if all(_clean(tc.get("featureGroup")).upper() == "API" for tc in items) else "web"

    web_group_order = {
        "UI": 1,
        "VALIDATE": 2,
        "FUNCTION": 3,
        "POPUP": 4,
        "DATA_GRID": 5,
        "EXCEPTION": 6,
    }

    decorated = list(enumerate(items))
    if resolved_scope == "api":
        decorated.sort(key=lambda pair: (
            API_REVIEW_CATEGORY_ORDER.get(_clean(pair[1].get("category")).upper(), 999),
            core._normalize_text(pair[1].get("screen", "")),
            int(pair[1].get("sortOrder") or pair[0]),
            pair[0],
        ))
    else:
        decorated.sort(key=lambda pair: (
            web_group_order.get(_clean(pair[1].get("featureGroup")).upper(), 999),
            _web_feature_sort_rank(pair[1]),
            core._normalize_text(pair[1].get("featureName", "Khác")),
            WEB_REVIEW_CATEGORY_ORDER.get(_clean(pair[1].get("category")).upper(), 999),
            core._normalize_text(pair[1].get("screen", "")),
            int(pair[1].get("sortOrder") or pair[0]),
            pair[0],
        ))

    ordered: list[dict] = []
    for idx, (_, item) in enumerate(decorated, start=1):
        item["id"] = f"TC_{idx:03d}"
        item["sortOrder"] = idx - 1
        item["name"] = _humanize_generated_text(item.get("name"))
        item["preCondition"] = _humanize_generated_text(item.get("preCondition"))
        item["testData"] = _humanize_generated_text(item.get("testData"))
        item["expectedResult"] = _humanize_generated_text(item.get("expectedResult"))
        item["steps"] = [_humanize_generated_text(x) for x in (item.get("steps") or []) if _humanize_generated_text(x)]
        item["validation"] = validate_testcase(item)
        ordered.append(item)
    return ordered

def map_web_matrix_to_testcases(matrix: dict) -> list[dict]:
    result: list[dict] = []
    for screen_index, screen in enumerate(matrix.get("screens", [])):
        screen_name = _clean(screen.get("screen_name")) or f"Màn hình {screen_index + 1}"
        indexed = list(enumerate(screen.get("test_rules", [])))
        indexed.sort(
            key=lambda item: (
                core.WEB_FEATURE_GROUP_ORDER.get(_clean(item[1].get("feature_group")).upper(), 999),
                core._normalize_text(item[1].get("feature_name", "")),
                WEB_REVIEW_CATEGORY_ORDER.get(_clean(item[1].get("category")).upper(), 999),
                item[0],
            )
        )
        for _, rule in indexed:
            source_rule_id = _clean(rule.get("rule_id")) or f"RULE_{len(result)+1:03d}"
            visible_tc_id = f"TC_{len(result)+1:03d}"
            name = _display_mapping_term(_clean(rule.get("rule_name")) or _clean(rule.get("test_objective")) or visible_tc_id)
            tc = _base_case(
                tc_id=visible_tc_id,
                name=name,
                pre_condition=_precondition(rule, api=False),
                steps=_steps(rule, screen_name, api=False),
                test_data=_test_data(rule),
                expected_result=_clean(rule.get("expected_result")),
                tc_type=_web_type(rule),
                source_rule_id=source_rule_id,
                source_requirement=_clean(rule.get("source_requirement")),
                feature_group=_clean(rule.get("feature_group")),
                feature_name=_web_feature_name(rule),
                category=_clean(rule.get("category")).upper(),
                screen=screen_name,
            )
            result.append(tc)
    return order_and_renumber_testcases(result, "web")


def map_api_matrix_to_testcases(matrix: dict) -> list[dict]:
    result: list[dict] = []
    for module_index, module in enumerate(matrix.get("api_modules", [])):
        module_name = _clean(module.get("module_name")) or f"API Module {module_index + 1}"
        for endpoint in module.get("endpoints", []):
            method = _clean(endpoint.get("method")).upper()
            endpoint_path = _clean(endpoint.get("endpoint_path") or endpoint.get("path"))
            endpoint_summary = _clean(endpoint.get("summary"))
            display_method = "" if method == "UNMAPPED" else method
            display_path = "" if endpoint_path.upper() == "UNMAPPED" else endpoint_path
            endpoint_target = " ".join(x for x in (display_method, display_path) if x).strip() or endpoint_summary or module_name
            indexed = list(enumerate(endpoint.get("test_rules", [])))
            indexed.sort(
                key=lambda item: (
                    API_REVIEW_CATEGORY_ORDER.get(_clean(item[1].get("category")).upper(), 999),
                    item[0],
                )
            )
            for _, rule in indexed:
                source_rule_id = _clean(rule.get("rule_id")) or f"RULE_API_{len(result)+1:03d}"
                visible_tc_id = f"TC_{len(result)+1:03d}"
                if not _clean(rule.get("target")) and endpoint_target:
                    rule = {**rule, "target": endpoint_target}
                name = _clean(rule.get("rule_name")) or _clean(rule.get("test_objective")) or visible_tc_id
                category = _clean(rule.get("category")).upper()
                category_title = API_CATEGORY_TITLES.get(category, category or "API")
                tc = _base_case(
                    tc_id=visible_tc_id,
                    name=name,
                    pre_condition=_precondition(rule, api=True),
                    steps=_api_steps(rule, method, endpoint_path, endpoint_summary),
                    test_data=_test_data(rule),
                    expected_result=_api_expected_result(rule),
                    tc_type=_api_type(rule),
                    source_rule_id=source_rule_id,
                    source_requirement=_clean(rule.get("source_requirement")),
                    feature_group="API",
                    feature_name=category_title,
                    category=category,
                    screen=endpoint_target,
                )
                tc["apiModule"] = module_name
                tc["reconciliationStatus"] = _clean(rule.get("reconciliation_status"))
                tc["appliedQaRule"] = _clean(rule.get("applied_qa_rule"))
                result.append(tc)
    return order_and_renumber_testcases(result, "api")


def build_tree_text(testcases: list[dict]) -> str:
    """Create deterministic XMind-compatible bullet tree without AI."""
    lines: list[str] = []
    current_screen = None
    current_group = None
    current_feature = None

    group_title = {
        "UI": "1. UI",
        "VALIDATE": "2. VALIDATE",
        "FUNCTION": "3. FUNCTION",
        "POPUP": "4. POPUP",
        "DATA_GRID": "5. DATA GRID",
        "EXCEPTION": "6. NGOẠI LỆ",
        "API": "API TEST CASE",
    }

    for tc in testcases:
        screen = _clean(tc.get("screen")) or "Test Cases"
        group = _clean(tc.get("featureGroup")) or "FUNCTION"
        feature = _clean(tc.get("featureName")) or "Khác"

        if screen != current_screen:
            lines.append(f"- {screen}")
            current_screen = screen
            current_group = None
            current_feature = None

        if group != current_group:
            lines.append(f"  - {group_title.get(group, group)}")
            current_group = group
            current_feature = None

        if feature != current_feature:
            lines.append(f"    - {feature}")
            current_feature = feature

        lines.append(f"      - {tc.get('id','')} - {tc.get('name','')}")
        indent = "        "
        if _clean(tc.get("preCondition")):
            lines.append(f"{indent}- Pre-condition: {tc['preCondition']}")
            indent += "  "
        step_text = "\\n".join(f"{i+1}. {_clean(step)}" for i, step in enumerate(tc.get("steps", [])))
        lines.append(f"{indent}- Các bước thực hiện: {step_text}")
        indent += "  "
        if _clean(tc.get("testData")):
            lines.append(f"{indent}- Dữ liệu kiểm thử: {tc['testData']}")
            indent += "  "
        lines.append(f"{indent}- Kết quả mong đợi: {tc.get('expectedResult','')}")

    return "\n".join(lines)
