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
    "VALIDATION": 1,
    "UI": 2,
    "ACTION": 3,
    "DATA_GRID": 4,
    "BUSINESS_FLOW": 5,
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


def _web_type(rule: dict) -> str:
    text = " ".join(
        _clean(rule.get(k))
        for k in ("target", "rule_name", "test_objective", "test_condition", "expected_result", "feature_name")
    ).casefold()
    if any(k in text for k in ("popup", "modal", "hộp thoại", "dialog", "xác nhận")):
        return "Popup"
    return {
        "UI": "Giao diện",
        "VALIDATION": "Kiểm tra dữ liệu",
        "ACTION": "Chức năng",
        "DATA_GRID": "Chức năng",
        "BUSINESS_FLOW": "Luồng",
        "EXCEPTION": "Ngoại lệ",
    }.get(_clean(rule.get("category")).upper(), "Chức năng")


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
    if feature_group == "PRECONDITION_PERMISSION":
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


def _steps(rule: dict, screen_name: str, *, api: bool = False) -> list[str]:
    target = _clean(rule.get("target"))
    objective = _clean(rule.get("test_objective")) or _clean(rule.get("rule_name"))
    condition = _clean(rule.get("test_condition"))
    pre = _precondition(rule, api=api)
    data = _test_data(rule)

    steps: list[str] = []
    if api:
        if target:
            steps.append(f"Chuẩn bị đối tượng/request kiểm thử cho {target}.")
        if objective:
            steps.append(objective.rstrip(".") + ".")
        if condition and condition not in {pre, data}:
            steps.append(condition.rstrip(".") + ".")
        steps.append("Gửi request và quan sát response.")
    else:
        if screen_name:
            steps.append(f"Mở màn hình {screen_name}.")
        if target:
            steps.append(f"Thao tác với {target} theo mục tiêu kiểm thử.")
        if objective:
            steps.append(objective.rstrip(".") + ".")
        if condition and condition not in {pre, data}:
            steps.append(condition.rstrip(".") + ".")
        steps.append("Quan sát kết quả.")

    # Preserve order but avoid exact duplicates.
    result: list[str] = []
    seen = set()
    for step in steps:
        key = step.casefold().strip()
        if key and key not in seen:
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

    steps = ["Chuẩn bị request theo PreConditions và Dữ liệu kiểm thử."]
    if target:
        steps.append(f"Áp dụng điều kiện kiểm thử cho {target}.")
    call_target = " ".join(x for x in (effective_method, effective_path) if x).strip()
    if not call_target:
        call_target = _clean(endpoint_summary) or "API tương ứng"
    steps.append(f"Gửi request {call_target}.")
    steps.append("Kiểm tra response theo Kết quả mong đợi.")
    if _clean(rule.get("business_result")):
        steps.append("Kiểm tra Business Result/side effect theo Kết quả mong đợi.")
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
        "name": name,
        "preCondition": pre_condition,
        "steps": steps,
        "testData": test_data,
        "expectedResult": expected_result,
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
            name = _clean(rule.get("rule_name")) or _clean(rule.get("test_objective")) or visible_tc_id
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
                feature_name=_clean(rule.get("feature_name")) or "Khác",
                category=_clean(rule.get("category")).upper(),
                screen=screen_name,
            )
            result.append(tc)
    return result


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
    return result


def build_tree_text(testcases: list[dict]) -> str:
    """Create deterministic XMind-compatible bullet tree without AI."""
    lines: list[str] = []
    current_screen = None
    current_group = None
    current_feature = None

    group_title = {
        "PRECONDITION_PERMISSION": "1. KIỂM TRA TIỀN ĐIỀU KIỆN - PHÂN QUYỀN",
        "GENERAL_UI": "2. KIỂM TRA GIAO DIỆN CHUNG",
        "FILTER": "3. KIỂM TRA BỘ LỌC",
        "DATA_GRID": "4. KIỂM TRA LƯỚI DỮ LIỆU",
        "FUNCTION": "5. KIỂM TRA CHỨC NĂNG",
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
