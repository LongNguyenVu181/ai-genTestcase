# -*- coding: utf-8 -*-
from __future__ import annotations

import re
from typing import Any

import pipeline_core as core

ALLOWED_TYPES = {"Giao diện", "Kiểm tra dữ liệu", "Chức năng", "Ngoại lệ", "Popup", "Luồng"}

WEB_REVIEW_CATEGORY_ORDER = {
    "VALIDATION": 1,
    "UI": 2,
    "ACTION": 3,
    "DATA_GRID": 4,
    "BUSINESS_FLOW": 5,
    "EXCEPTION": 6,
}

API_REVIEW_CATEGORY_ORDER = {
    "REQUEST_VALIDATION": 1,
    "RESPONSE_VALIDATION": 2,
    "AUTHENTICATION": 3,
    "AUTHORIZATION": 4,
    "METHOD_URL": 5,
    "HAPPY_PATH": 6,
    "BUSINESS_RULE": 7,
    "INTEGRATION": 8,
    "EXCEPTION": 9,
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
    if category == "EXCEPTION":
        return "Ngoại lệ"
    if category in {"REQUEST_VALIDATION", "RESPONSE_VALIDATION"}:
        return "Kiểm tra dữ liệu"
    if category in {"AUTHENTICATION", "AUTHORIZATION", "INTEGRATION", "BUSINESS_RULE"}:
        return "Luồng"
    return "Chức năng"


def _precondition(rule: dict, *, api: bool = False) -> str:
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
            endpoint_target = " ".join(
                x for x in (_clean(endpoint.get("method")), _clean(endpoint.get("path"))) if x
            ).strip()
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
                tc = _base_case(
                    tc_id=visible_tc_id,
                    name=name,
                    pre_condition=_precondition(rule, api=True),
                    steps=_steps(rule, module_name, api=True),
                    test_data=_test_data(rule),
                    expected_result=_clean(rule.get("expected_result")),
                    tc_type=_api_type(rule),
                    source_rule_id=source_rule_id,
                    source_requirement=_clean(rule.get("source_requirement")),
                    feature_group="API",
                    feature_name=endpoint_target or module_name,
                    category=category,
                    screen=module_name,
                )
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
