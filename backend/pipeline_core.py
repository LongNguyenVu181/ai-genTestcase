# -*- coding: utf-8 -*-
"""TestPilot AI document-analysis pipelines used by the FastAPI backend."""
import json
import io
import os
import re
import sys
import tempfile
import time
import traceback
import zipfile
import hashlib
import copy
import unicodedata
import yaml
from dataclasses import dataclass
from difflib import SequenceMatcher
from collections import OrderedDict, Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from pypdf import PdfReader
from datetime import datetime
from openai import OpenAI
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


def _env_int(name: str, default: int, minimum: int = 1) -> int:
    """Read integer env config without making app import fail on a bad local value."""
    try:
        return max(minimum, int(str(os.getenv(name, str(default))).strip()))
    except (TypeError, ValueError):
        return max(minimum, int(default))


def _safe_float(value, default: float = 0.0, minimum: float | None = None, maximum: float | None = None) -> float:
    try:
        result = float(value)
    except (TypeError, ValueError):
        result = float(default)
    if minimum is not None:
        result = max(minimum, result)
    if maximum is not None:
        result = min(maximum, result)
    return result

# ==========================================
# 0. CẤU HÌNH WEB PIPELINE V3.6 — DISCOVERY -> CANONICAL INVENTORY -> QA RULES
# ==========================================
# Agent 1: chia tài liệu nguyên văn thành các semantic chunk nhỏ, có context ở biên.
AGENT1_CHUNK_TARGET_CHARS = 12000
AGENT1_CHUNK_MAX_CHARS = 15000
AGENT1_CONTEXT_CHARS = 1200
AGENT1_MIN_RECURSIVE_CHARS = 2500
AGENT1_MAX_RECURSION_DEPTH = 5
AGENT1_API_RETRIES = 1
# Process independent Web chunks in parallel. Set TESTPILOT_WEB_PARALLEL_WORKERS=1 if provider rate-limit is tight.
AGENT1_PARALLEL_WORKERS = _env_int("TESTPILOT_WEB_PARALLEL_WORKERS", 2)
WEB_QA_PARALLEL_WORKERS = _env_int("TESTPILOT_WEB_QA_PARALLEL_WORKERS", 2)
WEB_DISCOVERY_MAX_OUTPUT_TOKENS = 12288
WEB_QA_MAX_OUTPUT_TOKENS = 24576
WEB_DISCOVERY_VERSION = "1.0"

# Qwen output budget. Nếu vẫn chạm length, pipeline sẽ tự chia nhỏ và retry.
QWEN_MAX_OUTPUT_TOKENS = 32768

# API pipeline: targeted scope scan + deep analysis + recursive split.
API_AGENT1_CHUNK_TARGET_CHARS = 7200
API_AGENT1_CHUNK_MAX_CHARS = 9000
API_AGENT1_CONTEXT_CHARS = 900
# V1.11.3: BA doc is scanned cheaply and may enrich ONLY the API selected by API Design.
# Distinct continuation/sibling APIs are never promoted into separate testcase workspaces.
API_SCOPE_SCAN_CHUNK_TARGET_CHARS = 5000
API_SCOPE_SCAN_CHUNK_MAX_CHARS = 6500
API_SCOPE_SCAN_MAX_TOKENS = 1400
API_SCOPE_SCAN_WORKERS = _env_int("TESTPILOT_API_SCOPE_WORKERS", 2)
API_DEEP_MAX_OUTPUT_TOKENS = 12288
API_JSON_REPAIR_MAX_OUTPUT_TOKENS = 8192
API_AGENT1_MIN_RECURSIVE_CHARS = 2200
API_AGENT1_MAX_RECURSION_DEPTH = 6
API_AGENT1_API_RETRIES = 1
API_ENDPOINT_HINT_LIMIT = 8
AI_REQUEST_TIMEOUT_SECONDS = _env_int("TESTPILOT_AI_REQUEST_TIMEOUT_SECONDS", 600, minimum=60)

# Web Rule Matrix schema — single source of truth for prompt + validator + renderer.
WEB_TEST_DESIGN_VERSION = "3.6"
WEB_RULE_FIELDS = frozenset({
    "rule_id", "target", "category", "feature_group", "feature_name",
    "rule_type", "rule_name", "test_objective", "test_condition",
    "expected_result", "source_requirement", "applied_qa_rule", "generation_reason",
})
WEB_ALLOWED_CATEGORIES = frozenset({
    "UI", "VALIDATION", "ACTION", "DATA_GRID", "BUSINESS_FLOW", "EXCEPTION",
})
WEB_ALLOWED_RULE_TYPES = frozenset({"EXPLICIT", "DERIVED"})

WEB_ALLOWED_FEATURE_GROUPS = frozenset({
    "UI", "VALIDATE", "FUNCTION", "POPUP", "DATA_GRID", "EXCEPTION",
})
WEB_FEATURE_GROUP_ORDER = {
    "UI": 1,
    "VALIDATE": 2,
    "FUNCTION": 3,
    "POPUP": 4,
    "DATA_GRID": 5,
    "EXCEPTION": 6,
}
WEB_FEATURE_GROUP_TITLES = {
    "UI": "1. UI",
    "VALIDATE": "2. VALIDATE",
    "FUNCTION": "3. FUNCTION",
    "POPUP": "4. POPUP",
    "DATA_GRID": "5. DATA GRID",
    "EXCEPTION": "6. NGOẠI LỆ",
}
WEB_CATEGORY_ORDER = {
    "UI": 1,
    "VALIDATION": 2,
    "ACTION": 3,
    "DATA_GRID": 4,
    "BUSINESS_FLOW": 5,
    "EXCEPTION": 6,
}
WEB_TC_TITLE_MAX_CHARS = 120

# Cache namespace: bump when deterministic pipeline semantics change.
WEB_AGENT1_CACHE_NAMESPACE = "web-qa-v3.6-from-canonical-inventory-r1"
WEB_DISCOVERY_CACHE_NAMESPACE = "web-discovery-v1.0-r1"
API_AGENT1_CACHE_NAMESPACE = "api-agent1-v3.11-source-role-r2"

# ==========================================

# ==========================================
# 2. KHAI BÁO CÁC HÀM XỬ LÝ (HELPER FUNCTIONS)
# ==========================================
def log_info(msg: str):
    timestamp = datetime.now().strftime("%H:%M:%S")
    log_msg = f"[LOG {timestamp}] ℹ️ {msg}"
    print(log_msg)
    sys.stdout.flush()

def log_error(msg: str, exc: Exception = None):
    timestamp = datetime.now().strftime("%H:%M:%S")
    err_msg = f"[ERROR {timestamp}] ❌ {msg}"
    if exc:
        err_msg += f"\nDetails: {str(exc)}\n{traceback.format_exc()}"
    print(err_msg)
    sys.stdout.flush()
    return err_msg

def extract_text_from_file(uploaded_file) -> str:
    """Read an uploaded file without consuming its cursor; supports PDF and text-like files."""
    if uploaded_file is None:
        return ""

    if hasattr(uploaded_file, "getvalue"):
        file_bytes = uploaded_file.getvalue()
    else:
        current_pos = None
        try:
            current_pos = uploaded_file.tell()
        except Exception:
            pass
        file_bytes = uploaded_file.read()
        if current_pos is not None:
            try:
                uploaded_file.seek(current_pos)
            except Exception:
                pass

    filename = str(getattr(uploaded_file, "name", "")).lower()
    if filename.endswith(".pdf"):
        reader = PdfReader(io.BytesIO(file_bytes))
        pages = []
        for idx, page in enumerate(reader.pages, start=1):
            page_text = page.extract_text() or ""
            if page_text.strip():
                pages.append(f"--- TRANG {idx} ---\n{page_text}")
        return "\n\n".join(pages)

    for encoding in ("utf-8", "utf-8-sig", "latin-1"):
        try:
            return file_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue
    return file_bytes.decode("utf-8", errors="replace")


def _log_response_summary(agent_name: str, elapsed: float, result: str, chunk_count: int, finish_reason: str | None, reasoning_chunk_count: int = 0):
    """Log chẩn đoán đầy đủ cho request Qwen mà không dump toàn bộ output ra console."""
    log_info(
        f"[{agent_name}] ✅ Request kết thúc | "
        f"Time={elapsed:.2f}s | Chunks={chunk_count} | "
        f"ReasoningChunks={reasoning_chunk_count} | "
        f"OutputChars={len(result):,} | "
        f"FinishReason={finish_reason or 'UNKNOWN'}"
    )


@dataclass
class QwenCallResult:
    ok: bool
    text: str
    finish_reason: str | None
    elapsed: float
    content_chunks: int
    reasoning_chunks: int
    error: str | None = None

    @property
    def complete(self) -> bool:
        return self.ok and self.finish_reason != "length" and bool(self.text.strip())


def _provider_option_error(exc: Exception) -> bool:
    """Only downgrade request options for provider capability/validation errors.

    Network errors, 429s and 5xx must bubble to the normal API retry path instead of
    silently issuing several almost-identical paid requests.
    """
    status = getattr(exc, "status_code", None)
    response = getattr(exc, "response", None)
    if status is None and response is not None:
        status = getattr(response, "status_code", None)
    text = str(exc or "").casefold()
    option_markers = (
        "response_format", "json_schema", "json object", "extra_body",
        "enable_thinking", "unknown parameter", "unsupported parameter",
        "not support", "unsupported", "invalid parameter", "unrecognized",
    )
    return status in {400, 404, 422} and any(marker in text for marker in option_markers)


def _response_format_variants(response_format: dict | None) -> list[tuple[str, dict | None]]:
    """Prefer strict schema, then JSON object, then plain text as compatibility fallbacks."""
    variants: list[tuple[str, dict | None]] = []
    if response_format is not None:
        variants.append(("requested", response_format))
        if str(response_format.get("type", "")).strip() == "json_schema":
            variants.append(("json_object", {"type": "json_object"}))
    variants.append(("plain", None))
    unique: list[tuple[str, dict | None]] = []
    seen = set()
    for label, value in variants:
        key = json.dumps(value, sort_keys=True, ensure_ascii=False) if value is not None else "<plain>"
        if key not in seen:
            seen.add(key)
            unique.append((label, value))
    return unique


def call_qwen_max_agent_detailed(
    content: str,
    api_key: str,
    base_url: str,
    model: str,
    prompt_template: str,
    max_tokens: int = QWEN_MAX_OUTPUT_TOKENS,
    agent_name: str = "QWEN_MAX_AGENT",
    enable_thinking: bool | None = None,
    response_format: dict | None = None,
) -> QwenCallResult:
    """Qwen caller có metadata đầy đủ để pipeline tự xử lý length/retry/split.

    Tầng transport KHÔNG parse JSON và KHÔNG tự cứu output bị truncate.
    """
    client = OpenAI(
        api_key=api_key,
        base_url=base_url,
        timeout=float(AI_REQUEST_TIMEOUT_SECONDS),
        max_retries=0,
    )

    # Chỉ thay placeholder {content}; prompt có nhiều JSON literal với { }.
    full_prompt = prompt_template.replace("{content}", content)
    start_time = time.time()
    log_info(
        f"[{agent_name}] 🚀 Request | Model={model} | MaxTokens={max_tokens} | "
        f"PromptChars={len(full_prompt):,} | ContentChars={len(content):,}"
    )

    try:
        base_request_kwargs = {
            "model": model,
            "messages": [{"role": "user", "content": full_prompt}],
            "temperature": 0.1,
            "max_tokens": max_tokens,
            "stream": True,
        }

        # Prefer strict structured output + non-thinking for extraction. Compatibility
        # fallback is allowed ONLY for provider option errors (400/404/422). Transport,
        # 429 and 5xx errors are handled by the outer retry and never fan out paid calls.
        request_variants: list[tuple[str, bool, dict]] = []
        optional_extra = {"enable_thinking": bool(enable_thinking)} if enable_thinking is not None else None
        for format_label, format_value in _response_format_variants(response_format):
            extra_choices = [True, False] if optional_extra is not None else [False]
            for use_extra_body in extra_choices:
                kwargs = dict(base_request_kwargs)
                if format_value is not None:
                    kwargs["response_format"] = format_value
                if use_extra_body and optional_extra is not None:
                    kwargs["extra_body"] = optional_extra
                request_variants.append((format_label, use_extra_body, kwargs))

        response = None
        last_request_exc = None
        for variant_index, (format_label, use_extra_body, request_kwargs) in enumerate(request_variants, start=1):
            try:
                response = client.chat.completions.create(**request_kwargs)
                if variant_index > 1:
                    log_info(
                        f"[{agent_name}] Provider option fallback thành công | "
                        f"format={format_label} | enable_thinking_flag={use_extra_body}"
                    )
                break
            except Exception as request_exc:
                last_request_exc = request_exc
                if not _provider_option_error(request_exc):
                    raise
                if variant_index < len(request_variants):
                    log_info(
                        f"[{agent_name}] Provider không hỗ trợ option hiện tại; thử compatibility fallback | "
                        f"format={format_label} | enable_thinking_flag={use_extra_body}"
                    )
                    continue
                raise

        if response is None:
            raise last_request_exc or RuntimeError("Không khởi tạo được response từ provider")

        full_response = []
        chunk_count = 0
        reasoning_chunk_count = 0
        finish_reason = None
        last_log_time = time.time()

        for chunk in response:
            if not chunk.choices:
                continue

            choice = chunk.choices[0]
            delta = choice.delta

            if getattr(delta, "reasoning_content", None):
                reasoning_chunk_count += 1

            if delta and delta.content:
                full_response.append(delta.content)
                chunk_count += 1

            if choice.finish_reason:
                finish_reason = choice.finish_reason

            now = time.time()
            if now - last_log_time >= 10:
                elapsed = int(now - start_time)
                log_info(
                    f"[{agent_name}] ⏳ Streaming | Elapsed={elapsed}s | "
                    f"ContentChunks={chunk_count} | ReasoningChunks={reasoning_chunk_count}"
                )
                last_log_time = now

        elapsed_total = time.time() - start_time
        result = "".join(full_response)
        _log_response_summary(
            agent_name,
            elapsed_total,
            result,
            chunk_count,
            finish_reason,
            reasoning_chunk_count
        )

        if finish_reason == "length":
            error = (
                f"OUTPUT BỊ CẮT DO MAX TOKENS | OutputChars={len(result):,} | "
                f"MaxTokens={max_tokens}"
            )
            log_error(f"[{agent_name}] ❌ {error}")
            return QwenCallResult(
                ok=False,
                text=result,
                finish_reason=finish_reason,
                elapsed=elapsed_total,
                content_chunks=chunk_count,
                reasoning_chunks=reasoning_chunk_count,
                error=error,
            )

        if not result.strip():
            error = (
                f"Request kết thúc nhưng Content rỗng | "
                f"FinishReason={finish_reason or 'UNKNOWN'} | "
                f"ReasoningChunks={reasoning_chunk_count}"
            )
            log_error(f"[{agent_name}] ⚠️ {error}")
            return QwenCallResult(
                ok=False,
                text=result,
                finish_reason=finish_reason,
                elapsed=elapsed_total,
                content_chunks=chunk_count,
                reasoning_chunks=reasoning_chunk_count,
                error=error,
            )

        return QwenCallResult(
            ok=True,
            text=result,
            finish_reason=finish_reason,
            elapsed=elapsed_total,
            content_chunks=chunk_count,
            reasoning_chunks=reasoning_chunk_count,
            error=None,
        )

    except Exception as e:
        elapsed_total = time.time() - start_time
        err_details = log_error(
            f"[{agent_name}] ❌ Request failed/timeout after {elapsed_total:.2f}s | "
            f"Model={model} | MaxTokens={max_tokens}",
            e
        )
        return QwenCallResult(
            ok=False,
            text="",
            finish_reason=None,
            elapsed=elapsed_total,
            content_chunks=0,
            reasoning_chunks=0,
            error=err_details,
        )


def extract_json_from_model_response(raw_text: str) -> tuple[bool, dict | list | None, str]:
    """Parse JSON root an toàn từ model response.

    KHÔNG scan các object con. Nếu root JSON bị truncate, phải báo parse fail thật sự,
    tránh nhặt nhầm một test_rule con rồi báo sai schema.
    """
    if raw_text is None:
        return False, None, "Raw response = None"

    text = raw_text.strip()
    if not text:
        return False, None, "Raw response rỗng"

    # Ưu tiên code fence hoàn chỉnh nếu model vẫn trả fence.
    fenced = re.search(r"```(?:json)?\s*(.*?)\s*```", text, flags=re.IGNORECASE | re.DOTALL)
    candidate = fenced.group(1).strip() if fenced else text

    try:
        return True, json.loads(candidate), "JSON parse OK (direct)"
    except json.JSONDecodeError as direct_err:
        direct_error = direct_err

    # Chỉ thử ROOT JSON đầu tiên để hỗ trợ text thừa trước JSON.
    decoder = json.JSONDecoder()
    root_obj = candidate.find("{")
    root_arr = candidate.find("[")
    root_positions = [p for p in (root_obj, root_arr) if p >= 0]

    if root_positions:
        root_start = min(root_positions)
        try:
            parsed, end = decoder.raw_decode(candidate[root_start:])
            trailing = candidate[root_start + end:].strip()
            diagnostic = "JSON parse OK (root raw_decode)"
            if trailing:
                diagnostic += f" | Ignored trailing chars={len(trailing)}"
            return True, parsed, diagnostic
        except json.JSONDecodeError:
            pass

    pos = getattr(direct_error, "pos", None)
    if pos is not None:
        line = getattr(direct_error, "lineno", None)
        col = getattr(direct_error, "colno", None)
        start = max(0, pos - 180)
        end = min(len(candidate), pos + 180)
        preview = candidate[start:end].replace("\n", "\\n")
        diagnostic = (
            f"JSON parse FAILED | {direct_error.msg} | "
            f"pos={pos} | line={line} | col={col} | "
            f"candidate_chars={len(candidate):,} | "
            f"context={preview!r}"
        )
    else:
        diagnostic = f"JSON parse FAILED | candidate_chars={len(candidate):,}"

    return False, None, diagnostic


PROMPT_API_JSON_SYNTAX_REPAIR = r"""
You are a JSON syntax repair tool.

Repair ONLY JSON syntax/serialization problems in INPUT.
STRICT RULES:
- Preserve every business rule and every factual value from INPUT.
- Do NOT add, remove, merge, split, infer, summarize, or rewrite business logic.
- Do NOT invent fields, codes, messages, statuses, endpoints, conditions, or test data.
- You may only fix JSON syntax such as quoting, escaping, commas, brackets/braces, or surrounding prose/fences.
- Return exactly one valid JSON object and nothing else.

INPUT:
{content}
"""


def repair_api_json_syntax(
    raw_text: str,
    *,
    api_key: str,
    base_url: str,
    model: str,
    agent_name: str,
) -> tuple[bool, dict | list | None, str]:
    """One bounded syntax-only repair attempt before source splitting.

    Repair is intentionally isolated from QA generation: it sees the malformed
    output, not the original BA source, so it cannot re-design the testcase set.
    """
    if not str(raw_text or "").strip():
        return False, None, "JSON repair skipped: raw output rỗng"

    repaired = call_qwen_max_agent_detailed(
        content=raw_text,
        api_key=api_key,
        base_url=base_url,
        model=model,
        prompt_template=PROMPT_API_JSON_SYNTAX_REPAIR,
        max_tokens=API_JSON_REPAIR_MAX_OUTPUT_TOKENS,
        agent_name=f"{agent_name}/JSON-REPAIR",
        enable_thinking=False,
        response_format=json_object_response_format(),
    )
    if not repaired.ok or repaired.finish_reason == "length":
        return False, None, (
            f"JSON repair request failed | finish_reason={repaired.finish_reason} | "
            f"error={repaired.error or ''}"
        )

    ok, parsed, diag = extract_json_from_model_response(repaired.text)
    if not ok:
        return False, None, f"JSON repair parse failed | {diag}"
    return True, parsed, f"JSON repair OK | {diag}"


def validate_rule_matrix_schema(data, strict: bool = False) -> tuple[bool, str]:
    """Validate Web Rule Matrix using the centralized Web schema constants."""
    if not isinstance(data, dict):
        return False, f"Root phải là object, nhận {type(data).__name__}"

    version = data.get("test_design_version")
    if strict and version != WEB_TEST_DESIGN_VERSION:
        return False, f"test_design_version không hợp lệ: {version!r}; expected={WEB_TEST_DESIGN_VERSION!r}"

    screens = data.get("screens")
    if not isinstance(screens, list):
        return False, "Thiếu field 'screens' hoặc 'screens' không phải array"

    total_rules = 0

    for sidx, screen in enumerate(screens):
        if not isinstance(screen, dict):
            return False, f"screens[{sidx}] phải là object"

        screen_name = screen.get("screen_name")
        if strict and (not isinstance(screen_name, str) or not screen_name.strip()):
            return False, f"screens[{sidx}].screen_name rỗng/không hợp lệ"

        rules = screen.get("test_rules")
        if not isinstance(rules, list):
            return False, f"screens[{sidx}] thiếu 'test_rules' hoặc không phải array"

        for ridx, rule in enumerate(rules):
            total_rules += 1
            if not isinstance(rule, dict):
                return False, f"screens[{sidx}].test_rules[{ridx}] phải là object"

            missing = sorted(WEB_RULE_FIELDS - set(rule.keys()))
            if missing:
                return False, f"Rule {sidx + 1}.{ridx + 1} thiếu field: {', '.join(missing)}"

            category = str(rule.get("category", "")).strip().upper()
            feature_group = str(rule.get("feature_group", "")).strip().upper()
            rule_type = str(rule.get("rule_type", "")).strip().upper()
            if category not in WEB_ALLOWED_CATEGORIES:
                return False, (
                    f"Rule {sidx + 1}.{ridx + 1} category không hợp lệ: {rule.get('category')!r}; "
                    f"allowed={sorted(WEB_ALLOWED_CATEGORIES)}"
                )
            if feature_group not in WEB_ALLOWED_FEATURE_GROUPS:
                return False, (
                    f"Rule {sidx + 1}.{ridx + 1} feature_group không hợp lệ: {rule.get('feature_group')!r}; "
                    f"allowed={sorted(WEB_ALLOWED_FEATURE_GROUPS)}"
                )
            if not str(rule.get("feature_name", "")).strip():
                return False, f"Rule {sidx + 1}.{ridx + 1} feature_name rỗng"
            if rule_type not in WEB_ALLOWED_RULE_TYPES:
                return False, (
                    f"Rule {sidx + 1}.{ridx + 1} rule_type không hợp lệ: {rule.get('rule_type')!r}; "
                    f"allowed={sorted(WEB_ALLOWED_RULE_TYPES)}"
                )

            if not str(rule.get("source_requirement", "")).strip():
                return False, f"Rule {sidx + 1}.{ridx + 1} source_requirement rỗng"

            if not str(rule.get("expected_result", "")).strip():
                return False, f"Rule {sidx + 1}.{ridx + 1} expected_result rỗng"

            if rule_type == "DERIVED" and not str(rule.get("applied_qa_rule", "")).strip():
                return False, f"Rule {sidx + 1}.{ridx + 1} DERIVED nhưng applied_qa_rule rỗng"

            if strict:
                for field in WEB_RULE_FIELDS:
                    if not isinstance(rule.get(field), str):
                        return False, f"Rule {sidx + 1}.{ridx + 1}.{field} phải là string"

                # rule_id is bookkeeping only. It is assigned deterministically after
                # merge and must never block document analysis/test design.

    return True, f"Schema OK | Screens={len(screens)} | TestRules={total_rules}"


# ==========================================
# 2.1 WEB PIPELINE V3.5 — LONG DOCUMENT BATCHING
# ==========================================

def _schema_failure_can_benefit_from_split(schema_diag: str | None) -> bool:
    """Only split for structural/model-shape failures that may improve with a smaller chunk.

    Semantic enum/traceability failures should fail fast instead of recursively spending tokens.
    """
    diag = str(schema_diag or "")
    non_retryable_markers = (
        "category không hợp lệ",
        "feature_group không hợp lệ",
        "feature_name rỗng",
        "rule_type không hợp lệ",
        "source_requirement rỗng",
        "expected_result rỗng",
        "DERIVED nhưng applied_qa_rule rỗng",
        "source_document không hợp lệ",
        "method không hợp lệ",
    )
    if any(marker in diag for marker in non_retryable_markers):
        return False

    retryable_markers = (
        "Root phải là object",
        "Thiếu field 'screens'",
        "Thiếu field 'api_modules'",
        "không phải array",
        "phải là object",
        "thiếu field:",
        "thiếu endpoints array",
        "thiếu test_rules array",
    )
    return any(marker in diag for marker in retryable_markers)


def _canonical_web_mapping_term(value: str) -> str:
    """Keep the QA term 'Mapping' instead of the awkward Vietnamese literal 'Ánh xạ'."""
    text = str(value or "")
    return re.sub(r"(?i)\bánh\s+xạ\b", "Mapping", text).strip()


def _extract_exact_length_base(rule: dict) -> int | None:
    """Find an explicit exact-length N without guessing unrelated numbers."""
    parts = [
        str(rule.get("source_requirement", "")),
        str(rule.get("rule_name", "")),
        str(rule.get("test_objective", "")),
    ]
    text = " | ".join(parts)
    patterns = (
        r"(?i)(?:độ\s+dài|length)[^0-9]{0,80}(\d+)\s*(?:ký\s*tự|kí\s*tự|characters?|chars?)",
        r"(?i)(?:đúng|chính\s+xác|exact(?:\s+length)?)\s*(\d+)\s*(?:ký\s*tự|kí\s*tự|characters?|chars?)",
        r"(?i)(\d+)\s*(?:ký\s*tự|kí\s*tự|characters?|chars?)[^.|;]{0,50}(?:độ\s+dài|length)",
    )
    for pattern in patterns:
        m = re.search(pattern, text)
        if m:
            try:
                n = int(m.group(1))
                if n > 0:
                    return n
            except ValueError:
                pass
    return None


def _rule_contains_combined_length_boundary(rule: dict, n: int) -> bool:
    """Detect the anti-pattern where N-1/N/N+1 were merged into one Web Rule."""
    text = " | ".join(
        str(rule.get(k, ""))
        for k in ("rule_name", "test_objective", "test_condition", "expected_result")
    )
    norm = _normalize_text(text)
    if "n-1" in norm and "n+1" in norm:
        return True
    values = {int(x) for x in re.findall(r"\b\d+\b", text)}
    return n > 1 and {n - 1, n, n + 1}.issubset(values)


def _split_exact_length_boundary_rule(rule: dict, n: int) -> list[dict]:
    """Deterministically enforce one boundary value = one Rule = one Testcase.

    This intentionally does not invent blocking/truncation/messages. It only states
    whether each value satisfies the explicit exact-length constraint.
    """
    target = str(rule.get("target") or "Trường dữ liệu").strip()
    source_requirement = str(rule.get("source_requirement") or "").strip()
    original_id = str(rule.get("rule_id") or "RULE").strip()
    values = ((n - 1, "N-1"), (n, "N"), (n + 1, "N+1"))
    result: list[dict] = []
    for idx, (value, label) in enumerate(values, start=1):
        item = copy.deepcopy(rule)
        item["rule_id"] = f"{original_id}_B{idx}"
        item["rule_name"] = f"Ràng buộc độ dài {target}: {value} ký tự ({label})"
        item["test_objective"] = f"Kiểm tra {target} với độ dài {value} ký tự ({label})"
        item["test_condition"] = f"Nhập {target} có độ dài {value} ký tự."
        if value == n:
            item["expected_result"] = f"{target} có độ dài {n} ký tự thỏa ràng buộc độ dài chính xác {n} ký tự."
            item["rule_type"] = "EXPLICIT"
            item["applied_qa_rule"] = "EXPLICIT FROM BA"
            item["generation_reason"] = ""
        else:
            item["expected_result"] = (
                f"{target} có độ dài {value} ký tự không thỏa ràng buộc độ dài chính xác {n} ký tự. "
                ""
            )
            item["rule_type"] = "DERIVED"
            item["applied_qa_rule"] = "Boundary Value Analysis"
            item["generation_reason"] = f"Biên {label} được suy ra trực tiếp từ ràng buộc độ dài chính xác {n} ký tự."
        item["source_requirement"] = source_requirement
        result.append(item)
    return result


def normalize_web_rule_matrix_enums(data: dict) -> dict:
    """Canonicalize Web output into the agreed six tester-facing sections.

    Internal QA category is preserved. feature_group is only presentation/ownership:
    UI | VALIDATE | FUNCTION | POPUP | DATA_GRID | EXCEPTION.
    Legacy groups are accepted and deterministically migrated so cached/older output
    can still be rendered without regenerating the document.
    """
    normalized = copy.deepcopy(data)

    def _has_popup_semantics(rule: dict) -> bool:
        text = _normalize_text(" ".join(
            str(rule.get(k, "")) for k in
            ("target", "rule_name", "test_objective", "test_condition", "expected_result", "feature_name")
        ))
        return any(token in text for token in ("popup", "modal", "dialog", "hop thoai", "hộp thoại"))

    for screen in normalized.get("screens", []) if isinstance(normalized, dict) else []:
        if not isinstance(screen, dict):
            continue
        raw_rules = screen.get("test_rules", []) if isinstance(screen.get("test_rules"), list) else []
        canonical_rules: list[dict] = []
        for rule in raw_rules:
            if not isinstance(rule, dict):
                continue

            category = str(rule.get("category", "") or "").strip().upper()
            legacy_group = str(rule.get("feature_group", "") or "").strip().upper()
            feature_name = str(rule.get("feature_name", "") or "").strip()
            rule_type = str(rule.get("rule_type", "") or "").strip().upper()

            rule["category"] = category
            rule["rule_type"] = rule_type

            # Six-section presentation template. Keep QA category independent from section.
            if legacy_group in WEB_ALLOWED_FEATURE_GROUPS:
                group = legacy_group
            elif legacy_group == "PRECONDITION_PERMISSION":
                group = "UI"
                feature_name = "Permission"
            elif legacy_group == "GENERAL_UI":
                group = "UI"
                feature_name = "Giao diện chung"
            elif legacy_group == "FILTER":
                if category == "VALIDATION":
                    group = "VALIDATE"
                elif category == "EXCEPTION":
                    group = "EXCEPTION"
                else:
                    group = "FUNCTION"
            elif legacy_group == "FUNCTION":
                if _has_popup_semantics(rule):
                    group = "POPUP"
                elif category == "EXCEPTION":
                    group = "EXCEPTION"
                elif category == "VALIDATION":
                    group = "VALIDATE"
                else:
                    group = "FUNCTION"
            elif legacy_group == "DATA_GRID":
                group = "DATA_GRID"
            else:
                if _has_popup_semantics(rule):
                    group = "POPUP"
                elif category == "VALIDATION":
                    group = "VALIDATE"
                elif category == "DATA_GRID":
                    group = "DATA_GRID"
                elif category == "EXCEPTION":
                    group = "EXCEPTION"
                elif category == "UI":
                    group = "UI"
                else:
                    group = "FUNCTION"

            # Popup owns its own UI / validation / action rules when the source clearly
            # describes a popup. This is a presentation decision only.
            if _has_popup_semantics(rule) and group not in {"DATA_GRID", "EXCEPTION"}:
                group = "POPUP"

            if group == "UI":
                if not feature_name or legacy_group == "GENERAL_UI":
                    feature_name = "Giao diện chung"
                if legacy_group == "PRECONDITION_PERMISSION":
                    feature_name = "Permission"
            elif group == "VALIDATE":
                feature_name = feature_name or str(rule.get("target") or "Validation").strip() or "Validation"
            elif group == "FUNCTION":
                feature_name = feature_name or str(rule.get("target") or "Chức năng").strip() or "Chức năng"
            elif group == "POPUP":
                feature_name = feature_name or str(rule.get("target") or "Popup").strip() or "Popup"
            elif group == "EXCEPTION":
                feature_name = feature_name or str(rule.get("target") or "Ngoại lệ").strip() or "Ngoại lệ"
            elif group == "DATA_GRID":
                feature_norm = _normalize_text(feature_name)
                if (
                    not feature_norm
                    or feature_norm.startswith(("cot ", "cột "))
                    or "mapping" in feature_norm
                    or ("anh xa" in feature_norm or "ánh xạ" in feature_norm)
                ):
                    feature_name = "Data Grid"

            rule["feature_group"] = group
            rule["feature_name"] = feature_name

            # Keep the technical tester term Mapping; do not translate it to 'Ánh xạ'.
            for field in ("rule_name", "test_objective"):
                if field in rule:
                    rule[field] = _canonical_web_mapping_term(rule.get(field, ""))

            # Hard guardrail: exact-length N-1/N/N+1 must never remain one testcase.
            n = _extract_exact_length_base(rule) if rule.get("category") == "VALIDATION" else None
            if n and _rule_contains_combined_length_boundary(rule, n):
                split_rules = _split_exact_length_boundary_rule(rule, n)
                for item in split_rules:
                    item["feature_group"] = "VALIDATE"
                    item["feature_name"] = feature_name
                canonical_rules.extend(split_rules)
            else:
                canonical_rules.append(rule)

        screen["test_rules"] = canonical_rules
    return normalized


def normalize_api_rule_matrix_enums(data: dict) -> dict:
    """Canonicalize harmless enum spelling drift without changing business meaning."""
    normalized = copy.deepcopy(data)
    category_aliases = {
        "AUTHENTICATION": "AUTH", "AUTHORIZATION": "PERMISSION",
        "VALIDATE": "VALIDATION", "VALIDATE_INPUT": "VALIDATION",
        "HAPPY PATH": "HAPPY_PATH", "HAPPYPATH": "HAPPY_PATH",
        "BUSINESS RULE": "BUSINESS_RULE", "BUSINESSRULE": "BUSINESS_RULE",
        # Legacy/free-form API categories are folded into the agreed Senior 5-group taxonomy.
        "EXCEPTION": "BUSINESS_RULE", "INTEGRATION": "BUSINESS_RULE",
        "METHOD_URL": "AUTH", "METHOD/URL": "AUTH",
    }
    reconciliation_aliases = {
        "DOCUMENT_ONLY": "DOC_ONLY", "DOC ONLY": "DOC_ONLY",
        "COMPLEMENT": "COMPLEMENTARY",
    }
    for module in normalized.get("api_modules", []) if isinstance(normalized, dict) else []:
        if not isinstance(module, dict):
            continue
        for endpoint in module.get("endpoints", []) if isinstance(module.get("endpoints"), list) else []:
            if not isinstance(endpoint, dict):
                continue
            if "method" in endpoint:
                method = str(endpoint.get("method", "") or "").strip().upper()
                endpoint["method"] = method if method in API_HTTP_METHODS else "UNMAPPED"
            for rule in endpoint.get("test_rules", []) if isinstance(endpoint.get("test_rules"), list) else []:
                if not isinstance(rule, dict):
                    continue
                category = str(rule.get("category", "") or "").strip().upper().replace("-", "_")
                rule["category"] = category_aliases.get(category, category)
                rule_type = str(rule.get("rule_type", "") or "").strip().upper()
                rule_type = {"INFERRED": "DERIVED", "GENERATED": "DERIVED", "SOURCE": "EXPLICIT"}.get(rule_type, rule_type)
                rule["rule_type"] = rule_type
                source_document = str(rule.get("source_document", "") or "").strip().upper()
                rule["source_document"] = {"SPEC": "API_SPEC", "API": "API_SPEC", "BUSINESS": "BA"}.get(source_document, source_document)
                rec = str(rule.get("reconciliation_status", "") or "").strip().upper().replace("-", "_")
                rule["reconciliation_status"] = reconciliation_aliases.get(rec, rec)
    return normalized


def normalize_web_rule_shape(data: dict) -> dict:
    """Fill missing serialization keys without inventing requirement meaning."""
    normalized = copy.deepcopy(data) if isinstance(data, dict) else {}
    normalized["test_design_version"] = str(normalized.get("test_design_version") or WEB_TEST_DESIGN_VERSION)
    screens = normalized.get("screens")
    if not isinstance(screens, list):
        return normalized
    for screen in screens:
        if not isinstance(screen, dict):
            continue
        if "screen_name" in screen and not isinstance(screen.get("screen_name"), str):
            screen["screen_name"] = str(screen.get("screen_name") or "")
        rules = screen.get("test_rules")
        if not isinstance(rules, list):
            continue
        for rule in rules:
            if not isinstance(rule, dict):
                continue
            for field in WEB_RULE_FIELDS:
                value = rule.get(field, "")
                rule[field] = "" if value is None else (value if isinstance(value, str) else str(value))
    return normalized


def _infer_web_qa_technique(rule: dict) -> str:
    text = _normalize_search_text(" ".join(str(rule.get(k, "") or "") for k in (
        "rule_name", "test_objective", "test_condition", "source_requirement",
        "generation_reason", "target", "category",
    )))
    if any(marker in text for marker in (
        "n-1", "n+1", "do dai", "length", "minimum", "maximum", "toi thieu", "toi da",
        "ky tu", "character", "boundary", "bien",
    )):
        return "Boundary Value Analysis"
    if any(marker in text for marker in (
        "chi khi", "only when", "dong thoi", "workflow", "trang thai", "status", "visible", "hien thi khi",
    )):
        return "Decision Table"
    return "Equivalence Partitioning"


def normalize_web_rule_semantics(data: dict) -> tuple[dict, list[dict]]:
    """Recover QA-technique metadata only; never repair business/source facts."""
    repairs: list[dict] = []
    if not isinstance(data, dict):
        return data, repairs
    for screen in data.get("screens", []) if isinstance(data.get("screens"), list) else []:
        for rule in screen.get("test_rules", []) if isinstance(screen, dict) and isinstance(screen.get("test_rules"), list) else []:
            if not isinstance(rule, dict):
                continue
            rule_type = str(rule.get("rule_type", "") or "").upper().strip()
            current = str(rule.get("applied_qa_rule", "") or "").strip()
            if rule_type == "EXPLICIT" and not current:
                rule["applied_qa_rule"] = "EXPLICIT FROM BA"
                repairs.append({"rule_id": str(rule.get("rule_id", "")), "field": "applied_qa_rule", "value": "EXPLICIT FROM BA"})
            elif rule_type == "DERIVED" and not current:
                technique = _infer_web_qa_technique(rule)
                rule["applied_qa_rule"] = technique
                if not str(rule.get("generation_reason", "") or "").strip():
                    rule["generation_reason"] = f"Sinh testcase từ ràng buộc nguồn bằng kỹ thuật {technique}."
                repairs.append({"rule_id": str(rule.get("rule_id", "")), "field": "applied_qa_rule", "value": technique})
    return data, repairs


def normalize_api_rule_shape(data: dict) -> dict:
    """Fill missing API serialization keys without inventing source requirements/outcomes."""
    normalized = copy.deepcopy(data) if isinstance(data, dict) else {}
    normalized["api_test_design_version"] = str(normalized.get("api_test_design_version") or "3.3")
    modules = normalized.get("api_modules")
    if not isinstance(modules, list):
        return normalized
    for module in modules:
        if not isinstance(module, dict):
            continue
        if "module_name" in module and not isinstance(module.get("module_name"), str):
            module["module_name"] = str(module.get("module_name") or "")
        endpoints = module.get("endpoints")
        if not isinstance(endpoints, list):
            continue
        for endpoint in endpoints:
            if not isinstance(endpoint, dict):
                continue
            for field in ("method", "endpoint_path", "summary"):
                value = endpoint.get(field, "")
                endpoint[field] = "" if value is None else (value if isinstance(value, str) else str(value))
            rules = endpoint.get("test_rules")
            if not isinstance(rules, list):
                continue
            for rule in rules:
                if not isinstance(rule, dict):
                    continue
                for field in API_RULE_FIELDS:
                    value = rule.get(field, "")
                    rule[field] = "" if value is None else (value if isinstance(value, str) else str(value))
    return normalized


def _normalize_text(value: str) -> str:
    value = unicodedata.normalize("NFKC", str(value or ""))
    value = re.sub(r"\s+", " ", value).strip().casefold()
    return value


def _normalize_search_text(value: str) -> str:
    """Accent-insensitive normalization for lexical routing only.

    Business/source text is never rewritten with this helper; it is used only for
    deterministic overlap/search scores where PDF extraction often splits Vietnamese accents.
    """
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = text.replace("đ", "d").replace("Đ", "D")
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"\s+", " ", text).strip().casefold()
    return text


def _compact_search_text(value: str) -> str:
    """Compact routing text resilient to PDF extraction that splits Vietnamese letters."""
    return re.sub(r"[^a-z0-9]+", "", _normalize_search_text(value))


def _api_operation_family(value: str) -> str:
    """Classify only obvious API-operation headings; unknown text stays UNKNOWN.

    This is a scope safety gate, not business inference. It prevents a selected Approve
    API from accidentally deep-analyzing clearly-labelled List/Reject/Download siblings.
    """
    compact = _compact_search_text(value)
    if not compact:
        return "UNKNOWN"
    # BA documents often call the same approval endpoint "duyệt", "xác nhận duyệt"
    # or "xác nhận duyệt giao dịch". Treat these wording variants as ONE operation family.
    if any(x in compact for x in ("xacnhanduyet", "confirmapproval", "confirmapprove")):
        return "APPROVE"
    if any(x in compact for x in ("tuchoi", "reject", "decline")):
        return "REJECT"
    if any(x in compact for x in ("vanti", "danhsach", "listtransaction", "getlist")):
        return "LIST"
    if any(x in compact for x in ("chitiet", "detail", "getdetail")):
        return "DETAIL"
    if any(x in compact for x in ("taiduthao", "download", "exportfile")):
        return "DOWNLOAD"
    if any(x in compact for x in ("inchungtu", "print")):
        return "PRINT"
    if any(x in compact for x in ("duyetgiaodich", "approval", "approve")):
        return "APPROVE"
    return "UNKNOWN"


def _gate_scope_match(ba_chunk: dict, primary_target: dict, match: dict) -> tuple[dict, dict | None]:
    """Deterministic post-router guard for strict single-target API analysis.

    API Design/Spec is the only source allowed to create an endpoint workspace. BA chunks
    may enrich that selected endpoint, but a distinct callable continuation/sibling API must
    never become a second workspace. The guard also repairs a common wording drift where
    the same approval endpoint is called "duyệt" vs "xác nhận duyệt" in BA headings.
    """
    guarded = dict(match)
    heading = str(ba_chunk.get("title") or "")
    heading_family = _api_operation_family(heading)
    target_family = _api_operation_family(
        f"{primary_target.get('summary', '')} {primary_target.get('endpoint_path', '')}"
    )
    original_role = _normalize_scope_role(guarded.get("scope_role"))
    new_role = original_role
    reason = None

    primary_method, primary_path = _api_endpoint_key(primary_target)
    related_method = str(guarded.get("related_method") or "UNMAPPED").upper().strip() or "UNMAPPED"
    related_path = str(guarded.get("related_endpoint_path") or "UNMAPPED").strip() or "UNMAPPED"
    explicit_distinct_endpoint = (
        related_path != "UNMAPPED"
        and (related_method, related_path) != (primary_method, primary_path)
    )

    # Known sibling heading -> always out of scope for this run.
    if target_family != "UNKNOWN" and heading_family != "UNKNOWN" and heading_family != target_family:
        new_role = "OUT_OF_SCOPE"
        reason = f"heading family {heading_family} differs from target {target_family}"

    # CONTINUATION is retained only for backward-compatible parsing. It is never allowed
    # to create another workspace. If it is clearly just a wording variant of the same
    # target operation and no distinct endpoint is visible, repair it to PRIMARY; otherwise
    # discard it as OUT_OF_SCOPE.
    elif original_role == "CONTINUATION":
        if (
            not explicit_distinct_endpoint
            and target_family != "UNKNOWN"
            and heading_family == target_family
        ):
            new_role = "PRIMARY"
            reason = "continuation label repaired to same target operation"
        else:
            new_role = "OUT_OF_SCOPE"
            reason = "distinct/ambiguous continuation is outside strict target endpoint"

    guarded["scope_role"] = new_role
    if reason and new_role != original_role:
        return guarded, {
            "chunk_id": ba_chunk.get("chunk_id"),
            "title": heading,
            "target_family": target_family,
            "heading_family": heading_family,
            "from_role": original_role,
            "to_role": new_role,
            "reason": reason,
        }
    return guarded, None


def _is_heading_line(line: str) -> bool:
    s = line.strip()
    if not s:
        return False
    if re.match(r"^#{1,6}\s+\S", s):
        return True
    if re.match(r"^\d+(?:\.\d+){1,6}\s+\S", s):
        return True
    if re.match(r"^\d+\.\s+[A-Za-zÀ-ỹ]", s):
        return True
    # Các heading thường gặp trong tài liệu BA/RSD.
    if re.match(
        r"^(màn hình|mh\d*\s*:|logic\s+|luồng\s+|quy tắc nghiệp vụ|mô tả màn hình|"
        r"đặc tả|tóm tắt usecase|ma trận phân quyền|các kết nối|sơ đồ luồng|mockup|"
        r"===\s*(?:source document|sheet)|api\s+design|api\s+spec|ba\s+business|"
        r"get\s+/|post\s+/|put\s+/|patch\s+/|delete\s+/|head\s+/|options\s+/)",
        s,
        flags=re.IGNORECASE,
    ):
        return True
    return False


def _is_screen_heading_line(line: str) -> bool:
    """Detect a real screen heading conservatively; subsection headings are excluded."""
    raw = re.sub(r"^[#\s]+", "", str(line or "").strip())
    raw = re.sub(r"^\d+(?:\.\d+)*[.)]?\s*", "", raw).strip()
    norm = _normalize_search_text(raw)
    if not norm:
        return False
    if norm.startswith(("mo ta man hinh", "logic man hinh", "danh sach man hinh")):
        return False
    return bool(
        re.match(r"(?i)^màn\s+hình\s+\S", raw)
        or re.match(r"(?i)^screen\s+\S", raw)
        or re.match(r"(?i)^mh\d*\s*:\s*\S", raw)
    )


def _safe_split_large_block(text: str, max_chars: int) -> list[str]:
    """Fallback split cho một block quá lớn; ưu tiên biên newline/<br>/câu hơn hard-cut."""
    text = text.strip()
    if len(text) <= max_chars:
        return [text] if text else []

    parts = []
    remaining = text
    min_cut = max(800, int(max_chars * 0.55))

    while len(remaining) > max_chars:
        window = remaining[:max_chars]
        candidates = [
            window.rfind("\n"),
            window.lower().rfind("<br>"),
            window.rfind(". "),
            window.rfind("; "),
            window.rfind(", "),
            window.rfind(" "),
        ]
        cut = max(candidates)
        if cut < min_cut:
            cut = max_chars
        else:
            # giữ delimiter ở phần trước nếu là <br>
            if window[max(0, cut-3):cut+1].lower().endswith("<br>"):
                cut += 1
        part = remaining[:cut].strip()
        if part:
            parts.append(part)
        remaining = remaining[cut:].strip()

    if remaining:
        parts.append(remaining)
    return parts


def _document_to_semantic_blocks(raw_text: str, base_title: str, max_block_chars: int) -> list[dict]:
    """Split into source-order blocks while retaining the nearest real screen ownership hint."""
    lines = raw_text.splitlines()
    blocks = []
    paragraph = []
    current_title = base_title
    current_screen_hint = ""

    def append_block(title: str, text: str, screen_hint: str):
        if text.strip():
            blocks.append({"title": title, "screen_hint": screen_hint, "text": text.strip()})

    def flush_paragraph():
        nonlocal paragraph
        if paragraph:
            block_text = "\n".join(paragraph).strip()
            if block_text:
                append_block(current_title, block_text, current_screen_hint)
            paragraph = []

    for line in lines:
        stripped = line.strip()
        if not stripped:
            flush_paragraph()
            continue

        if _is_heading_line(line):
            flush_paragraph()
            cleaned_title = re.sub(r"^[#\s]+", "", stripped).strip() or current_title
            current_title = cleaned_title
            if _is_screen_heading_line(line):
                current_screen_hint = cleaned_title
            append_block(current_title, line.rstrip(), current_screen_hint)
            continue

        # Markdown table: keep each row intact so a BA requirement row is not cut mid-row.
        if stripped.startswith("|"):
            flush_paragraph()
            append_block(current_title, line.rstrip(), current_screen_hint)
            continue

        paragraph.append(line.rstrip())

    flush_paragraph()

    if not blocks and raw_text.strip():
        blocks = [{"title": base_title, "screen_hint": "", "text": raw_text.strip()}]

    expanded = []
    for block in blocks:
        if len(block["text"]) <= max_block_chars:
            expanded.append(block)
        else:
            for part in _safe_split_large_block(block["text"], max_block_chars):
                expanded.append({
                    "title": block["title"],
                    "screen_hint": block.get("screen_hint", ""),
                    "text": part,
                })
    return expanded


def split_document_semantic(
    raw_text: str,
    base_title: str,
    target_chars: int = AGENT1_CHUNK_TARGET_CHARS,
    max_chars: int = AGENT1_CHUNK_MAX_CHARS,
    context_chars: int = AGENT1_CONTEXT_CHARS,
) -> list[dict]:
    """Pack source blocks without crossing an explicit screen boundary.

    A screen can contain many BA subsections (description, search logic, grid, popup...). Those
    subsections retain `screen_hint` so a later chunk still knows which screen owns its fields.
    """
    blocks = _document_to_semantic_blocks(raw_text, base_title, max_chars)
    core_chunks = []
    current_parts = []
    current_len = 0
    current_title = base_title
    current_screen_hint = ""

    def flush_current():
        nonlocal current_parts, current_len, current_title, current_screen_hint
        if not current_parts:
            return
        core = "\n\n".join(current_parts).strip()
        if core:
            core_chunks.append({
                "title": current_title,
                "screen_hint": current_screen_hint,
                "core_text": core,
                "char_count": len(core),
            })
        current_parts = []
        current_len = 0

    for block in blocks:
        block_text = block["text"].strip()
        if not block_text:
            continue
        block_screen_hint = block.get("screen_hint", "") or ""
        is_new_screen_heading = _is_screen_heading_line(block_text.splitlines()[0] if block_text.splitlines() else "")

        # Never mix two explicit screen roots in one model chunk.
        if is_new_screen_heading and current_parts:
            flush_current()
            current_title = block.get("title") or base_title
            current_screen_hint = block_screen_hint

        block_len = len(block_text) + (2 if current_parts else 0)
        if not current_parts:
            current_title = block.get("title") or base_title
            current_screen_hint = block_screen_hint
            current_parts = [block_text]
            current_len = len(block_text)
            continue

        proposed = current_len + block_len
        if proposed > max_chars or (current_len >= target_chars and proposed > target_chars):
            previous_screen_hint = current_screen_hint
            flush_current()
            current_title = block.get("title") or base_title
            current_screen_hint = block_screen_hint or previous_screen_hint
            current_parts = [block_text]
            current_len = len(block_text)
        else:
            current_parts.append(block_text)
            current_len = proposed
            if block_screen_hint:
                current_screen_hint = block_screen_hint
            if _is_heading_line(block_text.splitlines()[0] if block_text.splitlines() else ""):
                current_title = block.get("title") or current_title

    flush_current()

    chunks = []
    for idx, item in enumerate(core_chunks):
        prev_context = core_chunks[idx - 1]["core_text"][-context_chars:] if idx > 0 else ""
        next_context = core_chunks[idx + 1]["core_text"][:context_chars] if idx + 1 < len(core_chunks) else ""
        chunks.append({
            "chunk_id": f"C{idx + 1:03d}",
            "title": item["title"],
            "screen_hint": item.get("screen_hint", ""),
            "core_text": item["core_text"],
            "prev_context": prev_context,
            "next_context": next_context,
            "char_count": item["char_count"],
            "depth": 0,
        })

    return chunks


def build_agent1_chunk_payload(chunk: dict) -> str:
    """Đóng gói raw chunk + boundary context. Metadata không phải requirement."""
    return f"""=== CHUNK METADATA — KHÔNG PHẢI REQUIREMENT ===
chunk_id: {chunk.get('chunk_id')}
source_section_hint: {chunk.get('title')}
chunk_depth: {chunk.get('depth', 0)}

=== ACTIVE SCREEN HEADING — TRÍCH TỪ NGUỒN, CHỈ DÙNG ĐỂ GÁN OWNERSHIP ===
{chunk.get('screen_hint', '')}

=== PREVIOUS CONTEXT — CHỈ DÙNG ĐỂ HIỂU, KHÔNG SINH RULE CHỈ TỪ ĐOẠN NÀY ===
{chunk.get('prev_context', '')}

=== CURRENT SOURCE — PHẠM VI CHÍNH ĐƯỢC PHÉP SINH TEST RULE ===
{chunk.get('core_text', '')}

=== NEXT CONTEXT — CHỈ DÙNG ĐỂ HIỂU, KHÔNG SINH RULE CHỈ TỪ ĐOẠN NÀY ===
{chunk.get('next_context', '')}
"""



WEB_INVENTORY_REQ_FIELDS = ("property", "requirement", "source_requirement")


def _stringify(value) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    return str(value).strip()


def _normalize_inventory_requirement(item) -> dict:
    if isinstance(item, str):
        text = item.strip()
        return {"property": "OTHER", "requirement": text, "source_requirement": text}
    if not isinstance(item, dict):
        text = _stringify(item)
        return {"property": "OTHER", "requirement": text, "source_requirement": text}
    requirement = _stringify(item.get("requirement") or item.get("behavior") or item.get("description"))
    source_requirement = _stringify(item.get("source_requirement") or item.get("source") or requirement)
    return {
        "property": _stringify(item.get("property") or item.get("kind") or "OTHER").upper() or "OTHER",
        "requirement": requirement,
        "source_requirement": source_requirement,
    }


def _normalize_inventory_requirement_list(value) -> list[dict]:
    if value is None:
        return []
    if not isinstance(value, list):
        value = [value]
    result = []
    for item in value:
        req = _normalize_inventory_requirement(item)
        if req["requirement"] or req["source_requirement"]:
            result.append(req)
    return result


def normalize_web_inventory_shape(data: dict) -> dict:
    """Coerce small model shape deviations without inventing requirements."""
    if not isinstance(data, dict):
        return {"web_discovery_version": WEB_DISCOVERY_VERSION, "screens": []}
    raw_screens = data.get("screens")
    if not isinstance(raw_screens, list):
        raw_screens = []
    screens = []
    for raw in raw_screens:
        if not isinstance(raw, dict):
            continue
        screen = {
            "screen_name": _stringify(raw.get("screen_name") or raw.get("name")),
            "screen_code": _stringify(raw.get("screen_code") or raw.get("code")),
            "source_evidence": _stringify(raw.get("source_evidence") or raw.get("evidence")),
            "screen_requirements": _normalize_inventory_requirement_list(raw.get("screen_requirements") or raw.get("requirements")),
            "fields": [],
            "controls": [],
            "grids": [],
            "popups": [],
            "logic": [],
        }
        for item in raw.get("fields", []) if isinstance(raw.get("fields"), list) else []:
            if not isinstance(item, dict):
                continue
            screen["fields"].append({
                "field_name": _stringify(item.get("field_name") or item.get("name")),
                "control_type": _stringify(item.get("control_type") or item.get("type")),
                "container": _stringify(item.get("container") or "SCREEN") or "SCREEN",
                "requirements": _normalize_inventory_requirement_list(item.get("requirements") or item.get("rules")),
            })
        for item in raw.get("controls", []) if isinstance(raw.get("controls"), list) else []:
            if not isinstance(item, dict):
                continue
            screen["controls"].append({
                "control_name": _stringify(item.get("control_name") or item.get("name")),
                "control_type": _stringify(item.get("control_type") or item.get("type")),
                "container": _stringify(item.get("container") or "SCREEN") or "SCREEN",
                "requirements": _normalize_inventory_requirement_list(item.get("requirements") or item.get("rules")),
            })
        for item in raw.get("grids", []) if isinstance(raw.get("grids"), list) else []:
            if not isinstance(item, dict):
                continue
            columns = []
            for col in item.get("columns", []) if isinstance(item.get("columns"), list) else []:
                if isinstance(col, str):
                    columns.append({"column_name": col.strip(), "mapping": "", "presentation": "", "source_requirement": col.strip()})
                elif isinstance(col, dict):
                    columns.append({
                        "column_name": _stringify(col.get("column_name") or col.get("name")),
                        "mapping": _stringify(col.get("mapping") or col.get("value")),
                        "presentation": _stringify(col.get("presentation") or col.get("format")),
                        "source_requirement": _stringify(col.get("source_requirement") or col.get("source") or col.get("description")),
                    })
            screen["grids"].append({
                "grid_name": _stringify(item.get("grid_name") or item.get("name") or "Data Grid") or "Data Grid",
                "columns": columns,
                "requirements": _normalize_inventory_requirement_list(item.get("requirements") or item.get("rules")),
            })
        for item in raw.get("popups", []) if isinstance(raw.get("popups"), list) else []:
            if not isinstance(item, dict):
                continue
            screen["popups"].append({
                "popup_name": _stringify(item.get("popup_name") or item.get("name")),
                "requirements": _normalize_inventory_requirement_list(item.get("requirements") or item.get("rules")),
            })
        for item in raw.get("logic", []) if isinstance(raw.get("logic"), list) else []:
            if not isinstance(item, dict):
                continue
            targets = item.get("targets")
            if isinstance(targets, str):
                targets = [t.strip() for t in re.split(r"[,;|]", targets) if t.strip()]
            elif not isinstance(targets, list):
                targets = []
            screen["logic"].append({
                "logic_name": _stringify(item.get("logic_name") or item.get("name")),
                "logic_type": _stringify(item.get("logic_type") or item.get("type") or "BUSINESS_FLOW").upper() or "BUSINESS_FLOW",
                "targets": [_stringify(t) for t in targets if _stringify(t)],
                "condition": _stringify(item.get("condition")),
                "behavior": _stringify(item.get("behavior") or item.get("outcome") or item.get("requirement")),
                "source_requirement": _stringify(item.get("source_requirement") or item.get("source") or item.get("behavior") or item.get("outcome")),
            })
        if screen["screen_name"]:
            screens.append(screen)
    return {"web_discovery_version": WEB_DISCOVERY_VERSION, "screens": screens}


def validate_web_inventory_schema(data: dict) -> tuple[bool, str]:
    if not isinstance(data, dict):
        return False, "Inventory root phải là object"
    screens = data.get("screens")
    if not isinstance(screens, list):
        return False, "Inventory thiếu screens array"
    object_count = 0
    requirement_count = 0
    for sidx, screen in enumerate(screens):
        if not isinstance(screen, dict) or not _stringify(screen.get("screen_name")):
            return False, f"Inventory screens[{sidx}] thiếu screen_name"
        for key in ("screen_requirements", "fields", "controls", "grids", "popups", "logic"):
            if not isinstance(screen.get(key), list):
                return False, f"Inventory screens[{sidx}].{key} phải là array"
        requirement_count += len(screen.get("screen_requirements", []))
        for key in ("fields", "controls", "grids", "popups"):
            object_count += len(screen.get(key, []))
            for item in screen.get(key, []):
                if not isinstance(item, dict):
                    return False, f"Inventory screens[{sidx}].{key} chứa phần tử không phải object"
                reqs = item.get("requirements", [])
                if not isinstance(reqs, list):
                    return False, f"Inventory screens[{sidx}].{key}.requirements phải là array"
                requirement_count += len(reqs)
        for logic in screen.get("logic", []):
            if not isinstance(logic, dict):
                return False, f"Inventory screens[{sidx}].logic chứa phần tử không phải object"
            if not isinstance(logic.get("targets", []), list):
                return False, f"Inventory screens[{sidx}].logic.targets phải là array"
            requirement_count += 1
    return True, f"Inventory OK | Screens={len(screens)} | Objects={object_count} | Requirements={requirement_count}"


def _inventory_req_sig(req: dict) -> tuple:
    return (
        _normalize_text(req.get("property", "")),
        _normalize_text(req.get("requirement", "")),
        _normalize_text(req.get("source_requirement", "")),
    )


def _merge_req_lists(dst: list[dict], src: list[dict]) -> None:
    seen = {_inventory_req_sig(x) for x in dst}
    for req in src:
        sig = _inventory_req_sig(req)
        if sig not in seen:
            dst.append(copy.deepcopy(req))
            seen.add(sig)


def _merge_named_inventory_items(dst: list[dict], src: list[dict], name_key: str) -> None:
    by_key = {(_normalize_text(item.get(name_key, "")), _normalize_text(item.get("container", ""))): item for item in dst}
    for item in src:
        key = (_normalize_text(item.get(name_key, "")), _normalize_text(item.get("container", "")))
        if not key[0]:
            continue
        existing = by_key.get(key)
        if existing is None:
            dst.append(copy.deepcopy(item))
            by_key[key] = dst[-1]
            continue
        if not _stringify(existing.get("control_type")) and _stringify(item.get("control_type")):
            existing["control_type"] = item.get("control_type", "")
        _merge_req_lists(existing.setdefault("requirements", []), item.get("requirements", []))


def _merge_grid_items(dst: list[dict], src: list[dict]) -> None:
    by_key = {_normalize_text(item.get("grid_name", "")): item for item in dst}
    for item in src:
        key = _normalize_text(item.get("grid_name", "")) or "data grid"
        existing = by_key.get(key)
        if existing is None:
            dst.append(copy.deepcopy(item))
            by_key[key] = dst[-1]
            continue
        _merge_req_lists(existing.setdefault("requirements", []), item.get("requirements", []))
        col_seen = {
            (_normalize_text(c.get("column_name", "")), _normalize_text(c.get("mapping", "")), _normalize_text(c.get("presentation", "")))
            for c in existing.setdefault("columns", [])
        }
        for col in item.get("columns", []):
            sig = (_normalize_text(col.get("column_name", "")), _normalize_text(col.get("mapping", "")), _normalize_text(col.get("presentation", "")))
            if sig not in col_seen:
                existing["columns"].append(copy.deepcopy(col))
                col_seen.add(sig)


def _merge_logic_items(dst: list[dict], src: list[dict]) -> None:
    def sig(item: dict) -> tuple:
        return (
            _normalize_text(item.get("logic_name", "")),
            _normalize_text(item.get("logic_type", "")),
            tuple(sorted(_normalize_text(x) for x in item.get("targets", []) if _normalize_text(x))),
            _normalize_text(item.get("condition", "")),
            _normalize_text(item.get("behavior", "")),
        )
    seen = {sig(x) for x in dst}
    for item in src:
        key = sig(item)
        if key not in seen:
            dst.append(copy.deepcopy(item))
            seen.add(key)


def merge_web_inventories(inventories: list[dict]) -> tuple[dict, dict]:
    """Build one canonical screen/component/logic inventory before any QA rule is applied."""
    merged: OrderedDict[str, dict] = OrderedDict()
    meta: dict[str, dict] = {}
    aliases = 0
    for inv in inventories:
        normalized = normalize_web_inventory_shape(inv)
        for screen in normalized.get("screens", []):
            name = _stringify(screen.get("screen_name"))
            code = _stringify(screen.get("screen_code"))
            key = None
            if code:
                code_norm = _normalize_text(code)
                for k, m in meta.items():
                    if m.get("code_norm") == code_norm:
                        key = k
                        break
                if key is None:
                    key = f"code::{code_norm}"
            if key is None:
                exact = f"name::{_normalize_text(name)}"
                if exact in merged:
                    key = exact
            if key is None:
                for k, m in meta.items():
                    if _screen_alias_equivalent(name, m.get("screen_name", "")):
                        key = k
                        aliases += 1
                        break
            if key is None:
                key = f"name::{_normalize_text(name)}"
            if key not in merged:
                merged[key] = copy.deepcopy(screen)
                meta[key] = {"screen_name": name, "code_norm": _normalize_text(code)}
                continue
            cur = merged[key]
            cur["screen_name"] = _prefer_human_screen_name(cur.get("screen_name", ""), name)
            if not cur.get("screen_code") and code:
                cur["screen_code"] = code
                meta[key]["code_norm"] = _normalize_text(code)
            if len(_stringify(screen.get("source_evidence"))) > len(_stringify(cur.get("source_evidence"))):
                cur["source_evidence"] = screen.get("source_evidence", "")
            _merge_req_lists(cur.setdefault("screen_requirements", []), screen.get("screen_requirements", []))
            _merge_named_inventory_items(cur.setdefault("fields", []), screen.get("fields", []), "field_name")
            _merge_named_inventory_items(cur.setdefault("controls", []), screen.get("controls", []), "control_name")
            _merge_grid_items(cur.setdefault("grids", []), screen.get("grids", []))
            _merge_named_inventory_items(cur.setdefault("popups", []), screen.get("popups", []), "popup_name")
            _merge_logic_items(cur.setdefault("logic", []), screen.get("logic", []))
    screens = list(merged.values())
    stats = {
        "screens": len(screens),
        "screen_aliases_merged": aliases,
        "fields": sum(len(s.get("fields", [])) for s in screens),
        "controls": sum(len(s.get("controls", [])) for s in screens),
        "grids": sum(len(s.get("grids", [])) for s in screens),
        "popups": sum(len(s.get("popups", [])) for s in screens),
        "logic": sum(len(s.get("logic", [])) for s in screens),
        "requirements": sum(
            len(s.get("screen_requirements", []))
            + sum(len(x.get("requirements", [])) for x in s.get("fields", []))
            + sum(len(x.get("requirements", [])) for x in s.get("controls", []))
            + sum(len(x.get("requirements", [])) for x in s.get("grids", []))
            + sum(len(x.get("requirements", [])) for x in s.get("popups", []))
            + len(s.get("logic", []))
            for s in screens
        ),
    }
    return {"web_discovery_version": WEB_DISCOVERY_VERSION, "screens": screens}, stats


def audit_web_inventory_rule_coverage(inventory: dict, matrix: dict) -> dict:
    """Deterministic diagnostics: did discovered BA objects survive into tester rules?

    This does not invent missing rules. It flags suspicious omissions so the run diagnostics make
    silent coverage loss visible. Only objects that actually carry requirements/columns are audited.
    """
    output_screens = matrix.get("screens", []) if isinstance(matrix, dict) else []
    out_by_name = {_normalize_text(s.get("screen_name", "")): s for s in output_screens if isinstance(s, dict)}
    uncovered: list[dict] = []
    audited = 0
    covered = 0

    def has_name(rules: list[dict], name: str) -> bool:
        n = _normalize_search_text(name)
        if not n:
            return True
        n_tokens = [t for t in n.split() if len(t) >= 2]
        for rule in rules:
            text = _normalize_search_text(" ".join(str(rule.get(k, "") or "") for k in (
                "target", "feature_name", "rule_name", "test_objective", "test_condition",
                "expected_result", "source_requirement",
            )))
            if n in text:
                return True
            if n_tokens and len(n_tokens) >= 2 and all(t in text for t in n_tokens):
                return True
        return False

    for screen in inventory.get("screens", []) if isinstance(inventory, dict) else []:
        sname = _stringify(screen.get("screen_name"))
        out = out_by_name.get(_normalize_text(sname))
        if out is None:
            # Conservative alias fallback.
            for candidate in output_screens:
                if isinstance(candidate, dict) and _screen_alias_equivalent(sname, candidate.get("screen_name", "")):
                    out = candidate
                    break
        rules = out.get("test_rules", []) if isinstance(out, dict) else []
        for key, name_key in (("fields", "field_name"), ("controls", "control_name"), ("grids", "grid_name"), ("popups", "popup_name")):
            for item in screen.get(key, []) if isinstance(screen.get(key), list) else []:
                if not isinstance(item, dict):
                    continue
                meaningful = bool(item.get("requirements")) or (key == "grids" and bool(item.get("columns")))
                if not meaningful:
                    continue
                name = _stringify(item.get(name_key))
                audited += 1
                if has_name(rules, name):
                    covered += 1
                else:
                    uncovered.append({"screen": sname, "type": key[:-1] if key.endswith("s") else key, "name": name})
    return {
        "audited_objects": audited,
        "covered_objects": covered,
        "coverage_ratio": round(covered / audited, 4) if audited else 1.0,
        "uncovered_objects": uncovered[:50],
        "uncovered_count": len(uncovered),
    }


def process_web_discovery_chunk_recursive(
    chunk: dict,
    api_key: str,
    base_url: str,
    model: str,
    cache: dict | None = None,
    diagnostics: list | None = None,
) -> tuple[bool, list[dict]]:
    diagnostics = diagnostics if diagnostics is not None else []
    payload = build_agent1_chunk_payload(chunk)
    key = _cache_key(WEB_DISCOVERY_CACHE_NAMESPACE, model, PROMPT_WEB_DISCOVERY_INVENTORY, payload)
    if cache is not None and key in cache:
        diagnostics.append({"chunk_id": chunk.get("chunk_id"), "stage": "DISCOVERY", "status": "CACHE_HIT"})
        return True, [copy.deepcopy(cache[key])]

    result = None
    for attempt in range(AGENT1_API_RETRIES + 1):
        agent_name = f"WEB_DISCOVERY/{chunk.get('chunk_id')}"
        result = call_qwen_max_agent_detailed(
            content=payload,
            api_key=api_key,
            base_url=base_url,
            model=model,
            prompt_template=PROMPT_WEB_DISCOVERY_INVENTORY,
            max_tokens=WEB_DISCOVERY_MAX_OUTPUT_TOKENS,
            agent_name=agent_name,
            enable_thinking=False,
            response_format=json_object_response_format(),
        )
        if result.ok or result.finish_reason == "length":
            break
        if attempt < AGENT1_API_RETRIES:
            time.sleep(2)
    assert result is not None

    reason = None
    parse_diag = ""
    schema_diag = ""
    parsed = None
    if result.finish_reason == "length":
        reason = "MAX_TOKENS"
    elif not result.ok:
        diagnostics.append({"chunk_id": chunk.get("chunk_id"), "stage": "DISCOVERY", "status": "API_FAILED", "error": result.error})
        return False, []
    else:
        ok, parsed, parse_diag = extract_json_from_model_response(result.text)
        if not ok or parsed is None:
            reason = "JSON_PARSE_FAIL"
        else:
            parsed = normalize_web_inventory_shape(parsed)
            schema_ok, schema_diag = validate_web_inventory_schema(parsed)
            if not schema_ok:
                reason = "SCHEMA_FAIL"
            else:
                if cache is not None:
                    cache[key] = copy.deepcopy(parsed)
                diagnostics.append({
                    "chunk_id": chunk.get("chunk_id"), "stage": "DISCOVERY", "status": "OK",
                    "screens": len(parsed.get("screens", [])), "schema": schema_diag,
                })
                return True, [parsed]

    depth = int(chunk.get("depth", 0))
    can_split = depth < AGENT1_MAX_RECURSION_DEPTH and len(chunk.get("core_text", "")) > AGENT1_MIN_RECURSIVE_CHARS
    diagnostics.append({
        "chunk_id": chunk.get("chunk_id"), "stage": "DISCOVERY", "status": reason,
        "depth": depth, "parse": parse_diag, "schema": schema_diag,
    })
    if not can_split:
        return False, []
    children = _split_chunk_for_retry(chunk)
    if len(children) <= 1:
        return False, []
    log_info(f"[WEB_DISCOVERY/{chunk.get('chunk_id')}] ✂️ split do {reason}: {len(children)} child chunks")
    matrices = []
    for child in children:
        ok, child_results = process_web_discovery_chunk_recursive(child, api_key, base_url, model, cache, diagnostics)
        if not ok:
            return False, []
        matrices.extend(child_results)
    return True, matrices


def _inventory_screen_has_content(screen: dict) -> bool:
    if not isinstance(screen, dict):
        return False
    if screen.get("screen_requirements"):
        return True
    if screen.get("logic"):
        return True
    for key in ("fields", "controls", "grids", "popups"):
        for item in screen.get(key, []) if isinstance(screen.get(key), list) else []:
            if not isinstance(item, dict):
                continue
            if item.get("requirements"):
                return True
            if key == "grids" and item.get("columns"):
                return True
    return False


def _screen_inventory_payload(screen: dict) -> str:
    return json.dumps({
        "web_discovery_version": WEB_DISCOVERY_VERSION,
        "screen": screen,
    }, ensure_ascii=False, indent=2)


def _split_screen_inventory_for_qa(screen: dict) -> list[dict]:
    """Fallback only when one screen's QA output hits model length.

    Keep source-discovered ownership; split by component families rather than source chunks so
    fields are never reassigned to another screen. Logic is kept as its own work unit.
    """
    base = {
        "screen_name": screen.get("screen_name", ""),
        "screen_code": screen.get("screen_code", ""),
        "source_evidence": screen.get("source_evidence", ""),
        "screen_requirements": [], "fields": [], "controls": [], "grids": [], "popups": [], "logic": [],
    }
    parts = []
    groups = [
        ("screen_requirements", screen.get("screen_requirements", [])),
        ("fields", screen.get("fields", [])),
        ("controls", screen.get("controls", [])),
        ("grids", screen.get("grids", [])),
        ("popups", screen.get("popups", [])),
        ("logic", screen.get("logic", [])),
    ]
    # Pack roughly by serialized size, preserving each inventory item intact.
    current = copy.deepcopy(base)
    current_chars = len(_screen_inventory_payload(current))
    target_chars = 14000
    for key, items in groups:
        for item in items:
            item_chars = len(json.dumps(item, ensure_ascii=False)) + 10
            if current_chars + item_chars > target_chars and any(current[k] for k in ("screen_requirements", "fields", "controls", "grids", "popups", "logic")):
                parts.append(current)
                current = copy.deepcopy(base)
                current_chars = len(_screen_inventory_payload(current))
            current[key].append(copy.deepcopy(item))
            current_chars += item_chars
    if any(current[k] for k in ("screen_requirements", "fields", "controls", "grids", "popups", "logic")):
        parts.append(current)
    return parts or [copy.deepcopy(screen)]


def process_web_qa_screen_recursive(
    screen: dict,
    api_key: str,
    base_url: str,
    model: str,
    cache: dict | None = None,
    diagnostics: list | None = None,
    depth: int = 0,
) -> tuple[bool, list[dict]]:
    diagnostics = diagnostics if diagnostics is not None else []
    payload = _screen_inventory_payload(screen)
    key = _cache_key(WEB_AGENT1_CACHE_NAMESPACE, model, PROMPT_WEB_QA_FROM_INVENTORY, payload)
    screen_name = _stringify(screen.get("screen_name")) or "Unnamed Screen"
    if cache is not None and key in cache:
        diagnostics.append({"screen": screen_name, "stage": "QA", "status": "CACHE_HIT"})
        return True, [copy.deepcopy(cache[key])]

    result = None
    for attempt in range(AGENT1_API_RETRIES + 1):
        result = call_qwen_max_agent_detailed(
            content=payload,
            api_key=api_key,
            base_url=base_url,
            model=model,
            prompt_template=PROMPT_WEB_QA_FROM_INVENTORY,
            max_tokens=WEB_QA_MAX_OUTPUT_TOKENS,
            agent_name=f"WEB_QA/{screen_name[:36]}",
            enable_thinking=False,
            response_format=json_object_response_format(),
        )
        if result.ok or result.finish_reason == "length":
            break
        if attempt < AGENT1_API_RETRIES:
            log_info(f"[WEB_QA/{screen_name}] 🔁 Retry API lần {attempt + 2}/{AGENT1_API_RETRIES + 1}")
            time.sleep(2)
    assert result is not None
    reason = None
    parse_diag = ""
    schema_diag = ""
    parsed = None
    if result.finish_reason == "length":
        reason = "MAX_TOKENS"
    elif not result.ok:
        diagnostics.append({"screen": screen_name, "stage": "QA", "status": "API_FAILED", "error": result.error})
        return False, []
    else:
        ok, parsed, parse_diag = extract_json_from_model_response(result.text)
        if not ok or parsed is None:
            reason = "JSON_PARSE_FAIL"
        else:
            parsed = normalize_web_rule_shape(parsed)
            parsed = normalize_web_rule_matrix_enums(parsed)
            parsed, _ = normalize_web_rule_semantics(parsed)
            screens = parsed.get("screens", []) if isinstance(parsed, dict) else []
            if len(screens) > 1:
                reason = "MULTI_SCREEN_OUTPUT"
                schema_diag = f"Expected <=1 screen, got {len(screens)}"
            elif _inventory_screen_has_content(screen) and (
                len(screens) == 0 or not any(x.get("test_rules") for x in screens if isinstance(x, dict))
            ):
                reason = "EMPTY_QA_OUTPUT"
                schema_diag = "Inventory có requirement nhưng QA stage trả về 0 rule"
            else:
                if len(screens) == 1:
                    screens[0]["screen_name"] = screen_name
                schema_ok, schema_diag = validate_rule_matrix_schema(parsed, strict=True)
                if not schema_ok:
                    reason = "SCHEMA_FAIL"
                else:
                    if cache is not None:
                        cache[key] = copy.deepcopy(parsed)
                    diagnostics.append({
                        "screen": screen_name, "stage": "QA", "status": "OK",
                        "rules": sum(len(x.get("test_rules", [])) for x in screens), "schema": schema_diag,
                    })
                    return True, [parsed]

    diagnostics.append({
        "screen": screen_name, "stage": "QA", "status": reason,
        "depth": depth, "parse": parse_diag, "schema": schema_diag,
    })
    if depth >= 3:
        return False, []
    parts = _split_screen_inventory_for_qa(screen)
    if len(parts) <= 1:
        return False, []
    log_info(f"[WEB_QA/{screen_name}] ✂️ split inventory do {reason}: {len(parts)} parts")
    matrices = []
    for part in parts:
        ok, child = process_web_qa_screen_recursive(part, api_key, base_url, model, cache, diagnostics, depth + 1)
        if not ok:
            return False, []
        matrices.extend(child)
    return True, matrices


def _cache_key(prefix: str, model: str, prompt_template: str, payload: str) -> str:
    material = f"{prefix}\n{model}\n{prompt_template}\n{payload}".encode("utf-8", errors="ignore")
    return hashlib.sha256(material).hexdigest()


def _split_chunk_for_retry(chunk: dict) -> list[dict]:
    """Recursive split semantic cho một chunk bị length/parse/schema fail."""
    core = chunk.get("core_text", "")
    if len(core) <= AGENT1_MIN_RECURSIVE_CHARS:
        return []

    target = max(AGENT1_MIN_RECURSIVE_CHARS, min(7000, len(core) // 2))
    max_chars = max(target + 800, min(8500, len(core)))
    local = split_document_semantic(
        core,
        base_title=chunk.get("title", "RetryChunk"),
        target_chars=target,
        max_chars=max_chars,
        context_chars=min(AGENT1_CONTEXT_CHARS, 800),
    )

    if len(local) <= 1:
        # Last-resort 50/50 split tại whitespace gần giữa.
        mid = len(core) // 2
        left_cut = core.rfind("\n", 0, mid)
        if left_cut < AGENT1_MIN_RECURSIVE_CHARS:
            left_cut = core.rfind(" ", 0, mid)
        if left_cut < AGENT1_MIN_RECURSIVE_CHARS:
            left_cut = mid
        parts = [core[:left_cut].strip(), core[left_cut:].strip()]
        local = []
        for p in parts:
            if p:
                local.append({
                    "title": chunk.get("title", "RetryChunk"),
                    "screen_hint": chunk.get("screen_hint", ""),
                    "core_text": p,
                    "prev_context": "",
                    "next_context": "",
                    "char_count": len(p),
                    "depth": chunk.get("depth", 0) + 1,
                })

    parent_prev = chunk.get("prev_context", "")
    parent_next = chunk.get("next_context", "")
    children = []
    parent_id = chunk.get("chunk_id", "C")
    for idx, child in enumerate(local):
        child = dict(child)
        child["chunk_id"] = f"{parent_id}.{idx + 1}"
        child["depth"] = chunk.get("depth", 0) + 1
        if not child.get("screen_hint"):
            child["screen_hint"] = chunk.get("screen_hint", "")
        if idx == 0 and parent_prev:
            child["prev_context"] = (parent_prev + "\n" + child.get("prev_context", ""))[-AGENT1_CONTEXT_CHARS:]
        if idx == len(local) - 1 and parent_next:
            child["next_context"] = (child.get("next_context", "") + "\n" + parent_next)[:AGENT1_CONTEXT_CHARS]
        children.append(child)
    return children


def process_agent1_chunk_recursive(
    chunk: dict,
    api_key: str,
    base_url: str,
    model: str,
    prompt_template: str,
    cache: dict | None = None,
    diagnostics: list | None = None,
) -> tuple[bool, list[dict]]:
    """Agent 1 leaf processor.

    length hoặc JSON/schema fail do complexity => tự chia nhỏ và retry.
    Không bỏ qua chunk lỗi; nếu leaf vẫn fail, toàn pipeline fail để bảo toàn coverage.
    """
    diagnostics = diagnostics if diagnostics is not None else []
    payload = build_agent1_chunk_payload(chunk)
    key = _cache_key(WEB_AGENT1_CACHE_NAMESPACE, model, prompt_template, payload)

    if cache is not None and key in cache:
        diagnostics.append({
            "chunk_id": chunk.get("chunk_id"),
            "chars": len(chunk.get("core_text", "")),
            "depth": chunk.get("depth", 0),
            "status": "CACHE_HIT",
        })
        return True, [copy.deepcopy(cache[key])]

    call_result = None
    for attempt in range(AGENT1_API_RETRIES + 1):
        agent_name = f"AGENT1/{chunk.get('chunk_id')}"
        call_result = call_qwen_max_agent_detailed(
            content=payload,
            api_key=api_key,
            base_url=base_url,
            model=model,
            prompt_template=prompt_template,
            max_tokens=QWEN_MAX_OUTPUT_TOKENS,
            agent_name=agent_name,
            enable_thinking=False,
            response_format=json_object_response_format(),
        )
        if call_result.ok or call_result.finish_reason == "length":
            break
        if attempt < AGENT1_API_RETRIES:
            log_info(f"[{agent_name}] 🔁 Retry API lần {attempt + 2}/{AGENT1_API_RETRIES + 1}")
            time.sleep(2)

    assert call_result is not None

    reason_to_split = None
    parse_diag = None
    schema_diag = None

    if call_result.finish_reason == "length":
        reason_to_split = "MAX_TOKENS"
    elif not call_result.ok:
        diagnostics.append({
            "chunk_id": chunk.get("chunk_id"),
            "chars": len(chunk.get("core_text", "")),
            "depth": chunk.get("depth", 0),
            "status": "API_FAILED",
            "error": call_result.error,
        })
        return False, []
    else:
        parsed_ok, parsed_json, parse_diag = extract_json_from_model_response(call_result.text)
        if not parsed_ok:
            log_info(f"[{agent_name}] JSON parse fail; thử syntax-only repair trước khi split source.")
            repair_ok, repaired_json, repair_diag = repair_api_json_syntax(
                call_result.text,
                api_key=api_key,
                base_url=base_url,
                model=model,
                agent_name=agent_name,
            )
            parse_diag = f"{parse_diag} | {repair_diag}"
            if repair_ok:
                parsed_ok = True
                parsed_json = repaired_json
                log_info(f"[{agent_name}] ✅ JSON syntax repair thành công; không cần split source.")
            else:
                reason_to_split = "JSON_PARSE_FAIL"
        if parsed_ok:
            parsed_json = normalize_web_rule_shape(parsed_json)
            parsed_json = normalize_web_rule_matrix_enums(parsed_json)
            parsed_json, metadata_repairs = normalize_web_rule_semantics(parsed_json)
            schema_ok, schema_diag = validate_rule_matrix_schema(parsed_json, strict=False)
            if not schema_ok:
                reason_to_split = "SCHEMA_FAIL"
                log_error(f"[{agent_name}] Schema validation failed | {schema_diag}")
            else:
                diagnostics.append({
                    "chunk_id": chunk.get("chunk_id"),
                    "chars": len(chunk.get("core_text", "")),
                    "depth": chunk.get("depth", 0),
                    "status": "OK",
                    "elapsed": round(call_result.elapsed, 2),
                    "finish_reason": call_result.finish_reason,
                    "output_chars": len(call_result.text),
                    "screens": len(parsed_json.get("screens", [])),
                    "rules": sum(len(s.get("test_rules", [])) for s in parsed_json.get("screens", [])),
                    "json_repaired": "JSON repair OK" in str(parse_diag or ""),
                    "semantic_metadata_repairs": len(metadata_repairs),
                })
                if cache is not None:
                    cache[key] = copy.deepcopy(parsed_json)
                return True, [parsed_json]

    depth = int(chunk.get("depth", 0))
    split_reason_is_retryable = (
        reason_to_split in {"MAX_TOKENS", "JSON_PARSE_FAIL"}
        or (reason_to_split == "SCHEMA_FAIL" and _schema_failure_can_benefit_from_split(schema_diag))
    )
    can_split = (
        split_reason_is_retryable
        and depth < AGENT1_MAX_RECURSION_DEPTH
        and len(chunk.get("core_text", "")) > AGENT1_MIN_RECURSIVE_CHARS
    )

    diagnostics.append({
        "chunk_id": chunk.get("chunk_id"),
        "chars": len(chunk.get("core_text", "")),
        "depth": depth,
        "status": "SPLIT_RETRY" if can_split else "FAILED_LEAF",
        "reason": reason_to_split,
        "parse_diag": parse_diag,
        "schema_diag": schema_diag,
        "finish_reason": call_result.finish_reason,
        "output_chars": len(call_result.text),
        "raw_tail": call_result.text[-1200:] if call_result.text else "",
    })

    if not can_split:
        return False, []

    children = _split_chunk_for_retry(chunk)
    if len(children) < 2:
        return False, []

    log_info(
        f"[AGENT1/{chunk.get('chunk_id')}] ✂️ Recursive split do {reason_to_split} | "
        f"{len(chunk.get('core_text', '')):,} chars -> "
        + " + ".join(f"{len(c.get('core_text', '')):,}" for c in children)
    )

    all_matrices = []
    for child in children:
        child_ok, child_matrices = process_agent1_chunk_recursive(
            chunk=child,
            api_key=api_key,
            base_url=base_url,
            model=model,
            prompt_template=prompt_template,
            cache=cache,
            diagnostics=diagnostics,
        )
        if not child_ok:
            return False, []
        all_matrices.extend(child_matrices)
    return True, all_matrices


def _rule_signature(rule: dict) -> tuple:
    """Deterministic conservative dedup for repeated OCR/chunk requirements.

    Source wording is intentionally excluded: two repeated source fragments may differ in OCR
    formatting while producing the same test objective/condition/expected result.
    Python does not infer QA meaning; it only removes structurally equivalent generated rules.
    """
    return (
        # Conservative identity: final organization and rule identity are part of the signature.
        # This avoids silently dropping distinct Senior-QA rules during Python merge.
        _normalize_text(rule.get("feature_group")),
        _normalize_text(rule.get("feature_name")),
        _normalize_text(rule.get("rule_name")),
        _normalize_text(rule.get("target")),
        _normalize_text(rule.get("category")),
        _normalize_text(rule.get("rule_type")),
        _normalize_text(rule.get("test_objective")),
        _normalize_text(rule.get("test_condition")),
        _normalize_text(rule.get("expected_result")),
    )


def _infer_rule_prefix(rules: list[dict], fallback: str) -> str:
    prefixes = []
    for rule in rules:
        rid = str(rule.get("rule_id", "")).strip()
        m = re.match(r"^(.+?)[-_](\d+)$", rid)
        if m:
            p = m.group(1).strip()
            if p and p.upper() not in {"TMP", "RULE", "R"}:
                prefixes.append(p)
    if prefixes:
        prefix = Counter(prefixes).most_common(1)[0][0]
    else:
        prefix = fallback
    prefix = unicodedata.normalize("NFKD", prefix).encode("ascii", "ignore").decode("ascii")
    prefix = re.sub(r"[^A-Za-z0-9_]+", "_", prefix).strip("_").upper()
    return prefix or fallback


# ============================================================
# PATCH V2 — CANONICAL SCREEN MERGE
# Thay thế hàm merge_rule_matrices hiện tại bằng block này.
# Dùng helper/import sẵn có trong file:
# _normalize_text, _rule_signature, _infer_rule_prefix,
# OrderedDict, copy, re, unicodedata
# ============================================================

def _screen_core_tokens(name: str) -> tuple[str, ...]:
    s = unicodedata.normalize("NFKD", str(name or "")).encode("ascii", "ignore").decode("ascii").lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)

    generic = {
        "man", "hinh", "screen", "page", "module",
        "bidv", "trai", "phieu"
    }
    return tuple(t for t in s.split() if t and t not in generic)


def _stable_screen_code(screen: dict) -> str | None:
    name = str(screen.get("screen_name", "") or "").strip()

    # 1) Screen code xuất hiện trực tiếp trong screen_name.
    for code in re.findall(r"\b[A-Z][A-Z0-9]*(?:_[A-Z0-9]+){1,}\b", name):
        c = re.sub(r"_\d+$", "", code.upper())
        if c and not c.startswith("SCREEN"):
            return c

    # 2) Fallback từ rule_id, ví dụ:
    # BOND_ORDER_LIST-001 / BOND_ORDER_LIST_2-001 / BOND_ORDER_LIST_3-001
    # -> BOND_ORDER_LIST
    for rule in screen.get("test_rules", []) or []:
        rid = str(rule.get("rule_id", "") or "").strip().upper()
        if not rid:
            continue

        prefix = re.sub(r"-\d+$", "", rid)
        prefix = re.sub(r"_\d+$", "", prefix)

        if (
            prefix
            and "_" in prefix
            and not prefix.startswith("SCREEN")
            and re.fullmatch(r"[A-Z][A-Z0-9_]+", prefix)
        ):
            return prefix

    return None


def _screen_alias_equivalent(name_a: str, name_b: str) -> bool:
    """
    Merge bảo thủ:
      Danh sách lệnh đặt mua
      Màn hình Danh sách lệnh đặt mua Trái phiếu
    => cùng screen.

    Không merge:
      Danh sách lệnh đặt mua
      Màn hình xem chi tiết lệnh đặt mua
    """
    a = _screen_core_tokens(name_a)
    b = _screen_core_tokens(name_b)

    if not a or not b:
        return False

    if a == b:
        return True

    sa, sb = set(a), set(b)
    inter = len(sa & sb)
    union = len(sa | sb)
    jaccard = inter / max(1, union)

    # Token phân biệt loại màn hình.
    discriminators = {
        "danh", "sach", "chi", "tiet", "them", "moi",
        "chinh", "sua", "tao", "duyet", "xem"
    }

    if (sa & discriminators) != (sb & discriminators):
        return False

    return inter >= 4 and jaccard >= 0.85


def _is_technical_screen_code(name: str) -> bool:
    s = str(name or "").strip()
    return bool(re.fullmatch(r"[A-Z][A-Z0-9_]+", s)) and "_" in s


def _prefer_human_screen_name(current: str, candidate: str) -> str:
    """
    Giữ tên nghiệp vụ dễ đọc.
    Nếu current chỉ là code kỹ thuật và candidate là tên người đọc được -> dùng candidate.
    Nếu cả hai đều là tên người đọc được -> ưu tiên tên mô tả rõ hơn nhưng không quá dài.
    """
    cur = str(current or "").strip()
    cand = str(candidate or "").strip()
    if not cand:
        return cur
    if not cur:
        return cand

    if _is_technical_screen_code(cur) and not _is_technical_screen_code(cand):
        return cand
    if not _is_technical_screen_code(cur) and _is_technical_screen_code(cand):
        return cur

    # Ưu tiên tên có "màn hình" hoặc "trái phiếu" khi cùng alias,
    # vì thường mô tả domain rõ hơn.
    cur_score = (2 if "màn hình" in cur.lower() else 0) + (1 if "trái phiếu" in cur.lower() else 0)
    cand_score = (2 if "màn hình" in cand.lower() else 0) + (1 if "trái phiếu" in cand.lower() else 0)
    if cand_score > cur_score:
        return cand
    if cand_score == cur_score and len(cand) > len(cur) and len(cand) <= 80:
        return cand
    return cur


def merge_rule_matrices(matrices: list[dict]) -> tuple[dict, dict]:
    """
    Merge Python-only:
    - canonicalize alias cùng screen;
    - exact-dedup rule;
    - renumber rule_id.
    Không suy luận QA/business.
    """
    screens_map: OrderedDict[str, dict] = OrderedDict()
    seen_by_screen: dict[str, set] = {}
    screen_meta: OrderedDict[str, dict] = OrderedDict()

    raw_rules = 0
    exact_duplicates = 0
    screen_aliases_merged = 0

    for matrix in matrices:
        for screen in matrix.get("screens", []):
            name = str(screen.get("screen_name", "")).strip() or "Unnamed Screen"
            stable_code = _stable_screen_code(screen)

            canonical_key = None

            # A. Screen code là identity mạnh nhất.
            if stable_code:
                code_key = f"code::{stable_code.lower()}"
                if code_key in screens_map:
                    canonical_key = code_key
                else:
                    for k, meta in screen_meta.items():
                        if meta.get("stable_code") == stable_code:
                            canonical_key = k
                            break
                    if canonical_key is None:
                        canonical_key = code_key

            # B. Exact normalized name.
            if canonical_key is None:
                exact_key = f"name::{_normalize_text(name)}"
                if exact_key in screens_map:
                    canonical_key = exact_key

            # C. Alias gần như cùng tên, merge bảo thủ.
            if canonical_key is None:
                for k, meta in screen_meta.items():
                    if _screen_alias_equivalent(name, meta["screen_name"]):
                        canonical_key = k
                        screen_aliases_merged += 1
                        break

            # D. Screen mới.
            if canonical_key is None:
                canonical_key = f"name::{_normalize_text(name)}"

            if canonical_key not in screens_map:
                # stable_code chỉ dùng làm identity/key. Tên hiển thị vẫn là tên nghiệp vụ.
                display_name = name
                screens_map[canonical_key] = {
                    "screen_name": display_name,
                    "test_rules": []
                }
                seen_by_screen[canonical_key] = set()
                screen_meta[canonical_key] = {
                    "screen_name": name,
                    "stable_code": stable_code,
                }
            elif stable_code and not screen_meta[canonical_key].get("stable_code"):
                # Khi chunk sau tìm ra code thật, chỉ cập nhật identity nội bộ.
                # KHÔNG thay screen_name hiển thị thành code.
                screen_meta[canonical_key]["stable_code"] = stable_code

            # Cập nhật tên hiển thị nếu alias mới human-readable hơn.
            screens_map[canonical_key]["screen_name"] = _prefer_human_screen_name(
                screens_map[canonical_key].get("screen_name", ""),
                name,
            )

            for rule in screen.get("test_rules", []):
                raw_rules += 1
                sig = _rule_signature(rule)

                if sig in seen_by_screen[canonical_key]:
                    exact_duplicates += 1
                    # Preserve the more informative source wording when repeated OCR/chunks
                    # generated the same semantic rule.
                    current_rules = screens_map[canonical_key]["test_rules"]
                    for existing in current_rules:
                        if _rule_signature(existing) == sig:
                            old_src = str(existing.get("source_requirement", "")).strip()
                            new_src = str(rule.get("source_requirement", "")).strip()
                            if len(new_src) > len(old_src):
                                existing["source_requirement"] = new_src
                            break
                    continue

                seen_by_screen[canonical_key].add(sig)
                screens_map[canonical_key]["test_rules"].append(copy.deepcopy(rule))

    final_screens = list(screens_map.values())

    # IDs carry no business meaning. Generate them only after merge/dedup is final.
    final_rule_index = 0
    for screen in final_screens:
        for rule in screen.get("test_rules", []):
            final_rule_index += 1
            rule["rule_id"] = f"R_{final_rule_index:04d}"

    final = {
        "test_design_version": WEB_TEST_DESIGN_VERSION,
        "screens": final_screens,
    }

    stats = {
        "input_matrices": len(matrices),
        "screens": len(final_screens),
        "raw_rules": raw_rules,
        "screen_aliases_merged": screen_aliases_merged,
        "exact_duplicates_removed": exact_duplicates,
        "final_rules": sum(len(s.get("test_rules", [])) for s in final_screens),
    }
    return final, stats


def run_agent1_document_pipeline(
    raw_text: str,
    base_filename: str,
    api_key: str,
    base_url: str,
    model: str,
    prompt_template: str | None = None,
    cache: dict | None = None,
    progress_callback=None,
) -> tuple[bool, dict | None, dict]:
    """Web V3.6: BA source -> screen/component/logic inventory -> QA rules.

    The first AI pass is discovery-only and is forbidden from applying QA rules. After all
    source chunks are merged into one canonical inventory, the second pass generates QA rules
    screen-by-screen from that inventory. This prevents source chunk boundaries from becoming
    fake screens and prevents QA derivation before field/logic ownership is known.
    """
    started = time.time()
    initial_chunks = split_document_semantic(raw_text, base_filename)
    diagnostics: list[dict] = []
    inventory_parts: list[dict] = []
    discovery_workers = min(AGENT1_PARALLEL_WORKERS, max(1, len(initial_chunks)))

    log_info(
        f"[WEB_PIPELINE_V3.6] 📚 DocumentChars={len(raw_text):,} | "
        f"DiscoveryChunks={len(initial_chunks)} | DiscoveryWorkers={discovery_workers}"
    )
    if progress_callback:
        progress_callback(0, 100, "Bước 1/2: Đang đọc cấu trúc BA và lập danh sách màn hình / field / logic...")

    def _discover(index: int, chunk: dict):
        local: list[dict] = []
        ok, results = process_web_discovery_chunk_recursive(
            chunk=chunk,
            api_key=api_key,
            base_url=base_url,
            model=model,
            cache=cache,
            diagnostics=local,
        )
        return index, ok, results, local

    if discovery_workers <= 1 or len(initial_chunks) <= 1:
        for idx, chunk in enumerate(initial_chunks):
            ok, parts = process_web_discovery_chunk_recursive(
                chunk=chunk,
                api_key=api_key,
                base_url=base_url,
                model=model,
                cache=cache,
                diagnostics=diagnostics,
            )
            if not ok:
                return False, None, {
                    "ok": False,
                    "stage": "web_discovery",
                    "architecture": "DISCOVERY->CANONICAL_INVENTORY->QA_RULES",
                    "document_chars": len(raw_text),
                    "initial_chunks": len(initial_chunks),
                    "elapsed": round(time.time() - started, 2),
                    "diagnostics": diagnostics,
                }
            inventory_parts.extend(parts)
            if progress_callback:
                progress_callback(int(40 * (idx + 1) / max(1, len(initial_chunks))), 100,
                                  f"Bước 1/2: Đã đọc {idx + 1}/{len(initial_chunks)} phần tài liệu BA")
    else:
        ordered: dict[int, tuple[bool, list[dict], list[dict]]] = {}
        completed = 0
        with ThreadPoolExecutor(max_workers=discovery_workers, thread_name_prefix="web-discovery") as executor:
            futures = {executor.submit(_discover, idx, chunk): idx for idx, chunk in enumerate(initial_chunks)}
            for future in as_completed(futures):
                idx, ok, parts, local = future.result()
                ordered[idx] = (ok, parts, local)
                completed += 1
                if progress_callback:
                    progress_callback(int(40 * completed / max(1, len(initial_chunks))), 100,
                                      f"Bước 1/2: Đã đọc {completed}/{len(initial_chunks)} phần tài liệu BA")
        for idx in range(len(initial_chunks)):
            ok, parts, local = ordered[idx]
            diagnostics.extend(local)
            if not ok:
                return False, None, {
                    "ok": False,
                    "stage": "web_discovery",
                    "architecture": "DISCOVERY->CANONICAL_INVENTORY->QA_RULES",
                    "document_chars": len(raw_text),
                    "initial_chunks": len(initial_chunks),
                    "elapsed": round(time.time() - started, 2),
                    "diagnostics": diagnostics,
                }
            inventory_parts.extend(parts)

    inventory, inventory_stats = merge_web_inventories(inventory_parts)
    inventory_ok, inventory_diag = validate_web_inventory_schema(inventory)
    if not inventory_ok or not inventory.get("screens"):
        summary = {
            "ok": False,
            "stage": "web_inventory",
            "architecture": "DISCOVERY->CANONICAL_INVENTORY->QA_RULES",
            "document_chars": len(raw_text),
            "initial_chunks": len(initial_chunks),
            "inventory": inventory_stats,
            "inventory_schema": inventory_diag,
            "elapsed": round(time.time() - started, 2),
            "diagnostics": diagnostics,
        }
        log_error(f"[WEB_PIPELINE_V3.6] Không xác định được màn hình hợp lệ | {inventory_diag}")
        return False, None, summary

    inventory_preview = [
        {
            "screen_name": _stringify(screen.get("screen_name")),
            "screen_code": _stringify(screen.get("screen_code")),
            "fields": [_stringify(x.get("field_name")) for x in screen.get("fields", [])[:100]],
            "controls": [_stringify(x.get("control_name")) for x in screen.get("controls", [])[:100]],
            "grids": [_stringify(x.get("grid_name")) for x in screen.get("grids", [])[:30]],
            "popups": [_stringify(x.get("popup_name")) for x in screen.get("popups", [])[:30]],
            "logic": [
                {
                    "logic_name": _stringify(x.get("logic_name")),
                    "logic_type": _stringify(x.get("logic_type")),
                    "targets": list(x.get("targets", []))[:20],
                }
                for x in screen.get("logic", [])[:100]
            ],
        }
        for screen in inventory.get("screens", [])[:50]
    ]
    screens = inventory.get("screens", [])
    log_info(
        f"[WEB_DISCOVERY] ✅ Canonical inventory | Screens={inventory_stats['screens']} | "
        f"Fields={inventory_stats['fields']} | Controls={inventory_stats['controls']} | "
        f"Grids={inventory_stats['grids']} | Popups={inventory_stats['popups']} | "
        f"Logic={inventory_stats['logic']} | Requirements={inventory_stats['requirements']}"
    )
    if progress_callback:
        progress_callback(45, 100,
                          f"Bước 1/2 hoàn tất: {inventory_stats['screens']} màn hình, "
                          f"{inventory_stats['fields']} field, {inventory_stats['logic']} logic. Bắt đầu áp dụng QA Rule...")

    qa_workers = min(WEB_QA_PARALLEL_WORKERS, max(1, len(screens)))
    qa_matrices: list[dict] = []

    def _qa(index: int, screen: dict):
        local: list[dict] = []
        ok, results = process_web_qa_screen_recursive(
            screen=screen,
            api_key=api_key,
            base_url=base_url,
            model=model,
            cache=cache,
            diagnostics=local,
        )
        return index, ok, results, local

    if qa_workers <= 1 or len(screens) <= 1:
        for idx, screen in enumerate(screens):
            name = _stringify(screen.get("screen_name"))
            if progress_callback:
                progress_callback(45 + int(50 * idx / max(1, len(screens))), 100,
                                  f"Bước 2/2: Áp dụng QA Rule cho màn hình {idx + 1}/{len(screens)} — {name}")
            ok, matrices = process_web_qa_screen_recursive(
                screen=screen,
                api_key=api_key,
                base_url=base_url,
                model=model,
                cache=cache,
                diagnostics=diagnostics,
            )
            if not ok:
                return False, None, {
                    "ok": False,
                    "stage": "web_qa_rules",
                    "architecture": "DISCOVERY->CANONICAL_INVENTORY->QA_RULES",
                    "document_chars": len(raw_text),
                    "initial_chunks": len(initial_chunks),
                    "inventory": inventory_stats,
                    "inventory_preview": inventory_preview,
                    "elapsed": round(time.time() - started, 2),
                    "diagnostics": diagnostics,
                }
            qa_matrices.extend(matrices)
            if progress_callback:
                progress_callback(45 + int(50 * (idx + 1) / max(1, len(screens))), 100,
                                  f"Bước 2/2: Hoàn tất {idx + 1}/{len(screens)} màn hình")
    else:
        ordered_qa: dict[int, tuple[bool, list[dict], list[dict]]] = {}
        completed = 0
        with ThreadPoolExecutor(max_workers=qa_workers, thread_name_prefix="web-qa") as executor:
            futures = {executor.submit(_qa, idx, screen): idx for idx, screen in enumerate(screens)}
            for future in as_completed(futures):
                idx, ok, results, local = future.result()
                ordered_qa[idx] = (ok, results, local)
                completed += 1
                if progress_callback:
                    progress_callback(45 + int(50 * completed / max(1, len(screens))), 100,
                                      f"Bước 2/2: Hoàn tất {completed}/{len(screens)} màn hình")
        for idx in range(len(screens)):
            ok, results, local = ordered_qa[idx]
            diagnostics.extend(local)
            if not ok:
                return False, None, {
                    "ok": False,
                    "stage": "web_qa_rules",
                    "architecture": "DISCOVERY->CANONICAL_INVENTORY->QA_RULES",
                    "document_chars": len(raw_text),
                    "initial_chunks": len(initial_chunks),
                    "inventory": inventory_stats,
                    "inventory_preview": inventory_preview,
                    "elapsed": round(time.time() - started, 2),
                    "diagnostics": diagnostics,
                }
            qa_matrices.extend(results)

    final_matrix, merge_stats = merge_rule_matrices(qa_matrices)
    final_matrix = normalize_web_rule_shape(final_matrix)
    final_matrix = normalize_web_rule_matrix_enums(final_matrix)
    final_matrix, final_semantic_repairs = normalize_web_rule_semantics(final_matrix)
    merge_stats["semantic_metadata_repairs"] = len(final_semantic_repairs)
    schema_ok, schema_diag = validate_rule_matrix_schema(final_matrix, strict=True)
    coverage_audit = audit_web_inventory_rule_coverage(inventory, final_matrix) if schema_ok else {}
    elapsed = time.time() - started

    summary = {
        "ok": schema_ok,
        "stage": "completed" if schema_ok else "final_schema",
        "architecture": "DISCOVERY->CANONICAL_INVENTORY->QA_RULES",
        "document_chars": len(raw_text),
        "initial_chunks": len(initial_chunks),
        "discovery_workers": discovery_workers,
        "qa_workers": qa_workers,
        "inventory": inventory_stats,
        "inventory_schema": inventory_diag,
        "inventory_preview": inventory_preview,
        "qa_matrices": len(qa_matrices),
        "elapsed": round(elapsed, 2),
        "merge": merge_stats,
        "schema": schema_diag,
        "coverage_audit": coverage_audit,
        "diagnostics": diagnostics,
    }
    if not schema_ok:
        log_error(f"[WEB_PIPELINE_V3.6] ❌ Final schema fail | {schema_diag}")
        return False, None, summary

    if coverage_audit.get("uncovered_count"):
        log_info(
            f"[WEB_COVERAGE_AUDIT] ⚠️ {coverage_audit['uncovered_count']} object có requirement "
            f"chưa thấy target rõ trong Rule Matrix | Coverage={coverage_audit.get('coverage_ratio')}"
        )
    if progress_callback:
        progress_callback(100, 100, "Hoàn tất phân tích Web: đã xác định màn hình/field/logic và áp dụng QA Rule.")
    log_info(
        f"[WEB_PIPELINE_V3.6] ✅ Completed | Time={elapsed:.2f}s | "
        f"Screens={merge_stats['screens']} | FinalRules={merge_stats['final_rules']} | "
        f"InventoryFields={inventory_stats['fields']} | InventoryLogic={inventory_stats['logic']}"
    )
    return True, final_matrix, summary




def _parse_bullet_forest(text: str) -> list[dict]:
    """Parse indented bullet text. Literal \\n becomes a newline inside the same node title."""
    clean = re.sub(r"^```[a-zA-Z]*\s*", "", str(text or "").strip())
    clean = re.sub(r"\s*```$", "", clean).strip()
    root = {"title": "__ROOT__", "children": []}
    stack = [(-1, root)]

    for raw_line in clean.splitlines():
        if not raw_line.strip():
            continue
        line = raw_line.replace("\t", "    ")
        indent = len(line) - len(line.lstrip(" "))
        level = indent // 2
        content = line.strip()
        if content.startswith(("- ", "* ")):
            content = content[2:].strip()
        if not content:
            continue

        content = content.replace("\\n", "\n")
        node = {"title": content, "children": []}
        while stack and stack[-1][0] >= level:
            stack.pop()
        parent = stack[-1][1] if stack else root
        parent["children"].append(node)
        stack.append((level, node))
    return root["children"]


















def create_xmind_from_text(tree_data: str, output_path: str, root_title: str = "Kế hoạch & Kịch bản Kiểm thử"):
    """Create an XMind archive from the normalized bullet-tree parser."""
    try:
        forest = _parse_bullet_forest(tree_data)
        if not forest:
            raise ValueError("Không có node hợp lệ để tạo XMind.")

        node_id_counter = [1]

        def to_xmind_node(node: dict) -> dict:
            node_id = f"node_{node_id_counter[0]}"
            node_id_counter[0] += 1
            return {
                "id": node_id,
                "title": str(node.get("title", "")),
                "children": {
                    "attached": [to_xmind_node(child) for child in node.get("children", [])]
                },
            }

        root_topic = {
            "id": "root_node",
            "title": root_title,
            "children": {"attached": [to_xmind_node(node) for node in forest]},
        }
        content_json = [{
            "id": "sheet_1",
            "title": "Sơ đồ Kiểm thử QA",
            "rootTopic": root_topic,
        }]
        manifest_json = {"file-entries": {"content.json": {}, "metadata.json": {}}}
        metadata_json = {}

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zip_file:
            zip_file.writestr("content.json", json.dumps(content_json, ensure_ascii=False, indent=2))
            zip_file.writestr("manifest.json", json.dumps(manifest_json, ensure_ascii=False, indent=2))
            zip_file.writestr("metadata.json", json.dumps(metadata_json, ensure_ascii=False, indent=2))

        log_info(f"[XMIND_EXPORT] Created | Nodes={node_id_counter[0] - 1} | Path={output_path}")
        return True, None
    except Exception as e:
        err_msg = log_error("Lỗi khi đóng gói file XMind", e)
        return False, err_msg


# ==============================================================================
# HÀM BÓC TÁCH CÂY XMIND VÀ XUẤT EXCEL CHUẨN THEO NHÁNH NHÓM (DỌC THEO TÊN MÀN HÌNH)
# ==============================================================================
# ==============================================================================
# HÀM BÓC TÁCH CÂY XMIND VÀ XUẤT EXCEL CHUẨN (CẬP NHẬT DOCSTRING & QUÉT ĐỘ SÂU)
# ==============================================================================
def convert_tree_to_ui_excel(tree_text: str) -> bytes:
    """
    Chuyển toàn bộ cây Test Case Web/XMind thành workbook Excel.

    FIX V3.3:
    - KHÔNG còn chỉ lấy root/screen đầu tiên.
    - Mỗi Screen Root có testcase được xuất thành 1 worksheet riêng.
    - Thêm worksheet Summary để đối soát số testcase theo màn hình.
    - Chỉ nhận node bắt đầu bằng TC_... / TC_API_... là testcase;
      không nhận nhầm leaf text có dấu " - ".
    - Quét đệ quy mọi Target/Sub-group nên không mất testcase do độ sâu cây thay đổi.
    - Tự kiểm tra tổng TC trong tree == tổng TC đã ghi Excel; lệch là raise lỗi,
      tránh xuất file thiếu testcase một cách im lặng.
    """
    clean_text = re.sub(r'^```[a-zA-Z]*\s*', '', str(tree_text or '').strip())
    clean_text = re.sub(r'\s*```$', '', clean_text).strip()
    if not clean_text:
        raise ValueError("Không có dữ liệu Test Case để xuất Excel.")

    # ----------------------------------------------------------
    # 1. Parse normalized bullet tree once (shared with XMind/API exporters)
    # ----------------------------------------------------------
    root_node = {"title": "__ROOT__", "children": _parse_bullet_forest(clean_text)}

    # ----------------------------------------------------------
    # 2. Helpers
    # ----------------------------------------------------------
    tc_regex = re.compile(r'^TC(?:_API)?[_\- ]?\d+\s*(?:[-:]|$)', re.IGNORECASE)

    def is_test_case_node(node: dict) -> bool:
        return bool(tc_regex.match(str(node.get("title", "")).strip()))

    def count_testcases(node: dict) -> int:
        total = 1 if is_test_case_node(node) else 0
        for child in node.get("children", []):
            total += count_testcases(child)
        return total

    def has_test_cases(node: dict) -> bool:
        return count_testcases(node) > 0

    def clean_prefix(value: str, prefixes: list[str]) -> str:
        t = str(value or "").strip()
        low = t.lower()
        for prefix in prefixes:
            if low.startswith(prefix.lower()):
                return t[len(prefix):].strip()
        return t

    def extract_details(tc_node: dict) -> tuple[str, str, str, str]:
        """Lấy Pre-condition / Steps / Data test / Expected ở bất kỳ độ sâu con nào."""
        pre = ""
        steps = ""
        data_test = ""
        expected = ""

        def walk(node: dict):
            nonlocal pre, steps, data_test, expected
            for child in node.get("children", []):
                title = str(child.get("title", "")).strip()
                low = title.lower()

                if low.startswith(("pre-condition:", "precondition:", "tiền điều kiện:")):
                    if not pre:
                        pre = clean_prefix(title, ["Pre-condition:", "Precondition:", "Tiền điều kiện:"])
                elif low.startswith(("các bước thực hiện:", "các bước:", "steps:", "step:")):
                    if not steps:
                        steps = clean_prefix(title, ["Các bước thực hiện:", "Các bước:", "Steps:", "Step:"])
                elif low.startswith(("dữ liệu kiểm thử:", "du lieu kiem thu:", "data test:", "test data:")):
                    if not data_test:
                        data_test = clean_prefix(
                            title,
                            ["Dữ liệu kiểm thử:", "Du lieu kiem thu:", "Data test:", "Test data:"]
                        )
                elif low.startswith(("kết quả mong đợi:", "expected result:", "expected:", "response (kết quả mong đợi):")):
                    if not expected:
                        expected = clean_prefix(
                            title,
                            ["Kết quả mong đợi:", "Expected Result:", "Expected:", "Response (Kết quả mong đợi):"]
                        )
                walk(child)

        walk(tc_node)
        return pre, steps, data_test, expected

    def split_tc_title(title: str) -> tuple[str, str]:
        t = str(title or "").strip()
        m = re.match(r'^(TC(?:_API)?[_\- ]?\d+)\s*(?:[-:]\s*)?(.*)$', t, flags=re.IGNORECASE)
        if not m:
            return "", t
        return m.group(1).strip(), m.group(2).strip()

    def sanitize_sheet_title(title: str, used: set[str]) -> str:
        name = re.sub(r'[\\/*?:\[\]]', '_', str(title or "Screen")).strip()
        name = name or "Screen"
        name = name[:31]
        base = name
        idx = 2
        while name.lower() in used:
            suffix = f"_{idx}"
            name = (base[:31-len(suffix)] + suffix)[:31]
            idx += 1
        used.add(name.lower())
        return name

    # Screen roots = mọi top-level node có ít nhất một testcase.
    # Nếu model bọc thêm một root chung duy nhất, unwrap một cấp khi hợp lý.
    screen_nodes = [node for node in root_node["children"] if has_test_cases(node)]
    if len(screen_nodes) == 1:
        only = screen_nodes[0]
        child_roots = [child for child in only.get("children", []) if has_test_cases(child)]
        # Không unwrap nếu children trông giống category chuẩn.
        category_like = any(
            re.match(r'^\d+\.\s*(kiểm tra|kiem tra)', str(c.get("title", "")).strip(), flags=re.IGNORECASE)
            for c in child_roots
        )
        if child_roots and not category_like and all(count_testcases(c) > 0 for c in child_roots):
            screen_nodes = child_roots

    if not screen_nodes:
        raise ValueError("Không tìm thấy Test Case hợp lệ (TC_...) trong cây để xuất Excel.")

    expected_total = sum(count_testcases(screen) for screen in screen_nodes)

    # ----------------------------------------------------------
    # 3. Workbook + styles
    # ----------------------------------------------------------
    wb = Workbook()
    default_ws = wb.active
    wb.remove(default_ws)

    fill_green = PatternFill(start_color="81C784", fill_type="solid")
    fill_light_green = PatternFill(start_color="C8E6C9", fill_type="solid")
    fill_white = PatternFill(start_color="FFFFFF", fill_type="solid")
    fill_header_bg = PatternFill(start_color="4DD0E1", fill_type="solid")
    fill_pink = PatternFill(start_color="F8BBD0", fill_type="solid")
    fill_blue_exec = PatternFill(start_color="90CAF9", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color='B0BEC5'),
        right=Side(style='thin', color='B0BEC5'),
        top=Side(style='thin', color='B0BEC5'),
        bottom=Side(style='thin', color='B0BEC5')
    )

    font_title = Font(name="Segoe UI", size=11, bold=True)
    font_bold = Font(name="Segoe UI", size=10, bold=True)
    font_body = Font(name="Segoe UI", size=10)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # Senior-QA-style two-row header. Administrative/execution fields stay blank until manual execution.
    group_headers = {
        "A11": "Test Case",
        "E11": "Steps",
        "I11": "Chrome",
        "L11": "Kết quả hiện tại",
        "M11": "Ghi chú",
        "O11": "QC viết testcase",
        "P11": "Sprint viết testcase",
        "R11": "Sprint thực hiện test",
        "V11": "Case cần Auto (Yes/No)",
        "W11": "Case Đã Auto (Yes/No)",
        "X11": "Smoke test (Yes/No)",
        "Y11": "Regression test (Yes/No)",
        "Z11": "TCs Out of date (Yes/No)",
        "AA11": "Ngày TCs Out of date",
    }
    detail_headers = [
        "ID", "Name", "PreConditions", "Importance", "Step", "Data test",
        "Expected Result", "Actual Result", "Lần 1", "Lần 2", "Lần 3",
        "Kết quả hiện tại", "Ghi chú", "Mã lỗi", "QC viết testcase",
        "Sprint viết testcase", "QC thực hiện test", "Sprint thực hiện test",
        "Người review", "Ngày review", "Nội dung review",
        "Case cần Auto (Yes/No)", "Case Đã Auto (Yes/No)",
        "Smoke test (Yes/No)", "Regression test (Yes/No)",
        "TCs Out of date (Yes/No)", "Ngày TCs Out of date",
    ]
    col_widths = {
        'A': 14, 'B': 52, 'C': 38, 'D': 12, 'E': 52, 'F': 30, 'G': 58, 'H': 30,
        'I': 10, 'J': 10, 'K': 10, 'L': 16, 'M': 28, 'N': 16, 'O': 18, 'P': 18,
        'Q': 18, 'R': 18, 'S': 18, 'T': 16, 'U': 34, 'V': 18, 'W': 18, 'X': 16,
        'Y': 18, 'Z': 18, 'AA': 20,
    }

    used_sheet_names = set()
    exported_total = 0
    summary_rows = []

    def write_sheet(screen: dict):
        nonlocal exported_total

        screen_name = str(screen.get("title", "Tên Màn Hình")).strip() or "Tên Màn Hình"
        ws = wb.create_sheet(title=sanitize_sheet_title(screen_name, used_sheet_names))
        ws.views.sheetView[0].showGridLines = True

        # Metadata block — inspired by the Senior workbook while keeping unknown project fields blank.
        ws['D1'] = "KỊCH BẢN KIỂM THỬ *"
        ws['D1'].font = font_title
        ws['D1'].alignment = align_center
        ws['C2'] = "Tên màn hình/Tên chức năng"
        ws['C2'].font = font_bold
        ws['D2'] = screen_name
        ws['D2'].font = font_body
        ws['C4'] = "Mã Testcase"
        ws['C4'].font = font_bold
        ws['D4'] = "TC_"
        ws['D4'].font = font_body
        ws['B13'] = "Màn hình test"
        ws['C13'] = "Link test:"
        ws['B14'] = "Đường dẫn:"

        for r in (1, 2, 4):
            for col in ["C", "D"]:
                ws[f"{col}{r}"].border = thin_border

        # Row 11: high-level visual groups.
        for col_idx in range(1, 28):
            c = ws.cell(row=11, column=col_idx)
            c.fill = fill_header_bg
            c.border = thin_border
            c.alignment = align_center
            c.font = font_bold
        for pos, label in group_headers.items():
            ws[pos] = label

        # Row 12: detailed executable/review columns.
        for col_idx, label in enumerate(detail_headers, start=1):
            c = ws.cell(row=12, column=col_idx, value=label)
            c.font = font_bold
            c.alignment = align_center
            c.border = thin_border
            c.fill = fill_pink if col_idx <= 8 else fill_blue_exec

        for col, width in col_widths.items():
            ws.column_dimensions[col].width = width

        current_row = 15
        screen_exported = 0

        def write_group_header(title: str, indent_level: int, strong: bool):
            nonlocal current_row
            if not title:
                return
            prefix = "  " * max(0, indent_level)
            ws.cell(row=current_row, column=2, value=f"{prefix}{title}").font = font_bold
            fill = fill_green if strong else fill_light_green
            for col in range(1, 28):
                c = ws.cell(row=current_row, column=col)
                c.fill = fill
                c.border = thin_border
            current_row += 1

        def write_tc(node: dict):
            nonlocal current_row, screen_exported, exported_total
            tc_id, tc_name = split_tc_title(node.get("title", ""))
            pre, steps, data_test, expected = extract_details(node)

            ws.cell(row=current_row, column=1, value=tc_id).alignment = align_center
            ws.cell(row=current_row, column=2, value=tc_name)
            ws.cell(row=current_row, column=3, value=pre)
            ws.cell(row=current_row, column=4, value=None)  # Importance: do not infer without source/company policy.
            ws.cell(row=current_row, column=5, value=steps)
            ws.cell(row=current_row, column=6, value=data_test)
            ws.cell(row=current_row, column=7, value=expected)
            # H:AA are execution/review/automation columns and intentionally remain blank.

            for col in range(1, 28):
                c = ws.cell(row=current_row, column=col)
                c.font = font_body
                c.fill = fill_white
                c.border = thin_border
                c.alignment = align_center if col in (1, 4, 9, 10, 11, 12, 14, 22, 23, 24, 25, 26, 27) else align_left

            current_row += 1
            screen_exported += 1
            exported_total += 1

        def traverse(node: dict, depth: int = 0):
            if is_test_case_node(node):
                write_tc(node)
                return

            if not has_test_cases(node):
                return

            # Node dưới screen: Senior feature-group ở depth 0, feature_name/subgroup ở depth >=1
            write_group_header(str(node.get("title", "")).strip(), depth, strong=(depth == 0))
            for child in node.get("children", []):
                traverse(child, depth + 1)

        for child in screen.get("children", []):
            traverse(child, depth=0)

        ws.freeze_panes = "A13"
        ws.auto_filter.ref = f"A12:AA{max(12, current_row - 1)}"
        summary_rows.append((screen_name, screen_exported, ws.title))

    for screen in screen_nodes:
        write_sheet(screen)

    # ----------------------------------------------------------
    # 4. Safety invariant: XMind tree TC count == Excel TC count
    # ----------------------------------------------------------
    if exported_total != expected_total:
        raise ValueError(
            f"Excel export không đầy đủ: Tree có {expected_total} testcase nhưng Excel mới ghi {exported_total}."
        )

    # ----------------------------------------------------------
    # 5. Summary sheet
    # ----------------------------------------------------------
    summary = wb.create_sheet(title="Summary", index=0)
    summary['A1'] = "ĐỐI SOÁT TEST CASE"
    summary['A1'].font = Font(name="Segoe UI", size=12, bold=True)
    summary['A3'] = "Màn hình / Chức năng"
    summary['B3'] = "Số Test Case"
    summary['C3'] = "Worksheet"
    for cell in summary[3]:
        cell.font = font_bold
        cell.fill = fill_header_bg
        cell.border = thin_border
        cell.alignment = align_center

    row = 4
    for screen_name, tc_count, sheet_name in summary_rows:
        summary.cell(row=row, column=1, value=screen_name)
        summary.cell(row=row, column=2, value=tc_count)
        summary.cell(row=row, column=3, value=sheet_name)
        for col in range(1, 4):
            summary.cell(row=row, column=col).border = thin_border
            summary.cell(row=row, column=col).alignment = align_left if col != 2 else align_center
        row += 1

    summary.cell(row=row, column=1, value="TOTAL").font = font_bold
    summary.cell(row=row, column=2, value=exported_total).font = font_bold
    for col in range(1, 4):
        summary.cell(row=row, column=col).border = thin_border

    summary.column_dimensions['A'].width = 55
    summary.column_dimensions['B'].width = 18
    summary.column_dimensions['C'].width = 35

    try:
        log_info(
            f"[EXCEL_EXPORT] Screens={len(screen_nodes)} | TreeTC={expected_total} | ExcelTC={exported_total}"
        )
    except Exception:
        pass

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


# ==========================================
# 2.2 PIPELINE API V3.3 — SENIOR TEMPLATE / STRUCTURED TESTCASE DESIGN / MULTI-DOC
# ==========================================

API_RULE_FIELDS = {
    "rule_id", "target", "category", "rule_type", "rule_name",
    "test_objective", "test_condition", "precondition", "test_data",
    "expected_http_code", "expected_status", "expected_code", "expected_message",
    "expected_trace_id", "expected_data_body", "business_result",
    "source_requirement", "source_document", "reconciliation_status",
    "applied_qa_rule", "generation_reason"
}
API_ALLOWED_CATEGORIES = {
    "AUTH", "PERMISSION", "VALIDATION", "HAPPY_PATH", "BUSINESS_RULE"
}
API_CATEGORY_ORDER = {
    "AUTH": 1,
    "PERMISSION": 2,
    "VALIDATION": 3,
    "HAPPY_PATH": 4,
    "BUSINESS_RULE": 5,
}
API_ALLOWED_RULE_TYPES = {"EXPLICIT", "DERIVED"}
API_ALLOWED_SOURCE_DOCUMENTS = {"API_SPEC", "BA", "BOTH"}
API_ALLOWED_RECONCILIATION_STATUSES = {"CONSISTENT", "COMPLEMENTARY", "CONFLICT", "DOC_ONLY", "DERIVED"}
API_HTTP_METHODS = {"GET", "POST", "PUT", "PATCH", "DELETE", "HEAD", "OPTIONS", "UNMAPPED"}


def api_rule_matrix_response_format() -> dict:
    """Strict JSON Schema for Qwen structured output.

    This constrains syntax/shape only. Business semantics are still validated by
    validate_api_rule_matrix_schema() and deterministic scope enforcement.
    """
    string_prop = {"type": "string"}
    rule_properties = {field: dict(string_prop) for field in sorted(API_RULE_FIELDS)}
    rule_properties["category"] = {"type": "string", "enum": sorted(API_ALLOWED_CATEGORIES)}
    rule_properties["rule_type"] = {"type": "string", "enum": sorted(API_ALLOWED_RULE_TYPES)}
    rule_properties["source_document"] = {"type": "string", "enum": sorted(API_ALLOWED_SOURCE_DOCUMENTS)}
    rule_properties["reconciliation_status"] = {
        "type": "string", "enum": sorted(API_ALLOWED_RECONCILIATION_STATUSES)
    }

    schema = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "api_test_design_version": {"type": "string", "enum": ["3.3"]},
            "api_modules": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "module_name": {"type": "string"},
                        "endpoints": {
                            "type": "array",
                            "items": {
                                "type": "object",
                                "additionalProperties": False,
                                "properties": {
                                    "method": {"type": "string", "enum": sorted(API_HTTP_METHODS)},
                                    "endpoint_path": {"type": "string"},
                                    "summary": {"type": "string"},
                                    "test_rules": {
                                        "type": "array",
                                        "items": {
                                            "type": "object",
                                            "additionalProperties": False,
                                            "properties": rule_properties,
                                            "required": sorted(API_RULE_FIELDS),
                                        },
                                    },
                                },
                                "required": ["method", "endpoint_path", "summary", "test_rules"],
                            },
                        },
                    },
                    "required": ["module_name", "endpoints"],
                },
            },
        },
        "required": ["api_test_design_version", "api_modules"],
    }
    return {
        "type": "json_schema",
        "json_schema": {
            "name": "api_rule_matrix",
            "strict": True,
            "schema": schema,
        },
    }


def json_object_response_format() -> dict:
    return {"type": "json_object"}


def _try_load_structured_api_doc(text: str):
    if not text or not text.strip():
        return None
    try:
        obj = json.loads(text)
        if isinstance(obj, (dict, list)):
            return obj
    except Exception:
        pass
    try:
        obj = yaml.safe_load(text)
        if isinstance(obj, (dict, list)):
            return obj
    except Exception:
        pass
    return None


def extract_api_endpoint_index(design_text: str) -> list[dict]:
    """Trích endpoint index bằng Python, không suy luận QA.

    Hỗ trợ OpenAPI/Swagger JSON/YAML, Postman collection JSON và regex fallback.
    Index chỉ dùng làm metadata định danh endpoint cho Agent 1 chunk.
    """
    results = []
    seen = set()

    def add(method, path, summary="", module=""):
        method = str(method or "").upper().strip()
        path = str(path or "").strip()
        if method not in API_HTTP_METHODS - {"UNMAPPED"} or not path:
            return
        key = (method, _normalize_text(path))
        if key in seen:
            # Later text patterns may provide a human-readable title/description.
            for existing in results:
                if (existing.get("method"), _normalize_text(existing.get("path"))) == key:
                    if not str(existing.get("summary", "")).strip() and str(summary or "").strip():
                        existing["summary"] = str(summary or "").strip()
                    if not str(existing.get("module", "")).strip() and str(module or "").strip():
                        existing["module"] = str(module or "").strip()
                    break
            return
        seen.add(key)
        results.append({
            "method": method,
            "path": path,
            "summary": str(summary or "").strip(),
            "module": str(module or "").strip(),
        })

    obj = _try_load_structured_api_doc(design_text)
    if isinstance(obj, dict):
        # OpenAPI / Swagger
        paths = obj.get("paths")
        if isinstance(paths, dict):
            for path, path_item in paths.items():
                if not isinstance(path_item, dict):
                    continue
                for method, operation in path_item.items():
                    if str(method).lower() not in {"get", "post", "put", "patch", "delete", "head", "options"}:
                        continue
                    operation = operation if isinstance(operation, dict) else {}
                    tags = operation.get("tags") or []
                    module = tags[0] if isinstance(tags, list) and tags else ""
                    add(
                        method,
                        path,
                        operation.get("summary") or operation.get("operationId") or operation.get("description") or "",
                        module,
                    )

        # Postman collection
        if isinstance(obj.get("item"), list):
            def walk_items(items, parents=None):
                parents = parents or []
                for item in items or []:
                    if not isinstance(item, dict):
                        continue
                    name = item.get("name", "")
                    if isinstance(item.get("item"), list):
                        walk_items(item.get("item"), parents + ([name] if name else []))
                        continue
                    req = item.get("request")
                    if not isinstance(req, dict):
                        continue
                    method = req.get("method")
                    url = req.get("url")
                    if isinstance(url, dict):
                        raw_url = url.get("raw") or ""
                        path_parts = url.get("path")
                        if isinstance(path_parts, list) and path_parts:
                            path = "/" + "/".join(str(x) for x in path_parts)
                        else:
                            path = raw_url
                    else:
                        path = str(url or "")
                    module = " / ".join([p for p in parents if p])
                    add(method, path, name, module)
            walk_items(obj.get("item"))

    # Fallback từ text Markdown / raw Swagger / tài liệu mô tả.
    for m in re.finditer(r"(?im)^\s*(GET|POST|PUT|PATCH|DELETE|HEAD|OPTIONS)\s+([/][^\s|`\"']+)", design_text or ""):
        add(m.group(1), m.group(2))
    for m in re.finditer(r'(?im)^\s*[\"\']?(/[^\"\']+)[\"\']?\s*:\s*\{?\s*$', design_text or ""):
        path = m.group(1).strip()
        tail = (design_text or "")[m.end():m.end()+800]
        mm = re.search(r'(?im)[\"\']?(get|post|put|patch|delete|head|options)[\"\']?\s*:', tail)
        if mm:
            add(mm.group(1), path)

    # Human-readable API design docs often use "Endpoint: /..." + "Method: POST".
    text = design_text or ""
    for m in re.finditer(r'(?im)\bEndpoint\s*[:：]\s*([/][^\s|`\"\']+)', text):
        path = m.group(1).strip().rstrip('.,;)')
        window = text[max(0, m.start()-500):min(len(text), m.end()+900)]
        mm = re.search(r'(?im)\bMethod\s*[:：]\s*(GET|POST|PUT|PATCH|DELETE|HEAD|OPTIONS)\b', window)
        if mm:
            # Nearby heading/description is useful for BA scope routing.
            desc = ''
            dm = re.search(r'(?im)\b(?:Mô tả|Description)\s*[:：]\s*(.+)', window)
            if dm:
                desc = dm.group(1).strip()[:240]
            add(mm.group(1), path, desc)

    # Also support a title beginning with /path followed by Method later in the section.
    for m in re.finditer(r'(?im)^\s*([/][^\s:]+)\s*:\s*(.+)$', text):
        path = m.group(1).strip()
        tail = text[m.end():m.end()+1200]
        mm = re.search(r'(?im)\bMethod\s*[:：]\s*(GET|POST|PUT|PATCH|DELETE|HEAD|OPTIONS)\b', tail)
        if mm:
            add(mm.group(1), path, m.group(2).strip()[:240])

    return results


def _tokenize_for_api_hint(text: str) -> set[str]:
    norm = _normalize_search_text(text)
    return {
        t for t in re.findall(r"[a-z0-9_./-]+", norm)
        if len(t) >= 3 and t not in {"api", "the", "and", "for", "with", "request", "response"}
    }


def select_api_endpoint_hints(chunk_text: str, endpoint_index: list[dict], limit: int = API_ENDPOINT_HINT_LIMIT) -> list[dict]:
    """Chọn endpoint hints bằng lexical matching bảo thủ, không dùng LLM."""
    if not endpoint_index:
        return []
    raw = chunk_text or ""
    chunk_norm = _normalize_search_text(raw)
    chunk_tokens = _tokenize_for_api_hint(raw)
    scored = []
    for ep in endpoint_index:
        path = str(ep.get("path", ""))
        method = str(ep.get("method", ""))
        summary = str(ep.get("summary", ""))
        module = str(ep.get("module", ""))
        score = 0.0
        if path and _normalize_search_text(path) in chunk_norm:
            score += 100.0
        if method and re.search(rf"\b{re.escape(method.casefold())}\b", chunk_norm):
            score += 6.0
        ep_tokens = _tokenize_for_api_hint(" ".join([path, summary, module]))
        overlap = len(chunk_tokens & ep_tokens)
        if overlap:
            score += overlap * 2.0
        if score > 0:
            scored.append((score, ep))
    scored.sort(key=lambda x: (-x[0], str(x[1].get("path", ""))))
    return [copy.deepcopy(ep) for _, ep in scored[:limit]]



def _api_endpoint_key(ep: dict) -> tuple[str, str]:
    return (str(ep.get("method", "UNMAPPED")).upper().strip(), str(ep.get("path") or ep.get("endpoint_path") or "UNMAPPED").strip())


def extract_target_api_excerpt(design_text: str, endpoint: dict, radius: int = 6000) -> str:
    """Small source excerpt around a target endpoint; used for routing, not testcase generation."""
    text = design_text or ""
    path = str(endpoint.get("path") or endpoint.get("endpoint_path") or "").strip()
    if not text:
        return ""
    # Short API Design files should be read in full. The previous 3.2k-radius window
    # truncated valid response fields even in the 4-page benchmark spec.
    if len(text) <= 24000:
        return text.strip()
    pos = text.find(path) if path and path != "UNMAPPED" else -1
    if pos < 0:
        # API design files are usually short; cap fallback so BA routing stays cheap.
        return text[: min(len(text), radius * 2)]
    start = max(0, pos - radius)
    end = min(len(text), pos + len(path) + radius)
    return text[start:end].strip()


def build_target_api_descriptors(design_text: str, endpoint_index: list[dict]) -> list[dict]:
    targets = []
    for ep in endpoint_index:
        targets.append({
            "method": str(ep.get("method", "UNMAPPED")).upper().strip() or "UNMAPPED",
            "endpoint_path": str(ep.get("path", "")).strip() or "UNMAPPED",
            "summary": str(ep.get("summary", "")).strip(),
            "module": str(ep.get("module", "")).strip(),
            "spec_excerpt": extract_target_api_excerpt(design_text, ep),
        })
    if not targets:
        # Keep pipeline usable for informal specs without a parseable endpoint.
        targets.append({
            "method": "UNMAPPED",
            "endpoint_path": "UNMAPPED",
            "summary": "API được mô tả trong API Design/Spec",
            "module": "",
            "spec_excerpt": (design_text or "")[:6400],
        })
    return targets


def select_primary_api_targets(design_text: str, targets: list[dict], document_hint: str = "") -> list[dict]:
    """Conservatively choose the primary endpoint when a Design document exposes many.

    No business rule is inferred here. We only use document title/intro and exact endpoint
    occurrences. If the result is ambiguous, keep all candidates rather than guessing.
    """
    if len(targets) <= 1:
        return targets
    intro = (design_text or "")[:6000]
    haystack = _normalize_search_text(f"{document_hint}\n{intro}")
    scored: list[tuple[float, int, dict]] = []
    for idx, target in enumerate(targets):
        path = str(target.get("endpoint_path") or "")
        summary = str(target.get("summary") or "")
        score = 0.0
        if path and path != "UNMAPPED":
            exact_count = (design_text or "").count(path)
            score += min(30.0, exact_count * 5.0)
            if _normalize_search_text(path) in _normalize_search_text(intro):
                score += 15.0
        target_tokens = _tokenize_for_api_hint(summary)
        hint_tokens = _tokenize_for_api_hint(haystack)
        score += min(18.0, 3.0 * len(target_tokens & hint_tokens))
        scored.append((score, idx, target))
    scored.sort(key=lambda item: (-item[0], item[1]))
    if not scored:
        return targets
    top_score = scored[0][0]
    second_score = scored[1][0] if len(scored) > 1 else -1.0
    if top_score >= 10.0 and top_score >= second_score + 6.0:
        selected = [copy.deepcopy(scored[0][2])]
        log_info(
            f"[API_TARGET_SELECTOR] Chọn 1/{len(targets)} endpoint theo Design title/intro | "
            f"{selected[0].get('method')} {selected[0].get('endpoint_path')} | score={top_score:.1f}"
        )
        return selected

    # Product contract: one API Design upload represents ONE API under test. If the
    # document contains sibling/downstream endpoints and lexical scores are ambiguous,
    # choose the endpoint that appears earliest in the design text instead of creating
    # multiple testcase folders. Score is used only as a secondary tie-breaker.
    positions = {}
    for score, idx, target in scored:
        path = str(target.get("endpoint_path") or "")
        pos = (design_text or "").find(path) if path and path != "UNMAPPED" else -1
        positions[id(target)] = pos if pos >= 0 else 10**12
    scored_by_position = sorted(
        scored,
        key=lambda item: (positions.get(id(item[2]), 10**12), -item[0], item[1]),
    )
    selected = [copy.deepcopy(scored_by_position[0][2])]
    log_info(
        f"[API_TARGET_SELECTOR] ⚠️ Design có {len(targets)} endpoint và điểm gần nhau; "
        f"strict-single-target chọn endpoint xuất hiện sớm nhất: "
        f"{selected[0].get('method')} {selected[0].get('endpoint_path')}"
    )
    return selected


def build_api_scope_router_payload(ba_chunk: dict, targets: list[dict]) -> str:
    compact_targets = []
    for target in targets:
        compact_targets.append({
            "method": target.get("method", "UNMAPPED"),
            "endpoint_path": target.get("endpoint_path", "UNMAPPED"),
            "summary": target.get("summary", ""),
            "spec_excerpt": target.get("spec_excerpt", "")[:5000],
        })
    return f"""=== TARGET API(S) FROM API DESIGN / SPEC ===
{json.dumps(compact_targets, ensure_ascii=False, indent=2)}

=== BA CHUNK METADATA ===
chunk_id: {ba_chunk.get('chunk_id')}
section_hint: {ba_chunk.get('title')}

=== BA CURRENT SOURCE ===
{ba_chunk.get('core_text', '')}
"""


def fallback_route_ba_chunk_by_spec_overlap(ba_chunk: dict, targets: list[dict], min_overlap: int = 4) -> list[dict]:
    """Conservative lexical fallback when the scope-router call fails.

    Uses target API spec excerpt rather than endpoint path alone, so human BA docs can
    still match on business terms such as duyệt/workflow/ngày hiệu lực.
    """
    chunk_tokens = _tokenize_for_api_hint(ba_chunk.get("core_text", ""))
    generic = {
        "thong", "giao", "dich", "system", "server", "client", "user",
        "api", "service", "request", "response", "thanh", "cong", "khong",
    }
    chunk_tokens = {t for t in chunk_tokens if t not in generic}
    matches = []
    for target in targets:
        spec_tokens = _tokenize_for_api_hint(" ".join([
            str(target.get("summary", "")), str(target.get("spec_excerpt", ""))[:6000]
        ]))
        spec_tokens = {t for t in spec_tokens if t not in generic}
        overlap = len(chunk_tokens & spec_tokens)
        if overlap >= min_overlap:
            matches.append({
                "method": target.get("method", "UNMAPPED"),
                "endpoint_path": target.get("endpoint_path", "UNMAPPED"),
                "summary": target.get("summary", ""),
                "confidence": min(0.75, 0.35 + overlap / 40.0),
                "reason": f"lexical spec overlap={overlap}",
            })
    return matches


def _normalize_scope_role(value: str) -> str:
    role = str(value or "").strip().upper()
    if role in {"PRIMARY", "CONTINUATION", "DEPENDENCY", "OUT_OF_SCOPE"}:
        return role
    return "OUT_OF_SCOPE"


def _build_continuation_target(primary_target: dict, match: dict) -> dict:
    """Build a stable target descriptor for a continuation API without inventing its contract.

    Exact method/path are kept only when the router extracted them from the BA source.
    Unknown identifiers remain UNMAPPED while summary carries the BA operation name so
    the testcase workspace can still be separated from the primary API.
    """
    name = str(match.get("related_api_name") or "API tiếp nối").strip() or "API tiếp nối"
    method = str(match.get("related_method") or "UNMAPPED").upper().strip() or "UNMAPPED"
    if method not in API_HTTP_METHODS:
        method = "UNMAPPED"
    path = str(match.get("related_endpoint_path") or "UNMAPPED").strip() or "UNMAPPED"
    if not path.startswith("/") and path != "UNMAPPED":
        # A non-path phrase is not an endpoint contract. Keep it only as the summary.
        path = "UNMAPPED"
    return {
        "method": method,
        "endpoint_path": path,
        "summary": name,
        "module": str(primary_target.get("module") or "").strip(),
        "spec_excerpt": "",
        "scope_role": "CONTINUATION",
        "parent_method": primary_target.get("method", "UNMAPPED"),
        "parent_endpoint_path": primary_target.get("endpoint_path", "UNMAPPED"),
    }


def route_ba_chunk_to_target_apis(
    ba_chunk: dict,
    targets: list[dict],
    api_key: str,
    base_url: str,
    model: str,
    cache: dict | None = None,
) -> tuple[bool, list[dict], dict]:
    """Cheap scope scan only; no testcase generation.

    V1.11.3 keeps only PRIMARY/DEPENDENCY context for the selected API and skips
    every distinct continuation/sibling operation. API Design owns endpoint identity.
    """
    payload = build_api_scope_router_payload(ba_chunk, targets)
    cache_key = _cache_key("api-scope-router-v1.11.3-strict-target", model, PROMPT_API_SCOPE_ROUTER, payload)
    if cache is not None and cache_key in cache:
        result = copy.deepcopy(cache[cache_key])
        return True, result.get("matches", []), {"status": "CACHE_HIT", "chunk_id": ba_chunk.get("chunk_id")}

    call = call_qwen_max_agent_detailed(
        content=payload,
        api_key=api_key,
        base_url=base_url,
        model=model,
        prompt_template=PROMPT_API_SCOPE_ROUTER,
        max_tokens=API_SCOPE_SCAN_MAX_TOKENS,
        agent_name=f"API_SCOPE/{ba_chunk.get('chunk_id')}",
        enable_thinking=False,
        response_format=json_object_response_format(),
    )
    if not call.ok:
        return False, [], {"status": "API_FAILED", "chunk_id": ba_chunk.get("chunk_id"), "error": call.error}
    ok, parsed, diag = extract_json_from_model_response(call.text)
    if not ok or not isinstance(parsed, dict):
        return False, [], {"status": "PARSE_FAILED", "chunk_id": ba_chunk.get("chunk_id"), "diag": diag}

    raw_matches = parsed.get("matches") if isinstance(parsed.get("matches"), list) else []
    target_keys = {_api_endpoint_key(t): t for t in targets}
    matches = []
    for item in raw_matches:
        if not isinstance(item, dict):
            continue
        method = str(item.get("method", "UNMAPPED")).upper().strip() or "UNMAPPED"
        path = str(item.get("endpoint_path", "UNMAPPED")).strip() or "UNMAPPED"
        target = target_keys.get((method, path))
        if target is None and len(targets) == 1:
            # Single primary spec is unambiguous even if the small router slightly drifts identifiers.
            target = targets[0]
            method, path = _api_endpoint_key(target)
        if target is None:
            continue

        role = _normalize_scope_role(item.get("scope_role"))

        match = {
            "method": method,
            "endpoint_path": path,
            "summary": target.get("summary", ""),
            "scope_role": role,
            "related_api_name": str(item.get("related_api_name") or "").strip(),
            "related_method": (
                str(item.get("related_method") or "UNMAPPED").upper().strip()
                if str(item.get("related_method") or "UNMAPPED").upper().strip() in API_HTTP_METHODS
                else "UNMAPPED"
            ),
            "related_endpoint_path": str(item.get("related_endpoint_path") or "UNMAPPED").strip() or "UNMAPPED",
            "confidence": _safe_float(item.get("confidence"), 0.0, 0.0, 1.0),
            "reason": str(item.get("reason") or "").strip(),
        }
        matches.append(match)

    result = {"matches": matches}
    if cache is not None:
        cache[cache_key] = copy.deepcopy(result)
    role_counts = {}
    for match in matches:
        role_counts[match.get("scope_role", "OUT_OF_SCOPE")] = role_counts.get(match.get("scope_role", "OUT_OF_SCOPE"), 0) + 1
    return True, matches, {
        "status": "OK", "chunk_id": ba_chunk.get("chunk_id"),
        "elapsed": round(call.elapsed, 2), "role_counts": role_counts,
    }

def build_api_combined_source(design_text: str, ba_text: str) -> str:
    parts = []
    if design_text and design_text.strip():
        parts.append("# === SOURCE DOCUMENT: API_SPEC ===\n" + design_text.strip())
    if ba_text and ba_text.strip():
        parts.append("# === SOURCE DOCUMENT: BA ===\n" + ba_text.strip())
    return "\n\n".join(parts).strip()


def build_api_agent1_chunk_payload(
    chunk: dict,
    endpoint_index: list[dict],
    *,
    source_document: str = "API_SPEC",
    target_endpoint: dict | None = None,
    spec_case_reference: str = "",
) -> str:
    hints = select_api_endpoint_hints(chunk.get("core_text", ""), endpoint_index)
    hint_text = json.dumps(hints, ensure_ascii=False, indent=2) if hints else "[]"
    target = target_endpoint or {}
    target_text = json.dumps({
        "method": target.get("method", "UNMAPPED"),
        "endpoint_path": target.get("endpoint_path") or target.get("path") or "UNMAPPED",
        "summary": target.get("summary", ""),
    }, ensure_ascii=False, indent=2)
    target_spec_reference = ""
    source_document = "BA" if str(source_document).upper() == "BA" else "API_SPEC"
    if source_document == "BA":
        # Prefer the testcase inventory already generated from API Spec. This is much safer than
        # asking BA extraction to rediscover technical cases from raw design prose and materially
        # reduces duplicate Permission/Validation/Happy-Path scenarios across documents.
        target_spec_reference = str(spec_case_reference or "").strip()
        if not target_spec_reference:
            target_spec_reference = str(target.get("spec_excerpt") or "").strip()[:7000]
        else:
            target_spec_reference = target_spec_reference[:14000]
        source_policy = (
            "BA POLICY: create standalone BUSINESS_RULE cases only. Do not create standalone AUTH/PERMISSION/VALIDATION/HAPPY_PATH. "
            "Use SPEC TESTCASE REFERENCE only to attach exact BA error code/message/outcome to an already-existing SPEC technical case. "
            "For such enrichment, mirror the SPEC category/target/condition instead of inventing a new scenario. "
            "Unmatched BA technical records are discarded."
        )
    else:
        source_policy = (
            "API_SPEC POLICY: create standalone AUTH/PERMISSION/VALIDATION/HAPPY_PATH cases only. "
            "Do not create standalone BUSINESS_RULE cases; BA owns business-rule testcase creation."
        )
    return f"""=== FORCED ANALYSIS SCOPE — MANDATORY ===
source_document: {source_document}
TARGET API:
{target_text}

SOURCE RESPONSIBILITY — MANDATORY:
{source_policy}

Rules for this call:
- Generate rules ONLY for TARGET API / operation shown above.
- Ignore sibling APIs/operations that are NOT the TARGET. If TARGET itself is a continuation such as Confirm Approval, analyze that continuation normally.
- source_document of every generated rule MUST be exactly {source_document}.
- If TARGET has an exact method/path, do not infer another endpoint from nearby BA text.
- If TARGET endpoint_path is UNMAPPED, NEVER invent a URL. Use TARGET summary to identify the operation and keep method/path UNMAPPED unless the current source explicitly states them.
- BA text may describe many APIs. A sibling/continuation API is context only and MUST NOT become a testcase endpoint.

=== SPEC TESTCASE REFERENCE — CHỈ DÙNG ĐỂ ĐỐI CHIẾU/ENRICH CASE KỸ THUẬT, KHÔNG TỰ SINH RULE TỪ ĐÂY ===
{target_spec_reference}

=== CHUNK METADATA — KHÔNG PHẢI REQUIREMENT ===
chunk_id: {chunk.get('chunk_id')}
source_section_hint: {chunk.get('title')}
chunk_depth: {chunk.get('depth', 0)}

=== ENDPOINT INDEX HINT — CHỈ ĐỊNH DANH, KHÔNG TỰ TẠO TEST RULE TỪ INDEX ===
{hint_text}

=== PREVIOUS CONTEXT — CHỈ DÙNG ĐỂ HIỂU, KHÔNG SINH RULE CHỈ TỪ ĐOẠN NÀY ===
{chunk.get('prev_context', '')}

=== CURRENT SOURCE — PHẠM VI CHÍNH ĐƯỢC PHÉP SINH API TEST RULE ===
{chunk.get('core_text', '')}

=== NEXT CONTEXT — CHỈ DÙNG ĐỂ HIỂU, KHÔNG SINH RULE CHỈ TỪ ĐOẠN NÀY ===
{chunk.get('next_context', '')}
"""



def force_api_matrix_scope(data: dict, *, source_document: str, target_endpoint: dict | None) -> dict:
    """Deterministically enforce scope metadata after LLM extraction."""
    if not isinstance(data, dict):
        return data
    target_endpoint = target_endpoint or {}
    method = str(target_endpoint.get("method", "UNMAPPED")).upper().strip() or "UNMAPPED"
    path = str(target_endpoint.get("endpoint_path") or target_endpoint.get("path") or "UNMAPPED").strip() or "UNMAPPED"
    summary = str(target_endpoint.get("summary", "")).strip()
    source_document = "BA" if str(source_document).upper() == "BA" else "API_SPEC"
    for module in data.get("api_modules", []) if isinstance(data.get("api_modules"), list) else []:
        for ep in module.get("endpoints", []) if isinstance(module.get("endpoints"), list) else []:
            ep["method"] = method
            ep["endpoint_path"] = path
            if summary and not str(ep.get("summary", "")).strip():
                ep["summary"] = summary
            for rule in ep.get("test_rules", []) if isinstance(ep.get("test_rules"), list) else []:
                rule["source_document"] = source_document
                if str(rule.get("rule_type", "")).upper() == "DERIVED":
                    rule["reconciliation_status"] = "DERIVED"
                else:
                    rule["reconciliation_status"] = "DOC_ONLY"
    return data


def validate_api_rule_matrix_schema(data, strict: bool = False) -> tuple[bool, str]:
    if not isinstance(data, dict):
        return False, f"Root phải là object, nhận {type(data).__name__}"
    modules = data.get("api_modules")
    if not isinstance(modules, list):
        return False, "Thiếu field 'api_modules' hoặc 'api_modules' không phải array"

    total_endpoints = 0
    total_rules = 0
    for midx, module in enumerate(modules):
        if not isinstance(module, dict):
            return False, f"api_modules[{midx}] phải là object"
        module_name = module.get("module_name")
        if strict and (not isinstance(module_name, str) or not module_name.strip()):
            return False, f"api_modules[{midx}].module_name rỗng"
        endpoints = module.get("endpoints")
        if not isinstance(endpoints, list):
            return False, f"api_modules[{midx}] thiếu endpoints array"
        for eidx, ep in enumerate(endpoints):
            total_endpoints += 1
            if not isinstance(ep, dict):
                return False, f"api_modules[{midx}].endpoints[{eidx}] phải là object"
            method = str(ep.get("method", "")).upper().strip()
            path = str(ep.get("endpoint_path", "")).strip()
            if method not in API_HTTP_METHODS:
                return False, f"Endpoint {midx}.{eidx} method không hợp lệ: {method!r}"
            if strict and not path:
                return False, f"Endpoint {midx}.{eidx} endpoint_path rỗng"
            rules = ep.get("test_rules")
            if not isinstance(rules, list):
                return False, f"Endpoint {midx}.{eidx} thiếu test_rules array"
            for ridx, rule in enumerate(rules):
                total_rules += 1
                if not isinstance(rule, dict):
                    return False, f"Rule {midx}.{eidx}.{ridx} phải là object"
                missing = sorted(API_RULE_FIELDS - set(rule.keys()))
                if missing:
                    return False, f"Rule {midx}.{eidx}.{ridx} thiếu field: {', '.join(missing)}"
                category = str(rule.get("category", "")).upper().strip()
                rule_type = str(rule.get("rule_type", "")).upper().strip()
                source_document = str(rule.get("source_document", "")).upper().strip()
                reconciliation_status = str(rule.get("reconciliation_status", "")).upper().strip()
                if category not in API_ALLOWED_CATEGORIES:
                    return False, f"Rule {midx}.{eidx}.{ridx} category không hợp lệ: {category!r}"
                if rule_type not in API_ALLOWED_RULE_TYPES:
                    return False, f"Rule {midx}.{eidx}.{ridx} rule_type không hợp lệ: {rule_type!r}"
                if source_document not in API_ALLOWED_SOURCE_DOCUMENTS:
                    return False, f"Rule {midx}.{eidx}.{ridx} source_document không hợp lệ: {source_document!r}"
                if reconciliation_status not in API_ALLOWED_RECONCILIATION_STATUSES:
                    return False, f"Rule {midx}.{eidx}.{ridx} reconciliation_status không hợp lệ: {reconciliation_status!r}"
                if strict:
                    for field in API_RULE_FIELDS:
                        if not isinstance(rule.get(field), str):
                            return False, f"Rule {midx}.{eidx}.{ridx}.{field} phải là string"
                    # rule_id is bookkeeping only and is assigned after merge/dedup.
                    # It must never block semantic analysis because of empty/duplicate model IDs.
                    if not rule.get("source_requirement", "").strip():
                        return False, f"Rule {midx}.{eidx}.{ridx} thiếu source_requirement"
                    if rule_type == "DERIVED" and not rule.get("applied_qa_rule", "").strip():
                        return False, f"Rule {midx}.{eidx}.{ridx} DERIVED thiếu applied_qa_rule"

    return True, f"API Schema OK | Modules={len(modules)} | Endpoints={total_endpoints} | TestRules={total_rules}"


def _api_rule_signature(rule: dict) -> tuple:
    return (
        _normalize_text(rule.get("target")),
        _normalize_text(rule.get("category")),
        _normalize_text(rule.get("rule_type")),
        _normalize_text(rule.get("rule_name")),
        _normalize_text(rule.get("test_condition")),
        _normalize_text(rule.get("expected_http_code")),
        _normalize_text(rule.get("expected_status")),
        _normalize_text(rule.get("expected_code")),
        _normalize_text(rule.get("expected_message")),
        _normalize_text(rule.get("expected_data_body")),
        _normalize_text(rule.get("business_result")),
        _normalize_text(rule.get("applied_qa_rule")),
    )


def _infer_api_qa_technique(rule: dict) -> str:
    """Deterministically recover missing QA-technique metadata for DERIVED rules.

    This is intentionally conservative: it never changes the testcase condition or
    expected business behavior. It only classifies the derivation technique from
    wording already present in the generated rule/source traceability.
    """
    text = " ".join(
        str(rule.get(k, "") or "")
        for k in (
            "rule_name", "test_objective", "test_condition", "source_requirement",
            "generation_reason", "target", "category",
        )
    ).lower()

    # Boundary Value Analysis: exact/min/max/limit and neighbouring values.
    boundary_markers = (
        "n-1", "n + 1", "n+1", "boundary", "biên", "độ dài", "length",
        "minlength", "maxlength", "minimum", "maximum", "tối thiểu", "tối đa",
        "ký tự", "character", "decimal", "số thập phân", "file size", "kích thước",
        "selection count", "số lượng",
    )
    if any(marker in text for marker in boundary_markers):
        return "Boundary Value Analysis"

    # Decision Table: compound/branching conditions and visibility/availability rules.
    decision_markers = (
        "only when", "chỉ khi", "đồng thời", " and ", " or ", " && ", " || ",
        "visible", "visibility", "hiển thị khi", "workflow", "checker",
        "trạng thái", "status", "điều kiện a", "điều kiện b",
    )
    if any(marker in text for marker in decision_markers):
        return "Decision Table"

    # Remaining catalog derivations are positive/negative input partitions:
    # required, null/empty, type/character class, format, enum, etc.
    return "Equivalence Partitioning"


def normalize_api_rule_semantics(data: dict) -> tuple[dict, list[dict]]:
    """Repair non-business semantic metadata before final strict validation.

    A single missing `applied_qa_rule` is metadata damage, not a reason to discard an
    otherwise valid multi-minute AI run. The function never invents response codes,
    messages, business outcomes, source requirements, or testcase conditions.
    """
    repairs: list[dict] = []
    if not isinstance(data, dict):
        return data, repairs

    for module in data.get("api_modules", []) if isinstance(data.get("api_modules"), list) else []:
        for ep in module.get("endpoints", []) if isinstance(module.get("endpoints"), list) else []:
            for rule in ep.get("test_rules", []) if isinstance(ep.get("test_rules"), list) else []:
                if not isinstance(rule, dict):
                    continue
                rule_type = str(rule.get("rule_type", "") or "").upper().strip()
                current = str(rule.get("applied_qa_rule", "") or "").strip()

                if rule_type == "EXPLICIT" and not current:
                    rule["applied_qa_rule"] = "EXPLICIT FROM SPEC"
                    repairs.append({
                        "rule_id": str(rule.get("rule_id", "")),
                        "field": "applied_qa_rule",
                        "value": "EXPLICIT FROM SPEC",
                        "reason": "explicit_rule_metadata_normalized",
                    })

                elif rule_type == "DERIVED" and not current:
                    technique = _infer_api_qa_technique(rule)
                    rule["applied_qa_rule"] = technique
                    if not str(rule.get("generation_reason", "") or "").strip():
                        rule["generation_reason"] = (
                            f"Sinh testcase từ ràng buộc nguồn bằng kỹ thuật {technique}."
                        )
                    repairs.append({
                        "rule_id": str(rule.get("rule_id", "")),
                        "field": "applied_qa_rule",
                        "value": technique,
                        "reason": "derived_rule_metadata_recovered",
                    })

                # Traceability is mandatory for review, but one missing text field must not
                # discard an otherwise valid multi-minute run. Use an explicit marker rather
                # than inventing a requirement. The marker is visible to reviewers/export.
                if not str(rule.get("source_requirement", "") or "").strip():
                    source_document = str(rule.get("source_document") or "UNKNOWN").strip().upper()
                    marker = f"[TRACEABILITY_MISSING:{source_document}]"
                    rule["source_requirement"] = marker
                    repairs.append({
                        "rule_id": str(rule.get("rule_id", "")),
                        "field": "source_requirement",
                        "value": marker,
                        "reason": "missing_traceability_marked_for_review",
                    })

    return data, repairs


def _api_rule_concept_signature(rule: dict) -> tuple:
    """Stable logical identity for exact reconciliation across Spec/BA/chunks.

    `applied_qa_rule` is deliberately excluded because it is metadata and frequently
    differs between an API_SPEC extraction and a BA enrichment of the same testcase.
    """
    return (
        _normalize_text(rule.get("target")),
        _normalize_text(rule.get("category")),
        _normalize_text(rule.get("rule_name")),
        _normalize_text(rule.get("test_condition")),
    )


def _api_rule_match_text(rule: dict) -> str:
    return _normalize_search_text(" ".join(
        str(rule.get(k, "") or "")
        for k in ("target", "rule_name", "test_condition")
    ))


def _api_rules_semantically_equivalent(a: dict, b: dict) -> bool:
    """Conservative fuzzy reconciliation for the same testcase worded differently.

    This is only used cross-document (API_SPEC vs BA) and only inside one endpoint.
    It never merges different QA categories or obvious boundary values.
    """
    if str(a.get("source_document", "")).upper() == str(b.get("source_document", "")).upper():
        return False
    if _normalize_text(a.get("category")) != _normalize_text(b.get("category")):
        return False

    at = _normalize_text(a.get("target"))
    bt = _normalize_text(b.get("target"))
    if at and bt and at != bt:
        # Allow minor target wording drift only when one contains the other.
        if at not in bt and bt not in at:
            return False

    a_text = _api_rule_match_text(a)
    b_text = _api_rule_match_text(b)
    if not a_text or not b_text:
        return False

    # Never fuzzy-merge cases that carry different explicit numeric boundary values.
    a_numbers = re.findall(r"(?<![a-z])\d+(?![a-z])", a_text)
    b_numbers = re.findall(r"(?<![a-z])\d+(?![a-z])", b_text)
    if a_numbers and b_numbers and a_numbers != b_numbers:
        return False

    seq = SequenceMatcher(None, a_text, b_text).ratio()
    a_tokens = {t for t in re.findall(r"[a-z0-9_./-]+", a_text) if len(t) >= 2}
    b_tokens = {t for t in re.findall(r"[a-z0-9_./-]+", b_text) if len(t) >= 2}
    union = a_tokens | b_tokens
    jaccard = (len(a_tokens & b_tokens) / len(union)) if union else 0.0

    a_condition = _normalize_search_text(a.get("test_condition", ""))
    b_condition = _normalize_search_text(b.get("test_condition", ""))

    # Guard against merging opposite/independent partitions that happen to share most words.
    neg_markers = (" khong ", " invalid ", " sai ", " thieu ", " missing ", " expired ", " fail ", " timeout ")
    a_pad = f" {a_condition} "
    b_pad = f" {b_condition} "
    if any(x in a_pad for x in neg_markers) != any(x in b_pad for x in neg_markers):
        return False

    partition_markers = {"null", "empty", "rong", "blank", "whitespace", "missing", "thieu"}
    a_parts = partition_markers & set(re.findall(r"[a-z0-9_]+", a_condition))
    b_parts = partition_markers & set(re.findall(r"[a-z0-9_]+", b_condition))
    if a_parts and b_parts and a_parts != b_parts:
        return False

    condition_seq = (
        SequenceMatcher(None, a_condition, b_condition).ratio()
        if a_condition and b_condition else 0.0
    )

    # Same target/category plus nearly identical condition is enough for BA enrichment;
    # otherwise require very strong whole-rule similarity/token overlap.
    same_target = bool(at and bt and at == bt)
    return (
        (same_target and condition_seq >= 0.82)
        or seq >= 0.84
        or (len(a_tokens & b_tokens) >= 4 and jaccard >= 0.68)
    )


def _api_rules_have_contract_conflict(a: dict, b: dict) -> bool:
    """Conservative conflict detection: only exact contract/result fields can conflict.

    Longer business_result prose is not compared because two docs often describe the
    same outcome at different levels of detail.
    """
    for fld in ("expected_http_code", "expected_status", "expected_code", "expected_message"):
        av = str(a.get(fld, "")).strip()
        bv = str(b.get(fld, "")).strip()
        if av and bv and _normalize_text(av) != _normalize_text(bv):
            return True
    return False


def _merge_api_source_document(a: str, b: str) -> str:
    aa, bb = str(a or "").upper(), str(b or "").upper()
    if aa == bb:
        return aa or bb or "API_SPEC"
    if "BOTH" in {aa, bb}:
        return "BOTH"
    if {aa, bb} == {"API_SPEC", "BA"}:
        return "BOTH"
    return aa or bb or "API_SPEC"


def _merge_api_reconciliation_status(a: str, b: str, merged_source: str) -> str:
    aa, bb = str(a or "").upper().strip(), str(b or "").upper().strip()
    if "CONFLICT" in {aa, bb}:
        return "CONFLICT"
    if "DERIVED" in {aa, bb}:
        return "DERIVED"
    if merged_source == "BOTH":
        if aa == bb == "CONSISTENT":
            return "CONSISTENT"
        return "COMPLEMENTARY"
    return aa or bb or "DOC_ONLY"


def merge_api_rule_matrices(matrices: list[dict]) -> tuple[dict, dict]:
    """Merge by TARGET ENDPOINT first, not by AI-generated module label.

    Endpoint identity comes from API Design. Module-name drift between API Spec and BA is
    ignored; the same method+path is one logical API and reconciles into one workspace.
    """
    endpoints_map: OrderedDict[tuple, dict] = OrderedDict()
    concept_maps: dict[tuple, dict] = {}
    raw_rules = 0
    duplicates = 0
    conflicts = 0

    for matrix in matrices:
        for module in matrix.get("api_modules", []):
            module_name = str(module.get("module_name", "")).strip() or "Target API"
            for ep in module.get("endpoints", []):
                method = str(ep.get("method", "UNMAPPED")).upper().strip() or "UNMAPPED"
                path = str(ep.get("endpoint_path", "UNMAPPED")).strip() or "UNMAPPED"
                summary_text = str(ep.get("summary", "")).strip()
                # Continuation APIs may be named by BA without exposing an exact endpoint path.
                # Keep distinct UNMAPPED operations separated by summary rather than collapsing them.
                if path == "UNMAPPED":
                    ekey = (method, "UNMAPPED::" + _normalize_text(summary_text or module_name))
                else:
                    ekey = (method, _normalize_text(path))
                if ekey not in endpoints_map:
                    endpoints_map[ekey] = {
                        "module_name": module_name,
                        "method": method,
                        "endpoint_path": path,
                        "summary": str(ep.get("summary", "")).strip(),
                        "test_rules": [],
                    }
                    concept_maps[ekey] = {}
                target_ep = endpoints_map[ekey]
                if not target_ep.get("summary") and ep.get("summary"):
                    target_ep["summary"] = str(ep.get("summary", "")).strip()
                if target_ep.get("module_name") in {"Target API", "UNMAPPED"} and module_name not in {"Target API", "UNMAPPED"}:
                    target_ep["module_name"] = module_name

                concept_map = concept_maps[ekey]
                for incoming in ep.get("test_rules", []):
                    raw_rules += 1
                    rule = copy.deepcopy(incoming)
                    concept = _api_rule_concept_signature(rule)
                    existing_idx = concept_map.get(concept)

                    # Cross-document wording often differs slightly even when BA only adds
                    # an error code/message to a testcase already defined by API Spec.
                    # Reconcile such cases conservatively instead of creating duplicate cases.
                    if existing_idx is None:
                        for candidate_idx, candidate in enumerate(target_ep["test_rules"]):
                            if _api_rules_semantically_equivalent(candidate, rule):
                                existing_idx = candidate_idx
                                break

                    if existing_idx is not None:
                        existing = target_ep["test_rules"][existing_idx]
                        if _api_rules_have_contract_conflict(existing, rule):
                            conflicts += 1
                            existing["reconciliation_status"] = "CONFLICT"
                            rule["reconciliation_status"] = "CONFLICT"
                            target_ep["test_rules"].append(rule)
                            continue

                        duplicates += 1
                        merged_source = _merge_api_source_document(existing.get("source_document"), rule.get("source_document"))
                        existing["source_document"] = merged_source
                        existing["reconciliation_status"] = _merge_api_reconciliation_status(
                            existing.get("reconciliation_status"), rule.get("reconciliation_status"), merged_source
                        )
                        old_src = str(existing.get("source_requirement", "")).strip()
                        new_src = str(rule.get("source_requirement", "")).strip()
                        if new_src and _normalize_text(new_src) != _normalize_text(old_src):
                            existing["source_requirement"] = (old_src + " | " + new_src).strip(" |")
                        for fld in (
                            "precondition", "test_data", "expected_http_code", "expected_status",
                            "expected_code", "expected_message", "expected_trace_id",
                            "expected_data_body", "business_result", "generation_reason",
                        ):
                            if not str(existing.get(fld, "")).strip() and str(rule.get(fld, "")).strip():
                                existing[fld] = rule.get(fld, "")
                        concept_map.setdefault(concept, existing_idx)
                        continue

                    concept_map[concept] = len(target_ep["test_rules"])
                    target_ep["test_rules"].append(rule)

    final_modules = []
    global_rule_counter = 0
    unmapped_endpoints = 0
    for ep_data in endpoints_map.values():
        if ep_data.get("method") == "UNMAPPED" or ep_data.get("endpoint_path") == "UNMAPPED":
            unmapped_endpoints += 1
        # IDs are bookkeeping only. Assign them after business merge/dedup is final.
        for rule in ep_data.get("test_rules", []):
            global_rule_counter += 1
            rule["rule_id"] = f"R_{global_rule_counter:04d}"
        final_modules.append({
            "module_name": ep_data.pop("module_name", "Target API"),
            "endpoints": [ep_data],
        })

    final = {"api_test_design_version": "3.3", "api_modules": final_modules}
    stats = {
        "input_matrices": len(matrices),
        "modules": len(final_modules),
        "endpoints": len(endpoints_map),
        "unmapped_endpoints": unmapped_endpoints,
        "raw_rules": raw_rules,
        "duplicates_removed": duplicates,
        "conflicts_detected": conflicts,
        "final_rules": global_rule_counter,
    }
    return final, stats



def build_api_spec_case_reference(
    matrices: list[dict],
    target_endpoint: dict | None,
    *,
    max_chars: int = 14000,
) -> str:
    """Build a compact testcase inventory from the *actual* API_SPEC Agent1 output.

    BA deep-analysis uses this only as reconciliation context. It prevents BA from
    independently regenerating Permission/Validation/Happy-Path cases just to contribute
    an error code/message. No rule may be created solely from this reference.
    """
    target_endpoint = target_endpoint or {}
    target_method = str(target_endpoint.get("method", "UNMAPPED")).upper().strip() or "UNMAPPED"
    target_path = str(target_endpoint.get("endpoint_path") or target_endpoint.get("path") or "UNMAPPED").strip() or "UNMAPPED"
    technical_categories = {"AUTH", "PERMISSION", "VALIDATION", "HAPPY_PATH"}
    items: list[dict] = []
    seen: set[tuple] = set()

    for matrix in matrices:
        if not isinstance(matrix, dict):
            continue
        for module in matrix.get("api_modules", []) if isinstance(matrix.get("api_modules"), list) else []:
            for ep in module.get("endpoints", []) if isinstance(module.get("endpoints"), list) else []:
                method = str(ep.get("method", "UNMAPPED")).upper().strip() or "UNMAPPED"
                path = str(ep.get("endpoint_path", "UNMAPPED")).strip() or "UNMAPPED"
                if (method, _normalize_text(path)) != (target_method, _normalize_text(target_path)):
                    continue
                for rule in ep.get("test_rules", []) if isinstance(ep.get("test_rules"), list) else []:
                    if not isinstance(rule, dict):
                        continue
                    category = str(rule.get("category", "") or "").upper().strip()
                    if category not in technical_categories:
                        continue
                    identity = _api_rule_concept_signature(rule)
                    if identity in seen:
                        continue
                    seen.add(identity)
                    items.append({
                        "category": category,
                        "target": str(rule.get("target", "") or ""),
                        "rule_name": str(rule.get("rule_name", "") or ""),
                        "test_condition": str(rule.get("test_condition", "") or ""),
                        "expected_http_code": str(rule.get("expected_http_code", "") or ""),
                        "expected_code": str(rule.get("expected_code", "") or ""),
                        "expected_message": str(rule.get("expected_message", "") or ""),
                    })

    if not items:
        return "[]"

    # Keep valid JSON while respecting prompt size. Earlier cases normally represent the
    # primary contract order from the Spec and are the most useful reconciliation anchors.
    selected: list[dict] = []
    for item in items:
        candidate = selected + [item]
        encoded = json.dumps(candidate, ensure_ascii=False, separators=(",", ":"))
        if len(encoded) > max_chars:
            break
        selected.append(item)

    return json.dumps(selected, ensure_ascii=False, separators=(",", ":"))


def enforce_api_source_responsibility(data: dict) -> tuple[dict, dict]:
    """Enforce the multi-document source contract after Spec/BA reconciliation.

    Product contract:
    - API_SPEC owns technical testcase creation: AUTH/PERMISSION/VALIDATION/HAPPY_PATH.
    - BA owns BUSINESS_RULE testcase creation.
    - BA technical records are enrichment-only. They survive only when they reconciled with
      a SPEC testcase (source_document=BOTH) or when a real contract conflict must stay visible.
    - API_SPEC-only BUSINESS_RULE records are dropped so the Design document cannot duplicate
      business cases that BA is responsible for.

    This is intentionally applied *after* merge so BA can still enrich exact error code/message
    on a matching SPEC testcase without creating a second testcase.
    """
    if not isinstance(data, dict):
        return data, {"dropped_total": 0, "kept_total": 0, "drop_reasons": {}}

    technical_categories = {"AUTH", "PERMISSION", "VALIDATION", "HAPPY_PATH"}
    dropped_total = 0
    kept_total = 0
    reasons = Counter()

    for module in data.get("api_modules", []) if isinstance(data.get("api_modules"), list) else []:
        if not isinstance(module, dict):
            continue
        for endpoint in module.get("endpoints", []) if isinstance(module.get("endpoints"), list) else []:
            if not isinstance(endpoint, dict):
                continue
            kept_rules = []
            for rule in endpoint.get("test_rules", []) if isinstance(endpoint.get("test_rules"), list) else []:
                if not isinstance(rule, dict):
                    continue
                category = str(rule.get("category", "") or "").upper().strip()
                source = str(rule.get("source_document", "") or "").upper().strip()
                rec = str(rule.get("reconciliation_status", "") or "").upper().strip()

                keep = True
                reason = ""
                if rec == "CONFLICT":
                    # A documented conflict is review-worthy and must not be hidden merely
                    # because it crossed the normal source responsibility boundary.
                    keep = True
                elif category == "BUSINESS_RULE" and source == "API_SPEC":
                    keep = False
                    reason = "API_SPEC_BUSINESS_RULE"
                elif category in technical_categories and source == "BA":
                    keep = False
                    reason = "BA_TECHNICAL_UNMATCHED"

                if keep:
                    kept_rules.append(rule)
                    kept_total += 1
                else:
                    dropped_total += 1
                    reasons[reason] += 1

            endpoint["test_rules"] = kept_rules

    # Filtering can create rule-id gaps. IDs are bookkeeping only, so renumber once here.
    counter = 0
    for module in data.get("api_modules", []) if isinstance(data.get("api_modules"), list) else []:
        for endpoint in module.get("endpoints", []) if isinstance(module.get("endpoints"), list) else []:
            for rule in endpoint.get("test_rules", []) if isinstance(endpoint.get("test_rules"), list) else []:
                counter += 1
                rule["rule_id"] = f"R_{counter:04d}"

    return data, {
        "dropped_total": dropped_total,
        "kept_total": kept_total,
        "drop_reasons": dict(reasons),
    }


def _api_split_chunk_for_retry(chunk: dict) -> list[dict]:
    core = chunk.get("core_text", "")
    if len(core) <= API_AGENT1_MIN_RECURSIVE_CHARS:
        return []
    target = max(API_AGENT1_MIN_RECURSIVE_CHARS, min(6000, len(core) // 2))
    max_chars = max(target + 700, min(7600, len(core)))
    local = split_document_semantic(
        core,
        base_title=chunk.get("title", "API_RetryChunk"),
        target_chars=target,
        max_chars=max_chars,
        context_chars=min(API_AGENT1_CONTEXT_CHARS, 700),
    )
    if len(local) <= 1:
        mid = len(core) // 2
        cut = core.rfind("\n", 0, mid)
        if cut < API_AGENT1_MIN_RECURSIVE_CHARS:
            cut = core.rfind(" ", 0, mid)
        if cut < API_AGENT1_MIN_RECURSIVE_CHARS:
            cut = mid
        local = []
        for part in [core[:cut].strip(), core[cut:].strip()]:
            if part:
                local.append({
                    "title": chunk.get("title", "API_RetryChunk"),
                    "core_text": part,
                    "prev_context": "",
                    "next_context": "",
                    "char_count": len(part),
                    "depth": chunk.get("depth", 0) + 1,
                })
    children = []
    parent_id = chunk.get("chunk_id", "A")
    for idx, child in enumerate(local):
        child = dict(child)
        child["chunk_id"] = f"{parent_id}.{idx+1}"
        child["depth"] = chunk.get("depth", 0) + 1
        children.append(child)
    return children


def process_api_agent1_chunk_recursive(
    chunk: dict,
    endpoint_index: list[dict],
    api_key: str,
    base_url: str,
    model: str,
    prompt_template: str,
    cache: dict | None = None,
    diagnostics: list | None = None,
    source_document: str = "API_SPEC",
    target_endpoint: dict | None = None,
    spec_case_reference: str = "",
) -> tuple[bool, list[dict]]:
    diagnostics = diagnostics if diagnostics is not None else []
    payload = build_api_agent1_chunk_payload(
        chunk,
        endpoint_index,
        source_document=source_document,
        target_endpoint=target_endpoint,
        spec_case_reference=spec_case_reference,
    )
    key = _cache_key(API_AGENT1_CACHE_NAMESPACE, model, prompt_template, payload)
    if cache is not None and key in cache:
        diagnostics.append({"chunk_id": chunk.get("chunk_id"), "status": "CACHE_HIT", "chars": len(chunk.get("core_text", ""))})
        return True, [copy.deepcopy(cache[key])]

    call_result = None
    for attempt in range(API_AGENT1_API_RETRIES + 1):
        agent_name = f"API_AGENT1/{chunk.get('chunk_id')}"
        call_result = call_qwen_max_agent_detailed(
            content=payload,
            api_key=api_key,
            base_url=base_url,
            model=model,
            prompt_template=prompt_template,
            max_tokens=API_DEEP_MAX_OUTPUT_TOKENS,
            agent_name=agent_name,
            enable_thinking=False,
            response_format=api_rule_matrix_response_format(),
        )
        if call_result.ok or call_result.finish_reason == "length":
            break
        if attempt < API_AGENT1_API_RETRIES:
            log_info(f"[{agent_name}] 🔁 Retry API lần {attempt+2}/{API_AGENT1_API_RETRIES+1}")
            time.sleep(2)

    assert call_result is not None
    reason_to_split = None
    parse_diag = None
    schema_diag = None
    parsed_json = None

    if call_result.finish_reason == "length":
        reason_to_split = "MAX_TOKENS"
    elif not call_result.ok:
        diagnostics.append({"chunk_id": chunk.get("chunk_id"), "status": "API_FAILED", "error": call_result.error})
        return False, []
    else:
        parsed_ok, parsed_json, parse_diag = extract_json_from_model_response(call_result.text)

        # Do not split the BA source immediately for a serialization mistake.
        # First run one syntax-only repair on the model OUTPUT. This preserves
        # the business section boundary and avoids losing context across children.
        if not parsed_ok:
            log_info(f"[{agent_name}] JSON parse fail; thử 1 lần syntax-only repair trước khi split source.")
            repair_ok, repaired_json, repair_diag = repair_api_json_syntax(
                call_result.text,
                api_key=api_key,
                base_url=base_url,
                model=model,
                agent_name=agent_name,
            )
            parse_diag = f"{parse_diag} | {repair_diag}"
            if repair_ok:
                parsed_ok = True
                parsed_json = repaired_json
                log_info(f"[{agent_name}] ✅ JSON syntax repair thành công; không cần split source.")
            else:
                reason_to_split = "JSON_PARSE_FAIL"

        if parsed_ok:
            parsed_json = normalize_api_rule_shape(parsed_json)
            parsed_json = normalize_api_rule_matrix_enums(parsed_json)
            parsed_json = force_api_matrix_scope(
                parsed_json, source_document=source_document, target_endpoint=target_endpoint
            )
            parsed_json, leaf_semantic_repairs = normalize_api_rule_semantics(parsed_json)
            schema_ok, schema_diag = validate_api_rule_matrix_schema(parsed_json, strict=False)
            if not schema_ok:
                reason_to_split = "SCHEMA_FAIL"
                log_error(f"[{agent_name}] Schema validation failed | {schema_diag}")
            else:
                modules = parsed_json.get("api_modules", [])
                rules = sum(len(ep.get("test_rules", [])) for m in modules for ep in m.get("endpoints", []))
                endpoints = sum(len(m.get("endpoints", [])) for m in modules)
                diagnostics.append({
                    "chunk_id": chunk.get("chunk_id"), "status": "OK",
                    "chars": len(chunk.get("core_text", "")), "depth": chunk.get("depth", 0),
                    "elapsed": round(call_result.elapsed, 2), "output_chars": len(call_result.text),
                    "endpoints": endpoints, "rules": rules,
                    "json_repaired": "JSON repair OK" in str(parse_diag or ""),
                    "semantic_metadata_repairs": len(leaf_semantic_repairs),
                })
                if cache is not None:
                    cache[key] = copy.deepcopy(parsed_json)
                return True, [parsed_json]

    depth = int(chunk.get("depth", 0))
    split_reason_is_retryable = (
        reason_to_split in {"MAX_TOKENS", "JSON_PARSE_FAIL"}
        or (reason_to_split == "SCHEMA_FAIL" and _schema_failure_can_benefit_from_split(schema_diag))
    )
    can_split = (
        split_reason_is_retryable
        and depth < API_AGENT1_MAX_RECURSION_DEPTH
        and len(chunk.get("core_text", "")) > API_AGENT1_MIN_RECURSIVE_CHARS
    )
    diagnostics.append({
        "chunk_id": chunk.get("chunk_id"), "status": "SPLIT_RETRY" if can_split else "FAILED_LEAF",
        "reason": reason_to_split, "parse_diag": parse_diag, "schema_diag": schema_diag,
        "finish_reason": call_result.finish_reason, "output_chars": len(call_result.text),
        "raw_tail": call_result.text[-1200:] if call_result.text else "",
    })
    if not can_split:
        return False, []
    children = _api_split_chunk_for_retry(chunk)
    if len(children) < 2:
        return False, []
    log_info(
        f"[API_AGENT1/{chunk.get('chunk_id')}] ✂️ Recursive split do {reason_to_split} | "
        f"{len(chunk.get('core_text', '')):,} chars -> " + " + ".join(f"{len(c.get('core_text', '')):,}" for c in children)
    )
    all_matrices = []
    for child in children:
        ok, mats = process_api_agent1_chunk_recursive(
            child, endpoint_index, api_key, base_url, model, prompt_template, cache, diagnostics,
            source_document=source_document,
            target_endpoint=target_endpoint,
            spec_case_reference=spec_case_reference,
        )
        if not ok:
            return False, []
        all_matrices.extend(mats)
    return True, all_matrices


def run_api_agent1_document_pipeline(
    design_text: str,
    ba_text: str,
    base_filename: str,
    api_key: str,
    base_url: str,
    model: str,
    prompt_template: str,
    cache: dict | None = None,
    progress_callback=None,
) -> tuple[bool, dict | None, dict]:
    """V1.11.4 source-role single-target multi-document API pipeline.

    API Design/Spec defines exactly one API under test. The BA document is scanned only for
    relevance; deep QA generation runs only on BA chunks that enrich that same target API.
    Distinct sibling/continuation APIs are excluded so they cannot create extra folders.
    """
    started = time.time()
    endpoint_index = extract_api_endpoint_index(design_text)
    targets = build_target_api_descriptors(design_text, endpoint_index)
    targets = select_primary_api_targets(design_text, targets, document_hint=base_filename)
    diagnostics: list[dict] = []
    matrices: list[dict] = []
    spec_matrices: list[dict] = []

    log_info(
        f"[API_TARGET_PIPELINE] 🎯 TargetAPIs={len(targets)} | "
        + ", ".join(f"{t.get('method')} {t.get('endpoint_path')}" for t in targets)
    )

    # Stage 1 — deep analysis of API contract itself, target by target.
    if progress_callback:
        progress_callback(0, 100, f"Đã nhận diện {len(targets)} API mục tiêu từ API Design")
    for t_idx, target in enumerate(targets, start=1):
        spec_excerpt = target.get("spec_excerpt") or design_text
        spec_chunks = split_document_semantic(
            spec_excerpt,
            base_title=f"API_SPEC::{target.get('endpoint_path')}",
            target_chars=API_AGENT1_CHUNK_TARGET_CHARS,
            max_chars=API_AGENT1_CHUNK_MAX_CHARS,
            context_chars=API_AGENT1_CONTEXT_CHARS,
        )
        for c_idx, chunk in enumerate(spec_chunks, start=1):
            chunk["chunk_id"] = f"SPEC-{t_idx}.{c_idx}"
            if progress_callback:
                progress_callback(
                    5 + int(15 * ((t_idx-1 + c_idx/max(1,len(spec_chunks))) / max(1,len(targets)))),
                    100,
                    f"Đang đọc contract API {target.get('method')} {target.get('endpoint_path')} ({c_idx}/{len(spec_chunks)})",
                )
            ok, mats = process_api_agent1_chunk_recursive(
                chunk, endpoint_index, api_key, base_url, model, prompt_template,
                cache, diagnostics, source_document="API_SPEC", target_endpoint=target,
            )
            if not ok:
                return False, None, {
                    "ok": False, "stage": "api_spec", "endpoint_index": len(endpoint_index),
                    "targets": targets, "elapsed": round(time.time()-started,2), "diagnostics": diagnostics,
                }
            spec_matrices.extend(mats)
            matrices.extend(mats)

    # Build reconciliation anchors from the real Spec-generated testcase inventory, not
    # from raw design prose. BA may use these anchors only to enrich matching technical cases.
    spec_case_references = {
        _api_endpoint_key(target): build_api_spec_case_reference(spec_matrices, target)
        for target in targets
    }

    # Stage 2 — cheap BA scope scan. It reads chunks for routing only, no testcase generation.
    # PRIMARY/DEPENDENCY enrich the target API; sibling/continuation operations are OUT_OF_SCOPE.
    ba_chunks = split_document_semantic(
        ba_text,
        base_title=f"{base_filename}::BA",
        target_chars=API_SCOPE_SCAN_CHUNK_TARGET_CHARS,
        max_chars=API_SCOPE_SCAN_CHUNK_MAX_CHARS,
        context_chars=450,
    )
    routed_pairs: list[tuple[dict, dict, str]] = []
    scope_diagnostics = []
    # BA must never create endpoint workspaces. Keep the counter only for diagnostics
    # in case an older/model response still emits CONTINUATION.
    continuation_targets: OrderedDict[tuple, dict] = OrderedDict()
    scope_role_counts = {"PRIMARY": 0, "CONTINUATION": 0, "DEPENDENCY": 0, "OUT_OF_SCOPE": 0}
    for idx, chunk in enumerate(ba_chunks, start=1):
        chunk["chunk_id"] = f"BA-SCAN-{idx}"

    scope_gate_overrides: list[dict] = []

    def _scan_one(chunk):
        try:
            ok, matches, diag = route_ba_chunk_to_target_apis(
                chunk, targets, api_key, base_url, model, cache
            )
        except Exception as exc:
            ok, matches = False, []
            diag = {
                "status": "ROUTER_EXCEPTION",
                "chunk_id": chunk.get("chunk_id"),
                "error": str(exc),
            }
        if not ok:
            # Fallback is intentionally PRIMARY-only. It must never guess a continuation API.
            fallback = fallback_route_ba_chunk_by_spec_overlap(chunk, targets)
            matches = [{**m, "scope_role": "PRIMARY"} for m in fallback]
            diag = {**(diag or {}), "fallback_matches": len(matches)}
        return chunk, matches, diag

    def _accept_match(chunk, match):
        primary_target = next(
            (t for t in targets if _api_endpoint_key(t) == (match.get("method"), match.get("endpoint_path"))),
            None,
        )
        if primary_target is None:
            return
        guarded_match, gate_diag = _gate_scope_match(chunk, primary_target, match)
        if gate_diag:
            scope_gate_overrides.append(gate_diag)
        role = _normalize_scope_role(guarded_match.get("scope_role"))
        scope_role_counts[role] = scope_role_counts.get(role, 0) + 1
        if role == "PRIMARY":
            routed_pairs.append((chunk, primary_target, role))
        elif role == "DEPENDENCY":
            # Dependency context may explain a direct outcome of the target API, but remains
            # inside the target endpoint. It never gets its own endpoint workspace.
            routed_pairs.append((chunk, primary_target, role))
        elif role == "CONTINUATION":
            # Strict target mode: never create a second endpoint from BA. The gate normally
            # repairs same-operation wording to PRIMARY and turns the rest OUT_OF_SCOPE.
            scope_role_counts["OUT_OF_SCOPE"] = scope_role_counts.get("OUT_OF_SCOPE", 0) + 1
        # OUT_OF_SCOPE: intentionally ignored.

    completed_scans = 0
    if API_SCOPE_SCAN_WORKERS > 1 and len(ba_chunks) > 1:
        with ThreadPoolExecutor(max_workers=min(API_SCOPE_SCAN_WORKERS, len(ba_chunks))) as executor:
            futures = [executor.submit(_scan_one, chunk) for chunk in ba_chunks]
            for future in as_completed(futures):
                chunk, matches, diag = future.result()
                completed_scans += 1
                scope_diagnostics.append(diag)
                if progress_callback:
                    progress_callback(
                        20 + int(30 * completed_scans / max(1, len(ba_chunks))), 100,
                        f"Đang phân loại phạm vi BA ({completed_scans}/{len(ba_chunks)})",
                    )
                for match in matches:
                    _accept_match(chunk, match)
    else:
        for chunk in ba_chunks:
            chunk, matches, diag = _scan_one(chunk)
            completed_scans += 1
            scope_diagnostics.append(diag)
            if progress_callback:
                progress_callback(
                    20 + int(30 * completed_scans / max(1, len(ba_chunks))), 100,
                    f"Đang phân loại phạm vi BA ({completed_scans}/{len(ba_chunks)})",
                )
            for match in matches:
                _accept_match(chunk, match)

    # Deduplicate same BA chunk/target/role pair. BA never creates a second endpoint.
    dedup = OrderedDict()
    for chunk, target, role in routed_pairs:
        target_identity = _api_endpoint_key(target)
        dedup[(chunk.get("chunk_id"), role) + target_identity] = (chunk, target, role)
    routed_pairs = list(dedup.values())

    if not routed_pairs and ba_chunks:
        # Never deep-analyze all BA as a fallback. Pick only a very small lexical PRIMARY shortlist.
        scored = []
        for chunk in ba_chunks:
            fallback_matches = fallback_route_ba_chunk_by_spec_overlap(chunk, targets, min_overlap=2)
            for match in fallback_matches:
                target = next(
                    (t for t in targets if _api_endpoint_key(t) == (match.get("method"), match.get("endpoint_path"))),
                    None,
                )
                if target is not None:
                    scored.append((float(match.get("confidence") or 0), chunk, target))
        scored.sort(key=lambda item: item[0], reverse=True)
        routed_pairs = [(chunk, target, "PRIMARY") for _, chunk, target in scored[:3]]
        if routed_pairs:
            log_info(
                f"[API_SCOPE_ROUTER] ⚠️ AI router không chọn chunk; dùng lexical PRIMARY shortlist "
                f"{len(routed_pairs)} chunk thay vì deep-analyze toàn BA."
            )

    selected_chunk_ids = {c.get("chunk_id") for c, _, _ in routed_pairs}
    log_info(
        f"[API_SCOPE_ROUTER] BAChunks={len(ba_chunks)} | RelevantPairs={len(routed_pairs)} | "
        f"Continuations={len(continuation_targets)} | Roles={scope_role_counts} | "
        f"GateOverrides={len(scope_gate_overrides)} | "
        f"Skipped={max(0, len(ba_chunks)-len(selected_chunk_ids))}"
    )

    # Stage 3 — deep analysis ONLY on selected BA chunks.
    # PRIMARY + DEPENDENCY are forced into the API Design target. Distinct continuation/sibling
    # APIs are already excluded and can never create another workspace.
    for idx, (chunk, target, scope_role) in enumerate(routed_pairs, start=1):
        deep_chunk = dict(chunk)
        deep_chunk["chunk_id"] = f"BA-DEEP-{idx}"
        if progress_callback:
            progress_callback(
                50 + int(32 * idx / max(1, len(routed_pairs))), 100,
                f"Đang phân tích BA [{scope_role}] {target.get('summary') or target.get('endpoint_path')} "
                f"({idx}/{len(routed_pairs)})",
            )
        ok, mats = process_api_agent1_chunk_recursive(
            deep_chunk, endpoint_index, api_key, base_url, model, prompt_template,
            cache, diagnostics,
            source_document="BA",
            target_endpoint=target,
            spec_case_reference=spec_case_references.get(_api_endpoint_key(target), ""),
        )
        if not ok:
            return False, None, {
                "ok": False, "stage": "ba_deep", "endpoint_index": len(endpoint_index),
                "targets": targets, "continuation_targets": list(continuation_targets.values()),
                "ba_chunks": len(ba_chunks), "relevant_pairs": len(routed_pairs),
                "scope_role_counts": scope_role_counts,
                "elapsed": round(time.time()-started,2), "diagnostics": diagnostics,
                "scope_diagnostics": scope_diagnostics,
                "scope_gate_overrides": scope_gate_overrides,
            }
        matrices.extend(mats)

    if progress_callback:
        progress_callback(85, 100, "Đang đối soát API Spec và BA, loại trùng và giữ conflict...")
    final, merge_stats = merge_api_rule_matrices(matrices)
    final = normalize_api_rule_shape(final)
    final = normalize_api_rule_matrix_enums(final)

    # Enforce document ownership after reconciliation, not before it. This lets BA enrich
    # a matching SPEC testcase with exact code/message while preventing BA from creating
    # a second technical testcase and preventing API Spec from duplicating BA business rules.
    final, source_policy_stats = enforce_api_source_responsibility(final)
    merge_stats["source_policy"] = source_policy_stats
    merge_stats["final_rules"] = sum(
        len(ep.get("test_rules", []))
        for module in final.get("api_modules", [])
        for ep in module.get("endpoints", [])
    )

    # Final semantic metadata recovery. Missing QA-technique metadata must not kill
    # an otherwise valid run after all Spec/BA deep-analysis calls have completed.
    final, semantic_repairs = normalize_api_rule_semantics(final)
    if semantic_repairs:
        merge_stats["semantic_metadata_repairs"] = len(semantic_repairs)
        log_info(
            f"[API_TARGET_PIPELINE] 🩹 Semantic metadata repaired | "
            f"Rules={len(semantic_repairs)} | "
            + ", ".join(
                f"{r.get('rule_id') or '?'}:{r.get('value')}" for r in semantic_repairs[:8]
            )
        )
    else:
        merge_stats["semantic_metadata_repairs"] = 0

    schema_ok, schema_diag = validate_api_rule_matrix_schema(final, strict=True)
    summary = {
        "ok": schema_ok,
        "document_chars": len(design_text or "") + len(ba_text or ""),
        "endpoint_index": len(endpoint_index),
        "target_apis": [{k: t.get(k) for k in ("method", "endpoint_path", "summary")} for t in targets],
        "continuation_apis": [{k: t.get(k) for k in ("method", "endpoint_path", "summary")} for t in continuation_targets.values()],
        "scope_role_counts": scope_role_counts,
        "ba_chunks_scanned": len(ba_chunks),
        "ba_relevant_pairs": len(routed_pairs),
        "leaf_matrices": len(matrices),
        "elapsed": round(time.time()-started,2),
        "merge": merge_stats,
        "semantic_repairs": semantic_repairs,
        "schema": schema_diag,
        "scope_diagnostics": scope_diagnostics,
        "scope_gate_overrides": scope_gate_overrides,
        "diagnostics": diagnostics,
    }
    if not schema_ok:
        summary["stage"] = "final_schema"
        log_error(f"[API_TARGET_PIPELINE] ❌ Final schema fail after normalization | {schema_diag}")
        return False, None, summary
    if progress_callback:
        progress_callback(90, 100, f"Đã hoàn tất Rule Matrix cho {merge_stats.get('endpoints', 0)} API/luồng tiếp nối")
    log_info(
        f"[API_TARGET_PIPELINE] ✅ Completed | Time={summary['elapsed']:.2f}s | "
        f"Targets={len(targets)} | Continuations={len(continuation_targets)} | "
        f"BARelevant={len(routed_pairs)}/{len(ba_chunks)} | "
        f"FinalRules={merge_stats['final_rules']}"
    )
    return True, final, summary














# ============================================================
# ENGLISH PROMPTS — RECOMMENDED CLEAN VERSION
# Prompt instructions: ENGLISH
# Generated Rule Matrix / Test Cases: VIETNAMESE
# PROMPT_WEB_UI intentionally removed.
# Normalized: grid ownership, traceable TC title, parser-safe multiline steps.
# ============================================================

PROMPT_WEB_DISCOVERY_INVENTORY = """
You are WEB_BA_DISCOVERY_AGENT — a structure analyst for Banking / Enterprise Web requirements.

THIS IS DISCOVERY ONLY. DO NOT DESIGN TESTCASES YET.
You MUST NOT apply QA techniques, derive boundaries, create negative cases, or produce a Rule Matrix.
Your only job is to reconstruct the BA document into a canonical inventory in this exact order:
1) identify real screens;
2) list fields/controls/grids/popups that belong to each screen;
3) attach the source-described logic to the correct screen/object;
4) preserve source evidence for the next QA stage.

LANGUAGE:
- Human-readable inventory values MUST be Vietnamese.
- Keep exact technical identifiers, field labels, screen codes, API names, enum values, messages and formats when the BA source uses them.
- JSON keys stay exactly as defined below.

==================================================
1. CHUNK OWNERSHIP
==================================================
Input has ACTIVE SCREEN HEADING, PREVIOUS CONTEXT, CURRENT SOURCE and NEXT CONTEXT.
- CURRENT SOURCE owns extracted requirements.
- ACTIVE SCREEN HEADING is copied from the nearest explicit screen heading in the BA source. You MAY use it to assign CURRENT SOURCE fields/logic to that screen when the original screen section spans multiple chunks. It is ownership evidence, not a standalone testcase requirement.
- Context may identify the screen that CURRENT SOURCE belongs to or complete a cut sentence/table.
- Do NOT extract a requirement that exists only in context.
- `=== SOURCE DOCUMENT: ... ===` is metadata, never a screen.

==================================================
2. SCREEN DISCOVERY — DO THIS FIRST
==================================================
Create a screen only when source evidence shows a real Web screen/page, for example:
- explicit "Màn hình ..." / screen title / screen code;
- a dedicated screen description section;
- a table/mockup that clearly specifies controls of that screen.

NEVER create a new screen from subsection names such as:
- Mô tả màn hình
- Logic tìm kiếm
- Điều kiện tìm kiếm
- Luồng xử lý
- Quy tắc nghiệp vụ
- Hold / Confirm / Cancel
- API mapping
- Request / Response
- Data Grid / Danh sách kết quả
- Popup / Dialog
These are parts of the owning screen.

A screen mentioned only as a navigation destination/reference is NOT enough to create a screen. It becomes a real screen only when the BA source also describes its own fields/controls/behavior.
A technical code is identity metadata; `screen_name` must be the human-readable business name.
If naming variants clearly refer to the same screen, use one stable business name.
If CURRENT SOURCE contains a requirement but the owning screen cannot be determined even with boundary context, OMIT it rather than inventing a screen.

==================================================
3. FIELD / CONTROL INVENTORY — BEFORE LOGIC
==================================================
For each discovered screen, inventory every source-described object. Do NOT invent generic controls.

FIELDS:
- textbox / textarea / input
- numeric / amount / percentage
- dropdown / combobox / autocomplete / single-select / multi-select
- date / date range / time / datetime
- checkbox / radio / toggle / switch
- upload field
- readonly/display field when it has source-defined behavior

For each field preserve EACH described property as a separate requirement item, e.g.:
- placeholder
- default value / default selection
- required / optional
- enabled / disabled / readonly / editable / visible / hidden
- length / min / max
- allowed characters / format / pattern / mask
- options / order / selection mode / select all / clear / +N
- search attributes / contains/exact / debounce
- date format / range relation / min/max date
- dependency/cascade with another field
Do NOT derive N-1/N/N+1 here. Just preserve the BA constraint.

CONTROLS:
Buttons, links, icons and direct actions such as Search, Reset, Add, Edit, Copy, View, Approve, Hold, Confirm, Cancel, Upload.
Preserve label/icon/tooltip, visibility, enable state, click/open/navigation/trigger behavior only when documented.

DATA GRID:
- one grid object per real grid;
- preserve complete active column set/order when source provides it;
- each column may preserve mapping/display meaning and presentation/format;
- preserve source-defined pagination/sort/loading/empty-state behavior as grid requirements.

POPUP:
Popup/dialog/modal belongs to its owning screen. Do NOT create a screen for it.
Preserve popup title/content/field/action/visibility behavior as popup requirements.

SCREEN-LEVEL REQUIREMENTS:
Use `screen_requirements` for screen title, breadcrumb, static sections/labels, access/permission or other requirements that are not naturally owned by one field/control/grid/popup.

==================================================
4. LOGIC BINDING — AFTER OBJECT INVENTORY
==================================================
After listing objects, extract business/interaction logic and bind it to its owner through `targets`.
Examples:
- Search uses CIF + Branch + Status and displays matching rows.
- Selecting Product A limits values of Package B.
- Button Approve is visible only in status PENDING.
- Clicking Confirm sends documented parameters and success reloads the list.
- Grid column Status displays response/business status.

For every logic item preserve:
- `logic_name`: stable business behavior name;
- `logic_type`: SEARCH | RESET | DEPENDENCY | VISIBILITY | ENABLEMENT | NAVIGATION | REQUEST_MAPPING | RESPONSE_HANDLING | BUSINESS_RULE | STATE_TRANSITION | POPUP_FLOW | GRID_MAPPING | OTHER;
- `targets`: exact field/control/grid/popup names involved;
- `condition`: source condition/input/state;
- `behavior`: exact source-described result/action;
- `source_requirement`: concise source evidence.

Do NOT turn API mentions inside a Web BA document into API-test inventory. They are Web logic only when they explain FE request/response behavior.
Do NOT invent status codes, error messages, permissions, DB behavior, timeouts or API parameters not stated by source.

==================================================
5. ACTIVE REQUIREMENT / DUPLICATE RULE
==================================================
- Later explicit revision overrides older generic wording when they conflict.
- Deleted/strikethrough/removed requirements are inactive unless explicitly reintroduced. Text wrapped in `[STRIKETHROUGH]...[/STRIKETHROUGH]` was struck through in the DOCX and MUST be treated as inactive.
- Repeated OCR/table text must not duplicate inventory items.
- Preserve details that affect behavior: exact value, length, format, role, state, mapping, message, debounce, condition.

==================================================
6. OUTPUT JSON — MANDATORY
==================================================
Return ONLY valid JSON. No markdown or commentary.

{
  "web_discovery_version": "1.0",
  "screens": [
    {
      "screen_name": "",
      "screen_code": "",
      "source_evidence": "",
      "screen_requirements": [
        {
          "property": "",
          "requirement": "",
          "source_requirement": ""
        }
      ],
      "fields": [
        {
          "field_name": "",
          "control_type": "",
          "container": "SCREEN | <popup name>",
          "requirements": [
            {
              "property": "",
              "requirement": "",
              "source_requirement": ""
            }
          ]
        }
      ],
      "controls": [
        {
          "control_name": "",
          "control_type": "BUTTON | LINK | ICON | ACTION | OTHER",
          "container": "SCREEN | <popup name>",
          "requirements": [
            {
              "property": "",
              "requirement": "",
              "source_requirement": ""
            }
          ]
        }
      ],
      "grids": [
        {
          "grid_name": "",
          "columns": [
            {
              "column_name": "",
              "mapping": "",
              "presentation": "",
              "source_requirement": ""
            }
          ],
          "requirements": [
            {
              "property": "",
              "requirement": "",
              "source_requirement": ""
            }
          ]
        }
      ],
      "popups": [
        {
          "popup_name": "",
          "requirements": [
            {
              "property": "",
              "requirement": "",
              "source_requirement": ""
            }
          ]
        }
      ],
      "logic": [
        {
          "logic_name": "",
          "logic_type": "",
          "targets": [],
          "condition": "",
          "behavior": "",
          "source_requirement": ""
        }
      ]
    }
  ]
}

Do not create empty fake screens. If CURRENT SOURCE has no requirement that can be assigned to a real screen:
{"web_discovery_version":"1.0","screens":[]}

=== BA SOURCE CHUNK ===
{content}
"""


PROMPT_WEB_QA_FROM_INVENTORY = """
You are WEB_QA_RULE_DESIGNER — a Senior QA Test Design Lead for Banking / Enterprise Web systems.

IMPORTANT ARCHITECTURE:
The BA document has ALREADY been read by a discovery stage.
The input is a CANONICAL SCREEN INVENTORY containing the real screen, its fields/controls/grids/popups, and corresponding source logic.
You MUST NOT rediscover document structure and MUST NOT create a new screen/object/logic that is absent from the inventory.

MANDATORY THINKING ORDER:
1) read the screen identity;
2) scan EVERY inventoried field/control/grid/popup;
3) read ALL logic and connect it to its listed targets;
4) create source-explicit test objectives;
5) only AFTER that, apply allowed QA rules to real constraints;
6) remove semantic duplicates before output.

ALL human-readable output MUST be Vietnamese. Keep technical identifiers/values/messages/formats exactly when needed.
`screen_name` in output MUST be exactly the canonical screen_name provided in input.

==================================================
1. SOURCE OF TRUTH / NO INVENTION
==================================================
The canonical inventory is the ONLY evidence source for this call.
- Do not add undocumented fields/buttons/APIs/roles/messages/statuses/DB behavior.
- Do not copy a behavior from one field to another.
- Do not turn an API mentioned by Web logic into API testing.
- Do not create generic responsive/cross-browser/tab-order/security/session/logging testcases unless present in inventory.
- One inventory object may have zero or many test objectives depending on its requirements.

==================================================
2. INTERNAL CATEGORY
==================================================
`category` MUST be exactly one of:
UI | VALIDATION | ACTION | DATA_GRID | BUSINESS_FLOW | EXCEPTION

UI:
- screen-level static presentation, title/breadcrumb/static labels/sections;
- pure static/readonly presentation outside a grid.

VALIDATION:
- field/control-local state and input behavior: placeholder/default/required/enable/readonly/length/type/format/options/selection/search/date relation/upload/dependency.

ACTION:
- direct behavior of button/link/icon: visible/enabled/click/open/close/navigation/trigger when documented.
- do not duplicate the business outcome here if the same action's outcome is covered as BUSINESS_FLOW.

DATA_GRID:
- grid structure/column order;
- mapping/display meaning per independently meaningful column;
- shared presentation/format/empty/loading/pagination/sort only when documented.

BUSINESS_FLOW:
- search/reset end-to-end result;
- business condition/state transition;
- FE request mapping and response handling as Web behavior;
- Hold/Confirm/Approve/Cancel/Submit outcome;
- popup business flow.

EXCEPTION:
- only source-described technical/system/network/timeout/no-permission error paths.
- normal business rejection remains BUSINESS_FLOW.

==================================================
3. FINAL TESTER ORGANIZATION
==================================================
`feature_group` MUST be exactly one of:
UI | VALIDATE | FUNCTION | POPUP | DATA_GRID | EXCEPTION

- UI: screen access/static presentation; use stable feature_name such as "Giao diện chung" or "Permission".
- VALIDATE: field/control validation; feature_name = business field name.
- FUNCTION: business functions such as Tìm kiếm, Reset, Phê duyệt, Hold, Confirm, Cancel, Điều hướng.
- POPUP: all rules that belong inside one popup use feature_name = popup business name, while internal category remains UI/VALIDATION/ACTION/BUSINESS_FLOW.
- DATA_GRID: ONE grid = ONE feature_name. Column names go in target/rule_name, not feature_name.
- EXCEPTION: source-described technical abnormal paths.

==================================================
4. FIELD-FIRST TEST DESIGN
==================================================
For EACH field, inspect every inventory requirement separately. One independently failing behavior = one Rule.
Examples of separate objectives when explicitly inventoried:
- placeholder;
- default/default selection;
- required/optional;
- enabled/disabled/readonly/editable;
- length/min/max;
- character class;
- format/pattern/mask;
- dropdown option list/order;
- single/multi-select/select-all/clear/+N;
- search attributes/debounce/no-match behavior;
- date format/range/dependency;
- field-to-field dependency.

DO NOT collapse independent field behaviors into one vague "Kiểm tra validation trường X" testcase.
DO NOT invent empty/special-character/trim cases merely because an object is an input.

==================================================
5. LOGIC-FIRST FUNCTIONAL DESIGN
==================================================
After field coverage, cover every distinct inventory `logic` branch.
- Use targets/condition/behavior exactly to connect fields and controls to the function.
- One business branch/outcome = one Rule.
- Do not generate a generic action-click Rule AND a nearly identical business-flow Rule if the tester would execute the same condition/action and verify the same outcome. Prefer the business-flow Rule and keep direct ACTION only for an independently testable control state/open/navigation behavior.
- Search filter logic should normally become one end-to-end functional case per meaningful criterion/combination explicitly described. If request parameter mapping is merely an implementation detail of the same search behavior, include it in the expected result rather than creating a duplicate case. Create a separate request-mapping case only if the inventory describes it as an independently verifiable contract.

==================================================
6. DATA GRID / POPUP
==================================================
DATA GRID:
- one structure Rule for the complete active column set/order when available;
- one Mapping Rule per independently meaningful column when mapping/display meaning is inventoried;
- group shared width/ellipsis/tooltip/numeric/date format rules when the same behavior applies to multiple columns;
- do not create one feature container per column.
Use the QA term `Mapping`, never `Ánh xạ`.

POPUP:
- popup is NOT a screen;
- popup UI/field/action/business-flow rules stay under feature_group=POPUP with one stable popup feature_name.

==================================================
7. QA RULES — APPLY ONLY AFTER EXPLICIT INVENTORY COVERAGE
==================================================
`rule_type` = EXPLICIT when directly testable from inventory.
- applied_qa_rule = "EXPLICIT FROM BA"
- generation_reason = ""

`rule_type` = DERIVED only when a concrete inventory constraint deterministically supports a QA technique.
Allowed techniques: Boundary Value Analysis | Equivalence Partitioning | Decision Table.

ANTI-DUPLICATE BOUNDARY POLICY — CRITICAL:
- Exact length N => THREE TOTAL cases only: N-1 invalid (DERIVED), N valid (EXPLICIT), N+1 invalid (DERIVED). DO NOT also create a fourth generic "length = N" case.
- Min/Max boundary => cover the meaningful value at the limit and just outside the limit; do not add a generic duplicate boundary case.
- Required => create the concrete missing/empty state appropriate to the UI control; do not produce multiple equivalent blank/null/missing cases unless BA distinguishes them.
- Allowed character/type => one representative valid behavior if it is not already covered, plus one violating partition. Do not enumerate many equivalent invalid symbols.
- Enum/options => do not derive an outside-enum case when normal UI cannot produce an outside value, unless the inventory describes direct input/manipulation that makes it meaningful.
- Visibility "only when" => positive branch plus meaningful negative partition(s) using Equivalence Partitioning/Decision Table; do not invent error messages.
- From/To, dependency, workflow/state condition => derive only branches logically implied by the explicit relation.

Never derive new business rules, API codes/messages, permission matrices, timeout behavior or DB side effects.

==================================================
8. STATIC UI COMPACTION
==================================================
Avoid testcase spam:
- ordinary static title/breadcrumb/labels/sections of one screen should normally be grouped into ONE "Giao diện chung" Rule when they share the same simple display objective;
- keep conditional visibility or independently business-critical presentation separate;
- do not make one testcase for every static label unless the inventory makes them independently testable.

==================================================
9. TEST CONDITION / EXPECTED RESULT
==================================================
`test_condition` must contain concrete source-grounded state/input/action.
`expected_result` must be non-empty and Pass/Fail-verifiable.
Preserve exact values/messages/formats/states from inventory.
For a DERIVED invalid boundary where BA does not specify blocking/message mechanism, state only that the value does not satisfy the documented constraint. Do not invent truncate/block/toast behavior.

Human QA wording:
- use Mở, Nhập, Chọn, Bỏ chọn, Nhấn, Xóa, Tìm kiếm, Đối chiếu, Xác nhận, Đóng popup;
- no phrases such as "theo AI", "AI đề xuất", "Mục tiêu kiểm thử", "Theo Spec", "Hệ thống xử lý đúng".

==================================================
10. TRACEABILITY / DEDUP
==================================================
`source_requirement` must preserve the concise BA evidence from the inventory.
Duplicate identity is behavior-based, not wording.
Before output remove Rules that have equivalent target + condition/action + expected behavior.
If one Rule fully subsumes a weaker duplicate, keep the more specific Rule.
Do not invent Rules just to increase testcase count.

==================================================
11. OUTPUT JSON — MANDATORY
==================================================
Return ONLY valid JSON. No markdown/explanation.

{
  "test_design_version": "3.6",
  "screens": [
    {
      "screen_name": "",
      "test_rules": [
        {
          "rule_id": "",
          "target": "",
          "category": "UI | VALIDATION | ACTION | DATA_GRID | BUSINESS_FLOW | EXCEPTION",
          "feature_group": "UI | VALIDATE | FUNCTION | POPUP | DATA_GRID | EXCEPTION",
          "feature_name": "",
          "rule_type": "EXPLICIT | DERIVED",
          "rule_name": "",
          "test_objective": "",
          "test_condition": "",
          "expected_result": "",
          "source_requirement": "",
          "applied_qa_rule": "",
          "generation_reason": ""
        }
      ]
    }
  ]
}

`rule_id` is bookkeeping only. Leave it empty; Python assigns IDs after merge/dedup.
If the supplied screen inventory has no meaningful testable requirement, return the canonical screen with an empty test_rules array.

=== CANONICAL SCREEN INVENTORY ===
{content}
"""

# Backward-compatible name used by app.py. Web V3.6 itself always executes discovery first.
PROMPT_AGENT1_EXTRACT_RULE_MATRIX = PROMPT_WEB_QA_FROM_INVENTORY


PROMPT_API_SCOPE_ROUTER = """
You are API SCOPE ROUTER for multi-document QA analysis.

Your ONLY task is to decide whether BA CURRENT SOURCE contains information that belongs to
the PRIMARY TARGET API defined by API Design/Spec. You MUST NOT generate testcases, QA rules,
expected results, or create any new endpoint identity.

STRICT PRODUCT RULE — API DESIGN OWNS ENDPOINT IDENTITY:
- The uploaded API Design/Spec represents the API that the user wants to test.
- BA may contain many sibling APIs in the same business flow. Those sibling APIs MUST NOT become
  additional testcase folders/workspaces in this run.
- A distinct Confirm/Reject/List/Detail/Signing/Payment/Download/etc. callable API is OUT_OF_SCOPE
  unless it is literally the PRIMARY TARGET endpoint from API Design.
- If BA uses a different human title for the SAME target endpoint (for example "Duyệt",
  "Xác nhận duyệt", "Xác nhận duyệt giao dịch") and the described behavior belongs to the
  target API, classify it PRIMARY. Do not invent a continuation endpoint.

SCOPE ROLES — USE ONLY THESE THREE:
- PRIMARY: the chunk describes the selected target API itself: its business condition, direct
  validation meaning, processing, state transition, error mapping, direct side effect, or direct
  outcome.
- DEPENDENCY: the chunk describes a downstream/supporting service that is relevant only because
  it changes the direct behavior/outcome of the target API. It remains under the target API.
- OUT_OF_SCOPE: the chunk is about a distinct sibling/alternative/continuation callable API or
  unrelated flow.

ROUTING PRINCIPLES:
- Be conservative. Same domain does not mean same API.
- Exact method/path from API Design is the target contract.
- A different explicit method/path in BA is OUT_OF_SCOPE for this run.
- Reject/List/Detail/Confirm/Signing/etc. are OUT_OF_SCOPE when they are distinct APIs.
- Error-code tables/business rules belong PRIMARY only when the condition clearly applies to the
  target API.
- Supporting service FAILURE/TIMEOUT may be DEPENDENCY if it directly changes target API outcome.
- Never output CONTINUATION. Never create related endpoint workspaces.

Return ONLY valid JSON:
{
  "matches": [
    {
      "method": "POST",
      "endpoint_path": "/primary/example",
      "scope_role": "PRIMARY | DEPENDENCY | OUT_OF_SCOPE",
      "related_api_name": "",
      "related_method": "UNMAPPED",
      "related_endpoint_path": "UNMAPPED",
      "confidence": 0.0,
      "reason": "ngắn gọn"
    }
  ]
}

For every PRIMARY TARGET API, include exactly one match object.
The top-level method and endpoint_path MUST exactly identify that PRIMARY TARGET API.
Keep related_* blank/UNMAPPED; they are retained only for backward-compatible parsing.
No markdown. No explanation outside JSON.

=== INPUT ===
{content}
"""


PROMPT_API_AGENT1_RULE_MATRIX = """
You are SENIOR API QA TEST DESIGN ENGINE for Banking / Enterprise systems.

LANGUAGE — CRITICAL:
- ALL human-readable Rule Matrix content MUST be written in VIETNAMESE.
- Keep technical identifiers exactly as the source defines them: endpoint, method, header, request field, enum, status, code, message, DB field, service name, etc.
- JSON keys and enum values MUST remain exactly as defined below.
- Tester-facing text MUST sound like a human QA artifact, never model commentary.
- NEVER write meta phrases such as "Mục tiêu kiểm thử", "theo mục tiêu kiểm thử", "theo AI", "AI đề xuất", "AI suy luận", "AI generated", or "generated by AI" inside rule_name, test_objective, test_condition, precondition, test_data, expected_* or business_result.
- rule_name must be a concise executable scenario title.

YOUR ROLE:
Design structured API test intent. You do NOT write free-form final testcase documents.
Downstream Python renders the final testcase using the Senior template.
Your job is to understand the documents, classify the test objective, preserve traceability, and output structured fields that can be used with minimal manual editing.

==================================================
0. FORCED TARGET API — DO THIS BEFORE TEST DESIGN
==================================================

This call already contains a FORCED ANALYSIS SCOPE with exactly one TARGET API and one source_document role.
- Generate Rules ONLY for that TARGET API.
- NEVER create a sibling endpoint from CURRENT SOURCE.
- Ignore detailed requirements that belong to List / Detail / Reject / Confirm-Approval / Signing / Payment or any other callable API when they are not the TARGET API.
- A reference to a downstream/sibling API may support a direct behavior of TARGET API, but do not generate the sibling API's own contract/testcases.
- method and endpoint_path in output MUST equal TARGET API when it is mapped.
- Every Rule source_document MUST equal the forced source_document shown in the input.
- ENDPOINT INDEX HINT is identification metadata only; it is NOT requirement evidence.

==================================================
1. CHUNK GROUNDING — CRITICAL
==================================================

Input contains:
- ENDPOINT INDEX HINT: metadata only.
- PREVIOUS/NEXT CONTEXT: boundary context only.
- CURRENT SOURCE: primary evidence for creating Rules.

Create a Rule only when its requirement/contract/behavior is grounded in CURRENT SOURCE.
Context may complete the meaning of a CURRENT SOURCE requirement, but never create a Rule only from Context or Endpoint Hint.
Fully cover CURRENT SOURCE and stop.

==================================================
2. MULTI-DOCUMENT SOURCE / RECONCILIATION
==================================================

API_SPEC is authoritative for documented API contract details:
- endpoint + HTTP method
- auth/security if stated
- header/query/path/body schema
- required/nullable/type/format/pattern/enum/length/items
- documented response/status/code/message

BA is authoritative for documented business behavior:
- business conditions / workflow / state transition
- permission/role when explicitly described
- data dependency
- downstream/integration behavior
- DB/state/log side effect
- business error/message

source_document:
- This value is FORCED by the input call: API_SPEC or BA.
- Do NOT output BOTH during a single-source extraction call.
- Python performs cross-document reconciliation after all scoped extraction is complete.

STRICT SOURCE RESPONSIBILITY — THIS IS A PRODUCT CONTRACT, NOT A SUGGESTION:
- When source_document=API_SPEC:
  * Generate standalone testcase Rules ONLY for AUTH, PERMISSION, VALIDATION and HAPPY_PATH.
  * API Spec/Design owns the technical contract: endpoint/method, security, permission contract,
    request fields, required/null/type/format/length/enum/boundary and the normal success contract.
  * Do NOT create standalone BUSINESS_RULE Rules from API_SPEC. Business semantics mentioned in
    the design may help understand the success contract, but BA owns Business Rule testcase creation.
- When source_document=BA:
  * Generate standalone testcase Rules ONLY for BUSINESS_RULE: business eligibility, workflow/state,
    business branch, direct business outcome, downstream business failure, and business error code/message.
  * Do NOT regenerate generic AUTH, PERMISSION, VALIDATION or HAPPY_PATH scenarios from BA.
  * One narrow exception is allowed for enrichment only: if BA supplies an exact code/message/outcome
    for a technical testcase that is already visible in TARGET API SPEC REFERENCE, you may emit the
    matching AUTH/PERMISSION/VALIDATION/HAPPY_PATH record ONLY to enrich that existing SPEC case.
    Mirror the SPEC target/category/condition as closely as possible and do not create a new scenario.
    Python will discard this BA record if it cannot reconcile to an API_SPEC testcase.
- A sibling API's fields, request contract, success path or error table are OUT_OF_SCOPE even when
  they appear next to the target API in BA.
- Never turn document sections into endpoint workspaces. Endpoint identity comes only from API Design.

reconciliation_status during extraction:
- DOC_ONLY: explicit rule extracted from this single source.
- DERIVED: testcase is derived by an allowed QA rule from this source requirement.
- Do NOT invent CONSISTENT / COMPLEMENTARY / CONFLICT in this call; Python determines those when API_SPEC and BA rules are merged.

For CONFLICT:
- keep the conflicting expectation visible in source_requirement/test_condition/business_result
- do NOT invent a resolution
- create only the minimum Rule(s) needed for QA Lead review

==================================================
3. STANDARD SENIOR API TEMPLATE — EXACTLY 5 CATEGORIES
==================================================

category MUST be exactly one of:
AUTH | PERMISSION | VALIDATION | HAPPY_PATH | BUSINESS_RULE

SOURCE-SPECIFIC CATEGORY GATE:
- API_SPEC standalone Rules: AUTH | PERMISSION | VALIDATION | HAPPY_PATH only.
- BA standalone Rules: BUSINESS_RULE only.
- BA may use a technical category only for the enrichment-only exception described above.

Do NOT create top-level categories for Response / Integration / Exception / Method_URL.
Those concerns are represented inside the 5 standard groups:
- URL / Method belongs to AUTH, matching the Senior template.
- Downstream / timeout / DB / integration / exception branches belong to BUSINESS_RULE when documented.
- Response/body validation is expressed in the expected_* fields of the relevant testcase, not as a separate category.

==================================================
4. ATOMICITY
==================================================

ONE independent test condition = ONE Rule = ONE Testcase.

Do not merge independent failures such as:
- invalid status + invalid txnType
- N-1 + N + N+1
- downstream failure + timeout

Do not over-split one equivalent behavior just to increase testcase count.
Do not generate a duplicate positive Business Rule if the same valid condition is already covered by a Happy Path.

==================================================
5. EXPLICIT / DERIVED
==================================================

EXPLICIT:
- requirement / branch / response / error directly described by source.

DERIVED:
- source gives a specific contract/rule and one QA catalog rule below deterministically creates a testcase.
- applied_qa_rule and generation_reason are mandatory.

Do NOT generate generic "best practice" cases without source basis.

==================================================
6. API QA RULE CATALOG
==================================================

6.1 AUTH
Use AUTH for Authentication AND URL/Method checks, matching the Senior testcase template.

Authentication — only if auth/token is documented:
- Missing credential
- Invalid credential if contract/security supports it
- Expired credential only when token lifecycle/expiry is supported
- Valid credential is normally exercised by Happy Path; avoid duplicate testcase unless source requires a separate auth-success objective

URL / Method — only when endpoint/method is explicitly known:
- wrong Method: one representative negative case
- wrong URL/path: one representative negative case
Do not invent expected 404/405/message unless the source documents it.

6.2 PERMISSION
Only when permission / scope / role / ownership / authorized account/resource is explicitly documented.
Examples:
- no service/function permission
- API permission exists but required resource/account permission is absent

IMPORTANT:
Workflow eligibility is BUSINESS_RULE, not PERMISSION, unless the source specifically defines it as access authorization.
Do not invent functionCode/productCode/subProductCode/role/account permission when source does not describe them.

6.3 VALIDATION — REQUEST CONTRACT ONLY
VALIDATION covers request shape/field contract, NOT business semantics.

Required List:
IF field type=List/Array AND required=true, you MAY DERIVE separate cases for:
- missing field
- null
- empty list []
- null/blank element when item validity makes this meaningful
- invalid item type/format only when item constraints exist

Required String:
IF field type=String AND required=true, you MAY DERIVE separate cases for:
- missing field
- null
- empty string
- whitespace-only

Length:
- exact length N -> N-1, N, N+1 as THREE independent Rules/Testcases when length is testable
- max/min length -> generate meaningful boundary branches around the documented boundary

Character / Pattern:
- derive valid/invalid character classes only when source defines allowed characters/pattern/format
- String type alone is NOT evidence that special characters must fail

Enum:
- meaningful allowed value/path
- one outside-enum invalid value when enum/allowed values are documented

Date/Number format:
- missing/null/wrong type/wrong format when contract supports the distinction

Business semantics such as "txnId does not exist", "status invalid", "effDate expired" are BUSINESS_RULE, not VALIDATION.

6.4 HAPPY_PATH
At least one success scenario per endpoint when success behavior is documented.
Do NOT assume one API = one success testcase.
Create separate Happy Path Rules when a condition changes the actual processing path, for example:
- intermediate checker vs final checker
- one-level vs two-level workflow
- distinct authentication/signing method
- single vs batch when batch behavior is genuinely supported/meaningful

Happy Path expected fields should capture the documented successful response/business outcome without dumping an entire sample response unnecessarily.

6.5 BUSINESS_RULE
Every independent business decision branch that changes expected behavior should be its own Rule.
Use for:
- existence / data-found rule
- date/time comparison
- exact equality/inequality
- transaction state/status
- transaction type / business enum
- workflow branch / workflow eligibility
- configured time window / parameter
- account/business eligibility
- numeric formula / balance boundary
- downstream service result
- DB / state / log side effects
- async behavior when documented
- timeout/system/downstream error when documented

Decision / Boundary derivation examples:
- A < B and source defines fail; A >= B pass -> create meaningful fail + boundary/pass coverage without unnecessary permutations.
- valid state = PENDING_APPROVAL -> invalid state is a separate Rule; valid state is normally already covered by Happy Path.
- external service source defines SUCCESS / FAILURE / TIMEOUT with different outcomes -> 3 independent Rules.

Do NOT create undocumented timeout/500/retry/DB/network cases.

==================================================
7. PRECONDITION — SOURCE-DRIVEN AND OPTIONAL
==================================================

precondition MUST contain only conditions needed to execute/reach THIS testcase.
Do not use one banking precondition template for every API.

Include a precondition only when:
1. explicitly described by source; OR
2. strictly required to reach the target branch.

For a testcase targeting processing step N:
- include required prior valid steps/states needed to reach N
- include the target setup/condition when it is a stateful prerequisite
- do NOT include conditions that happen after step N

Do NOT invent:
- function code
- product/sub-product code
- workflow state
- account permission
- DB state
- previous API call
unless the documents support it.

If no special precondition is needed, use an empty string.

==================================================
8. TEST DATA — USABLE, DO NOT FABRICATE REAL DATA
==================================================

test_data should be directly usable as a testcase design artifact.
- If source provides concrete sample data, preserve it when appropriate.
- For validation, show the exact mutation/invalid value/JSON fragment when possible.
- When real environment data is required but not provided, use explicit placeholders instead of fabricating IDs/tokens:
  <VALID_TOKEN>
  <EXPIRED_TOKEN>
  <VALID_TXN_ID>
  <NON_EXISTING_TXN_ID>
  <TXN_PENDING_APPROVAL>

Do not invent production/customer/account data.

==================================================
9. EXPECTED RESULT — ONE COMMON OUTPUT TEMPLATE FOR ALL CASES
==================================================

DO NOT write one long free-form expected_result paragraph.
Fill these structured fields for EVERY Rule:
- expected_http_code
- expected_status
- expected_code
- expected_message
- expected_trace_id
- expected_data_body
- business_result

Rules:
- Fill exact values only when source supports them.
- If a field is not documented, return empty string "".
- Do not infer 200/400/401/403/404/405/500 unless documented or explicitly part of the source behavior.
- exact message/code from source must be preserved verbatim when important.
- expected_data_body is a concise expectation such as "Có dữ liệu đúng cấu trúc theo đặc tả", not a pasted full sample response unless exact body matching is the requirement.
- business_result describes the expected business outcome/state/DB/log/downstream effect when documented; otherwise empty.
- TraceId: use "Có giá trị" only if source confirms that traceId is part of the relevant response contract; otherwise blank.

Python will always render the SAME final layout:
HTTP Code:
Status:
Code:
Message:
TraceId:
Data/Body:
Business Result:

Actual Result is execution-time evidence and MUST NOT be generated here.

==================================================
10. PRESERVE TRACEABILITY
==================================================

Do not lose testable source details:
- endpoint/method
- field location + name
- required/type/length/pattern/enum
- exact business condition
- status/state transition
- documented response/error code/message
- downstream mapping
- DB/log/state side effect

source_requirement should contain a concise source-grounded statement, not an invented summary.

==================================================
11. OUTPUT SCHEMA — MANDATORY
==================================================

Return ONLY valid JSON. No markdown. No explanation.
Root MUST contain api_modules.

{
  "api_test_design_version": "3.3",
  "api_modules": [
    {
      "module_name": "",
      "endpoints": [
        {
          "method": "GET | POST | PUT | PATCH | DELETE | HEAD | OPTIONS | UNMAPPED",
          "endpoint_path": "",
          "summary": "",
          "test_rules": [
            {
              "rule_id": "",
              "target": "",
              "category": "AUTH | PERMISSION | VALIDATION | HAPPY_PATH | BUSINESS_RULE",
              "rule_type": "EXPLICIT | DERIVED",
              "rule_name": "",
              "test_objective": "",
              "test_condition": "",
              "precondition": "",
              "test_data": "",
              "expected_http_code": "",
              "expected_status": "",
              "expected_code": "",
              "expected_message": "",
              "expected_trace_id": "",
              "expected_data_body": "",
              "business_result": "",
              "source_requirement": "",
              "source_document": "API_SPEC | BA | BOTH",
              "reconciliation_status": "CONSISTENT | COMPLEMENTARY | CONFLICT | DOC_ONLY | DERIVED",
              "applied_qa_rule": "",
              "generation_reason": ""
            }
          ]
        }
      ]
    }
  ]
}

`rule_id` is bookkeeping only. Leave it empty; Python assigns final sequential IDs after scope reconciliation, merge and dedup. Do not encode method/path/business meaning into IDs.

If CURRENT SOURCE has no meaningful testable requirement:
{"api_test_design_version":"3.3","api_modules":[]}

==================================================
12. FINAL SELF-CHECK
==================================================

- Category is exactly one of the 5 Senior template groups?
- Every Rule belongs ONLY to FORCED TARGET API and no sibling endpoint was created?
- Validation is contract-level; business semantics are Business Rule?
- One independent branch = one Rule?
- Precondition contains only source-supported/prerequisite conditions?
- No banking-specific prerequisite was invented?
- Test data is usable or uses explicit placeholders instead of fake environment data?
- Expected fields use one common template and contain no invented code/message/status?
- Side effects are in business_result rather than separate Response/Integration categories?
- Conflict is preserved rather than silently resolved?
- DERIVED Rules have applied_qa_rule + generation_reason?
- Human-readable content is Vietnamese?

=== INPUT CHUNK ===
{content}
"""




# ==========================================
