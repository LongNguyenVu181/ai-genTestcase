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
# 0. CẤU HÌNH WEB PIPELINE V3.5 — SENIOR-HYBRID / VALIDATION-EXPANDED / SCALABLE LONG DOCUMENT
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

# Qwen output budget. Nếu vẫn chạm length, pipeline sẽ tự chia nhỏ và retry.
QWEN_MAX_OUTPUT_TOKENS = 32768

# API pipeline: targeted scope scan + deep analysis + recursive split.
API_AGENT1_CHUNK_TARGET_CHARS = 7200
API_AGENT1_CHUNK_MAX_CHARS = 9000
API_AGENT1_CONTEXT_CHARS = 900
# V1.9.3: BA doc is scanned cheaply; PRIMARY/CONTINUATION/DEPENDENCY are routed before deep analysis.
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
WEB_TEST_DESIGN_VERSION = "3.5"
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
WEB_AGENT1_CACHE_NAMESPACE = "web-agent1-v3.5-human-template-r1"
API_AGENT1_CACHE_NAMESPACE = "api-agent1-v3.8-human-output"

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
    rule_ids = set()

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

                rule_id = str(rule.get("rule_id", "")).strip()
                if not rule_id:
                    return False, f"Rule {sidx + 1}.{ridx + 1} rule_id rỗng"
                if rule_id in rule_ids:
                    return False, f"Duplicate rule_id: {rule_id}"
                rule_ids.add(rule_id)

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
            item["applied_qa_rule"] = "EXPLICIT FROM SPEC"
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
                rule["applied_qa_rule"] = "EXPLICIT FROM SPEC"
                repairs.append({"rule_id": str(rule.get("rule_id", "")), "field": "applied_qa_rule", "value": "EXPLICIT FROM SPEC"})
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
    if any(x in compact for x in ("xacnhanduyet", "confirmapproval", "confirmapprove")):
        return "CONFIRM_APPROVE"
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
    """Deterministic post-router guard for obvious sibling headings.

    The LLM remains responsible for semantic routing inside ambiguous sections. This guard
    only overrides cases where the BA section title itself clearly names another operation.
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

    if target_family == "APPROVE" and heading_family == "CONFIRM_APPROVE":
        new_role = "CONTINUATION"
        if not str(guarded.get("related_api_name") or "").strip():
            guarded["related_api_name"] = heading.strip() or "API Xác nhận duyệt"
        reason = "heading confirms required approval continuation"
    elif target_family != "UNKNOWN" and heading_family != "UNKNOWN" and heading_family != target_family:
        # Known sibling operation. CONTINUATION exception above is the only automatic bridge.
        new_role = "OUT_OF_SCOPE"
        reason = f"heading family {heading_family} differs from target {target_family}"

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
        r"===\s*source document|api\s+design|api\s+spec|ba\s+business|"
        r"get\s+/|post\s+/|put\s+/|patch\s+/|delete\s+/|head\s+/|options\s+/)",
        s,
        flags=re.IGNORECASE,
    ):
        return True
    return False


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
    """Split a document into heading/table/paragraph blocks without exceeding caller max size."""
    lines = raw_text.splitlines()
    blocks = []
    paragraph = []
    current_title = base_title

    def flush_paragraph():
        nonlocal paragraph
        if paragraph:
            block_text = "\n".join(paragraph).strip()
            if block_text:
                blocks.append({"title": current_title, "text": block_text})
            paragraph = []

    for line in lines:
        stripped = line.strip()
        if not stripped:
            flush_paragraph()
            continue

        if _is_heading_line(line):
            flush_paragraph()
            current_title = re.sub(r"^[#\s]+", "", stripped).strip() or current_title
            blocks.append({"title": current_title, "text": line.rstrip()})
            continue

        # Markdown table: keep each row intact so a requirement row is not cut mid-row.
        if stripped.startswith("|"):
            flush_paragraph()
            blocks.append({"title": current_title, "text": line.rstrip()})
            continue

        paragraph.append(line.rstrip())

    flush_paragraph()

    if not blocks and raw_text.strip():
        blocks = [{"title": base_title, "text": raw_text.strip()}]

    expanded = []
    for block in blocks:
        if len(block["text"]) <= max_block_chars:
            expanded.append(block)
        else:
            for part in _safe_split_large_block(block["text"], max_block_chars):
                expanded.append({"title": block["title"], "text": part})
    return expanded


def split_document_semantic(
    raw_text: str,
    base_title: str,
    target_chars: int = AGENT1_CHUNK_TARGET_CHARS,
    max_chars: int = AGENT1_CHUNK_MAX_CHARS,
    context_chars: int = AGENT1_CONTEXT_CHARS,
) -> list[dict]:
    """Pack semantic blocks thành chunk động.

    Mỗi chunk có CURRENT SOURCE riêng + context trước/sau để giữ dependency ở biên.
    Context chỉ dùng để hiểu, prompt cấm sinh rule chỉ từ context.
    """
    blocks = _document_to_semantic_blocks(raw_text, base_title, max_chars)
    core_chunks = []
    current_parts = []
    current_len = 0
    current_title = base_title

    def flush_current():
        nonlocal current_parts, current_len, current_title
        if not current_parts:
            return
        core = "\n\n".join(current_parts).strip()
        if core:
            core_chunks.append({
                "title": current_title,
                "core_text": core,
                "char_count": len(core),
            })
        current_parts = []
        current_len = 0

    for block in blocks:
        block_text = block["text"].strip()
        if not block_text:
            continue
        block_len = len(block_text) + (2 if current_parts else 0)

        if not current_parts:
            current_title = block.get("title") or base_title
            current_parts = [block_text]
            current_len = len(block_text)
            continue

        proposed = current_len + block_len
        # target là điểm flush ưu tiên; max là hard ceiling.
        if proposed > max_chars or (current_len >= target_chars and proposed > target_chars):
            flush_current()
            current_title = block.get("title") or base_title
            current_parts = [block_text]
            current_len = len(block_text)
        else:
            current_parts.append(block_text)
            current_len = proposed
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

=== PREVIOUS CONTEXT — CHỈ DÙNG ĐỂ HIỂU, KHÔNG SINH RULE CHỈ TỪ ĐOẠN NÀY ===
{chunk.get('prev_context', '')}

=== CURRENT SOURCE — PHẠM VI CHÍNH ĐƯỢC PHÉP SINH TEST RULE ===
{chunk.get('core_text', '')}

=== NEXT CONTEXT — CHỈ DÙNG ĐỂ HIỂU, KHÔNG SINH RULE CHỈ TỪ ĐOẠN NÀY ===
{chunk.get('next_context', '')}
"""


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

    used_prefixes = set()
    for screen_idx, screen in enumerate(final_screens, start=1):
        fallback = f"SCREEN_{screen_idx:02d}"
        prefix = _infer_rule_prefix(screen.get("test_rules", []), fallback)

        base_prefix = prefix
        suffix = 2
        while prefix in used_prefixes:
            prefix = f"{base_prefix}_{suffix}"
            suffix += 1
        used_prefixes.add(prefix)

        for idx, rule in enumerate(screen.get("test_rules", []), start=1):
            rule["rule_id"] = f"{prefix}-{idx:03d}"

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
    prompt_template: str,
    cache: dict | None = None,
    progress_callback=None,
) -> tuple[bool, dict | None, dict]:
    """Run the scalable Web document-analysis pipeline."""
    started = time.time()
    initial_chunks = split_document_semantic(raw_text, base_filename)
    diagnostics: list[dict] = []
    all_matrices: list[dict] = []

    workers = min(AGENT1_PARALLEL_WORKERS, max(1, len(initial_chunks)))
    log_info(
        f"[AGENT1_PIPELINE] 📚 DocumentChars={len(raw_text):,} | "
        f"InitialChunks={len(initial_chunks)} | Target≈{AGENT1_CHUNK_TARGET_CHARS:,} | "
        f"Workers={workers}"
    )

    # Small documents keep the original sequential path.
    if workers <= 1 or len(initial_chunks) <= 1:
        for idx, chunk in enumerate(initial_chunks, start=1):
            if progress_callback:
                progress_callback(
                    idx - 1,
                    len(initial_chunks),
                    f"Agent 1 đang xử lý chunk {idx}/{len(initial_chunks)} — {chunk.get('title')} ({chunk.get('char_count', 0):,} chars)"
                )

            ok, matrices = process_agent1_chunk_recursive(
                chunk=chunk,
                api_key=api_key,
                base_url=base_url,
                model=model,
                prompt_template=prompt_template,
                cache=cache,
                diagnostics=diagnostics,
            )
            if not ok:
                summary = {
                    "ok": False,
                    "document_chars": len(raw_text),
                    "initial_chunks": len(initial_chunks),
                    "parallel_workers": workers,
                    "elapsed": round(time.time() - started, 2),
                    "diagnostics": diagnostics,
                }
                return False, None, summary

            all_matrices.extend(matrices)
            if progress_callback:
                progress_callback(idx, len(initial_chunks), f"Hoàn tất chunk {idx}/{len(initial_chunks)}")
    else:
        # Each worker uses its own diagnostics list so output can be merged
        # deterministically in original chunk order.
        def _work(index: int, chunk: dict):
            local_diagnostics: list[dict] = []
            ok, matrices = process_agent1_chunk_recursive(
                chunk=chunk,
                api_key=api_key,
                base_url=base_url,
                model=model,
                prompt_template=prompt_template,
                cache=cache,
                diagnostics=local_diagnostics,
            )
            return index, ok, matrices, local_diagnostics

        results: dict[int, tuple[bool, list[dict], list[dict]]] = {}
        completed = 0
        with ThreadPoolExecutor(max_workers=workers, thread_name_prefix="testpilot-web") as executor:
            futures = {
                executor.submit(_work, idx, chunk): idx
                for idx, chunk in enumerate(initial_chunks)
            }
            for future in as_completed(futures):
                idx, ok, matrices, local_diagnostics = future.result()
                results[idx] = (ok, matrices, local_diagnostics)
                completed += 1
                if progress_callback:
                    progress_callback(
                        completed,
                        len(initial_chunks),
                        f"Hoàn tất {completed}/{len(initial_chunks)} chunk"
                    )

        for idx in range(len(initial_chunks)):
            ok, matrices, local_diagnostics = results[idx]
            diagnostics.extend(local_diagnostics)
            if not ok:
                summary = {
                    "ok": False,
                    "document_chars": len(raw_text),
                    "initial_chunks": len(initial_chunks),
                    "parallel_workers": workers,
                    "elapsed": round(time.time() - started, 2),
                    "diagnostics": diagnostics,
                }
                return False, None, summary
            all_matrices.extend(matrices)

    final_matrix, merge_stats = merge_rule_matrices(all_matrices)
    final_matrix = normalize_web_rule_shape(final_matrix)
    final_matrix = normalize_web_rule_matrix_enums(final_matrix)
    final_matrix, final_semantic_repairs = normalize_web_rule_semantics(final_matrix)
    merge_stats["semantic_metadata_repairs"] = len(final_semantic_repairs)
    schema_ok, schema_diag = validate_rule_matrix_schema(final_matrix, strict=True)
    elapsed = time.time() - started

    summary = {
        "ok": schema_ok,
        "document_chars": len(raw_text),
        "initial_chunks": len(initial_chunks),
        "parallel_workers": workers,
        "leaf_matrices": len(all_matrices),
        "elapsed": round(elapsed, 2),
        "merge": merge_stats,
        "schema": schema_diag,
        "diagnostics": diagnostics,
    }

    if not schema_ok:
        log_error(f"[AGENT1_PIPELINE] ❌ Final schema fail | {schema_diag}")
        return False, None, summary

    log_info(
        f"[AGENT1_PIPELINE] ✅ Completed | Time={elapsed:.2f}s | "
        f"Screens={merge_stats['screens']} | FinalRules={merge_stats['final_rules']} | "
        f"ExactDupRemoved={merge_stats['exact_duplicates_removed']} | Workers={workers}"
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
    log_info(
        f"[API_TARGET_SELECTOR] Design có {len(targets)} endpoint nhưng không đủ tín hiệu chọn 1; "
        "giữ tất cả để tránh đoán sai."
    )
    return targets


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

    V1.9.3 keeps PRIMARY and required CONTINUATION operations, preserves supporting
    DEPENDENCY context under the primary API, and skips OUT_OF_SCOPE sibling flows.
    """
    payload = build_api_scope_router_payload(ba_chunk, targets)
    cache_key = _cache_key("api-scope-router-v1.10.0-gated", model, PROMPT_API_SCOPE_ROUTER, payload)
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
) -> str:
    hints = select_api_endpoint_hints(chunk.get("core_text", ""), endpoint_index)
    hint_text = json.dumps(hints, ensure_ascii=False, indent=2) if hints else "[]"
    target = target_endpoint or {}
    target_text = json.dumps({
        "method": target.get("method", "UNMAPPED"),
        "endpoint_path": target.get("endpoint_path") or target.get("path") or "UNMAPPED",
        "summary": target.get("summary", ""),
    }, ensure_ascii=False, indent=2)
    return f"""=== FORCED ANALYSIS SCOPE — MANDATORY ===
source_document: {source_document}
TARGET API:
{target_text}

Rules for this call:
- Generate rules ONLY for TARGET API / operation shown above.
- Ignore sibling APIs/operations that are NOT the TARGET. If TARGET itself is a continuation such as Confirm Approval, analyze that continuation normally.
- source_document of every generated rule MUST be exactly {source_document}.
- If TARGET has an exact method/path, do not infer another endpoint from nearby BA text.
- If TARGET endpoint_path is UNMAPPED, NEVER invent a URL. Use TARGET summary to identify the operation and keep method/path UNMAPPED unless the current source explicitly states them.

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
    seen_ids = set()
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
                    rid = rule.get("rule_id", "").strip()
                    if not rid:
                        return False, f"Rule {midx}.{eidx}.{ridx}.rule_id rỗng"
                    if rid in seen_ids:
                        return False, f"Trùng rule_id: {rid}"
                    seen_ids.add(rid)
                    if not rule.get("source_requirement", "").strip():
                        return False, f"Rule {rid} thiếu source_requirement"
                    if rule_type == "DERIVED" and not rule.get("applied_qa_rule", "").strip():
                        return False, f"Rule {rid} DERIVED thiếu applied_qa_rule"

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
    """Logical identity used to reconcile the same rule across multiple documents."""
    return (
        _normalize_text(rule.get("target")),
        _normalize_text(rule.get("category")),
        _normalize_text(rule.get("rule_name")),
        _normalize_text(rule.get("test_condition")),
        _normalize_text(rule.get("applied_qa_rule")),
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

    V1.9.3 intentionally ignores module-name drift between API Spec and BA. The same
    method+path is one logical API and must reconcile into one endpoint workspace.
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
                    if concept in concept_map:
                        existing_idx = concept_map[concept]
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
                        continue

                    concept_map[concept] = len(target_ep["test_rules"])
                    target_ep["test_rules"].append(rule)

    final_modules = []
    global_rule_counter = 0
    unmapped_endpoints = 0
    for e_idx, ep_data in enumerate(endpoints_map.values(), start=1):
        if ep_data.get("method") == "UNMAPPED" or ep_data.get("endpoint_path") == "UNMAPPED":
            unmapped_endpoints += 1
        method = ep_data.get("method", "API")
        slug = unicodedata.normalize("NFKD", ep_data.get("endpoint_path", "API")).encode("ascii", "ignore").decode("ascii")
        slug = re.sub(r"[^A-Za-z0-9]+", "_", slug).strip("_").upper()[-28:] or f"EP{e_idx:02d}"
        prefix = f"API_{method}_{slug}" if method != "UNMAPPED" else f"API_UNMAPPED_{e_idx:02d}"
        for local_idx, rule in enumerate(ep_data.get("test_rules", []), start=1):
            global_rule_counter += 1
            rule["rule_id"] = f"{prefix}-{local_idx:03d}"
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
) -> tuple[bool, list[dict]]:
    diagnostics = diagnostics if diagnostics is not None else []
    payload = build_api_agent1_chunk_payload(
        chunk, endpoint_index, source_document=source_document, target_endpoint=target_endpoint
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
            source_document=source_document, target_endpoint=target_endpoint
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
    """V1.9.3 targeted multi-document API pipeline with continuation-aware routing.

    API Design/Spec defines the target API(s). The BA document is scanned only for
    relevance first; deep QA generation runs only on BA chunks mapped to a target API.
    This prevents a 30+ page BA document from generating cases for unrelated APIs.
    """
    started = time.time()
    endpoint_index = extract_api_endpoint_index(design_text)
    targets = build_target_api_descriptors(design_text, endpoint_index)
    targets = select_primary_api_targets(design_text, targets, document_hint=base_filename)
    diagnostics: list[dict] = []
    matrices: list[dict] = []

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
            matrices.extend(mats)

    # Stage 2 — cheap BA scope scan. It reads chunks for routing only, no testcase generation.
    # V1.9.3 classifies each BA chunk as PRIMARY / CONTINUATION / DEPENDENCY / OUT_OF_SCOPE.
    ba_chunks = split_document_semantic(
        ba_text,
        base_title=f"{base_filename}::BA",
        target_chars=API_SCOPE_SCAN_CHUNK_TARGET_CHARS,
        max_chars=API_SCOPE_SCAN_CHUNK_MAX_CHARS,
        context_chars=450,
    )
    routed_pairs: list[tuple[dict, dict, str]] = []
    scope_diagnostics = []
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
            # Dependency behavior remains under the parent API as BUSINESS_RULE/expected outcome.
            # It does not become a new endpoint workspace by itself.
            routed_pairs.append((chunk, primary_target, role))
        elif role == "CONTINUATION":
            cont = _build_continuation_target(primary_target, guarded_match)
            ckey = (
                _normalize_text(cont.get("summary", "")),
                str(cont.get("method", "UNMAPPED")).upper(),
                _normalize_text(cont.get("endpoint_path", "UNMAPPED")),
                _api_endpoint_key(primary_target),
            )
            continuation_targets.setdefault(ckey, cont)
            routed_pairs.append((chunk, continuation_targets[ckey], role))
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

    # Deduplicate same BA chunk/analysis-target/role pair.
    # Unknown continuation endpoints are separated by BA operation name, never collapsed together.
    dedup = OrderedDict()
    for chunk, target, role in routed_pairs:
        target_identity = _api_endpoint_key(target)
        if target_identity[1] == "UNMAPPED":
            target_identity = (
                target_identity[0],
                "UNMAPPED::" + _normalize_text(target.get("summary", "API tiếp nối")),
            )
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
    # PRIMARY + DEPENDENCY are forced into the primary API.
    # CONTINUATION gets a separate API workspace; unknown method/path stay UNMAPPED rather than invented.
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
            cache, diagnostics, source_document="BA", target_endpoint=target,
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

PROMPT_AGENT1_EXTRACT_RULE_MATRIX = """
You are WEB_AGENT1_QA_BRAIN — a Senior QA Test Design Lead for Banking / Enterprise systems.

LANGUAGE REQUIREMENT — CRITICAL:
- ALL generated Rule Matrix textual content MUST be written in VIETNAMESE.
- Keep technical identifiers exactly as they appear in the source when needed: screen code, field name, API name, endpoint, parameter, status code, enum value, message, etc.
- JSON keys and enum values MUST remain exactly as defined by the schema below.
- Do NOT translate business labels, field names, messages, or source values if doing so would alter the original requirement.
- Tester-facing text MUST sound like a human QA artifact, not model commentary.
- NEVER write meta phrases such as "Mục tiêu kiểm thử", "theo mục tiêu kiểm thử", "theo AI", "AI đề xuất", "AI suy luận", "AI generated", or "generated by AI" inside rule_name, test_objective, test_condition, expected_result.
- rule_name must be a concise behavior/scenario title, not an explanation of how it was generated.

YOUR ONLY TASK:
Read the CURRENT SOURCE, identify source-grounded WEB requirements, and create a TEST DESIGN / RULE MATRIX.
You are the ONLY agent allowed to analyze requirements and apply QA Test Design Techniques.
Downstream Python code validates, maps, persists and exports your Rule Matrix deterministically. No second AI agent is allowed to add/remove/reason about Test Rules.

IMPORTANT DESIGN PRINCIPLE:
- INTERNAL QA CLASSIFICATION and FINAL TESTER ORGANIZATION are two different dimensions.
- `category` is the internal QA reasoning type: UI | VALIDATION | ACTION | DATA_GRID | BUSINESS_FLOW | EXCEPTION.
- `feature_group` + `feature_name` organize the final output in a Senior-QA-style feature-oriented structure.
- Do NOT force a rule into a different internal category merely to place it under a desired final section.

==================================================
1. CHUNK OWNERSHIP — CRITICAL
==================================================

The input contains:
1. PREVIOUS CONTEXT: only used to understand requirements near the chunk boundary.
2. CURRENT SOURCE: the PRIMARY source that owns Test Rules.
3. NEXT CONTEXT: only used to understand requirements near the chunk boundary.

ONLY create a Rule when the tested behavior is grounded in CURRENT SOURCE.
You MAY use PREVIOUS/NEXT CONTEXT to complete the meaning of a requirement that belongs to CURRENT SOURCE.
DO NOT create a Rule when all supporting evidence appears only in CONTEXT.

This is one chunk of a large document:
- Fully cover CURRENT SOURCE.
- Do not attempt whole-document coverage from one chunk.
- Do not repeat requirements owned by another chunk.
- Repeated OCR/table/image text must not create duplicate Rules.

==================================================
2. SCOPE — WEB FUNCTIONAL TESTING ONLY
==================================================

WEB scope includes source-grounded requirements for:
- screen/UI controls and static presentation;
- textbox/numeric/dropdown/date picker/checkbox/input behavior;
- button/icon/link/action;
- search/filter/reset;
- Data Grid/column/mapping/pagination;
- popup/dialog/toast;
- Web display/access permission when explicitly documented;
- business flow initiated from the Web UI;
- FE API request parameters and FE response handling as part of a Web feature.

CRITICAL:
- An API mentioned inside a Web SRS does NOT automatically become an API Test Scope.
- FE API calls for Dropdown/Search/Hold/Confirm/Cancel/Load data belong to WEB FEATURE testing.
- Do NOT create Authentication/HTTP Method/Header/Schema tests merely because an API is mentioned.
- Do NOT retrieve/infer content from URLs or “Request access” links that are not included in CURRENT SOURCE.
- Do NOT invent API status/message/DB/permission/timeout/edge-case behavior.
- Do NOT add generic QA checklist cases such as responsive/zoom/tab-order/cross-browser unless CURRENT SOURCE explicitly requires them.

DEFERRED FROM THE CURRENT WEB GENERATOR — DO NOT CREATE RULES FOR THESE IN THIS VERSION:
- close tab / close browser lifecycle;
- session-expiry lifecycle;
- technical/audit logging of user actions or API request-response;
- pagination Next/Previous navigation behavior inferred only from the presence of > / < controls. Keep only pagination behavior explicitly stated by source (for example page-size values/default, first/last-page visibility, disabled states, request parameters).
These may be handled by another pipeline later.

==================================================
3. SCREEN OWNERSHIP & ACTIVE REQUIREMENT PRECEDENCE
==================================================

Only create a screen when there is clear evidence such as a dedicated screen section, screen code, screen description table, or mockup/control specification.

SCREEN-NAME STABILITY:
- A code such as BOND_ORDER_LIST is a technical identity, NOT the displayed screen name.
- `screen_name` MUST be a human-readable business name.
- Different naming variants for the same business screen MUST remain ONE screen.
- Sub-headings such as “Mô tả màn hình”, “Logic tìm kiếm”, “Hold lại tiền” are NOT separate screens.
- A screen mentioned only as a navigation destination/reference MUST NOT become a new screen root.

ACTIVE REQUIREMENT PRECEDENCE:
- Later explicit revision/change note overrides older/general wording.
- Detailed active requirement overrides a generic summary when they conflict.
- Strikethrough/deleted/removed content is INACTIVE and MUST NOT create a Rule unless explicitly reintroduced later.
- Do NOT create a generic Action Rule claiming all buttons are always available when detailed rules define separate visibility/state conditions.
- Never invent a replacement behavior when resolving old vs new source wording.

==================================================
4. INTERNAL QA CATEGORY — EXACTLY ONE PER RULE
==================================================

`category` MUST be one of:
UI | VALIDATION | ACTION | DATA_GRID | BUSINESS_FLOW | EXCEPTION

4.1 UI
- Static display/presentation only: screen title, breadcrumb, static label/text, layout/section, visual icon, pure visibility, read-only display outside Data Grid, popup title/content/icon.
- Placeholder/default/dropdown options/input behavior are NOT UI; they are VALIDATION.
- Data Grid presentation/mapping is NOT UI; it is DATA_GRID.
- IMPORTANT GROUPING: all columns/row-mapping/pagination/sort/empty-state rules of the same grid MUST share ONE stable feature_name for that grid (for example "Data Grid" or "Danh sách kết quả"). Do NOT create feature_name="Cột ..." for each column. Put the column name in target/rule_name instead.

4.2 VALIDATION — FIELD / INPUT CONTROL TEST DESIGN
VALIDATION owns ALL source-grounded behavior of Field/Input Controls, including initial state, allowed input, selection behavior, control-local interaction, and field-level dependency.

CRITICAL PRINCIPLE:
- Do NOT create one vague Rule such as “Kiểm tra validation trường X” when Placeholder / Default / Length / Character / Selection / Search / Dependency can fail independently.
- EACH independently failing behavior MUST become a separate Rule. Exact-length N-1 / N / N+1 MUST be three independent Rules/Testcases.
- Only create a behavior when CURRENT SOURCE provides the control type, constraint, state, value, or relationship needed to support it.

A. INITIAL STATE / BASIC CONTROL STATE
For every Field/Input Control, inspect whether source explicitly defines:
- Placeholder;
- Default Value;
- Default Selection;
- default “Tất cả” / blank / empty state;
- Required / Optional;
- Enable / Disable;
- Readonly / Editable / Input-enabled;
- Visible / Hidden when this is a FIELD state rather than a business-action visibility rule;
- initial checked/unchecked/on/off state for Checkbox/Radio/Toggle/Switch.

B. TEXTBOX / TEXTAREA / GENERIC INPUT
When explicitly specified, inspect independently:
- exact Length;
- MinLength / MaxLength;
- allowed character class: numeric / alphabetic / alphanumeric / other explicitly described characters;
- disallowed character/type implied by an explicit allowed-type constraint;
- input Format / Pattern / Mask;
- case rule (upper/lower) ONLY when source specifies it;
- whitespace/trim/leading-zero behavior ONLY when source specifies it;
- multiline behavior / line count ONLY when source specifies it.

C. NUMERIC / AMOUNT / CURRENCY / PERCENTAGE INPUT
When source defines the constraint, inspect independently:
- minimum / maximum / exact value;
- zero / negative / positive allowance ONLY when source establishes the rule;
- integer vs decimal;
- decimal scale / precision / maximum decimal places;
- thousand separator / decimal separator / display-input format when applicable;
- unit/currency/percentage suffix or prefix when explicitly defined;
- rounding behavior ONLY when explicitly defined.
Do NOT invent financial rounding, currency scale, or negative-number rules from domain knowledge.

D. DROPDOWN / COMBOBOX / AUTOCOMPLETE / SINGLE-SELECT / MULTI-SELECT
When source defines them, inspect independently:
- option list/content;
- option ordering;
- option display structure/label format;
- Single Select capability;
- Multi Select capability;
- default selection;
- Select All;
- deselect/unselect Select All;
- Clear Selection / clear icon;
- selected-value display after choosing one value;
- selected-values display after choosing multiple values;
- overflow presentation such as “+N” ONLY when source specifies it;
- disabled/non-selectable option ONLY when source specifies it;
- maximum/minimum selection count ONLY when source specifies it;
- dependency/cascade on another field;
- loading source/API behavior belongs to BUSINESS_FLOW, but the options/selection behavior itself remains VALIDATION.

E. SEARCH INSIDE DROPDOWN / AUTOCOMPLETE
When explicitly described, inspect independently:
- search input availability;
- attributes used for matching (e.g. code/name);
- exact / contains / fuzzy behavior ONLY as described by source;
- debounce / delay / timing constraint;
- result display structure;
- no-match behavior ONLY when source specifies it;
- selected value after search;
- search reset/clear behavior ONLY when source specifies it.

F. DATE / DATE RANGE / TIME / DATETIME CONTROL
When explicitly specified, inspect independently:
- Placeholder / Default / Empty state;
- manual input vs picker selection capability;
- input/display format;
- From/To relationship;
- minimum / maximum allowed date/time;
- disabled/unavailable dates ONLY when source specifies them;
- automatic reorder/swap of From/To ONLY when source specifies it;
- clear/reset behavior;
- date/time dependency on another field;
- exact range presentation after selection.

G. CHECKBOX / RADIO / TOGGLE / SWITCH
When explicitly specified, inspect independently:
- default checked/unchecked/on/off state;
- selectable/toggleable state;
- mutually exclusive Radio behavior when source defines the Radio group;
- single vs multiple Checkbox selection rule when source defines it;
- enable/disable/readonly state;
- dependent field/state changed by the selection ONLY when source specifies the dependency.

H. FILE UPLOAD CONTROL — ONLY WHEN PRESENT IN CURRENT SOURCE
When source explicitly describes upload constraints, inspect independently:
- allowed file extension/type;
- maximum/minimum file size;
- maximum/minimum file count;
- single/multiple upload capability;
- filename/display after upload;
- remove/replace/re-upload behavior;
- duplicate-file behavior ONLY when specified.
Do NOT create generic upload security/file cases when source is silent.

I. FIELD-TO-FIELD / CONDITIONAL VALIDATION
When source explicitly defines a relation, inspect independently:
- field B required only when field A has value/state X;
- field B enabled/disabled based on field A;
- allowed values of B depend on A;
- numeric/date relation between fields;
- mutually exclusive field combinations;
- conditional default/value propagation.
If the relationship is a business rule executed after Submit/Confirm rather than control-local validation, classify it as BUSINESS_FLOW instead.

CONTROL-TYPE COVERAGE — MANDATORY:
- If source explicitly declares Textbox/Textarea/Numeric/Dropdown/Combobox/Autocomplete/Single-Select/Multi-Select/DatePicker/DateRange/Checkbox/Radio/Toggle/File Upload/Readonly/Input, the declared capability/state is testable and MUST NOT be silently dropped.
- Do NOT copy behavior from another control. Example: Select All / +N / search / debounce on Dropdown A MUST NOT be applied to Dropdown B unless source explicitly defines it for B.
- Do NOT infer generic mandatory/empty/special-character/trim cases merely because a control is an input. A source constraint is required.

VALIDATION ATOMICITY EXAMPLES:
- “CIF: Placeholder + numeric-only + length 10” => at least three independent objectives: Placeholder / Character constraint / Length boundary.
- “Dropdown: default Tất cả + Multi-Select + Select All + search 0.5s” => separate objectives for Default / Multi-Select / Select All behavior / Search+timing, when each is explicitly described.
- “Date range: format dd/MM/yyyy + From > To is auto-swapped” => separate Format objective and Date-relation/auto-swap objective.

4.3 ACTION
Direct control behavior only:
- label/icon/tooltip when explicitly required;
- visibility condition;
- enable/disable;
- click/open/close popup;
- direct navigation;
- trigger the correct action/API when explicitly described.
Do NOT place response/result/message/DB/business outcome under ACTION.

4.4 DATA_GRID
All Data Grid presentation + data mapping:
- structure/header/column set/order;
- width/alignment;
- value format;
- null/empty presentation;
- wrap/ellipsis/tooltip;
- loading/empty state when documented;
- semantic field/response -> correct displayed column/value;
- row/data count mapping when documented.

DATA GRID — SENIOR/HYBRID RULE DESIGN:
A. STRUCTURE:
- ONE Rule for the complete active column set/order.
- If a Mockup/viewport contains only a subset and a later active table gives a fuller list, keep ONLY the fuller active structure Rule.

B. MAPPING:
- Use the tester term `Mapping` in human-readable rule names. NEVER translate it to `Ánh xạ`.
- Preferred rule name form: `Mapping cột <Tên cột>`.
- Mapping failures are independently debuggable.
- DEFAULT: create ONE Mapping Rule per independently meaningful semantic column when source describes what that column displays.
- Example: ID, Chi nhánh, CIF, Tên khách hàng, Mã trái phiếu, Ngày ghi nhận, SL đặt mua, Giá mua, Số tiền đặt mua, Tài khoản đặt mua, Ngân hàng, Trạng thái, Người cập nhật... may each have an independent Mapping Rule.
- Only group multiple columns into one Mapping Rule when the source explicitly defines one inseparable shared mapping rule and separate failures would not be meaningful.
- Width/format/tooltip coverage NEVER replaces mapping coverage.
- Do NOT invent API field names when source only provides semantic display meaning.

C. SHARED PRESENTATION:
To avoid spam, group common presentation behavior when the same rule applies to many columns:
- one Width Rule may cover all explicitly defined widths;
- one Ellipsis + Tooltip Rule may cover all columns sharing that behavior;
- one Numeric Format Rule may cover columns sharing the same numeric format;
- one Date/DateTime Format Rule may cover columns sharing the same format;
- one Scroll/Fixed-column Rule may cover the shared Grid behavior.

4.5 BUSINESS_FLOW
Includes source-grounded feature behavior such as:
- search/filter/reset execution;
- pagination behavior documented by source;
- business rules/state transition;
- FE request parameter/value sent to API;
- successful response -> UI;
- toast/message/navigation after processing;
- Hold/Confirm/Cancel/Copy/Edit/Submit business result.

FILTER END-TO-END COVERAGE — SENIOR STYLE:
When source explicitly states that a filter participates in search:
- Preserve the FE request/parameter Rule when parameters are documented.
- ALSO create a separate end-to-end search-result Rule when source explicitly says the list is queried/displayed according to that criterion.
- Example: select Chi nhánh A -> Search -> displayed records satisfy Chi nhánh A, ONLY if that relationship is supported by source.
- Do NOT create result logic for a control that source does not state is a search criterion.

4.6 EXCEPTION
Only source-grounded abnormal/technical paths:
- explicit server/system error;
- timeout/no response;
- network/technical error;
- no-permission error when source describes it;
- duplicate/technical failure when documented.
Normal business rejection belongs to BUSINESS_FLOW, not EXCEPTION.

==================================================
5. FINAL TESTER ORGANIZATION — SIX WEB SECTIONS
==================================================

Every Rule MUST also have exactly one `feature_group` and one stable human-readable `feature_name`.
This is FINAL TESTER ORGANIZATION only; do NOT change the internal QA `category` merely to fit a section.

Allowed `feature_group` values — EXACTLY SIX:
UI | VALIDATE | FUNCTION | POPUP | DATA_GRID | EXCEPTION

5.1 UI
Contains screen access/permission and screen-level static presentation.
- Permission/access/prerequisite Rules: feature_group=UI, feature_name="Permission".
- Breadcrumb, screen title, static labels, sections, icons, static visibility: feature_group=UI, feature_name="Giao diện chung".
- Do NOT create separate UI feature containers for Breadcrumb/Header/Label. Keep those as target/rule_name inside "Giao diện chung".

5.2 VALIDATE
Contains field/control behavior and independently failing input rules:
- Placeholder, Default, Required/Optional, Enable/Disable, Readonly/Editable;
- Length/Min/Max, character class, format/pattern/mask;
- dropdown options/selection/search, date relations, checkbox/radio/toggle, upload constraints;
- field-to-field validation/dependency when it is control-local.
Use a stable feature_name for the control/business field, e.g. "CIF", "Chi nhánh", "Trạng thái".

5.3 FUNCTION
Contains tester-visible business functions/actions:
- Search, Reset, Create/Add, View detail, Edit, Copy, Approve, Cancel, Hold, Confirm, Upload, navigation;
- end-to-end search/filter execution and source-grounded business flows initiated from the Web UI.
Keep all Rules for the same function under the same feature_name.

5.4 POPUP
Popup/dialog is its own tester section when source describes a popup/modal/dialog.
- Popup UI, popup validation and popup actions all use feature_group=POPUP.
- Preserve the internal category (UI / VALIDATION / ACTION / BUSINESS_FLOW) for QA reasoning.
- feature_name is the popup business name, e.g. "Popup xác nhận phê duyệt".

5.5 DATA_GRID
- ONE GRID = ONE feature_name, e.g. "Data Grid" or "Danh sách kết quả".
- Structure, Mapping, presentation, empty-state, explicit sort/pagination behavior stay under the same grid container.
- Column names belong to target/rule_name, NOT feature_name.
- Use the tester term `Mapping`; NEVER output `Ánh xạ`.

5.6 EXCEPTION
Contains ONLY source-grounded abnormal/technical behavior:
- explicit timeout/no response, server/system/network failure, source-described technical failure;
- source-described no-permission error when it is an error path.
- Normal business rejection remains BUSINESS_FLOW internally, but is organized under FUNCTION unless it is explicitly a technical exception.
feature_name should identify the affected feature, e.g. "Tìm kiếm", "Chi nhánh", "Phê duyệt".

OWNERSHIP EXAMPLES:
- Screen access permission: category=UI or BUSINESS_FLOW as appropriate, feature_group=UI, feature_name="Permission".
- Breadcrumb/title/labels: category=UI, feature_group=UI, feature_name="Giao diện chung".
- CIF placeholder/length/numeric-only: category=VALIDATION, feature_group=VALIDATE, feature_name="CIF".
- Search by CIF: category=BUSINESS_FLOW, feature_group=FUNCTION, feature_name="Tìm kiếm".
- Search timeout: category=EXCEPTION, feature_group=EXCEPTION, feature_name="Tìm kiếm".
- Popup confirmation title: category=UI, feature_group=POPUP, feature_name="Popup xác nhận".
- Popup confirmation action: category=ACTION, feature_group=POPUP, feature_name="Popup xác nhận".
- Cột ID Mapping: category=DATA_GRID, feature_group=DATA_GRID, feature_name="Data Grid", target="Cột ID".

==================================================
6. EXPLICIT / DERIVED QA RULES
==================================================

EXPLICIT:
- Behavior directly described by source.
- applied_qa_rule = "EXPLICIT FROM SPEC".
- generation_reason = "".

DERIVED is allowed ONLY from a real source constraint.
Allowed — ONLY when the corresponding source constraint really exists:
- Exact Length = N -> MUST create THREE independent Rules/Testcases: N-1, N, N+1. Never combine those values into one Rule.
- MinLength / MaxLength -> boundary values around the stated limit.
- Numeric Minimum / Maximum -> boundary values around the stated limit.
- Decimal scale / maximum decimal places = N -> valid scale and one value exceeding N decimals.
- Required -> missing/empty when appropriate for the declared control type.
- Allowed character/type -> valid value + value violating the explicit type/character constraint.
- Format/pattern/mask -> valid format + invalid format.
- Enum/allowed option set -> in-enum + outside-enum ONLY when the control/input can realistically receive an outside value.
- Date/time minimum/maximum -> boundary values around the stated date/time limit.
- Explicit From/To relationship -> valid relation + violating relation; preserve exact source handling such as auto-swap if stated.
- Maximum selection count = N -> N-1 / N / N+1 are independent Rules when the UI can reach those states.
- File size/count limit = N -> boundary around N when an upload control and explicit limit are present.

VISIBILITY CONDITION PARTITION — SENIOR STYLE:
If source explicitly says a control is visible/available ONLY WHEN a condition is true:
- Keep the positive condition Rule.
- You MAY derive negative partition Rule(s) for independently meaningful condition failures.
- Expected result is ONLY the logical opposite visibility/availability supported by the "only when" requirement.
- Do NOT invent an error message, API response, permission matrix, or alternative business behavior.
- For a compound condition A AND B, a negative partition for not-A and/or not-B is allowed when each is independently testable and grounded in the stated condition.
- applied_qa_rule may be "Equivalence Partitioning" or "Decision Table".

DERIVED requires:
- original source constraint in source_requirement;
- actual QA technique in applied_qa_rule;
- short Vietnamese generation_reason.

DO NOT DERIVE:
- new business rule/API/message/status/permission/DB behavior/technical exception.

==================================================
7. ATOMICITY — SENIOR-LIKE BUT NOT SPAMMY
==================================================

One Rule = one independently failing test objective.
Split when condition/expected behavior/target/parameter/business branch is independently testable.

DO NOT over-split unrelated equivalent behavior. HOWEVER boundary values are atomic:
- N-1, N and N+1 MUST be separate Rules/Testcases when derived from an exact/boundary constraint;
- static UI-only elements may be separate Rules but the UI section uses one shared `feature_name="Giao diện chung"` for general screen UI;
- common Grid width/format/tooltip behavior may be grouped as defined above.

DO split:
- Data Grid Mapping per independent semantic column by default;
- independently failing action visibility branches;
- API request vs success reload vs success toast;
- explicit server error vs timeout/no response.

==================================================
8. TEST CONDITION & EXPECTED RESULT — MANDATORY
==================================================

`test_condition`:
- state concrete input/state/role/value/branch required to execute the objective;
- may be empty only when no special condition/data is needed;
- grouped boundaries must include the full data set.

`expected_result`:
- MUST be non-empty, specific, and Pass/Fail-verifiable for every Rule;
- preserve exact source values/messages/formats/states;
- for DERIVED boundary/negative Rules where source does not define the UI rejection mechanism, state the result at CONTRACT LEVEL and do NOT invent truncate/block/message behavior.

Example for source "CIF length = 10":
- test_condition: "Nhập lần lượt CIF có độ dài 9, 10 và 11 ký tự."
- expected_result: "CIF 10 ký tự thỏa ràng buộc độ dài; CIF 9 và 11 ký tự không thỏa ràng buộc độ dài 10. Không tự khẳng định cơ chế chặn/cắt/message nếu Spec không mô tả."

Do NOT use vague Expected such as "Hệ thống xử lý đúng" or "Theo Spec" when exact behavior exists.

==================================================
9. REQUIREMENT PRESERVATION & DUPLICATE CONTROL
==================================================

`source_requirement` must be compact but preserve all details that affect the test:
- default/min/max/length/condition/role/state;
- request parameter/value;
- response/message;
- dependency/time/debounce/select behavior/format/mapping.

Duplicate identity is behavior-based, not wording.
Do not create two Rules with equivalent target + objective + condition + expected behavior.
Repeated OCR/table fragments MUST NOT duplicate Rules.
If one same-objective requirement is only a subset of a fuller active requirement, keep the fuller active requirement.

==================================================
9.5 TESTER-FACING WORDING — HUMAN QA STYLE
==================================================

Rule content must support deterministic Steps that read like a tester wrote them.
- Prefer concrete verbs: Mở, Nhập, Chọn, Bỏ chọn, Nhấn, Xóa, Tải lên, Tìm kiếm, Chuyển trang, Sắp xếp, Đối chiếu, Xác nhận, Đóng popup.
- `test_condition` should describe the actual action/data condition, e.g. "Nhập CIF gồm 9 ký tự số", NOT "Thiết lập dữ liệu cho CIF".
- `rule_name` should be an executable scenario title, e.g. "CIF mặc định để trống", "Hiển thị placeholder CIF", "Tìm kiếm theo CIF hợp lệ", "Đặt lại điều kiện tìm kiếm", "Mapping cột Người duyệt".
- NEVER use vague/meta wording such as "Thiết lập dữ liệu kiểm thử", "Áp dụng điều kiện kiểm thử", "Thực hiện thao tác tương ứng", "Quan sát kết quả", "Kiểm tra theo yêu cầu", "Mục tiêu kiểm thử", or any wording referring to AI/model generation.
- Do not invent a click/blur/submit trigger when the source does not describe which event triggers validation. In that case state only the concrete input/action that is source-grounded.

==================================================
10. LOCAL QUALITY GATE
==================================================

Before output, verify:
1. All meaningful CURRENT SOURCE requirements are mapped.
2. Every Rule has exactly one valid category.
3. Every Rule has exactly one valid feature_group and a stable feature_name.
4. Same Filter/Function is not scattered under inconsistent feature_name variants.
5. No OCR/source repetition created duplicates.
6. No old/strikethrough requirement overrode active revision.
7. Important parameter/value/message/format/mapping details were preserved.
8. Every DERIVED Rule has a real source constraint and QA technique.
9. No generic best-practice testcase was invented.
10. Data Grid structure uses the complete active column list, not a partial duplicate.
11. Data Grid semantic Mapping covers every described column, preferably one independently debuggable Rule per column.
12. Shared Grid presentation is compact rather than one width/tooltip TC per column.
13. Every explicitly declared Input-Control capability/state was scanned and not silently dropped: Placeholder/Default/Required/Enable/Readonly/Length-Type-Format/Selection/Search/Dependency as applicable.
14. Explicit Single-Select/Multi-Select/Select-All/Clear/+N/search/debounce behavior was preserved only for the control that owns it.
15. Text/Numeric/Date/Selection/File-upload constraints were converted into separate independently failing Validation objectives instead of one vague “validation” Rule.
16. Boundary/negative Validation Rules were created only from real source constraints and do not invent message/reject/truncate/rounding behavior.
17. If source explicitly says each filter participates in search, end-to-end result coverage is not silently replaced only by API-parameter coverage.
18. `expected_result` is non-empty and verifiable.
19. Visibility "only when" conditions have appropriate source-grounded positive/negative partitions without invented messages.

Do not invent Rules just to increase testcase count.

==================================================
11. OUTPUT JSON — MANDATORY
==================================================

Return ONLY valid JSON. NO markdown/explanation/text outside JSON.

Required root:
{
  "test_design_version": "3.5",
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

All human-readable values MUST be in VIETNAMESE; technical identifiers and established QA terms such as `Mapping` remain exact. NEVER output `Ánh xạ`; use `Mapping`.
If CURRENT SOURCE has no sufficiently grounded testable requirement:
{"test_design_version":"3.5","screens":[]}

==================================================
CHUNK SOURCE
==================================================

{content}
"""




PROMPT_API_SCOPE_ROUTER = """
You are API SCOPE ROUTER for multi-document QA analysis.

Your task is ONLY to classify how the BA CURRENT SOURCE relates to each PRIMARY TARGET API defined by API Design/Spec.
You MUST NOT generate testcases, QA rules, expected results, or rewrite requirements.
Thinking should be minimal and deterministic.

SCOPE ROLES — MANDATORY:
- PRIMARY: the chunk describes the target API itself, including its request validation, business eligibility, processing steps, direct side effects, or direct checks performed inside that API.
- CONTINUATION: the chunk describes a DISTINCT callable API/operation that is a required next step to COMPLETE the same business capability/flow started by the PRIMARY API. Example: Approve -> Confirm Approval after authentication/signing.
- DEPENDENCY: the chunk describes a supporting service/API/plugin/storage/payment/workflow call used by PRIMARY or CONTINUATION, but it is not itself the user-facing continuation that should become a full testcase endpoint workspace.
- OUT_OF_SCOPE: sibling/alternative operation not required to complete the selected flow, e.g. List, Detail, Reject when PRIMARY is Approve; unrelated Print/Download/Reference APIs.

ROUTING PRINCIPLES:
- API Design/Spec defines the PRIMARY target API(s).
- Do NOT deep-analyze all 30+ BA pages. Classify each BA chunk first.
- A same-domain sibling is NOT automatically relevant.
- CONTINUATION must be a necessary next callable step in the selected business flow, not merely a nearby API.
- Reject is usually an alternative branch to Approve, therefore OUT_OF_SCOPE when Approve is the target unless the spec explicitly selects Reject.
- Confirm Approval after Approve/authentication is CONTINUATION when the BA shows it is required to complete approval.
- Signing/auth init/verify, workflow service, S3/storage, account inquiry, payment, etc. are normally DEPENDENCY unless the BA clearly presents one as the selected continuation API itself.
- For CONTINUATION, preserve an exact method/path ONLY if explicitly visible in CURRENT SOURCE. Never invent an endpoint. If unknown, use "UNMAPPED".
- Be conservative about scope contamination, but do not cut off required continuation steps.

Return ONLY valid JSON:
{
  "matches": [
    {
      "method": "POST",
      "endpoint_path": "/primary/example",
      "scope_role": "PRIMARY | CONTINUATION | DEPENDENCY | OUT_OF_SCOPE",
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
For PRIMARY/DEPENDENCY/OUT_OF_SCOPE, related_* may be blank/UNMAPPED.
For CONTINUATION, related_api_name is mandatory; related_method/path are exact only when source explicitly states them.
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
              "rule_id": "TMP-001",
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
