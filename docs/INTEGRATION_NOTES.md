# Integration notes

## Mục tiêu
Ghép UI React/Vite mới của TestPilot AI với Multi-Agent pipeline đang chạy tốt trong `multiA.txt`.

## Phần AI được giữ nguyên
`backend/pipeline_core.py` được trích từ `multiA.txt`, bỏ duy nhất phần giao diện Streamlit và `import streamlit`.
Các phần sau được giữ nguyên:
- cấu hình chunk/batch/recursive split;
- Qwen OpenAI-compatible streaming caller;
- Agent 1 Web Rule Matrix prompt + schema validator;
- Agent 2 Web renderer prompt + invariant 1 Rule = 1 Test Case;
- API Agent 1/2 pipeline;
- XMind / Excel renderer.

## Lớp mới
- `backend/app.py`: FastAPI adapter quanh pipeline cũ.
- `frontend/`: UI React/Vite pastel đã thiết kế trước đó.
- API Key page lưu full key ở `sessionStorage` chỉ phục vụ local testing; model/base URL lưu local config.
- Upload page chạy Agent 1 -> review Rule Matrix -> Agent 2 -> Kho testcase.
- Kho testcase parse output Agent 2 thành accordion và vẫn tải XMind/Excel từ output gốc.

## API endpoints
- `GET /api/health`
- `POST /api/config/test`
- `POST /api/web/analyze`
- `POST /api/web/generate`
- `POST /api/api/analyze`
- `POST /api/api/generate`
- `GET /api/runs/{run_id}`
- `GET /api/runs/{run_id}/rule-matrix`
- `GET /api/runs/{run_id}/excel`
- `GET /api/runs/{run_id}/xmind`

## Lưu ý local test
`RUN_STORE` và cache đang là in-memory để test prompt nhanh. Restart backend sẽ mất run hiện tại. Production nên thay bằng Redis/DB/object storage.
