# TestPilot AI — V1.9.3 Clean Source

Clean production-oriented source for TestPilot AI. Historical prototypes, demo seed data, generated caches, old renderer code, benchmark files, and compatibility-only artifacts are intentionally excluded.

## Structure

```text
TestPilot AI
├─ backend/
│  ├─ app.py
│  ├─ db.py
│  ├─ pipeline_core.py
│  ├─ testcase_mapper.py
│  ├─ excel_export.py
│  └─ requirements.txt
├─ frontend/
│  ├─ src/
│  ├─ index.html
│  ├─ package.json
│  └─ vite.config.js
├─ build-prod.ps1
├─ setup.ps1
├─ start-dev.ps1
└─ .gitignore
```

## Current behavior

- Web and API testcase generation with a session-scoped process-memory store (no SQL database).
- API Spec determines the target API. BA/Service documents are scanned to retain only PRIMARY, CONTINUATION, and relevant DEPENDENCY sections.
- API testcase taxonomy: Auth, Permission, Validation, Happy Path, Business Rule.
- Project metadata, analysis jobs, runs, and testcases exist only for the active browser session.
  Refreshing the page clears the browser session and its matching in-memory backend data.
- AI API keys are not persisted in the database.

## First setup

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\setup.ps1
```

## Production build

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\build-prod.ps1
cd backend
py -m uvicorn app:app --host 127.0.0.1 --port 8000
```

Open `http://127.0.0.1:8000`.

## Deploy on Render

The repository includes a Docker definition and `render.yaml` for a single Render
web service. Create a Blueprint from this GitHub repository and deploy the
`ai-gentestcase` service. It builds React and FastAPI together and does not require
a persistent disk because all app data is session-scoped.

## Development

```powershell
.\start-dev.ps1
```

Backend: `http://127.0.0.1:8000`  
Frontend dev server: `http://localhost:5173`

## Environment options

- `TESTPILOT_WEB_PARALLEL_WORKERS` — Web chunk worker count.
- `TESTPILOT_API_SCOPE_WORKERS` — API BA scope-scan worker count.
