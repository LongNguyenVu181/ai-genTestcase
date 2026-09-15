# TestPilot AI V1.9.4 Responsive Hotfix

Copy the contents of this patch over the V1.9.3 clean-source root.

Fixes:
- Dashboard `Dự án mới` same-route no-op.
- Visible create-project validation and immediate navigation.
- Remove redundant/expensive metadata sync on reload.
- One-time browser-cache migration to SQLite backend.
- Faster SQLite project list query.
- GZip + immutable caching for hashed frontend assets.
- setup/build scripts return to project root.

After copying:

```powershell
cd "F:\VCB Packet\AI\testpilot-ai-v1.9.3-clean-source\testpilot-ai-v1.9.3-clean-source"
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\build-prod.ps1
cd backend
py -m uvicorn app:app --host 127.0.0.1 --port 8000
```
