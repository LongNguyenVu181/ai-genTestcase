$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

try {
    Set-Location "$root\frontend"
    npm run build
    Write-Host "Frontend build xong. Chay backend de serve ca API + UI:" -ForegroundColor Green
    Write-Host "cd '$root\backend'" -ForegroundColor Yellow
    Write-Host "py -m uvicorn app:app --host 127.0.0.1 --port 8000" -ForegroundColor Yellow
}
finally {
    Set-Location $root
}
