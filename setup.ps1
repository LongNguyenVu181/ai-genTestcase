$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "[1/2] Cai Python backend dependencies..." -ForegroundColor Cyan
Set-Location "$root\backend"
py -m pip install -r requirements.txt

Write-Host "[2/2] Cai frontend dependencies..." -ForegroundColor Cyan
Set-Location "$root\frontend"
npm install

Write-Host "Hoan tat setup." -ForegroundColor Green
