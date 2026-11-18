<#
  Perbarui aplikasi di VPS ke versi terbaru dari GitHub lalu nyalakan ulang service.
  Jalankan di PowerShell "Run as Administrator":
    powershell -ExecutionPolicy Bypass -File C:\shigeru-hr\deploy\windows\update.ps1
#>
param(
  [string]$AppDir = "C:\shigeru-hr",
  [string]$ServiceName = "ShigeruHR"
)
$ErrorActionPreference = "Stop"
Set-Location $AppDir

& (Join-Path $PSScriptRoot "backup-db.ps1") -AppDir $AppDir
git pull --ff-only
if ($LASTEXITCODE -ne 0) { throw "git pull gagal" }
npm ci --omit=dev
if ($LASTEXITCODE -ne 0) { throw "npm ci gagal" }
Restart-Service $ServiceName
Start-Sleep -Seconds 5
Get-Service $ServiceName | Format-Table -AutoSize
