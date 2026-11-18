<#
  Pasang Shigeru HR sebagai Windows Service (otomatis jalan saat server menyala dan
  dinyalakan ulang bila berhenti). Jalankan di PowerShell "Run as Administrator":

    powershell -ExecutionPolicy Bypass -File C:\shigeru-hr\deploy\windows\install-service.ps1 -NssmPath C:\nssm\nssm.exe

  NSSM (Non-Sucking Service Manager) diunduh sendiri dari https://nssm.cc/download.
#>
param(
  [Parameter(Mandatory = $true)][string]$NssmPath,
  [string]$AppDir = "C:\shigeru-hr",
  [string]$ServiceName = "ShigeruHR"
)
$ErrorActionPreference = "Stop"

if (-not (Test-Path $NssmPath)) { throw "nssm.exe tidak ditemukan di $NssmPath" }
if (-not (Test-Path (Join-Path $AppDir ".env"))) { throw "File $AppDir\.env belum ada. Salin dari deploy\windows\.env.vps.example lalu isi." }
$node = (Get-Command node -ErrorAction Stop).Source
$logDir = Join-Path $AppDir "logs"
New-Item -ItemType Directory -Force $logDir | Out-Null

if (Get-Service $ServiceName -ErrorAction SilentlyContinue) {
  Write-Host "Service $ServiceName sudah ada - diperbarui."
  & $NssmPath stop $ServiceName | Out-Null
} else {
  & $NssmPath install $ServiceName $node | Out-Null
}

& $NssmPath set $ServiceName Application $node | Out-Null
& $NssmPath set $ServiceName AppParameters "--env-file=.env server\index.js" | Out-Null
& $NssmPath set $ServiceName AppDirectory $AppDir | Out-Null
& $NssmPath set $ServiceName DisplayName "Shigeru HR" | Out-Null
& $NssmPath set $ServiceName Description "Aplikasi Shigeru HR (Node.js) - PT Mideli Makmur Jaya" | Out-Null
& $NssmPath set $ServiceName Start SERVICE_AUTO_START | Out-Null
& $NssmPath set $ServiceName AppStdout (Join-Path $logDir "shigeru-hr.log") | Out-Null
& $NssmPath set $ServiceName AppStderr (Join-Path $logDir "shigeru-hr.log") | Out-Null
& $NssmPath set $ServiceName AppRotateFiles 1 | Out-Null
& $NssmPath set $ServiceName AppRotateBytes 10485760 | Out-Null
& $NssmPath set $ServiceName AppExit Default Restart | Out-Null
& $NssmPath set $ServiceName AppRestartDelay 5000 | Out-Null
# Beri waktu menyimpan data ke database sebelum proses dihentikan.
& $NssmPath set $ServiceName AppStopMethodConsole 15000 | Out-Null

& $NssmPath start $ServiceName | Out-Null
Start-Sleep -Seconds 5
Get-Service $ServiceName | Format-Table -AutoSize

$port = (Get-Content (Join-Path $AppDir ".env") | Where-Object { $_ -match '^PORT=' }) -replace '^PORT=', ''
if (-not $port) { $port = "3000" }
try {
  $r = Invoke-WebRequest -UseBasicParsing "http://127.0.0.1:$port/healthz" -TimeoutSec 10
  Write-Host "OK: aplikasi menjawab di http://127.0.0.1:$port ($($r.Content))"
} catch {
  Write-Warning "Aplikasi belum menjawab di port $port. Cek log: $logDir\shigeru-hr.log"
}
