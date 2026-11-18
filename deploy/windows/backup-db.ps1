<#
  Backup database Shigeru HR (pg_dump format custom), simpan 14 hari terakhir.

  Sekali jalan (tes):
    powershell -ExecutionPolicy Bypass -File C:\shigeru-hr\deploy\windows\backup-db.ps1

  Daftarkan jadwal otomatis setiap hari jam 01:00 (Run as Administrator):
    powershell -ExecutionPolicy Bypass -File C:\shigeru-hr\deploy\windows\backup-db.ps1 -Register
#>
param(
  [string]$AppDir = "C:\shigeru-hr",
  [string]$BackupDir = "C:\shigeru-hr-backup",
  [string]$PgBin = "C:\Program Files\PostgreSQL\16\bin",
  [int]$KeepDays = 14,
  [switch]$Register
)
$ErrorActionPreference = "Stop"

if ($Register) {
  $action = New-ScheduledTaskAction -Execute "powershell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -AppDir `"$AppDir`" -BackupDir `"$BackupDir`" -PgBin `"$PgBin`" -KeepDays $KeepDays"
  $trigger = New-ScheduledTaskTrigger -Daily -At 1:00am
  Register-ScheduledTask -TaskName "ShigeruHR-Backup" -Action $action -Trigger $trigger -User "SYSTEM" -RunLevel Highest -Force | Out-Null
  Write-Host "Jadwal backup harian 01:00 terdaftar (Task Scheduler: ShigeruHR-Backup)."
  return
}

# Ambil user, password, host, port, database dari DATABASE_URL di .env
$line = Get-Content (Join-Path $AppDir ".env") | Where-Object { $_ -match '^DATABASE_URL=' } | Select-Object -First 1
if (-not $line) { throw "DATABASE_URL tidak ditemukan di $AppDir\.env" }
$uri = [Uri]($line -replace '^DATABASE_URL=', '')
$user, $pass = $uri.UserInfo.Split(':', 2)
$db = $uri.AbsolutePath.TrimStart('/')

New-Item -ItemType Directory -Force $BackupDir | Out-Null
$file = Join-Path $BackupDir ("shigeru-hr-" + (Get-Date -Format "yyyyMMdd-HHmm") + ".dump")
$env:PGPASSWORD = [Uri]::UnescapeDataString($pass)
& (Join-Path $PgBin "pg_dump.exe") -h $uri.Host -p $uri.Port -U $user -Fc -f $file $db
$code = $LASTEXITCODE
Remove-Item Env:PGPASSWORD
if ($code -ne 0) { throw "pg_dump gagal (kode $code)" }

Get-ChildItem $BackupDir -Filter "shigeru-hr-*.dump" |
  Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$KeepDays) } |
  Remove-Item -Force
Write-Host "Backup tersimpan: $file"
