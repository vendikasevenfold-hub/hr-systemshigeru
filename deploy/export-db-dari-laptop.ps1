<#
  Ambil isi database dari Docker di laptop ke file shigeru-hr.dump (untuk dipindah ke VPS).
  Jalankan dari folder repo:
    powershell -ExecutionPolicy Bypass -File deploy\export-db-dari-laptop.ps1
#>
$ErrorActionPreference = "Stop"
Set-Location (Join-Path $PSScriptRoot "..")
docker compose exec -T db sh -c 'pg_dump -U "$POSTGRES_USER" -Fc -f /tmp/shigeru-hr.dump "$POSTGRES_DB"'
if ($LASTEXITCODE -ne 0) { throw "pg_dump di container gagal" }
docker compose cp db:/tmp/shigeru-hr.dump .\shigeru-hr.dump
if ($LASTEXITCODE -ne 0) { throw "Gagal menyalin file dari container" }
docker compose exec -T db rm -f /tmp/shigeru-hr.dump
$f = Get-Item .\shigeru-hr.dump
Write-Host ("Selesai: {0} ({1:N1} MB). Salin file ini ke VPS lewat Remote Desktop." -f $f.FullName, ($f.Length / 1MB))
Write-Host "PERHATIAN: file ini berisi data karyawan & gaji. Hapus dari laptop setelah dipindah."
