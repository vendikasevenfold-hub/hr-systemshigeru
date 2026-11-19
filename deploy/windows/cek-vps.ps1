<#
  Cek kesiapan VPS untuk Shigeru HR. HANYA MEMBACA - tidak mengubah apa pun.
  Jalankan di PowerShell "Run as Administrator":
    powershell -ExecutionPolicy Bypass -File C:\cek-vps.ps1
  Hasil juga disimpan ke Desktop\cek-vps-hasil.txt (kirimkan isinya untuk langkah berikutnya).
#>
$ErrorActionPreference = "SilentlyContinue"
$out = New-Object System.Collections.Generic.List[string]
function Tulis($s) { $out.Add([string]$s); Write-Host $s }
function Judul($s) { Tulis ""; Tulis "=== $s ===" }

Judul "Sistem"
$os = Get-CimInstance Win32_OperatingSystem
Tulis ("OS        : {0} (build {1})" -f $os.Caption, $os.BuildNumber)
Tulis ("RAM       : {0:N1} GB total, {1:N1} GB bebas" -f ($os.TotalVisibleMemorySize / 1MB), ($os.FreePhysicalMemory / 1MB))
$c = Get-PSDrive C
Tulis ("Disk C    : {0:N1} GB bebas" -f ($c.Free / 1GB))
$tz = Get-TimeZone
Tulis ("Zona waktu: {0} ({1})" -f $tz.Id, $tz.DisplayName)

Judul "Program"
foreach ($cmd in "node", "npm", "git", "psql") {
  $g = Get-Command $cmd
  if ($g) {
    $v = (& $g.Source --version 2>&1 | Select-Object -First 1)
    Tulis ("{0,-5}: {1}  [{2}]" -f $cmd, $v, $g.Source)
  } else { Tulis ("{0,-5}: belum terpasang" -f $cmd) }
}
Get-ChildItem "C:\Program Files\PostgreSQL" -Directory | ForEach-Object { Tulis "PostgreSQL folder: $($_.FullName)" }
Get-Service | Where-Object { $_.Name -match 'postgres|mysql|maria|apache|httpd|nssm|ShigeruHR' } |
  ForEach-Object { Tulis ("Service   : {0} ({1}) - {2}" -f $_.Name, $_.DisplayName, $_.Status) }
foreach ($p in "C:\nssm\nssm.exe", "C:\win-acme\wacs.exe") { if (Test-Path $p) { Tulis "Ada       : $p" } }

Judul "Port yang dipakai (80, 443, 3000, 3100, 3306, 5432)"
Get-NetTCPConnection -State Listen | Where-Object { $_.LocalPort -in 80, 443, 3000, 3100, 3306, 5432 } |
  Sort-Object LocalPort -Unique | ForEach-Object {
    $pr = Get-Process -Id $_.OwningProcess
    Tulis ("{0,-5} {1,-15} {2} [{3}]" -f $_.LocalPort, $_.LocalAddress, $pr.ProcessName, $pr.Path)
  }

Judul "Apache"
$httpd = @()
$httpd += Get-Process httpd | ForEach-Object { $_.Path }
$httpd += "C:\xampp\apache\bin\httpd.exe", "C:\Apache24\bin\httpd.exe"
$httpd += Get-ChildItem "C:\wamp64\bin\apache\*\bin\httpd.exe", "C:\wamp\bin\apache\*\bin\httpd.exe" | ForEach-Object { $_.FullName }
$httpd = $httpd | Where-Object { $_ -and (Test-Path $_) } | Select-Object -Unique
if (-not $httpd) { Tulis "httpd.exe tidak ditemukan di lokasi umum. Catat lokasi Apache secara manual." }
foreach ($h in $httpd) {
  Tulis "httpd.exe : $h"
  Tulis ("Versi     : " + (& $h -v 2>&1 | Select-Object -First 1))
  Tulis "--- VirtualHost & konfigurasi (httpd -S) ---"
  & $h -S 2>&1 | ForEach-Object { Tulis "  $_" }
  $root = Split-Path (Split-Path $h)
  $conf = Join-Path $root "conf\httpd.conf"
  if (Test-Path $conf) {
    Tulis "--- Modul di $conf ---"
    Select-String -Path $conf -Pattern 'LoadModule\s+(proxy_module|proxy_http_module|headers_module|ssl_module|rewrite_module)' |
      ForEach-Object { Tulis ("  " + $_.Line.Trim()) }
    Tulis "--- DocumentRoot / Include ---"
    Select-String -Path $conf -Pattern '^\s*(DocumentRoot|Include|Listen)\b' | ForEach-Object { Tulis ("  " + $_.Line.Trim()) }
    Tulis "--- File berisi <VirtualHost / ServerName / SSLCertificateFile ---"
    Get-ChildItem (Join-Path $root "conf") -Recurse -Include *.conf | Where-Object { $_.FullName -notmatch '\\original\\' } |
      Select-String -Pattern '^\s*(<VirtualHost|ServerName|ServerAlias|SSLCertificateFile|DocumentRoot)' |
      ForEach-Object { Tulis ("  {0}:{1}: {2}" -f $_.Path.Replace($root, ''), $_.LineNumber, $_.Line.Trim()) }
  }
}

Judul "GLPI"
Get-ChildItem "C:\xampp\htdocs", "C:\Apache24\htdocs", "C:\wamp64\www", "C:\inetpub\wwwroot" -Directory -Filter "*glpi*" |
  ForEach-Object { Tulis "Folder GLPI: $($_.FullName)" }

Judul "Koneksi internet keluar"
foreach ($h in "github.com", "registry.npmjs.org", "acme-v02.api.letsencrypt.org") {
  $t = Test-NetConnection $h -Port 443 -WarningAction SilentlyContinue
  Tulis ("{0,-30} {1}" -f $h, $(if ($t.TcpTestSucceeded) { "OK" } else { "GAGAL" }))
}

$hasil = Join-Path ([Environment]::GetFolderPath("Desktop")) "cek-vps-hasil.txt"
$out | Set-Content -Encoding UTF8 $hasil -ErrorAction Stop
Write-Host ""
Write-Host "Selesai. Hasil disimpan di $hasil" -ForegroundColor Green
