# Dipanggil win-acme setelah sertifikat diperpanjang, supaya Apache memakai sertifikat baru.
$httpd = "C:\wamp64\bin\apache\apache2.4.65\bin\httpd.exe"
$cek = & $httpd -t 2>&1
if ($LASTEXITCODE -ne 0) { Write-Error "Konfigurasi Apache error, Apache TIDAK di-restart: $cek"; exit 1 }
Restart-Service wampapache64
