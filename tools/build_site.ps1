# Gom đúng các file cần đưa lên mạng vào thư mục "docs" (GitHub Pages phục vụ từ đây) (không có test, không có dữ liệu thật).
# Chạy: powershell -ExecutionPolicy Bypass -File tools\build_site.ps1
$root = Split-Path -Parent $PSScriptRoot
$site = Join-Path $root 'docs'
if (Test-Path $site) { Remove-Item $site -Recurse -Force }
New-Item -ItemType Directory $site | Out-Null
New-Item -ItemType Directory (Join-Path $site 'vendor') | Out-Null
foreach ($f in 'index.html','style.css','app.js','core.js','sw.js','manifest.webmanifest','icon-192.png','icon-512.png','apple-touch-icon.png') {
  Copy-Item (Join-Path $root $f) $site
}
Copy-Item (Join-Path $root 'vendor\xlsx.full.min.js') (Join-Path $site 'vendor')
Get-ChildItem $site -Recurse -File | Select-Object @{n='File';e={$_.FullName.Substring($site.Length+1)}}, Length | Format-Table -AutoSize

