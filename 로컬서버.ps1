# cheongwon 포털 로컬 미리보기
# → http://localhost:8080/청원지구_포털
# 탐색기에서 청원지구_포털.html 더블클릭 시에도 동일 URL로 자동 이동(서버 필요)
chcp 65001 | Out-Null
Set-Location $PSScriptRoot
$port = 8080
$url = "http://localhost:$port/청원지구_포털"

# 이미 8080에서 serve_portal이 떠 있으면 브라우저만 연다
$listening = Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue
if ($listening) {
    Write-Host "  (서버 실행 중) $url" -ForegroundColor Green
    Start-Process $url
    exit 0
}

Write-Host ""
Write-Host "  $url" -ForegroundColor Cyan
Write-Host "  (종료: Ctrl+C)" -ForegroundColor DarkGray
Write-Host ""
Start-Process $url
python -X utf8 tools/serve_portal.py
