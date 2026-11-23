Write-Host "===============================================================" -ForegroundColor Cyan
Write-Host "  AI DATA ANALYSIS AGENT - LOCALHOST LAUNCHER" -ForegroundColor Magenta
Write-Host "===============================================================" -ForegroundColor Cyan
Write-Host "Dang mo trinh duyet tai: http://localhost:8000" -ForegroundColor Yellow
Start-Process "http://localhost:8000"
python server.py
