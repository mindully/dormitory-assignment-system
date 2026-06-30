# PATH에 GitHub CLI 추가 (새 터미널 열면 자동 적용, 현재 세션용)
$env:Path = "$env:ProgramFiles\GitHub CLI;" + [Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [Environment]::GetEnvironmentVariable("Path","User")

# GitHub CLI 로그인 (브라우저 인증)
Write-Host "=== GitHub CLI 로그인 ===" -ForegroundColor Cyan
gh auth login --web -h github.com

Write-Host "`n"

# Vercel CLI 로그인 (브라우저 인증)
Write-Host "=== Vercel CLI 로그인 ===" -ForegroundColor Cyan
vercel login

Write-Host "`n"

# 로그인 상태 확인
Write-Host "=== 로그인 상태 ===" -ForegroundColor Cyan
Write-Host "GitHub CLI:" -NoNewline
gh auth status 2>&1 | Select-String "Logged in" | ForEach-Object { $_ -replace '.*(Logged in.*)', '$1' }

Write-Host "Vercel CLI:" -NoNewline
vercel whoami 2>&1

pause
