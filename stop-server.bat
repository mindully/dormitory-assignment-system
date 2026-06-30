@echo off
echo Stopping dormitory assignment server...
for /f "tokens=5" %%a in ('netstat -ano ^| findstr :3000 ^| findstr LISTENING') do (
  taskkill /PID %%a /F >nul 2>&1
  echo Server (PID %%a) stopped.
)
timeout /t 2 /nobreak >nul
echo Server stopped.
