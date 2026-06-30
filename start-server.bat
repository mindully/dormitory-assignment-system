@echo off
cd /d "%~dp0"
echo [1/3] Stopping any existing server on port 3000...
for /f "tokens=5" %%a in ('netstat -ano ^| findstr :3000 ^| findstr LISTENING') do (
  taskkill /PID %%a /F >nul 2>&1
  echo   Killed old server (PID %%a)
)
timeout /t 2 /nobreak >nul
echo [2/3] Starting new server...
start /B "" cmd /c "npx tsx server.ts >> server.log 2>&1"
timeout /t 3 /nobreak >nul
echo [3/3] Verifying server is running...
curl -s http://localhost:3000/api/health
echo.
echo Server started! Check http://localhost:3000
echo Log file: server.log
