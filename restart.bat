@echo off
rem Quick Restart for table2sheet Web Server
cd /d "%~dp0"
echo ============================================================
echo   Restarting table2sheet server...
echo ============================================================

for /f "tokens=5" %%a in ('netstat -aon ^| findstr :8765 ^| findstr LISTENING 2^>nul') do (
    echo Terminating old server (PID %%a)...
    taskkill /F /PID %%a >nul 2>&1
)
timeout /t 1 /nobreak >nul

call run.bat

