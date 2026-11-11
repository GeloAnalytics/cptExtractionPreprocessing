@echo off
rem =========================================================================
rem Double-click this file to start table2sheet.
rem No installation, no administrator rights, and no command prompt needed.
rem =========================================================================
cd /d "%~dp0"

title table2sheet Web Server (localhost:8765)

rem Terminate any old/lingering table2sheet server instances holding port 8765
for /f "tokens=5" %%a in ('netstat -aon ^| findstr :8765 ^| findstr LISTENING 2^>nul') do (
    taskkill /F /PID %%a >nul 2>&1
)

set "PY="

rem 1. Check if bundled in runtime/ folder (FIRST PRIORITY!)
if exist "%~dp0runtime\python\python.exe" (
    set "PY=%~dp0runtime\python\python.exe"
    goto :launch
)
if exist "%~dp0runtime\python\pythonw.exe" (
    set "PY=%~dp0runtime\python\pythonw.exe"
    goto :launch
)
if exist "%~dp0..\runtime\python\python.exe" (
    set "PY=%~dp0..\runtime\python\python.exe"
    goto :launch
)
if exist "%~dp0..\runtime\python\pythonw.exe" (
    set "PY=%~dp0..\runtime\python\pythonw.exe"
    goto :launch
)

rem 2. Check if standard python is available on system PATH
where python >nul 2>&1
if %errorlevel% equ 0 (
    set "PY=python"
    goto :launch
)

where pythonw >nul 2>&1
if %errorlevel% equ 0 (
    set "PY=pythonw"
    goto :launch
)

rem 3. Check common user-level Python installs (no admin required)
for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*") do (
    if exist "%%D\python.exe" (
        set "PY=%%D\python.exe"
        goto :launch
    )
    if exist "%%D\pythonw.exe" (
        set "PY=%%D\pythonw.exe"
        goto :launch
    )
)

for /d %%D in ("%USERPROFILE%\AppData\Local\Programs\Python\Python3*") do (
    if exist "%%D\python.exe" (
        set "PY=%%D\python.exe"
        goto :launch
    )
    if exist "%%D\pythonw.exe" (
        set "PY=%%D\pythonw.exe"
        goto :launch
    )
)

echo.
echo ============================================================
echo   Python was not detected on this computer.
echo ============================================================
echo.
echo   table2sheet requires Python or the bundled runtime/ folder.
echo.
pause
exit /b 1

:launch
echo ============================================================
echo   table2sheet v2.1 • Offline Web UI Server
echo   Running at http://localhost:8765
echo.
echo   * Keep this window open while using table2sheet.
echo   * To stop: close this window or click "Quit App" in browser.
echo ============================================================
echo.
if exist "table2sheet.py" (
    "%PY%" "table2sheet.py" --web
) else if exist "table2sheet_app\table2sheet.py" (
    "%PY%" "table2sheet_app\table2sheet.py" --web
) else (
    echo Error: table2sheet.py could not be found.
    pause
    exit /b 1
)
exit /b 0
