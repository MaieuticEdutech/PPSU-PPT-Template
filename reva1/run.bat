@echo off
setlocal
title PPT Designer - launcher

REM Always work from the folder this .bat lives in, so double-clicking works
REM no matter what the current directory happens to be.
cd /d "%~dp0"

REM ---------- make sure Python is available ----------
where python >nul 2>&1
if errorlevel 1 (
    echo.
    echo  [X] Python was not found on this machine.
    echo      Install Python and tick "Add python.exe to PATH", then run this again.
    echo.
    pause
    exit /b 1
)

REM ---------- work out this machine's LAN IP (for the shareable link) ----------
set "LAN_IP="
for /f "delims=" %%i in ('python "%~dp0show_ip.py" --bare 2^>nul') do set "LAN_IP=%%i"
if "%LAN_IP%"=="" set "LAN_IP=127.0.0.1"

cls
echo ==================================================
echo               P P T   D E S I G N E R
echo ==================================================
echo.
echo   On this computer : http://localhost:8002/
echo   Share on network : http://%LAN_IP%:8002/
echo.
echo ==================================================
echo.

REM ---------- (re)build the click-to-open shortcut for other people ----------
REM A .url internet shortcut opens the site in the default browser. Regenerated
REM on every launch so it always carries the CURRENT IP. Hand this one file to
REM anyone on the network - they just double-click it, nothing to install.
set "SHORTCUT=Open PPT Designer.url"
> "%SHORTCUT%" echo [InternetShortcut]
>>"%SHORTCUT%" echo URL=http://%LAN_IP%:8002/
>>"%SHORTCUT%" echo IconIndex=0
echo   Shareable shortcut refreshed: "%SHORTCUT%"
echo.

REM ---------- warn if the frontend is pointing at a different IP ----------
REM The backend address is hard-coded in frontend\index.html. DHCP can change
REM this machine's IP, which would silently break every API call.
findstr /c:"const BACKEND_HOST = '%LAN_IP%';" "frontend\index.html" >nul 2>&1
if errorlevel 1 (
    echo  [!] WARNING - frontend\index.html is NOT pointing at %LAN_IP%
    echo      People on the network will not be able to reach the backend.
    echo      Fix it by editing frontend\index.html and setting:
    echo.
    echo          const BACKEND_HOST = '%LAN_IP%';
    echo.
)

REM ---------- start the backend, from the backend folder ----------
REM 'start' inherits this directory, so "cd backend" keeps paths space-safe.
start "PPT Designer - BACKEND (do not close)" cmd /k "cd backend && python app.py"

REM ---------- start the frontend static server ----------
start "PPT Designer - FRONTEND (do not close)" cmd /k "python -m http.server 8002 --bind 0.0.0.0 --directory frontend"

REM ---------- give the servers a moment to boot, then open the site ----------
REM 'ping' is used as the delay rather than 'timeout' because timeout needs a
REM real console and fails when this script is run non-interactively.
ping -n 5 127.0.0.1 >nul 2>&1
start "" "http://localhost:8002/"

echo  Both servers are now running, each in its own window.
echo.
echo  To STOP the app: close those two server windows.
echo.
pause
