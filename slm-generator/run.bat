@echo off
setlocal
title PPSU SLM Generator - launcher
cd /d "%~dp0"

REM ---------- venv: create + install on first run ----------
if not exist "backend\venv\Scripts\python.exe" (
    echo First run: creating the Python environment...
    where python >nul 2>&1
    if errorlevel 1 (
        echo  [X] Python was not found. Install Python 3.11+ and tick
        echo      "Add python.exe to PATH", then run this again.
        pause
        exit /b 1
    )
    python -m venv backend\venv
    backend\venv\Scripts\python.exe -m pip install -r requirements.txt
)

REM ---------- Ollama must be running with the model pulled ----------
backend\venv\Scripts\python.exe -c "import sys;sys.path.insert(0,'backend');from ai_engine import get_engine;up,m=get_engine().available();sys.exit(0 if (up and m) else (2 if up else 1))"
if errorlevel 2 (
    echo  [!] Ollama is running but the model is not pulled yet.
    echo      Pulling qwen2.5:7b-instruct ^(about 4.7 GB, one-time^)...
    ollama pull qwen2.5:7b-instruct
) else if errorlevel 1 (
    echo  [X] Ollama is not running. Install it from https://ollama.com,
    echo      or start it from the Start Menu, then run this again.
    pause
    exit /b 1
)

REM ---------- LAN address for the shareable link ----------
set "LAN_IP="
for /f "delims=" %%i in ('backend\venv\Scripts\python.exe -c "import socket;s=socket.socket(socket.AF_INET,socket.SOCK_DGRAM);s.connect(('8.8.8.8',80));print(s.getsockname()[0])" 2^>nul') do set "LAN_IP=%%i"
if "%LAN_IP%"=="" set "LAN_IP=127.0.0.1"

cls
echo ==================================================
echo         P P S U   S L M   G E N E R A T O R
echo ==================================================
echo.
echo   On this computer : http://localhost:8010/
echo   Share on network : http://%LAN_IP%:8010/
echo.
echo   One server, one port - the page talks to the API
echo   on the same address, so the shared link needs no
echo   configuration even when this machine's IP changes.
echo ==================================================
echo.

REM ---------- click-to-open shortcut for other people ----------
set "SHORTCUT=Open SLM Generator.url"
> "%SHORTCUT%" echo [InternetShortcut]
>>"%SHORTCUT%" echo URL=http://%LAN_IP%:8010/
>>"%SHORTCUT%" echo IconIndex=0
echo   Shareable shortcut refreshed: "%SHORTCUT%"
echo.

start "PPSU SLM Generator - SERVER (do not close)" cmd /k "cd backend && venv\Scripts\python.exe -m uvicorn app:app --host 0.0.0.0 --port 8010"

ping -n 4 127.0.0.1 >nul 2>&1
start "" "http://localhost:8010/"

echo  Server running in its own window. To STOP: close that window.
echo.
pause
