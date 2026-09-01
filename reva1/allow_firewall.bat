@echo off
title PPT Designer - allow through Windows Firewall

REM ----------------------------------------------------------------
REM  RIGHT-CLICK THIS FILE  ->  "Run as administrator"
REM
REM  Opens ports 8000 (website) and 5000 (backend) for INBOUND traffic
REM  so other computers on your network can reach the app.
REM  Scoped to the local subnet only, so the ports are not exposed
REM  more widely than they need to be.
REM ----------------------------------------------------------------

REM --- confirm we are elevated; adding rules silently fails otherwise ---
net session >nul 2>&1
if errorlevel 1 (
    echo.
    echo  [X] This must be run as Administrator.
    echo.
    echo      Close this window, then RIGHT-CLICK "allow_firewall.bat"
    echo      and choose "Run as administrator".
    echo.
    pause
    exit /b 1
)

echo.
echo  Adding Windows Firewall rules...
echo.

REM remove any earlier copies so re-running never piles up duplicates
netsh advfirewall firewall delete rule name="PPT Designer frontend (TCP 8000)" >nul 2>&1
netsh advfirewall firewall delete rule name="PPT Designer backend (TCP 5000)"  >nul 2>&1

netsh advfirewall firewall add rule name="PPT Designer frontend (TCP 8000)" dir=in action=allow protocol=TCP localport=8000 profile=domain,private remoteip=localsubnet
netsh advfirewall firewall add rule name="PPT Designer backend (TCP 5000)"  dir=in action=allow protocol=TCP localport=5000 profile=domain,private remoteip=localsubnet

echo.
echo  ================================================
echo   Done. Other computers on your network can now
echo   reach the app while it is running.
echo  ================================================
echo.
echo  Verifying the rules exist:
echo.
netsh advfirewall firewall show rule name="PPT Designer frontend (TCP 8000)" | findstr /c:"Rule Name" /c:"Enabled" /c:"Action" /c:"Profiles"
netsh advfirewall firewall show rule name="PPT Designer backend (TCP 5000)"  | findstr /c:"Rule Name" /c:"Enabled" /c:"Action" /c:"Profiles"
echo.
pause
