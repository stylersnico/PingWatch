@echo off
REM Double-click launcher. Edit the two IPs below, or pass your own arguments:
REM   PingWatch.bat -Target1 10.0.0.1 -Target2 1.1.1.1
if "%~1"=="" (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0PingWatch.ps1" -Target1 192.168.1.1 -Target2 8.8.8.8 -DurationMinutes 75
) else (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0PingWatch.ps1" %*
)
pause
