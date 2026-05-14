@echo off
setlocal
set "SCRIPT=%~dp0Start-Assessment.ps1"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" %*
if %ERRORLEVEL% neq 0 pause
endlocal
