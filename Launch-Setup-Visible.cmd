@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "PS1=%SCRIPT_DIR%LabDash-Setup-Host.ps1"

rem Use Windows PowerShell for reliable hidden window
set "PSLEGACY=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

"%PSLEGACY%" -NoLogo -NoProfile -WindowStyle Hidden -Command ^
  Start-Process -WindowStyle Hidden -FilePath "%PSLEGACY%" -ArgumentList '-NoProfile -Sta -ExecutionPolicy Bypass -File "%PS1%" -NoPrefill'

exit /b 0
