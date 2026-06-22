@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\run_workhub_delivery_app.ps1"
if errorlevel 1 (
  echo.
  echo Workhub failed to start. Check workhub_run_error.log in this folder.
  pause
)
endlocal
