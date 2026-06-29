@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall_workhub_login_watchdog.ps1"
if errorlevel 1 (
  echo.
  echo Workhub auto-open removal failed.
  echo Please check the message above.
  pause
  exit /b 1
)
echo.
echo Workhub auto-open removed.
pause
endlocal
