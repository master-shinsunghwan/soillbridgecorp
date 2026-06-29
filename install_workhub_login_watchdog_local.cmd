@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install_workhub_login_watchdog.ps1"
if errorlevel 1 (
  echo.
  echo Workhub local auto-open setup failed.
  echo Please check the message above.
  pause
  exit /b 1
)
echo.
echo Workhub local auto-open setup complete.
pause
endlocal
