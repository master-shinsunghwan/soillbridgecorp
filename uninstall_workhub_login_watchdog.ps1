$ErrorActionPreference = "Stop"

$TaskName = "Soillbridge Workhub Login Watchdog"
$Task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
$StartupLauncher = Join-Path ([Environment]::GetFolderPath("Startup")) "Soillbridge Workhub Login Watchdog.vbs"
$StartupLauncherCmd = Join-Path ([Environment]::GetFolderPath("Startup")) "Soillbridge Workhub Login Watchdog.cmd"

if ($Task) {
  Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
  Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
  Write-Host "Workhub login watchdog removed."
} else {
  Write-Host "Workhub login watchdog is not installed."
}

if (Test-Path -LiteralPath $StartupLauncher) {
  Remove-Item -LiteralPath $StartupLauncher -Force
  Write-Host "Startup launcher removed."
}

if (Test-Path -LiteralPath $StartupLauncherCmd) {
  Remove-Item -LiteralPath $StartupLauncherCmd -Force
  Write-Host "Startup launcher command removed."
}
