param(
  [int]$Port = 8781,
  [int]$IntervalMinutes = 5,
  [string]$Url = ""
)

$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Watchdog = Join-Path $Root "scripts\workhub_login_watchdog.ps1"
$TaskName = "Soillbridge Workhub Login Watchdog"

if (-not (Test-Path -LiteralPath $Watchdog)) {
  throw "Watchdog script was not found: $Watchdog"
}

$PowerShell = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
$Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$Watchdog`" -Port $Port -IntervalMinutes $IntervalMinutes"
if (-not [string]::IsNullOrWhiteSpace($Url)) {
  $Arguments += " -Url `"$Url`""
}

$Action = New-ScheduledTaskAction -Execute $PowerShell -Argument $Arguments -WorkingDirectory $Root
$Trigger = New-ScheduledTaskTrigger -AtLogOn
$Settings = New-ScheduledTaskSettingsSet `
  -AllowStartIfOnBatteries `
  -DontStopIfGoingOnBatteries `
  -MultipleInstances IgnoreNew `
  -ExecutionTimeLimit (New-TimeSpan -Days 7)

Register-ScheduledTask `
  -TaskName $TaskName `
  -Action $Action `
  -Trigger $Trigger `
  -Settings $Settings `
  -Description "Open local Workhub on PC login and remind every 5 minutes until a login session exists." `
  -Force | Out-Null

Start-ScheduledTask -TaskName $TaskName

Write-Host "Workhub login watchdog installed and started."
Write-Host "Task name: $TaskName"
if ([string]::IsNullOrWhiteSpace($Url)) {
  Write-Host "Local URL: http://127.0.0.1:$Port"
} else {
  Write-Host "Target URL: $Url"
}
Write-Host "Reminder interval: $IntervalMinutes minutes"
