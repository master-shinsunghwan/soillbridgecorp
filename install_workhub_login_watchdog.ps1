param(
  [int]$Port = 8781,
  [int]$IntervalMinutes = 5,
  [string]$Url = "https://workhub.soilbridgecorp.cloud",
  [int]$ChromeDebugPort = 9231
)

$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Watchdog = Join-Path $Root "scripts\workhub_login_watchdog.ps1"
$TaskName = "Soillbridge Workhub Login Watchdog"
$StartupDir = [Environment]::GetFolderPath("Startup")
$StartupLauncher = Join-Path $StartupDir "Soillbridge Workhub Login Watchdog.vbs"
$OldStartupLauncherCmd = Join-Path $StartupDir "Soillbridge Workhub Login Watchdog.cmd"
$LocalLauncherDir = Join-Path $env:LOCALAPPDATA "SoillbridgeWorkhub"
$LocalLauncherCmd = Join-Path $LocalLauncherDir "Workhub Login Watchdog.cmd"

if (-not (Test-Path -LiteralPath $Watchdog)) {
  throw "Watchdog script was not found: $Watchdog"
}

Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" -ErrorAction SilentlyContinue |
  Where-Object { $_.ProcessId -ne $PID -and $_.CommandLine -like "*workhub_login_watchdog.ps1*" } |
  ForEach-Object {
    Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
  }

$PowerShell = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
$Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$Watchdog`" -Port $Port -IntervalMinutes $IntervalMinutes -ChromeDebugPort $ChromeDebugPort"
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

$InstallMode = "scheduled task"
try {
  Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $Action `
    -Trigger $Trigger `
    -Settings $Settings `
    -Description "Open Workhub on PC login and remind every 5 minutes until this PC has an active login." `
    -Force | Out-Null

  Remove-Item -LiteralPath $StartupLauncher -Force -ErrorAction SilentlyContinue
  Remove-Item -LiteralPath $OldStartupLauncherCmd -Force -ErrorAction SilentlyContinue
  Start-ScheduledTask -TaskName $TaskName
} catch {
  $InstallMode = "startup folder"
  New-Item -ItemType Directory -Path $LocalLauncherDir -Force | Out-Null
  $CmdLines = @(
    "@echo off",
    "cd /d `"$Root`"",
    "`"$PowerShell`" -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$Watchdog`" -Port $Port -IntervalMinutes $IntervalMinutes -ChromeDebugPort $ChromeDebugPort -Url `"$Url`""
  )
  Set-Content -LiteralPath $LocalLauncherCmd -Encoding ASCII -Value $CmdLines
  $VbsLines = @(
    "Set WshShell = CreateObject(""WScript.Shell"")",
    "WshShell.Run Chr(34) & ""$LocalLauncherCmd"" & Chr(34), 0, False"
  )
  Set-Content -LiteralPath $StartupLauncher -Encoding ASCII -Value $VbsLines
  Remove-Item -LiteralPath $OldStartupLauncherCmd -Force -ErrorAction SilentlyContinue

  Start-Process `
    -FilePath $PowerShell `
    -ArgumentList @(
      "-NoProfile",
      "-ExecutionPolicy", "Bypass",
      "-WindowStyle", "Hidden",
      "-File", "`"$Watchdog`"",
      "-Port", "$Port",
      "-IntervalMinutes", "$IntervalMinutes",
      "-ChromeDebugPort", "$ChromeDebugPort",
      "-Url", "`"$Url`""
    ) `
    -WorkingDirectory $Root `
    -WindowStyle Hidden | Out-Null
}

Write-Host "Workhub login watchdog installed and started."
Write-Host "Install mode: $InstallMode"
Write-Host "Task name: $TaskName"
Write-Host "Target URL: $Url"
Write-Host "Reminder interval: $IntervalMinutes minutes"
Write-Host "Chrome debug port: $ChromeDebugPort"
