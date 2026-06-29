$ErrorActionPreference = "Stop"

$TaskName = "Soillbridge Workhub Login Watchdog"
$Task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue

if ($Task) {
  Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
  Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
  Write-Host "Workhub login watchdog removed."
} else {
  Write-Host "Workhub login watchdog is not installed."
}
