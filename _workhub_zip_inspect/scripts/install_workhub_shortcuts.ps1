$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Root = Split-Path -Parent $ScriptDir
$Launcher = Join-Path $Root "workhub_run.cmd"

if (-not (Test-Path -LiteralPath $Launcher)) {
  throw "실행 파일을 찾지 못했습니다: $Launcher"
}

$Shell = New-Object -ComObject WScript.Shell

function New-Shortcut {
  param(
    [Parameter(Mandatory = $true)][string]$Path
  )

  $Shortcut = $Shell.CreateShortcut($Path)
  $Shortcut.TargetPath = $Launcher
  $Shortcut.WorkingDirectory = $Root
  $Shortcut.Description = "Workhub order automation"
  $Shortcut.IconLocation = "${env:SystemRoot}\System32\SHELL32.dll,46"
  $Shortcut.Save()
}

$Desktop = [Environment]::GetFolderPath("Desktop")
$StartMenuPrograms = [Environment]::GetFolderPath("Programs")

New-Shortcut -Path (Join-Path $Desktop "Workhub.lnk")
New-Shortcut -Path (Join-Path $StartMenuPrograms "Workhub.lnk")

Write-Host "Install complete"
Write-Host "Desktop shortcut: $(Join-Path $Desktop 'Workhub.lnk')"
Write-Host "Start menu shortcut: $(Join-Path $StartMenuPrograms 'Workhub.lnk')"
