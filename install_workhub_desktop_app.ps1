$ErrorActionPreference = "Stop"

$SourceRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$InstallRoot = Join-Path $env:LOCALAPPDATA "SoilbridgeWorkhubDesktop"
$ExeSource = Join-Path $SourceRoot "SoilbridgeWorkhub.exe"
$ExeTarget = Join-Path $InstallRoot "SoilbridgeWorkhub.exe"

if (-not (Test-Path -LiteralPath $ExeSource)) {
  throw "설치할 실행파일을 찾지 못했습니다: $ExeSource"
}

New-Item -ItemType Directory -Force -Path $InstallRoot | Out-Null
Copy-Item -LiteralPath $ExeSource -Destination $ExeTarget -Force

$Shell = New-Object -ComObject WScript.Shell
function New-WorkhubDesktopShortcut {
  param(
    [Parameter(Mandatory = $true)][string]$Path
  )

  $Shortcut = $Shell.CreateShortcut($Path)
  $Shortcut.TargetPath = $ExeTarget
  $Shortcut.WorkingDirectory = $InstallRoot
  $Shortcut.Description = "(주)소일브릿지 업무자동화 PC 앱"
  $Shortcut.IconLocation = $ExeTarget
  $Shortcut.Save()
}

$Desktop = [Environment]::GetFolderPath("Desktop")
$Programs = [Environment]::GetFolderPath("Programs")
$Startup = [Environment]::GetFolderPath("Startup")
$ShortcutName = "(주)소일브릿지 업무자동화.lnk"
New-WorkhubDesktopShortcut -Path (Join-Path $Desktop $ShortcutName)
New-WorkhubDesktopShortcut -Path (Join-Path $Programs $ShortcutName)

$StartupScript = Join-Path $Startup "SoilbridgeWorkhubDesktop_AutoStart.vbs"
$StartupCommand = ('"' + $ExeTarget + '"').Replace('"', '""')
$StartupContent = @"
Set shell = CreateObject("WScript.Shell")
shell.Run "$StartupCommand", 1, False
"@
Set-Content -LiteralPath $StartupScript -Value $StartupContent -Encoding Unicode

Write-Host "Install complete"
Write-Host "Install folder: $InstallRoot"
Write-Host "Desktop shortcut: $(Join-Path $Desktop $ShortcutName)"
Write-Host "Start menu shortcut: $(Join-Path $Programs $ShortcutName)"
Write-Host "Startup launcher: $StartupScript"
