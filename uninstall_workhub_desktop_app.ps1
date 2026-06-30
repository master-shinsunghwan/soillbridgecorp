$ErrorActionPreference = "Stop"

$InstallRoot = Join-Path $env:LOCALAPPDATA "SoilbridgeWorkhubDesktop"
$ShortcutName = "(주)소일브릿지 업무자동화.lnk"
$DesktopShortcut = Join-Path ([Environment]::GetFolderPath("Desktop")) $ShortcutName
$StartShortcut = Join-Path ([Environment]::GetFolderPath("Programs")) $ShortcutName

if (Test-Path -LiteralPath $DesktopShortcut) {
  Remove-Item -LiteralPath $DesktopShortcut -Force
}
if (Test-Path -LiteralPath $StartShortcut) {
  Remove-Item -LiteralPath $StartShortcut -Force
}
if (Test-Path -LiteralPath $InstallRoot) {
  Remove-Item -LiteralPath $InstallRoot -Recurse -Force
}

Write-Host "Uninstall complete"
