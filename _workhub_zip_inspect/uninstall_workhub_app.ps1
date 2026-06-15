$ErrorActionPreference = "Stop"

$InstallRoot = Join-Path $env:LOCALAPPDATA "Workhub"
$DesktopShortcut = Join-Path ([Environment]::GetFolderPath("Desktop")) "Workhub.lnk"
$StartShortcut = Join-Path ([Environment]::GetFolderPath("Programs")) "Workhub.lnk"

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
