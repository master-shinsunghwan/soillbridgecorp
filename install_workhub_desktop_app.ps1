$ErrorActionPreference = "Stop"

$SourceRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$InstallRoot = Join-Path $env:LOCALAPPDATA "SoilbridgeWorkhubDesktop"
$ExeSource = Join-Path $SourceRoot "SoilbridgeWorkhub.exe"
$ExeTarget = Join-Path $InstallRoot "SoilbridgeWorkhub.exe"

if (-not (Test-Path -LiteralPath $ExeSource)) {
  throw "설치할 실행파일을 찾지 못했습니다: $ExeSource"
}

function Stop-RunningWorkhub {
  $Processes = @(Get-Process -Name "SoilbridgeWorkhub" -ErrorAction SilentlyContinue)
  if ($Processes.Count -eq 0) {
    return
  }

  $Processes | Stop-Process -Force -ErrorAction SilentlyContinue
  for ($Attempt = 0; $Attempt -lt 40; $Attempt += 1) {
    if (-not (Get-Process -Name "SoilbridgeWorkhub" -ErrorAction SilentlyContinue)) {
      return
    }
    Start-Sleep -Milliseconds 250
  }

  throw "실행 중인 업무자동화 앱을 종료하지 못했습니다. 앱을 닫은 뒤 다시 설치해주세요."
}

New-Item -ItemType Directory -Force -Path $InstallRoot | Out-Null
Stop-RunningWorkhub

for ($Attempt = 1; $Attempt -le 10; $Attempt += 1) {
  try {
    Copy-Item -LiteralPath $ExeSource -Destination $ExeTarget -Force
    break
  } catch {
    if ($Attempt -eq 10) {
      throw
    }
    Start-Sleep -Milliseconds 300
  }
}

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
$ShortcutName = "(주)소일브릿지 업무자동화.lnk"
New-WorkhubDesktopShortcut -Path (Join-Path $Desktop $ShortcutName)
New-WorkhubDesktopShortcut -Path (Join-Path $Programs $ShortcutName)

Write-Host "Install complete"
Write-Host "Install folder: $InstallRoot"
Write-Host "Desktop shortcut: $(Join-Path $Desktop $ShortcutName)"
Write-Host "Start menu shortcut: $(Join-Path $Programs $ShortcutName)"
