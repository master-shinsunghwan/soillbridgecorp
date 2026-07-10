$ErrorActionPreference = "Stop"

$SourceRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$InstallRoot = Join-Path $env:LOCALAPPDATA "Workhub"
$InstallScripts = Join-Path $InstallRoot "scripts"
$InstallTemplates = Join-Path $InstallRoot "templates"
$InstallNodeModules = Join-Path $InstallRoot "node_modules"
$InstallLucide = Join-Path $InstallNodeModules "lucide"
$Launcher = Join-Path $InstallRoot "workhub_run.cmd"

$Python = Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if (-not (Test-Path -LiteralPath $Python)) {
  $PythonCommand = Get-Command python -ErrorAction SilentlyContinue
  if ($PythonCommand) {
    $Python = $PythonCommand.Source
  } else {
    throw "Python 실행 파일을 찾지 못했습니다."
  }
}

$RequiredScripts = @(
  "delivery_text_summary.py",
  "invoice_number_exporter.py",
  "lotte_order_form_converter.py",
  "vehicle_receipt_generator.py",
  "workhub_delivery_app.py",
  "workhub_desktop_launcher.py",
  "run_workhub_delivery_app.ps1",
  "workhub_delivery_app_README.md"
)

New-Item -ItemType Directory -Force -Path $InstallScripts | Out-Null
New-Item -ItemType Directory -Force -Path $InstallTemplates | Out-Null
New-Item -ItemType Directory -Force -Path $InstallNodeModules | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $InstallRoot "output") | Out-Null

foreach ($ScriptName in $RequiredScripts) {
  $Source = Join-Path (Join-Path $SourceRoot "scripts") $ScriptName
  if (-not (Test-Path -LiteralPath $Source)) {
    throw "필수 파일을 찾지 못했습니다: $Source"
  }
  Copy-Item -LiteralPath $Source -Destination (Join-Path $InstallScripts $ScriptName) -Force
}

& $Python -m pip install -r (Join-Path $SourceRoot "requirements.txt") -q

$TemplateSource = Join-Path (Join-Path $SourceRoot "templates") "vehicle_receipt_template.xlsx"
if (-not (Test-Path -LiteralPath $TemplateSource)) {
  throw "차량인수증 템플릿을 찾지 못했습니다: $TemplateSource"
}
Copy-Item -LiteralPath $TemplateSource -Destination (Join-Path $InstallTemplates "vehicle_receipt_template.xlsx") -Force

$LucideSource = Join-Path (Join-Path $SourceRoot "node_modules") "lucide"
if (Test-Path -LiteralPath $LucideSource) {
  if (Test-Path -LiteralPath $InstallLucide) {
    Remove-Item -LiteralPath $InstallLucide -Recurse -Force
  }
  Copy-Item -LiteralPath $LucideSource -Destination $InstallLucide -Recurse -Force
}

$LauncherContent = @"
@echo off
setlocal
"$Python" "%~dp0scripts\workhub_desktop_launcher.py"
endlocal
"@
Set-Content -LiteralPath $Launcher -Value $LauncherContent -Encoding ASCII

& $Python -m py_compile `
  (Join-Path $InstallScripts "delivery_text_summary.py") `
  (Join-Path $InstallScripts "invoice_number_exporter.py") `
  (Join-Path $InstallScripts "lotte_order_form_converter.py") `
  (Join-Path $InstallScripts "vehicle_receipt_generator.py") `
  (Join-Path $InstallScripts "workhub_delivery_app.py")

$Shell = New-Object -ComObject WScript.Shell
function New-WorkhubShortcut {
  param(
    [Parameter(Mandatory = $true)][string]$Path
  )

  $Shortcut = $Shell.CreateShortcut($Path)
  $Shortcut.TargetPath = $Launcher
  $Shortcut.WorkingDirectory = $InstallRoot
  $Shortcut.Description = "Workhub order automation"
  $Shortcut.IconLocation = "${env:SystemRoot}\System32\SHELL32.dll,46"
  $Shortcut.Save()
}

$Desktop = [Environment]::GetFolderPath("Desktop")
$Programs = [Environment]::GetFolderPath("Programs")
$Startup = [Environment]::GetFolderPath("Startup")
New-WorkhubShortcut -Path (Join-Path $Desktop "Workhub.lnk")
New-WorkhubShortcut -Path (Join-Path $Programs "Workhub.lnk")

$StartupScript = Join-Path $Startup "Workhub_AutoStart.vbs"
$StartupCommand = ('"' + $Launcher + '"').Replace('"', '""')
$StartupContent = @"
Set shell = CreateObject("WScript.Shell")
shell.Run "$StartupCommand", 1, False
"@
Set-Content -LiteralPath $StartupScript -Value $StartupContent -Encoding Unicode

Write-Host "Install complete"
Write-Host "Install folder: $InstallRoot"
Write-Host "Desktop shortcut: $(Join-Path $Desktop 'Workhub.lnk')"
Write-Host "Start menu shortcut: $(Join-Path $Programs 'Workhub.lnk')"
Write-Host "Startup launcher: $StartupScript"
