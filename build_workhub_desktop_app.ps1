$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$PythonCandidates = @(
  (Join-Path $Root ".venv\Scripts\python.exe"),
  (Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"),
  "python",
  "py"
)

function Find-Python {
  foreach ($Candidate in $PythonCandidates) {
    if ($Candidate -eq "python" -or $Candidate -eq "py") {
      $Command = Get-Command $Candidate -ErrorAction SilentlyContinue
      if ($Command) {
        return $Command.Source
      }
    } elseif (Test-Path -LiteralPath $Candidate) {
      return $Candidate
    }
  }

  throw "Python 실행 파일을 찾지 못했습니다. Python 3.11 이상을 설치한 뒤 다시 실행해주세요."
}

$Python = Find-Python
& $Python -m pip install --upgrade pyinstaller -q
& $Python -m pip install -r (Join-Path $Root "requirements.txt") -q

$BuildDir = Join-Path $Root ".build"
$DistDir = Join-Path $BuildDir "workhub_desktop_dist"
$WorkDir = Join-Path $BuildDir "workhub_desktop_build"
$SpecDir = Join-Path $BuildDir "workhub_desktop_spec"
$PackageDir = Join-Path $BuildDir "workhub_desktop_package"

New-Item -ItemType Directory -Force -Path $DistDir, $WorkDir, $SpecDir, $PackageDir | Out-Null

& $Python -m PyInstaller `
  --clean `
  --noconfirm `
  --onefile `
  --windowed `
  --name "SoilbridgeWorkhub" `
  --distpath $DistDir `
  --workpath $WorkDir `
  --specpath $SpecDir `
  (Join-Path $Root "scripts\workhub_vps_desktop_app.py")

$ExeSource = Join-Path $DistDir "SoilbridgeWorkhub.exe"
$ExeTarget = Join-Path $PackageDir "SoilbridgeWorkhub.exe"
Copy-Item -LiteralPath $ExeSource -Destination $ExeTarget -Force
Copy-Item -LiteralPath (Join-Path $Root "install_workhub_desktop_app.ps1") -Destination (Join-Path $PackageDir "install_workhub_desktop_app.ps1") -Force
Copy-Item -LiteralPath (Join-Path $Root "uninstall_workhub_desktop_app.ps1") -Destination (Join-Path $PackageDir "uninstall_workhub_desktop_app.ps1") -Force

$InstallCmd = @"
@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install_workhub_desktop_app.ps1"
pause
"@
Set-Content -LiteralPath (Join-Path $PackageDir "Install.cmd") -Value $InstallCmd -Encoding ASCII

$UninstallCmd = @"
@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall_workhub_desktop_app.ps1"
pause
"@
Set-Content -LiteralPath (Join-Path $PackageDir "Uninstall.cmd") -Value $UninstallCmd -Encoding ASCII

$Readme = @"
(주)소일브릿지 업무자동화 PC 앱

실행 방법
1. SoilbridgeWorkhub.exe를 더블클릭하면 주소창 없는 업무자동화 앱 창이 열립니다.
2. 한 번 실행하면 현재 사용자 시작프로그램에 자동 등록되어 다음 PC 부팅부터 자동 실행됩니다.
3. 설치해서 쓰려면 'Install.cmd'를 실행합니다.
4. 설치 후 바탕화면 또는 시작 메뉴의 '(주)소일브릿지 업무자동화' 바로가기를 실행합니다.

접속 대상
- https://workhub.soilbridgecorp.cloud/

참고
- 이 앱은 Windows WebView2 기반 앱 창으로 Workhub VPS를 띄웁니다.
- 로그인 세션은 PC 앱 안에 저장됩니다.
- 코드서명 인증서가 없으면 Windows가 첫 실행 때 알 수 없는 앱 경고를 띄울 수 있습니다.
"@
Set-Content -LiteralPath (Join-Path $PackageDir "README.txt") -Value $Readme -Encoding UTF8

$Stamp = Get-Date -Format "yyyyMMdd_HHmm"
$ZipPath = Join-Path (Join-Path $Root "output") "SoilbridgeWorkhub_Desktop_$Stamp.zip"
New-Item -ItemType Directory -Force -Path (Join-Path $Root "output") | Out-Null
Compress-Archive -Path (Join-Path $PackageDir "*") -DestinationPath $ZipPath -Force

Write-Host "Build complete"
Write-Host "Package folder: $PackageDir"
Write-Host "Zip file: $ZipPath"
