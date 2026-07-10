$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$PythonCandidates = @(
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
$DistDir = Join-Path $BuildDir "workhub_exe_dist"
$WorkDir = Join-Path $BuildDir "workhub_exe_build"
$PackageDir = Join-Path $BuildDir "workhub_package"

New-Item -ItemType Directory -Force -Path $DistDir, $WorkDir, $PackageDir | Out-Null

$AddData = @(
  "--add-data", "$(Join-Path $Root 'templates');templates"
)

$LucideDir = Join-Path $Root "node_modules\lucide"
if (Test-Path -LiteralPath $LucideDir) {
  $AddData += @("--add-data", "$LucideDir;node_modules\lucide")
}

& $Python -m PyInstaller `
  --clean `
  --noconfirm `
  --onefile `
  --windowed `
  --name "Workhub" `
  --paths (Join-Path $Root "scripts") `
  --distpath $DistDir `
  --workpath $WorkDir `
  @AddData `
  (Join-Path $Root "scripts\workhub_desktop_launcher.py")

Copy-Item -LiteralPath (Join-Path $DistDir "Workhub.exe") -Destination (Join-Path $PackageDir "Workhub.exe") -Force

$Usage = @"
Workhub 사용 방법

1. Workhub.exe를 실행합니다.
2. 한 번 실행하면 현재 사용자 시작프로그램에 자동 등록되어 다음 PC 부팅부터 자동 실행됩니다.
3. 독립 프로그램 창이 열리면 (주)소일브릿지 발주 업무자동화를 사용합니다.
4. 다른 PC에서도 Python 설치 없이 Workhub.exe만 실행하면 됩니다.

참고
- 첫 실행은 내부 파일을 준비하느라 몇 초 걸릴 수 있습니다.
- 결과 파일은 프로그램 창에서 다운로드됩니다.
- 포함 기능: 개별 택배건 정리, 송장번호 추출, 롯데택배 발주서 변환, 차량인수증 생성, 업체 CS 요청, CS 처리대장
- 프로그램을 완전히 종료하려면 작업 관리자에서 Workhub.exe를 종료하거나, 함께 들어있는 "Workhub 종료.cmd"를 실행하세요.
"@
Set-Content -LiteralPath (Join-Path $PackageDir "README.txt") -Value $Usage -Encoding UTF8

$StopCmd = @"
@echo off
taskkill /IM Workhub.exe /F >nul 2>nul
echo Workhub has been closed.
pause
"@
Set-Content -LiteralPath (Join-Path $PackageDir "Workhub_Stop.cmd") -Value $StopCmd -Encoding ASCII

$Stamp = Get-Date -Format "yyyyMMdd_HHmm"
$ZipPath = Join-Path (Join-Path $Root "output") "Workhub_PC_Desktop_$Stamp.zip"
New-Item -ItemType Directory -Force -Path (Join-Path $Root "output") | Out-Null
Compress-Archive -LiteralPath `
  (Join-Path $PackageDir "Workhub.exe"), `
  (Join-Path $PackageDir "Workhub_Stop.cmd"), `
  (Join-Path $PackageDir "README.txt") `
  -DestinationPath $ZipPath `
  -Force

Write-Host "Build complete"
Write-Host "Package folder: $PackageDir"
Write-Host "Zip file: $ZipPath"
