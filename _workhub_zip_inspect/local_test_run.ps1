$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$App = Join-Path $Root "scripts\workhub_delivery_app.py"
$Port = 8765
$Url = "http://127.0.0.1:$Port"
$DataDir = Join-Path $Root "local_data"
$VenvDir = Join-Path $Root ".venv"
$VenvPython = Join-Path $VenvDir "Scripts\python.exe"

function Find-Python {
  $Candidates = @("python", "py")
  foreach ($Candidate in $Candidates) {
    $Command = Get-Command $Candidate -ErrorAction SilentlyContinue
    if ($Command) {
      return $Command.Source
    }
  }
  throw "Python 3.10 이상을 찾지 못했습니다. Python 설치 후 다시 실행해주세요."
}

function Test-Workhub {
  try {
    $Response = Invoke-WebRequest -UseBasicParsing -Uri $Url -TimeoutSec 1
    return ($Response.StatusCode -eq 200)
  } catch {
    return $false
  }
}

if (-not (Test-Path -LiteralPath $App)) {
  throw "업무허브 앱 파일을 찾지 못했습니다: $App"
}

New-Item -ItemType Directory -Path $DataDir -Force | Out-Null
$env:WORKHUB_DATA_DIR = $DataDir

if (-not (Test-Path -LiteralPath $VenvPython)) {
  $SystemPython = Find-Python
  & $SystemPython -m venv $VenvDir
}

& $VenvPython -m pip install --upgrade pip
& $VenvPython -m pip install -r (Join-Path $Root "requirements.txt")

if (-not (Test-Workhub)) {
  Start-Process `
    -FilePath $VenvPython `
    -ArgumentList @("`"$App`"", "$Port") `
    -WorkingDirectory $Root `
    -WindowStyle Hidden

  $Ready = $false
  for ($i = 0; $i -lt 30; $i++) {
    Start-Sleep -Milliseconds 500
    if (Test-Workhub) {
      $Ready = $true
      break
    }
  }

  if (-not $Ready) {
    throw "업무허브를 시작하지 못했습니다. PowerShell 창의 오류 메시지를 확인해주세요."
  }
}

Start-Process $Url
Write-Host "업무허브 로컬 테스트 실행 완료: $Url"
Write-Host "테스트 데이터 저장 위치: $DataDir"
