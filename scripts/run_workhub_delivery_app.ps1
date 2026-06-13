$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Root = Split-Path -Parent $ScriptDir
$App = Join-Path $ScriptDir "workhub_delivery_app.py"
$Port = 8765
$Url = "http://127.0.0.1:$Port"

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

  throw "Python 실행 파일을 찾지 못했습니다."
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
  throw "업무허브 프로그램 파일을 찾지 못했습니다: $App"
}

if (-not (Test-Workhub)) {
  $Python = Find-Python
  Start-Process `
    -FilePath $Python `
    -ArgumentList @("`"$App`"", "$Port") `
    -WorkingDirectory $Root `
    -WindowStyle Hidden

  $Ready = $false
  for ($i = 0; $i -lt 20; $i++) {
    Start-Sleep -Milliseconds 500
    if (Test-Workhub) {
      $Ready = $true
      break
    }
  }

  if (-not $Ready) {
    throw "업무허브를 시작하지 못했습니다. 잠시 후 다시 실행해주세요."
  }
}

Start-Process $Url
