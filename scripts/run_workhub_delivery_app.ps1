$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Root = Split-Path -Parent $ScriptDir
$App = Join-Path $ScriptDir "workhub_delivery_app.py"
$Port = if ($env:WORKHUB_PORT) { [int]$env:WORKHUB_PORT } else { 8770 }
$Url = "http://127.0.0.1:$Port"
$LogPath = Join-Path $Root "workhub_run_error.log"

$PythonCandidates = @(
  (Join-Path $Root "runtime\python\python.exe"),
  (Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"),
  "python",
  "py"
)

function Write-RunLog($Message) {
  $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  Add-Content -LiteralPath $LogPath -Encoding UTF8 -Value "[$Timestamp] $Message"
}

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

  throw "Python executable was not found. Install Python 3.11+ or run this from Codex with the bundled Python runtime."
}

function Test-Workhub {
  try {
    $Response = Invoke-WebRequest -UseBasicParsing -Uri "$Url/login" -TimeoutSec 1
    return ($Response.StatusCode -eq 200 -and ($Response.Content -like "*소일브릿지*" -or $Response.Content -like "*로그인*"))
  } catch {
    return $false
  }
}

function Ensure-Requirements($Python) {
  $Requirements = Join-Path $Root "requirements.txt"
  if (-not (Test-Path -LiteralPath $Requirements)) {
    return
  }

  & $Python -c "import openpyxl" 2>$null
  if ($LASTEXITCODE -eq 0) {
    return
  }

  Write-Host "Installing required Python packages..."
  & $Python -m pip install -r $Requirements
  if ($LASTEXITCODE -ne 0) {
    throw "Required Python packages could not be installed. Run: python -m pip install -r requirements.txt"
  }
}

if (-not (Test-Path -LiteralPath $App)) {
  Write-RunLog "App file missing: $App"
  throw "Workhub app file was not found: $App"
}

if (-not (Test-Workhub)) {
  try {
    $Python = Find-Python
    Write-RunLog "Using Python: $Python"
    Ensure-Requirements $Python

    $Process = Start-Process `
      -FilePath $Python `
      -ArgumentList @("`"$App`"", "$Port") `
      -WorkingDirectory $Root `
      -WindowStyle Hidden `
      -PassThru
    Write-RunLog "Started Workhub process id: $($Process.Id), port: $Port"
  } catch {
    Write-RunLog "Startup failed: $($_.Exception.Message)"
    throw
  }

  $Ready = $false
  for ($i = 0; $i -lt 20; $i++) {
    Start-Sleep -Milliseconds 500
    if (Test-Workhub) {
      $Ready = $true
      break
    }
  }

  if (-not $Ready) {
    Write-RunLog "Workhub did not respond on $Url/login"
    throw "Workhub did not start in time. Wait a moment and run this file again."
  }
}

Write-RunLog "Opening $Url"
Start-Process $Url
