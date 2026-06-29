param(
  [int]$Port = 8781,
  [int]$IntervalMinutes = 5,
  [string]$Url = "https://workhub.soilbridgecorp.cloud",
  [int]$ChromeDebugPort = 9231
)

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Root = Split-Path -Parent $ScriptDir
$App = Join-Path $ScriptDir "workhub_delivery_app.py"
$UseLocalServer = [string]::IsNullOrWhiteSpace($Url)
$TargetUrl = if ($UseLocalServer) { "http://127.0.0.1:$Port" } else { $Url.TrimEnd("/") }
$LogPath = Join-Path $Root "workhub_login_watchdog.log"
$DataRoot = if ($env:WORKHUB_DATA_DIR) { $env:WORKHUB_DATA_DIR } else { $Root }
$DbPath = Join-Path $DataRoot "config\workhub.db"
$ChromeProfileDir = Join-Path $env:LOCALAPPDATA "SoillbridgeWorkhub\ChromeProfile"

$PythonCandidates = @(
  (Join-Path $Root "runtime\python\python.exe"),
  (Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"),
  "python",
  "py"
)

$ChromeCandidates = @(
  (Join-Path $env:ProgramFiles "Google\Chrome\Application\chrome.exe"),
  (Join-Path ${env:ProgramFiles(x86)} "Google\Chrome\Application\chrome.exe"),
  (Join-Path $env:LOCALAPPDATA "Google\Chrome\Application\chrome.exe"),
  "chrome.exe"
)

function Write-WatchLog {
  param([string]$Message)
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
  throw "Python executable was not found."
}

function Find-Chrome {
  foreach ($Candidate in $ChromeCandidates) {
    if ($Candidate -eq "chrome.exe") {
      $Command = Get-Command $Candidate -ErrorAction SilentlyContinue
      if ($Command) {
        return $Command.Source
      }
    } elseif ($Candidate -and (Test-Path -LiteralPath $Candidate)) {
      return $Candidate
    }
  }
  return ""
}

function Test-WorkhubServer {
  try {
    $Response = Invoke-WebRequest -UseBasicParsing -Uri "$TargetUrl/login" -TimeoutSec 2
    return ($Response.StatusCode -eq 200)
  } catch {
    return $false
  }
}

function Ensure-WorkhubServer {
  param([string]$Python)

  if (Test-WorkhubServer) {
    return
  }

  if (-not (Test-Path -LiteralPath $App)) {
    throw "Workhub app file was not found: $App"
  }

  Write-WatchLog "Starting Workhub local server on $TargetUrl"
  Start-Process `
    -FilePath $Python `
    -ArgumentList @("`"$App`"", "$Port") `
    -WorkingDirectory $Root `
    -WindowStyle Hidden | Out-Null

  for ($i = 0; $i -lt 120; $i++) {
    Start-Sleep -Milliseconds 500
    if (Test-WorkhubServer) {
      Write-WatchLog "Workhub local server is ready."
      return
    }
  }

  throw "Workhub did not start in time."
}

function Test-ActiveLoginSession {
  param([string]$Python)

  if (-not (Test-Path -LiteralPath $DbPath)) {
    return $false
  }

  $Code = @"
import sqlite3, time, sys
db = r'''$DbPath'''
try:
    con = sqlite3.connect(db)
    row = con.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='login_sessions'").fetchone()
    if not row:
        print("none")
        sys.exit(0)
    now = time.time()
    count = con.execute(
        "SELECT COUNT(*) FROM login_sessions WHERE expires_at > ? AND COALESCE(absolute_expires_at, expires_at) > ?",
        (now, now),
    ).fetchone()[0]
    print("active" if count else "none")
finally:
    try:
        con.close()
    except Exception:
        pass
"@

  try {
    $Result = & $Python -c $Code 2>$null
    return (($Result | Select-Object -Last 1) -eq "active")
  } catch {
    return $false
  }
}

function Test-RemoteWorkhubLoggedIn {
  if ($UseLocalServer) {
    return $false
  }

  try {
    $TargetHost = ([Uri]$TargetUrl).Host
    $Tabs = Invoke-RestMethod -UseBasicParsing -Uri "http://127.0.0.1:$ChromeDebugPort/json" -TimeoutSec 2
    foreach ($Tab in $Tabs) {
      $TabUrl = [string]$Tab.url
      if (-not $TabUrl) {
        continue
      }
      try {
        $Parsed = [Uri]$TabUrl
      } catch {
        continue
      }
      if ($Parsed.Host -eq $TargetHost -and -not $Parsed.AbsolutePath.StartsWith("/login")) {
        return $true
      }
    }
  } catch {
    return $false
  }

  return $false
}

function Open-WorkhubLogin {
  $Chrome = Find-Chrome
  if (-not $Chrome) {
    Write-WatchLog "Chrome was not found. Workhub was not opened."
    return
  }
  Write-WatchLog "Opening Workhub because no active login session was found: $TargetUrl"
  if ($UseLocalServer) {
    Start-Process -FilePath $Chrome -ArgumentList @("--new-window", $TargetUrl) | Out-Null
    return
  }

  New-Item -ItemType Directory -Path $ChromeProfileDir -Force | Out-Null
  Start-Process -FilePath $Chrome -ArgumentList @(
    "--remote-debugging-port=$ChromeDebugPort",
    "--user-data-dir=$ChromeProfileDir",
    "--new-window",
    $TargetUrl
  ) | Out-Null
}

$Python = if ($UseLocalServer) { Find-Python } else { "" }
Write-WatchLog "Watchdog started. Url=$TargetUrl IntervalMinutes=$IntervalMinutes Db=$DbPath Python=$Python LocalServer=$UseLocalServer ChromeDebugPort=$ChromeDebugPort"

while ($true) {
  try {
    if ($UseLocalServer) {
      Ensure-WorkhubServer -Python $Python
      if (-not (Test-ActiveLoginSession -Python $Python)) {
        Open-WorkhubLogin
      }
    } else {
      if (-not (Test-RemoteWorkhubLoggedIn)) {
        Open-WorkhubLogin
      } else {
        Write-WatchLog "Workhub is already logged in on this PC."
      }
    }
  } catch {
    Write-WatchLog "Watchdog error: $($_.Exception.Message)"
  }

  Start-Sleep -Seconds ([Math]::Max(60, $IntervalMinutes * 60))
}
