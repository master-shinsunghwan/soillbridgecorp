param(
  [Parameter(Mandatory = $true)][ValidatePattern("^\d+\.\d+\.\d+$")][string]$Version,
  [string]$Notes = "Soilbridge Workhub desktop update",
  [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$LauncherPath = Join-Path $Root "scripts\workhub_vps_desktop_app.py"
$BuildScript = Join-Path $Root "build_workhub_desktop_app.ps1"
$ExePath = Join-Path $Root ".build\workhub_desktop_dist\SoilbridgeWorkhub.exe"
$ManifestPath = Join-Path $Root "static\desktop_update.json"
$OutputDir = Join-Path $Root "output"
$Tag = "desktop-v$Version"
$Repository = "master-shinsunghwan/soillbridgecorp"
$AssetName = "SoilbridgeWorkhub.exe"

$LauncherSource = [IO.File]::ReadAllText($LauncherPath)
$VersionMatch = [regex]::Match($LauncherSource, 'APP_VERSION\s*=\s*"([^\"]+)"')
if (-not $VersionMatch.Success) {
  throw "APP_VERSION was not found in the desktop launcher."
}
if ($VersionMatch.Groups[1].Value -ne $Version) {
  throw "APP_VERSION ($($VersionMatch.Groups[1].Value)) does not match requested version ($Version)."
}

if (-not $SkipBuild) {
  & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $BuildScript
}
if (-not (Test-Path -LiteralPath $ExePath)) {
  throw "Desktop executable was not created: $ExePath"
}

$Hash = (Get-FileHash -Algorithm SHA256 -LiteralPath $ExePath).Hash.ToLowerInvariant()
$Size = (Get-Item -LiteralPath $ExePath).Length
$DownloadUrl = "https://github.com/$Repository/releases/download/$Tag/$AssetName"

$PreviousErrorActionPreference = $ErrorActionPreference
$ErrorActionPreference = "Continue"
gh release view $Tag --repo $Repository *> $null
$ReleaseExists = $LASTEXITCODE -eq 0
$ErrorActionPreference = $PreviousErrorActionPreference
if ($ReleaseExists) {
  gh release upload $Tag $ExePath --repo $Repository --clobber
} else {
  gh release create $Tag $ExePath --repo $Repository --title "Workhub Desktop $Version" --notes $Notes
}
if ($LASTEXITCODE -ne 0) {
  throw "GitHub Release upload failed."
}

$Manifest = [ordered]@{
  version = $Version
  url = $DownloadUrl
  sha256 = $Hash
  size = $Size
  published_at = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
  notes = $Notes
}
$ManifestJson = $Manifest | ConvertTo-Json
$Utf8WithoutBom = New-Object Text.UTF8Encoding($false)
[IO.File]::WriteAllText($ManifestPath, $ManifestJson + "`n", $Utf8WithoutBom)

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$LatestPackage = Get-ChildItem -LiteralPath $OutputDir -Filter "SoilbridgeWorkhub_Desktop_*.zip" |
  Sort-Object LastWriteTime -Descending |
  Select-Object -First 1
if ($LatestPackage) {
  $FriendlyPackage = Join-Path $OutputDir "SoilbridgeWorkhub_Desktop_${Version}_AutoUpdate.zip"
  Copy-Item -LiteralPath $LatestPackage.FullName -Destination $FriendlyPackage -Force
  Write-Host "Installer package: $FriendlyPackage"
}

Write-Host "Release tag: $Tag"
Write-Host "Executable SHA256: $Hash"
Write-Host "Manifest: $ManifestPath"
Write-Host "Commit and push static/desktop_update.json to activate this update."
