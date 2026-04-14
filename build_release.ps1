Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot

Write-Host "Installing packaging dependency..."
python -m pip install --upgrade pyinstaller

Write-Host "Removing previous build artifacts..."
Remove-Item -Recurse -Force -ErrorAction SilentlyContinue "$projectRoot\build"
Remove-Item -Recurse -Force -ErrorAction SilentlyContinue "$projectRoot\dist"
Remove-Item -Recurse -Force -ErrorAction SilentlyContinue "$projectRoot\release"

Write-Host "Building executable..."
python -m PyInstaller `
  --noconfirm `
  --clean `
  --onefile `
  --windowed `
  --name OutlookDesktopMailer `
  --add-data "outlook_mailer.ps1;." `
  --add-data "templates.json;." `
  outlook_desktop_mailer.py

$releaseDir = Join-Path $projectRoot "release"
New-Item -ItemType Directory -Force -Path $releaseDir | Out-Null

Copy-Item "$projectRoot\dist\OutlookDesktopMailer.exe" $releaseDir
Copy-Item "$projectRoot\README.md" $releaseDir

$zipPath = Join-Path $releaseDir "OutlookDesktopMailer-portable.zip"
Compress-Archive -Path "$releaseDir\OutlookDesktopMailer.exe", "$releaseDir\README.md" -DestinationPath $zipPath -Force

Write-Host ""
Write-Host "Build complete."
Write-Host "EXE: $releaseDir\OutlookDesktopMailer.exe"
Write-Host "ZIP: $zipPath"
