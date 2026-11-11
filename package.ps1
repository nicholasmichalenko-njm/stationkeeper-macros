# Zips the built kit for handing out: kit, import files, read-me.
# Usage: powershell -ExecutionPolicy Bypass -File package.ps1 -Version 0.7.0
param([string]$Root = $PSScriptRoot, [string]$Version = '0.7.0')
$ErrorActionPreference = 'Stop'
$stage = Join-Path $Root "build\Stationkeeper Macros $Version"
if (Test-Path $stage) { Remove-Item $stage -Recurse -Force }
New-Item -ItemType Directory -Force $stage | Out-Null
Copy-Item (Join-Path $Root 'build\Stationkeeper Macros Kit.dotm') $stage
Copy-Item (Join-Path $Root 'build\Import files') $stage -Recurse
Copy-Item (Join-Path $Root 'release\Read me first.txt') $stage
$zip = Join-Path $Root "build\StationkeeperMacros-$Version.zip"
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path $stage -DestinationPath $zip
Get-FileHash $zip -Algorithm SHA256 | ForEach-Object { "$($_.Hash)  $(Split-Path $zip -Leaf)" }
