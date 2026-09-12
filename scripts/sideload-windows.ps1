# Sideload into Windows PowerPoint by trusting a local network-share catalog.
# Windows desktop Office sideloads from a shared folder registered as a trusted
# catalog, not from a wef folder. Run this once, then add the folder in
#   PowerPoint ▸ File ▸ Options ▸ Trust Center ▸ Trusted Add-in Catalogs
# (or it is added here via the registry), tick "Show in Menu", restart
# PowerPoint, and pick the add-in under  Home ▸ Add-ins ▸ Shared Folder.
#
# The dev server (npm run dev) must be running and its cert trusted (npm run certs).
param(
  [string]$ShareDir = "$env:USERPROFILE\.office-addins\pptx-font-manager"
)
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $PSScriptRoot
New-Item -ItemType Directory -Force -Path $ShareDir | Out-Null
Copy-Item -Force (Join-Path $here 'manifest.xml') (Join-Path $ShareDir 'manifest.xml')

# Register the folder as a trusted catalog (per-user, no admin).
$guid = [guid]::NewGuid().ToString()
$base = 'HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs'
New-Item -Path (Join-Path $base $guid) -Force | Out-Null
$key = Join-Path $base $guid
New-ItemProperty -Path $key -Name 'Id'      -Value $guid            -PropertyType String -Force | Out-Null
New-ItemProperty -Path $key -Name 'Url'     -Value $ShareDir        -PropertyType String -Force | Out-Null
New-ItemProperty -Path $key -Name 'Flags'   -Value 1               -PropertyType DWord  -Force | Out-Null  # 1 = show in menu

Write-Host "Sideloaded manifest -> $ShareDir"
Write-Host "Registered trusted catalog $guid. Restart PowerPoint, then:"
Write-Host "  Home > Add-ins > Shared Folder > Font Manager"
