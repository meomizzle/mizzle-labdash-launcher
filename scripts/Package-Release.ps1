[CmdletBinding()]
param(
  [string]$Version,
  [switch]$OpenFolder
)

$ErrorActionPreference = 'Stop'

function Get-VersionFromHost {
  $hostPath = Join-Path $PSScriptRoot '..' 'LabDash-Setup-Host.ps1'
  if (-not (Test-Path -LiteralPath $hostPath)) {
    throw "Host file not found: $hostPath"
  }
  $text = Get-Content -LiteralPath $hostPath -Raw
  # Match: $AppVersion = '1.2.3' or "1.2.3" with arbitrary whitespace
  $re = '(?m)^\s*\$AppVersion\s*=\s*([\'\"])(?<v>[^\'\"]+)\1\s*$'
  $m = [regex]::Match($text, $re)
  if ($m.Success) { return $m.Groups['v'].Value }
  throw "Unable to detect version from LabDash-Setup-Host.ps1 using regex: $re"
}

try {
  Push-Location (Join-Path $PSScriptRoot '..')

  if (-not $Version -or [string]::IsNullOrWhiteSpace($Version)) {
    $Version = Get-VersionFromHost
  }

  $name = "Mizzle-AV-v$Version"
  $dist = Join-Path (Get-Location) 'dist'
  $stage = Join-Path $dist $name
  if (Test-Path $stage) { Remove-Item -Recurse -Force $stage }
  New-Item -ItemType Directory -Path $stage -Force | Out-Null

  $files = @(
    'LabDash-Setup-Host.ps1',
    'LabDash-Setup-Window.xaml',
    'Launch-Setup-Visible.cmd',
    'Handle-LabDash.ps1',
    'Register-LabDash-Protocol.ps1',
    'apps.example.json',
    'README.md',
    'DEV_BRIEF.md',
    'mizzle.ico',
    'labdash.ico'
  )

  foreach($f in $files){
    if (Test-Path -LiteralPath $f) {
      Copy-Item -LiteralPath $f -Destination (Join-Path $stage (Split-Path -Leaf $f)) -Force
    }
  }

  $zip = Join-Path $dist ("{0}.zip" -f $name)
  if (Test-Path $zip) { Remove-Item -Force $zip }
  Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $zip -Force

  $hash = (Get-FileHash -Algorithm SHA256 -LiteralPath $zip).Hash
  Set-Content -LiteralPath (Join-Path $dist ("{0}.sha256" -f (Split-Path -Leaf $zip))) -Value $hash

  Write-Host "Created: $zip" -ForegroundColor Green
  Write-Host "SHA256 : $hash"
  if ($OpenFolder) { Invoke-Item -LiteralPath $dist }
} finally {
  Pop-Location
}


