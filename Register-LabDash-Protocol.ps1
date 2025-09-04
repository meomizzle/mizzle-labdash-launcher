[CmdletBinding(SupportsShouldProcess=$true)]
param(
  [string]$Scheme = 'labdash',
  [string]$Description = 'Lab Dash URL Protocol',
  [string]$HandlerScript = 'Handle-LabDash.ps1',
  [switch]$KeepOpen
)

$ErrorActionPreference = 'Stop'

function Resolve-AbsPath([string]$p){
  (Resolve-Path -LiteralPath $p).ProviderPath
}

try {
  $root = Split-Path -Parent $MyInvocation.MyCommand.Path
  $scriptPath = Resolve-AbsPath (Join-Path $root $HandlerScript)
  if (-not (Test-Path -LiteralPath $scriptPath)) { throw "Handler script not found: $scriptPath" }

  $psExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
  if (-not (Test-Path -LiteralPath $psExe)) { throw "powershell.exe not found at $psExe" }

  if ($KeepOpen) {
    $cmd = '"{0}" -NoProfile -Sta -ExecutionPolicy Bypass -NoExit -File "{1}" "%1"' -f $psExe, $scriptPath
  } else {
    # Two-stage hidden launch to avoid visible console windows
    $escapedScript = $scriptPath.Replace('"','\"')
    $inner = '-NoProfile -WindowStyle Hidden -Sta -ExecutionPolicy Bypass -File "{0}" "%1"' -f $escapedScript
    $cmd = '"{0}" -NoLogo -NonInteractive -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -Command Start-Process -WindowStyle Hidden -FilePath "{0}" -ArgumentList ''{1}''' -f $psExe, $inner
  }

  $baseKey = "HKCU:\Software\Classes\$Scheme"
  if ($PSCmdlet.ShouldProcess($baseKey, 'Create/Update URL protocol registration')) {
    New-Item -Path $baseKey -Force | Out-Null
    New-ItemProperty -Path $baseKey -Name '(Default)' -Value $Description -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $baseKey -Name 'URL Protocol' -Value '' -PropertyType String -Force | Out-Null

    $cmdKey = Join-Path $baseKey 'shell\open\command'
    New-Item -Path $cmdKey -Force | Out-Null
    New-ItemProperty -Path $cmdKey -Name '(Default)' -Value $cmd -PropertyType String -Force | Out-Null
  }

  Write-Host "Registered '$Scheme' to: $cmd"
  Write-Host ("Test: Start-Process '{0}://open/test'" -f $Scheme)
} catch {
  Write-Error $_.Exception.Message
  exit 1
}
