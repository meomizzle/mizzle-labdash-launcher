[CmdletBinding()]
param(
  [Parameter(Mandatory=$true, Position=0)]
  [string]$UriArg
)

$ErrorActionPreference = 'Stop'

function Write-Log($msg){
  $stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
  $line = "[$stamp] $msg"
  # Always write to LocalAppData (reliable), also try script directory
  try {
    $fdir = Join-Path $env:LOCALAPPDATA 'LabDashLauncher'
    if (-not (Test-Path -LiteralPath $fdir)) { New-Item -ItemType Directory -Path $fdir -Force | Out-Null }
    $flog = Join-Path $fdir 'labdash-handler.log'
    Add-Content -LiteralPath $flog -Value $line -ErrorAction Stop
  } catch {}
  try {
    $dir = Split-Path -Parent $MyInvocation.MyCommand.Path
    $log = Join-Path $dir 'labdash-handler.log'
    Add-Content -LiteralPath $log -Value $line -ErrorAction SilentlyContinue
  } catch {}
}

function Get-KeyFromName([string]$name){
  if([string]::IsNullOrWhiteSpace($name)){ return '' }
  $t = $name.ToLowerInvariant()
  $t = [regex]::Replace($t,'[^a-z0-9]+','-').Trim('-')
  return $t
}

try {
  Write-Log "Invoked with: $UriArg"
  Write-Log "User: $env:USERNAME  Host: $env:COMPUTERNAME"
  Write-Log "Script: $($MyInvocation.MyCommand.Path)"
  $uri = [Uri]$UriArg
  if ($uri.Scheme -ne 'labdash') { throw "Unsupported scheme: $($uri.Scheme)" }

  # Accept labdash://open/{key} or labdash://{key}
  $segments = $uri.AbsolutePath.Trim('/').Split('/')
  if ($segments.Count -ge 2 -and $segments[0] -ieq 'open') {
    $key = $segments[1]
  } elseif ($segments.Count -ge 1 -and -not [string]::IsNullOrWhiteSpace($segments[0])) {
    $key = $segments[0]
  } else {
    throw "Invalid LabDash URI path: $($uri.AbsolutePath)"
  }
  
  # Optional query string args to append to configured args
  $qArgs = ''
  if (-not [string]::IsNullOrWhiteSpace($uri.Query)) {
    $q = $uri.Query.TrimStart('?')
    foreach($pair in ($q -split '&')){
      if ([string]::IsNullOrWhiteSpace($pair)) { continue }
      $kv = $pair -split '=',2
      $k = $kv[0]
      $v = ''
      if ($kv.Count -gt 1) { $v = $kv[1] }
      if ($k -ieq 'args') { $qArgs = [System.Uri]::UnescapeDataString($v) }
    }
  }

  $dir = Split-Path -Parent $MyInvocation.MyCommand.Path
  $jsonPath = Join-Path $dir 'apps.json'
  if (-not (Test-Path -LiteralPath $jsonPath)) { throw "apps.json not found: $jsonPath" }
  $apps = Get-Content -LiteralPath $jsonPath -Raw | ConvertFrom-Json
  if ($apps -isnot [System.Array]) { $apps = @($apps) }
  $apps = @($apps)
  Write-Log ("Loaded {0} app entries from apps.json" -f $apps.Count)

  # Try matching by explicit key, then by computed key from name, then by url suffix
  $match = $apps | Where-Object { $_.key -eq $key }
  if (-not $match) {
    $match = $apps | Where-Object { $_.name -and (Get-KeyFromName $_.name) -eq $key }
  }
  if (-not $match) {
    $match = $apps | Where-Object { $_.url -and ([string]$_.url).ToLower().EndsWith('/' + $key.ToLower()) }
  }
  if (-not $match) { throw "No app mapped for key '$key' in apps.json" }
  $app = $match | Select-Object -First 1
  Write-Log ("Matched key '{0}' to app '{1}'" -f $key, $app.name)

  if (-not $app.path -or -not (Test-Path -LiteralPath $app.path)) { throw "App path not found: $($app.path)" }
  $args = ''
  if ($app.args) { $args = [string]$app.args }
  if (-not [string]::IsNullOrWhiteSpace($qArgs)) {
    if ([string]::IsNullOrWhiteSpace($args)) { $args = $qArgs } else { $args = ($args + ' ' + $qArgs).Trim() }
  }

  Write-Log "Launching: '$($app.path)' $args"
  if ([string]::IsNullOrWhiteSpace($args)) {
    Start-Process -FilePath $app.path -WindowStyle Normal | Out-Null
  } else {
    Start-Process -FilePath $app.path -ArgumentList $args -WindowStyle Normal | Out-Null
  }
  Write-Log "Launch dispatched successfully."
} catch {
  Write-Log "ERROR: $($_.Exception.Message)"
  try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
    [System.Windows.Forms.MessageBox]::Show("LabDash handler error: $($_.Exception.Message)", 'LabDash Handler', 'OK', 'Error') | Out-Null
  } catch {}
  exit 1
}
