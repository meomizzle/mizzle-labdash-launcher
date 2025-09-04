# === File: LabDash-Setup-Host.ps1 ===
[CmdletBinding()]
param(
    [switch]$NoPrefill
)

# Version/build info
$AppVersion = '1.0.0'
$BuildStamp = (Get-Date).ToString('yyyyMMdd-HHmm')

# WPF bootstrap
Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase

# Resolve paths
$root      = Split-Path -Parent $MyInvocation.MyCommand.Path
$xamlPath  = Join-Path $root 'LabDash-Setup-Window.xaml'
$mizzleIco = Join-Path $root 'mizzle.ico'
$labdashIco= Join-Path $root 'labdash.ico'
$jsonPath  = Join-Path $root 'apps.json'

if(!(Test-Path -LiteralPath $xamlPath)){
    throw "XAML not found: $xamlPath"
}

# Load XAML (no Icon attribute in XAML to avoid parse errors)
[xml]$xaml = Get-Content -LiteralPath $xamlPath -Raw -Encoding UTF8
$reader    = New-Object System.Xml.XmlNodeReader $xaml
$win       = [Windows.Markup.XamlReader]::Load($reader)

# Stamp title with version (ASCII-only to avoid encoding issues)
try { $win.Title = 'Mizzle-AV - App Validation for Lab Dash - v' + $AppVersion } catch {}

# Try to set window icon programmatically (safe even if icon file is missing)
try {
    if(Test-Path -LiteralPath $mizzleIco){
        $icoUri = [Uri]::new((Get-Item -LiteralPath $mizzleIco).FullName)
        $win.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create($icoUri)
    }
} catch {}

function Get-Control([string]$name){
    $c = $win.FindName($name)
    if(-not $c){ throw "Missing control: $name" }
    return $c
}

# Controls
$PathBoxes   = 1..5 | ForEach-Object { Get-Control ("Path$_") }
$NameBoxes   = 1..5 | ForEach-Object { Get-Control ("Name$_") }
$ArgsBoxes   = 1..5 | ForEach-Object { Get-Control ("Args$_") }
$BrowseBtns  = 1..5 | ForEach-Object { Get-Control ("Browse$_") }
$ResultsRich = Get-Control 'ResultsRich'
$BtnValidate = Get-Control 'Validate'
$BtnClear    = Get-Control 'ClearEntries'
$BtnSave     = Get-Control 'SaveJson'
$BtnOpen     = Get-Control 'OpenJson'
$BtnOpenLog  = Get-Control 'OpenLog'
$BtnReport   = Get-Control 'AppsReport'
$BtnRemove   = Get-Control 'RemoveLink'
$BtnClose    = Get-Control 'Close'
$LogoLeft    = $win.FindName('LogoLeft')
$LogoRight   = $win.FindName('LogoRight')
$BuildInfoBl = $win.FindName('BuildInfo')
$RowTestBtns = 1..5 | ForEach-Object { $win.FindName(("Test{0}" -f $_)) }

# Optional: set left logo image to mizzle.ico if present
try {
    if($LogoLeft -and (Test-Path -LiteralPath $mizzleIco)){
        $LogoLeft.Source = [System.Windows.Media.Imaging.BitmapFrame]::Create([Uri]::new((Get-Item $mizzleIco).FullName))
    }
} catch {}

# Optional: set right logo image to labdash.ico if present and valid; fallback to mizzle.ico
try {
    if($LogoRight){
        if(Test-Path -LiteralPath $labdashIco){
            try {
                $LogoRight.Source = [System.Windows.Media.Imaging.BitmapFrame]::Create([Uri]::new((Get-Item $labdashIco).FullName))
            } catch {
                if(Test-Path -LiteralPath $mizzleIco){
                    $LogoRight.Source = [System.Windows.Media.Imaging.BitmapFrame]::Create([Uri]::new((Get-Item $mizzleIco).FullName))
                }
            }
        } elseif (Test-Path -LiteralPath $mizzleIco) {
            $LogoRight.Source = [System.Windows.Media.Imaging.BitmapFrame]::Create([Uri]::new((Get-Item $mizzleIco).FullName))
        }
    }
} catch {}

# Helpers
function Get-KeyFromName([string]$name){
    if([string]::IsNullOrWhiteSpace($name)){ return 'app' }
    $t = $name.ToLowerInvariant()
    $t = [regex]::Replace($t,'[^a-z0-9]+','-').Trim('-')
    if([string]::IsNullOrWhiteSpace($t)){ $t = 'app' }
    return $t
}

function Get-FocusedIndex {
    try {
        $fe = [System.Windows.Input.Keyboard]::FocusedElement
        if ($fe -and $fe.Name) {
            if ($fe.Name -match '^(?:Path|Name|Args)(\d+)$') { return [int]$matches[1] - 1 }
        }
    } catch {}
    for($i=0;$i -lt 5;$i++){
        if( -not [string]::IsNullOrWhiteSpace($PathBoxes[$i].Text) -or -not [string]::IsNullOrWhiteSpace($NameBoxes[$i].Text)){ return $i }
    }
    return 0
}

# Prefill UI from apps.json if present (unless -NoPrefill)
if (-not $NoPrefill) {
    try {
        if (Test-Path -LiteralPath $jsonPath) {
            $raw = Get-Content -LiteralPath $jsonPath -Raw -ErrorAction Stop
            if (-not [string]::IsNullOrWhiteSpace($raw)) {
                $data = $raw | ConvertFrom-Json -ErrorAction Stop
                if ($null -ne $data) {
                    $entries = @()
                    if ($data -is [System.Array]) { $entries = $data } else { $entries = @($data) }
                    for ($i = 0; $i -lt [Math]::Min(5, $entries.Count); $i++) {
                        $e = $entries[$i]
                        if ($null -ne $e) {
                            if ($e.PSObject.Properties.Name -contains 'path') { $PathBoxes[$i].Text = [string]$e.path }
                            if ($e.PSObject.Properties.Name -contains 'name') { $NameBoxes[$i].Text = [string]$e.name }
                            if ($e.PSObject.Properties.Name -contains 'args') { $ArgsBoxes[$i].Text = [string]$e.args }
                        }
                    }
                }
            }
        }
    } catch {}
}

# Wire hyperlink click for copyright link
try {
    $lnk = $win.FindName('LinkCopyright')
    if ($lnk) {
        $lnk.Add_RequestNavigate({ param($s,$e) Start-Process $e.Uri.AbsoluteUri; $e.Handled = $true })
    }
} catch {}

# Footer build info
try { if($BuildInfoBl){ $BuildInfoBl.Text = "v$AppVersion ($BuildStamp)" } } catch {}

# Startup info in Results panel: show resolved paths and version
try {
    if ($ResultsRich) {
        $para = New-Object System.Windows.Documents.Paragraph
        $para.Margin = '0,0,0,6'
        $para.Inlines.Add((New-Object System.Windows.Documents.Run ("Startup: v$AppVersion ($BuildStamp)")))
        $para.Inlines.Add((New-Object System.Windows.Documents.LineBreak))
        $para.Inlines.Add((New-Object System.Windows.Documents.Run ("Host:    " + $MyInvocation.MyCommand.Path)))
        $para.Inlines.Add((New-Object System.Windows.Documents.LineBreak))
        $para.Inlines.Add((New-Object System.Windows.Documents.Run ("XAML:    $xamlPath")))
        $para.Inlines.Add((New-Object System.Windows.Documents.LineBreak))
        $para.Inlines.Add((New-Object System.Windows.Documents.Run ("apps.json: $jsonPath")))
        $null = $ResultsRich.Document.Blocks.Add($para)
    }
} catch {}

function Append-PassLine([int]$index, [string]$name, [string]$key, [string]$args, [string]$path){
    $p = New-Object System.Windows.Documents.Paragraph
    $p.Margin = '0,0,0,4'

    $p.Inlines.Add( (New-Object System.Windows.Documents.Run ("PASS  [$index] $name -> ")) )

    # Teal phrase
    $teal = New-Object System.Windows.Documents.Run "URL to be entered into Lab Dash: "
    $teal.Foreground = New-Object System.Windows.Media.SolidColorBrush ([Windows.Media.ColorConverter]::ConvertFromString('#29CCCC'))
    $p.Inlines.Add($teal)

    # URL
    $p.Inlines.Add( (New-Object System.Windows.Documents.Run ("labdash://open/$key")) )

    if(-not [string]::IsNullOrWhiteSpace($args)){
        $p.Inlines.Add( (New-Object System.Windows.Documents.Run ("  [args: $args]")) )
    }

    $null = $ResultsRich.Document.Blocks.Add($p)
}

function Append-FailLine([int]$index, [string]$name, [string]$path){
    $p = New-Object System.Windows.Documents.Paragraph
    $p.Margin = '0,0,0,4'
    $p.Inlines.Add( (New-Object System.Windows.Documents.Run ("FAIL [$index] $name -> path not found: $path")) )
    $null = $ResultsRich.Document.Blocks.Add($p)
}

# Build entry objects from UI (optionally filter to valid only)
function Get-Entries([bool]$OnlyValid = $true){
    $entries = @()
    for($i=0;$i -lt 5;$i++){
        $p = ($PathBoxes[$i].Text).Trim()
        $n = ($NameBoxes[$i].Text).Trim()
        $a = ($ArgsBoxes[$i].Text).Trim()
        if([string]::IsNullOrWhiteSpace($p) -and [string]::IsNullOrWhiteSpace($n) -and [string]::IsNullOrWhiteSpace($a)){
            continue
        }
        if([string]::IsNullOrWhiteSpace($n)){
            $n = [IO.Path]::GetFileNameWithoutExtension($p)
        }
        $key = Get-KeyFromName $n
        $isValid = (Test-Path -LiteralPath $p) -and -not (Get-Item -LiteralPath $p).PSIsContainer
        if(-not $OnlyValid -or $isValid){
            $entry = [ordered]@{
                name = $n
                path = $p
                args = $a
                key  = $key
                url  = "labdash://open/$key"
            }
            $entries += $entry
        }
    }
    return ,$entries
}

# Wire Browse buttons
for($i=0;$i -lt 5;$i++){
    $BrowseBtns[$i].Tag = $i
    $BrowseBtns[$i].Add_Click({
        param($s,$e)
        $ii  = [int]$s.Tag
        $dlg = New-Object Microsoft.Win32.OpenFileDialog
        $dlg.Filter = "Applications (*.exe)|*.exe|All files (*.*)|*.*"
        $cur = $PathBoxes[$ii].Text
        if(-not [string]::IsNullOrWhiteSpace($cur)){
            try{ $dlg.InitialDirectory = Split-Path -Parent $cur }catch{}
        }
        if($dlg.ShowDialog()){
            $PathBoxes[$ii].Text = $dlg.FileName
            if([string]::IsNullOrWhiteSpace($NameBoxes[$ii].Text)){
                $NameBoxes[$ii].Text = [IO.Path]::GetFileNameWithoutExtension($dlg.FileName)
            }
        }
    })
}

# Validate click
$BtnValidate.Add_Click({
    $ResultsRich.Document.Blocks.Clear()
    $passed = 0; $failed = 0

    for($i=0;$i -lt 5;$i++){
        $p = ($PathBoxes[$i].Text).Trim()
        $n = ($NameBoxes[$i].Text).Trim()
        $a = ($ArgsBoxes[$i].Text).Trim()

        if([string]::IsNullOrWhiteSpace($p) -and [string]::IsNullOrWhiteSpace($n) -and [string]::IsNullOrWhiteSpace($a)){
            continue
        }

        if([string]::IsNullOrWhiteSpace($n)){
            $n = [IO.Path]::GetFileNameWithoutExtension($p)
            $NameBoxes[$i].Text = $n
        }
        $key = Get-KeyFromName $n

        if((Test-Path -LiteralPath $p) -and -not (Get-Item -LiteralPath $p).PSIsContainer){
            Append-PassLine ($i+1) $n $key $a $p
            $passed++
        } else {
            Append-FailLine ($i+1) $n $p
            $failed++
        }
    }

    $sum = New-Object System.Windows.Documents.Paragraph
    $sum.Inlines.Add( (New-Object System.Windows.Documents.Run ("Summary: $passed passed, $failed failed.")) )
    $null = $ResultsRich.Document.Blocks.Add($sum)
})

# Clear entries click
$BtnClear.Add_Click({
    for($i=0;$i -lt 5;$i++){
        $PathBoxes[$i].Text = ''
        $NameBoxes[$i].Text = ''
        $ArgsBoxes[$i].Text = ''
    }
    $ResultsRich.Document.Blocks.Clear()
})

# Save JSON click: write valid entries with generated labdash URLs
$BtnSave.Add_Click({
    try {
        # New entries from UI (valid only)
        $newEntries = Get-Entries -OnlyValid $true

        # Load existing entries (if any) so we append/merge instead of overwrite
        $existing = @()
        if (Test-Path -LiteralPath $jsonPath) {
            try {
                $raw = Get-Content -LiteralPath $jsonPath -Raw -ErrorAction Stop
                if (-not [string]::IsNullOrWhiteSpace($raw)) {
                    $tmp = $raw | ConvertFrom-Json -ErrorAction Stop
                    if ($tmp -is [System.Array]) { $existing = $tmp } elseif ($null -ne $tmp) { $existing = @($tmp) }
                }
            } catch {}
        }

        # Merge by unique key; if missing key on an existing record, compute from name
        $map = @{}
        foreach ($e in $existing) {
            if ($null -eq $e) { continue }
            $k = $e.key
            if ([string]::IsNullOrWhiteSpace($k) -and $e.name) { $k = (Get-KeyFromName $e.name) }
            if (-not [string]::IsNullOrWhiteSpace($k)) { $map[$k] = $e }
        }
        foreach ($n in $newEntries) {
            if ($null -eq $n) { continue }
            $map[$n.key] = $n
        }

        $merged = $map.GetEnumerator() | ForEach-Object { $_.Value }

        ($merged | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath $jsonPath -Encoding utf8

        $p = New-Object System.Windows.Documents.Paragraph
        $p.Inlines.Add( (New-Object System.Windows.Documents.Run ("Saved $($newEntries.Count) new/updated entries. Total in file: $($merged.Count).")) )
        $null = $ResultsRich.Document.Blocks.Add($p)
    } catch {
        $p = New-Object System.Windows.Documents.Paragraph
        $p.Inlines.Add( (New-Object System.Windows.Documents.Run ("Error saving JSON: $($_.Exception.Message)")) )
        $null = $ResultsRich.Document.Blocks.Add($p)
    }
})

# Open JSON click: ensure file exists then open with default editor
$BtnOpen.Add_Click({
    try {
        if (-not (Test-Path -LiteralPath $jsonPath)) {
            '[]' | Set-Content -LiteralPath $jsonPath -Encoding utf8
        }
        Invoke-Item -LiteralPath $jsonPath
    } catch {
        $p = New-Object System.Windows.Documents.Paragraph
        $p.Inlines.Add( (New-Object System.Windows.Documents.Run ("Error opening JSON: $($_.Exception.Message)")) )
        $null = $ResultsRich.Document.Blocks.Add($p)
    }
})

# Open Log button
$BtnOpenLog.Add_Click({
    try {
        $localApp = [Environment]::GetFolderPath('LocalApplicationData')
        if ([string]::IsNullOrWhiteSpace($localApp)) { $localApp = $env:LOCALAPPDATA }
        $appDir = Join-Path $localApp 'LabDashLauncher'
        $appLog = Join-Path $appDir 'labdash-handler.log'
        $scriptLog = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'labdash-handler.log'
        $target = $null
        if (Test-Path -LiteralPath $scriptLog) { $target = $scriptLog }
        elseif (Test-Path -LiteralPath $appLog) { $target = $appLog }
        else {
            if (-not (Test-Path -LiteralPath $appDir)) { New-Item -ItemType Directory -Path $appDir -Force | Out-Null }
            '' | Set-Content -LiteralPath $appLog -Encoding utf8
            $target = $appLog
        }
        if ([string]::IsNullOrWhiteSpace($target)) { throw 'Log path could not be resolved.' }
        if (-not (Test-Path -LiteralPath $target)) { '' | Set-Content -LiteralPath $target -Encoding utf8 }
        try { Invoke-Item -LiteralPath $target } catch { Start-Process notepad.exe -ArgumentList $target | Out-Null }
    } catch {
        $p = New-Object System.Windows.Documents.Paragraph
        $p.Inlines.Add( (New-Object System.Windows.Documents.Run ("Error opening log: $($_.Exception.Message)")) )
        $null = $ResultsRich.Document.Blocks.Add($p)
    }
})


# Remove AppLink: delete entry by key from apps.json
$BtnRemove.Add_Click({
    try {
        $i = Get-FocusedIndex
        $name = ($NameBoxes[$i].Text).Trim()
        if ([string]::IsNullOrWhiteSpace($name)) {
            $name = [IO.Path]::GetFileNameWithoutExtension(($PathBoxes[$i].Text).Trim())
        }
        if ([string]::IsNullOrWhiteSpace($name)) { throw "No app selected to remove." }
        $key = Get-KeyFromName $name
        if (-not (Test-Path -LiteralPath $jsonPath)) { throw "apps.json not found: $jsonPath" }
        $raw = Get-Content -LiteralPath $jsonPath -Raw -ErrorAction Stop
        $data = @()
        if (-not [string]::IsNullOrWhiteSpace($raw)) { $tmp = $raw | ConvertFrom-Json -ErrorAction Stop; if ($tmp -is [System.Array]) { $data = $tmp } elseif ($null -ne $tmp){ $data = @($tmp) } }
        $before = $data.Count
        $data = $data | Where-Object { $_.key -ne $key }
        $after = $data.Count
        ($data | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath $jsonPath -Encoding utf8
        $p = New-Object System.Windows.Documents.Paragraph
        $p.Inlines.Add( (New-Object System.Windows.Documents.Run ("Removed key '$key'. Entries: $before -> $after")) )
        $null = $ResultsRich.Document.Blocks.Add($p)
    } catch {
        $p = New-Object System.Windows.Documents.Paragraph
        $p.Inlines.Add( (New-Object System.Windows.Documents.Run ("Error removing link: $($_.Exception.Message)")) )
        $null = $ResultsRich.Document.Blocks.Add($p)
    }
})

# Per-row Test buttons
for($i=0;$i -lt 5;$i++){
    if($RowTestBtns[$i]){
        $RowTestBtns[$i].Tag = $i
        $RowTestBtns[$i].Add_Click({
            param($s,$e)
            $ii = [int]$s.Tag
            try {
                $name = ($NameBoxes[$ii].Text).Trim()
                $args = ($ArgsBoxes[$ii].Text).Trim()
                if ([string]::IsNullOrWhiteSpace($name)) {
                    $name = [IO.Path]::GetFileNameWithoutExtension(($PathBoxes[$ii].Text).Trim())
                }
                if ([string]::IsNullOrWhiteSpace($name)) { throw "No app selected to test." }
                $key = Get-KeyFromName $name
                $url = "labdash://open/$key"
                if (-not [string]::IsNullOrWhiteSpace($args)){
                    $enc = [Uri]::EscapeDataString($args)
                    $url = "$url?args=$enc"
                }
                Start-Process $url | Out-Null
                $p = New-Object System.Windows.Documents.Paragraph
                $p.Inlines.Add( (New-Object System.Windows.Documents.Run ("Tested link: $url")) )
                $null = $ResultsRich.Document.Blocks.Add($p)
            } catch {
                $p = New-Object System.Windows.Documents.Paragraph
                $p.Inlines.Add( (New-Object System.Windows.Documents.Run ("Error testing link: $($_.Exception.Message)")) )
                $null = $ResultsRich.Document.Blocks.Add($p)
            }
        })
    }
}

# Clickable copyright link
try {
    $lnk = $win.FindName('LinkCopyright')
    if ($lnk) {
        $lnk.Add_RequestNavigate({ param($s,$e) Start-Process $e.Uri.AbsoluteUri; $e.Handled = $true })
    }
} catch {}

# Apps Report: write a consolidated text report and open in Notepad
$BtnReport.Add_Click({
    try {
        if (-not (Test-Path -LiteralPath $jsonPath)) { throw "apps.json not found: $jsonPath" }
        $apps = Get-Content -LiteralPath $jsonPath -Raw | ConvertFrom-Json -ErrorAction Stop
        if ($apps -isnot [System.Array]) { $apps = @($apps) }
        $lines = @()
        $lines += "Apps Report (" + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') + ")"
        $lines += "Total: $($apps.Count)"
        $lines += ""
        $idx = 1
        foreach($a in $apps){
            $lines += ("[{0}] Name: {1}" -f $idx, $a.name)
            $lines += ("     Key : {0}" -f $a.key)
            $lines += ("     Path: {0}" -f $a.path)
            if ($a.args) { $lines += ("     Args: {0}" -f $a.args) }
            if ($a.url)  { $lines += ("     URL : {0}" -f $a.url) }
            $lines += ""
            $idx++
        }
        $tmp = Join-Path $env:TEMP 'apps-report.txt'
        $lines | Set-Content -LiteralPath $tmp -Encoding utf8
        Start-Process notepad.exe -ArgumentList $tmp | Out-Null
    } catch {
        $p = New-Object System.Windows.Documents.Paragraph
        $p.Inlines.Add( (New-Object System.Windows.Documents.Run ("Error opening report: $($_.Exception.Message)")) )
        $null = $ResultsRich.Document.Blocks.Add($p)
    }
})

# Close click
$BtnClose.Add_Click({ $win.Close() })

# Show
$null = $win.ShowDialog()
