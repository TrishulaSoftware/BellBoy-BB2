# MODE: Save to File (overwrite existing)
# TERMINAL: Trish Terminal	
# WORKDIR: D:\Trishula-Infra\Bell Boy\Balls Deep In Bullshit
# SAVE AS: Tony Montana 3

# === HEADLESS MODE ===
param(
    [switch]$Headless,
    [string]$Target,
    [ValidateSet('AppendEnd','AfterPattern','BeforePattern','ReplacePattern','InsideRegion')]
    [string]$Mode,

    # defaults avoid null fuss
    [string]$Pattern = '',
    [string]$Begin   = '',
    [string]$End     = '',

    # NEW: either -SnippetPath or -Snippet is required
    [string]$SnippetPath,
    [string]$Snippet,

    # NEW: override logs folder
    [string]$LogPath,

    # NEW: output controls
    [ValidateSet('UTF8','UTF8BOM','UTF16LE','ASCII')] [string]$Encoding = 'UTF8',
    [ValidateSet('CRLF','LF')] [string]$EOL = 'CRLF',

    [switch]$Literal,
    [switch]$CaseInsensitive,
    [switch]$DryRun
)

# --- AppRoot resolver (ISE-safe; must come BEFORE any Join-Path using $AppRoot) ---
# Guarantees $script:AppRoot is a valid folder even when running selections in ISE.
if (-not (Get-Variable -Name AppRoot -Scope Script -ErrorAction SilentlyContinue) -or
    [string]::IsNullOrWhiteSpace($script:AppRoot)) {

    $cand = @()
    if ($PSScriptRoot)      { $cand += $PSScriptRoot }
    if ($PSCommandPath)     { $cand += (Split-Path -Parent $PSCommandPath) }
    if ($MyInvocation.MyCommand.Path) { $cand += (Split-Path -Parent $MyInvocation.MyCommand.Path) }

    # Fallbacks for ISE "Run Selection" scenarios
    if (-not $cand -or $cand.Count -eq 0) { $cand += (Get-Location).Path }

    # Your canonical home (last resort)
    $cand += 'D:\Trishula-Infra\Bell Boy\BB-2'

    foreach ($c in $cand) {
        if ($c -and (Test-Path -LiteralPath $c)) { $script:AppRoot = $c; break }
    }
}

# --- Settings normalization (fix LastPath setter) ---
# Ensure $settings is a PSCustomObject, not a Hashtable
if (-not $settings) {
    $settings = [pscustomobject]@{}
} elseif ($settings -is [hashtable] -or $settings -is [System.Collections.IDictionary]) {
    $settings = [pscustomobject]$settings
}

# Ensure the LastPath property exists, set a sane default if missing/empty
$hasLastPathProp = $settings.PSObject.Properties.Match('LastPath').Count -gt 0
if (-not $hasLastPathProp) {
    $settings | Add-Member -NotePropertyName LastPath -NotePropertyValue $AppRoot -Force
} elseif ([string]::IsNullOrWhiteSpace($settings.LastPath)) {
    $settings.LastPath = $AppRoot
}

# (Optional) keep $txtFolder in sync on first load
if ($txtFolder -and [string]::IsNullOrWhiteSpace($txtFolder.Text)) {
    $txtFolder.Text = $settings.LastPath
}



Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# === ROOT & PATHS (PS 5.1-safe) ===
function Get-BB2Root {
    param([switch]$PreferAppData)
    # Prefer %APPDATA%\BellBoy-BB-2 for state (logs/config/backups), per project notes.
    if ($PreferAppData) { return (Join-Path $env:APPDATA 'BellBoy-BB-2') }

    # Try to infer script folder even when pasted in console
    $p = $MyInvocation.PSCommandPath
    if (-not $p) { $p = $MyInvocation.MyCommand.Path }
    if ($p) { return (Split-Path -Parent $p) }

    # Fallback to current directory
    return (Get-Location).Path
}

# Script location (for binaries/assets) vs AppData state (for config/logs/backups)
$AppRoot     = Get-BB2Root
$AppDataRoot = Get-BB2Root -PreferAppData

# Choose where state lives (logs/config/backups). Change to $AppRoot if you want it local.
$RootForState = $AppDataRoot

# Paths
$LogsDir    = Join-Path $RootForState 'Logs'
$BackupsDir = Join-Path $RootForState 'Backups'
$AppendBak  = Join-Path $BackupsDir 'Appends'
$InjectBak  = Join-Path $BackupsDir 'Injects'
$configPath = Join-Path $RootForState 'config.json'
$sessionLog = Join-Path $LogsDir 'session.log'

# Ensure directories exist
$null = New-Item -Force -ItemType Directory -Path $LogsDir, $BackupsDir, $AppendBak, $InjectBak | Out-Null

# Settings bootstrap (guarantee $settings & $settingsPath exist before UI uses them)
$settingsPath = $configPath
if (Test-Path $settingsPath) {
    try { $settings = Get-Content $settingsPath -Raw | ConvertFrom-Json } catch { $settings = [ordered]@{ LastPath = (Get-Location).Path } }
} else {
    $settings = [ordered]@{ LastPath = (Get-Location).Path }
    $settings | ConvertTo-Json | Set-Content $settingsPath -Encoding UTF8
}

# Allow headless -LogPath to override logs location
if ($PSBoundParameters.ContainsKey('LogPath') -and $LogPath) {
    $LogsDir    = $LogPath
    $sessionLog = Join-Path $LogsDir 'session.log'
}
$null = New-Item -Force -ItemType Directory -Path $LogsDir | Out-Null

# ISE-safe AppRoot resolver (must appear BEFORE any Join-Path using $AppRoot)
if (-not $AppRoot -or [string]::IsNullOrWhiteSpace($AppRoot)) {
    $AppRoot = $PSScriptRoot
    if ([string]::IsNullOrWhiteSpace($AppRoot)) {
        # ISE sometimes doesn't populate $PSScriptRoot; fallbacks:
        if ($MyInvocation.MyCommand.Path) {
            $AppRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
        } else {
            $AppRoot = (Get-Location).Path
        }
    }
}

# --- Settings normalization (fix LastPath setter) ---
# Ensure $settings is a PSCustomObject, not a Hashtable
if (-not $settings) {
    $settings = [pscustomobject]@{}
} elseif ($settings -is [hashtable] -or $settings -is [System.Collections.IDictionary]) {
    $settings = [pscustomobject]$settings
}

# Ensure the LastPath property exists, set a sane default if missing/empty
$hasLastPathProp = $settings.PSObject.Properties.Match('LastPath').Count -gt 0
if (-not $hasLastPathProp) {
    $settings | Add-Member -NotePropertyName LastPath -NotePropertyValue $AppRoot -Force
} elseif ([string]::IsNullOrWhiteSpace($settings.LastPath)) {
    $settings.LastPath = $AppRoot
}

# (Optional) keep $txtFolder in sync on first load
if ($txtFolder -and [string]::IsNullOrWhiteSpace($txtFolder.Text)) {
    $txtFolder.Text = $settings.LastPath
}



# === CONFIG ===
if (Test-Path $configPath) {
    try { $settings = Get-Content $configPath -Raw | ConvertFrom-Json } catch { $settings = @{ LastPath = $AppRoot; LastFile = '' } }
} else { $settings = @{ LastPath = $AppRoot; LastFile = '' } }

# === LOGGING ===
function Write-BBLog {
    param([string]$Message)
    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    "[${stamp}] $Message" | Add-Content -Path $sessionLog -Encoding UTF8
    if ($script:lstLog -and -not $script:lstLog.IsDisposed) {
        $script:lstLog.Items.Insert(0, "[${stamp}] $Message")
        while ($script:lstLog.Items.Count -gt 20) { $script:lstLog.Items.RemoveAt(20) }
    }
}


function Backup-File {
    param([string]$Path)
    $name = [IO.Path]::GetFileNameWithoutExtension($Path)
    $ext  = [IO.Path]::GetExtension($Path)
    $ts   = Get-Date -Format 'yyyyMMdd_HHmmss'
    $bak  = Join-Path $BackupsDir ("{0}_{1}{2}" -f $name,$ts,$ext)
    Copy-Item -LiteralPath $Path -Destination $bak -Force
    Write-BBLog "Backup created → $bak"
    return $bak
}

function Backup-AppendChunk {
    param([string]$TargetPath,[string]$Chunk,[int]$LineNumber,[string]$Mode)
    $base = [IO.Path]::GetFileName($TargetPath) -replace '[^\w\.-]','_'
    $ts   = Get-Date -Format 'yyyyMMdd_HHmmss'
    $file = "{0}__append_{1}_L{2}_{3}.txt" -f $base,$ts,$LineNumber,($Mode -replace ' ','_')
    $dest = Join-Path $AppendBak $file
    $Chunk | Set-Content -Path $dest -Encoding UTF8
    Write-BBLog "Append chunk saved → $dest"
    return $dest
}

function Backup-InjectChunk {
    param(
        [string]$TargetPath, [string]$Snippet, [string]$Mode,
        [string]$Pattern,    [string]$Begin,   [string]$End
    )

    # ---- determine backups root (same doctrine chain as Tab 2) ----
    $appData = [Environment]::GetFolderPath('ApplicationData')
    $root = if ($BackupsDir) {
        $BackupsDir
    } elseif ($RootForState) {
        Join-Path $RootForState 'Backups'
    } elseif ($appData) {
        Join-Path $appData 'BellBoy-BB-2\Backups'
    } else {
        Join-Path $PSScriptRoot 'Backups'
    }

    # Ensure Backups\Injects exists
    $injectDir = Join-Path $root 'Injects'
    if (-not (Test-Path -LiteralPath $injectDir)) {
        New-Item -ItemType Directory -Path $injectDir -Force | Out-Null
    }

    # ---- sanitize pieces for a safe filename (no slashes, etc.) ----
    $base     = ([IO.Path]::GetFileName($TargetPath) -replace '[^\w\.-]','_')
    $ts       = Get-Date -Format 'yyyyMMdd_HHmmss'
    $modeSafe = ($Mode -replace '[^\w\.-]','_')      # turns "Inside Region (Begin/End)" into "Inside_Region__Begin_End_"

    $file = "{0}__inject_{1}_{2}.txt" -f $base, $ts, $modeSafe
    $dest = Join-Path $injectDir $file

    # ---- write a small meta header + the snippet ----
    $meta = @"
# Bell Boy — Inject snippet backup
# Target : $TargetPath
# Time   : $ts
# Mode   : $Mode
# Pattern: $Pattern
# Begin  : $Begin
# End    : $End
# -----
$Snippet
"@

    $meta | Set-Content -LiteralPath $dest -Encoding UTF8
    Write-BBLog "Inject snippet saved → $dest"
    return $dest
}

# === BACKUP ROTATION ===
$script:MaxBackupsPerFile = 20  # change to taste

function Get-SafeBase([string]$p) { ([IO.Path]::GetFileName($p) -replace '[^\w\.-]','_') }

function Prune-BackupsFor {
    param([string]$TargetPath)
    if (-not $TargetPath) { return }
    $base = Get-SafeBase $TargetPath

    # 1) File backups (Backups\<name>_timestamp.ext)
    $filePattern = "$($base.Split('.')[0])*"  # loose match is fine; sorting handles order
    $files = Get-ChildItem -LiteralPath $BackupsDir -File -ErrorAction SilentlyContinue |
             Where-Object { $_.Name -like "$filePattern*" } |
             Sort-Object LastWriteTime -Descending
    if ($files.Count -gt $script:MaxBackupsPerFile) {
        $files | Select-Object -Skip $script:MaxBackupsPerFile | ForEach-Object {
            try { Remove-Item -LiteralPath $_.FullName -Force } catch {}
        }
    }

    # 2) Append chunks
    $app = Get-ChildItem -LiteralPath $AppendBak -File -ErrorAction SilentlyContinue |
           Where-Object { $_.Name -like "${base}__append_*" } |
           Sort-Object LastWriteTime -Descending
    if ($app.Count -gt $script:MaxBackupsPerFile) {
        $app | Select-Object -Skip $script:MaxBackupsPerFile | ForEach-Object {
            try { Remove-Item -LiteralPath $_.FullName -Force } catch {}
        }
    }

    # 3) Inject chunks
    $inj = Get-ChildItem -LiteralPath $InjectBak -File -ErrorAction SilentlyContinue |
           Where-Object { $_.Name -like "${base}__inject_*" } |
           Sort-Object LastWriteTime -Descending
    if ($inj.Count -gt $script:MaxBackupsPerFile) {
        $inj | Select-Object -Skip $script:MaxBackupsPerFile | ForEach-Object {
            try { Remove-Item -LiteralPath $_.FullName -Force } catch {}
        }
    }
}

# === HEADLESS RUNNER ===
if ($Headless) {
    try {
        if (-not $Target) { throw "Missing -Target" }
        if (-not $Mode)   { throw "Missing -Mode (AppendEnd|AfterPattern|BeforePattern|ReplacePattern|InsideRegion)" }
        if (-not (Test-Path -LiteralPath $Target)) { throw "Target not found: $Target" }
        $orig = Get-Content -LiteralPath $Target -Raw

        # pick snippet source
        if ($SnippetPath) {
            if (-not (Test-Path -LiteralPath $SnippetPath)) { throw "SnippetPath not found: $SnippetPath" }
            $snip = Get-Content -LiteralPath $SnippetPath -Raw
        } elseif ($Snippet) {
            $snip = $Snippet
        } else {
            throw "Provide -SnippetPath or -Snippet."
        }


        # map CLI mode names to Tab 3 names
        switch ($Mode) {
            'AppendEnd'      { $mm = 'Append to end' }
            'AfterPattern'   { $mm = 'After Pattern' }
            'BeforePattern'  { $mm = 'Before Pattern' }
            'ReplacePattern' { $mm = 'Replace Pattern' }
            'InsideRegion'   { $mm = 'Inside Region (Begin/End)' }
            default          { throw "Unsupported -Mode: $Mode" }
        }

        $useRegex = -not $Literal
        $ci       = [bool]$CaseInsensitive

        # Build preview using same engine as UI
        $p = CI-BuildPreview -Original $orig -Snippet $snip -Mode $mm `
                     -Pattern $Pattern -Begin $Begin -End $End `
                     -UseRegex $useRegex -CaseInsensitive $ci


        if ($p.Note -eq 'No match found.') {
            Write-Host "No match found. Nothing changed." -ForegroundColor Yellow
            exit 2
        }

        # Show diff when DryRun
        if ($DryRun) {
            Write-Host "=== DRY RUN ===" -ForegroundColor Cyan
            $p.Diff | Write-Output
            exit 0
        }

        # Backups + atomic write + rotation
        if (Test-Path -LiteralPath $Target) {
            $bk = Backup-File -Path $Target
            $script:LastBackupFile = $bk
            $script:LastTargetPath = $Target
        }
        [void](Backup-InjectChunk -TargetPath $Target -Snippet $snip -Mode $mm -Pattern $Pattern -Begin $Begin -End $End)

        $tmp = "$Target.bbtmp"
        $finalOut = Normalize-EOL -Text $p.Final -Style $EOL
        Out-WithEncoding -Path $tmp -Text $finalOut -Enc $Encoding
        Move-Item -LiteralPath $tmp -Destination $Target -Force



        Prune-BackupsFor -TargetPath $Target
        Write-Host "Injection complete → $Target" -ForegroundColor Green
        exit 0
    } catch {
        Write-Error $_.Exception.Message
        exit 1
    }
    return  # don't launch the form in headless mode
}



# === DIFF HELPERS (line-based, unified) ===
function Get-UnifiedDiffText {
    param([string[]]$Old,[string[]]$New)
    $max = [Math]::Max($Old.Count,$New.Count)
    $out = New-Object System.Collections.Generic.List[string]
    for ($i=0; $i -lt $max; $i++){
        $o = if ($i -lt $Old.Count) { $Old[$i] } else { $null }
        $n = if ($i -lt $New.Count) { $New[$i] } else { $null }
        if ($o -eq $n) { $out.Add(" $o") }
        elseif ($o -ne $null -and $n -ne $null) { $out.Add("-$o"); $out.Add("+$n") }
        elseif ($o -ne $null) { $out.Add("-$o") }
        elseif ($n -ne $null) { $out.Add("+$n") }
    }
    return $out -join [Environment]::NewLine
}

function Normalize-EOL {
    param([string]$Text,[ValidateSet('CRLF','LF')][string]$Style='CRLF')
    if ($null -eq $Text) { return '' }
    $t = $Text -replace "`r?`n","`n"      # normalize to LF
    if ($Style -eq 'CRLF') { return ($t -replace "`n","`r`n") }
    else { return $t }                    # leave as LF
}
function Out-WithEncoding {
    param([string]$Path,[string]$Text,[ValidateSet('UTF8','UTF8BOM','UTF16LE','ASCII')][string]$Enc='UTF8')
    switch ($Enc) {
        'UTF8'     { $Text | Set-Content -LiteralPath $Path -Encoding UTF8 }
        'UTF8BOM'  { $Text | Set-Content -LiteralPath $Path -Encoding UTF8BOM }
        'UTF16LE'  { $Text | Set-Content -LiteralPath $Path -Encoding Unicode }
        'ASCII'    { $Text | Set-Content -LiteralPath $Path -Encoding ASCII }
    }
}

# === FORM SHELL (minimal, safe) ===
$form = New-Object System.Windows.Forms.Form
$form.Text          = "🔱 Bell Boy v2 — BB-2 - Trishula Software - 'Pierce The Heavens' 🔱"
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
$form.Size          = New-Object System.Drawing.Size(1380,1080)   # default on open
$form.MinimumSize   = New-Object System.Drawing.Size(1200, 820)   # prevents too-small clipping
$form.StartPosition = "CenterScreen"
$form.Font          = New-Object System.Drawing.Font("Segoe UI", 9)

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = 'Fill'          # << key line: tabs fill the window
$form.Controls.Add($tabs)

function Enable-TabAutoScroll {
    param(
        [Parameter(Mandatory)]
        [System.Windows.Forms.TabPage[]]$TabPages
    )

    foreach ($tab in $TabPages) {
        if (-not $tab -or -not $tab.Controls) { continue }

        $maxR = 0; $maxB = 0
        foreach ($c in $tab.Controls) {
            $r = [int]$c.Left + [int]$c.Width
            $b = [int]$c.Top  + [int]$c.Height
            if ($r -gt $maxR) { $maxR = $r }
            if ($b -gt $maxB) { $maxB = $b }
        }

        $pad = 24
        $tab.AutoScroll = $true
        $minW = if ($maxR -gt 0) { $maxR + $pad } else { [int]$tab.Width  }
        $minH = if ($maxB -gt 0) { $maxB + $pad } else { [int]$tab.Height }
        $tab.AutoScrollMinSize = New-Object System.Drawing.Size($minW, $minH)
    }
}

function Prune-BackupsFor {
    param([string]$TargetPath)

    # Backups Doctrine: rotation disabled.
    # We keep this function so existing calls don't explode, but it does NOT delete anything.
    if ([string]::IsNullOrWhiteSpace($TargetPath)) { return }

    Write-BBLog ("Prune-BackupsFor called for {0} – rotation disabled; no backups deleted." -f $TargetPath)
}

# ===================================================================
# TAB 1 — Create
# ===================================================================
$tab1 = New-Object System.Windows.Forms.TabPage
$tab1.Text = "Tab 1 — Create"
$tabs.TabPages.Add($tab1)

# ——— Header
$lblHead = New-Object System.Windows.Forms.Label
$lblHead.Text = "Create File  •  Auto-Backup on Overwrite  •  Quick Actions"
$lblHead.AutoSize = $true
$lblHead.Location = New-Object System.Drawing.Point(20,20)
$tab1.Controls.Add($lblHead)

# ——— Save Folder row
$lblFolder = New-Object System.Windows.Forms.Label
$lblFolder.Text = "Save Folder:"
$lblFolder.AutoSize = $true
$lblFolder.Location = New-Object System.Drawing.Point(20,58)
$tab1.Controls.Add($lblFolder)

$txtFolder = New-Object System.Windows.Forms.TextBox
$txtFolder.Size = New-Object System.Drawing.Size(800,24)
$txtFolder.Location = New-Object System.Drawing.Point(120,56)
$txtFolder.Text = $settings.LastPath
$tab1.Controls.Add($txtFolder)

# Use EXPLORER-STYLE SAVE DIALOG (fills folder + name)
$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "Browse"
$btnBrowse.Size = New-Object System.Drawing.Size(90,26)
$btnBrowse.Location = New-Object System.Drawing.Point(930,54)
$tab1.Controls.Add($btnBrowse)

$saveDlg = New-Object System.Windows.Forms.SaveFileDialog
$saveDlg.Title  = "Choose save location / file"
$saveDlg.Filter = "All files (*.*)|*.*"
$saveDlg.OverwritePrompt = $false
$saveDlg.ValidateNames   = $true
try {
    if ($settings.LastPath -and (Test-Path $settings.LastPath)) {
        $saveDlg.InitialDirectory = $settings.LastPath
    }
} catch {}

$btnBrowse.Add_Click({
    if ($saveDlg.ShowDialog() -eq 'OK') {
        $txtFolder.Text = Split-Path -Path $saveDlg.FileName -Parent
        $txtName.Text   = Split-Path -Path $saveDlg.FileName -Leaf
    }
})

# ——— File Name
$lblName = New-Object System.Windows.Forms.Label
$lblName.Text = "File Name:"
$lblName.AutoSize = $true
$lblName.Location = New-Object System.Drawing.Point(20,92)
$tab1.Controls.Add($lblName)

$txtName = New-Object System.Windows.Forms.TextBox
$txtName.Size = New-Object System.Drawing.Size(400,24)
$txtName.Location = New-Object System.Drawing.Point(120,90)
$txtName.Text = "NewFile.txt"
$tab1.Controls.Add($txtName)

# ——— Template
$lblTpl = New-Object System.Windows.Forms.Label
$lblTpl.Text = "Template:"
$lblTpl.AutoSize = $true
$lblTpl.Location = New-Object System.Drawing.Point(540,92)
$tab1.Controls.Add($lblTpl)

$cmbTpl = New-Object System.Windows.Forms.ComboBox
$cmbTpl.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void]$cmbTpl.Items.AddRange(@('Blank (UTF-8)','README (txt)','PowerShell Script (.ps1)','JSON stub (.json)'))
$cmbTpl.SelectedIndex = 0
$cmbTpl.Size = New-Object System.Drawing.Size(300,24)
$cmbTpl.Location = New-Object System.Drawing.Point(610,90)
$tab1.Controls.Add($cmbTpl)

# ——— Content
$lblContent = New-Object System.Windows.Forms.Label
$lblContent.Text = "Content:"
$lblContent.AutoSize = $true
$lblContent.Location = New-Object System.Drawing.Point(20,128)
$tab1.Controls.Add($lblContent)

$txtContent = New-Object System.Windows.Forms.TextBox
$txtContent.Multiline = $true
$txtContent.ScrollBars = 'Vertical'
$txtContent.Size = New-Object System.Drawing.Size(1000,560)
$txtContent.Location = New-Object System.Drawing.Point(20,150)
$txtContent.Text = @"
# Bell Boy BB-2
# Created: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
"@
$tab1.Controls.Add($txtContent)

# ——— Actions Row
$btnCreate = New-Object System.Windows.Forms.Button
$btnCreate.Text = "Create / Overwrite (w/ Backup)"
$btnCreate.Size = New-Object System.Drawing.Size(220,32)
$btnCreate.Location = New-Object System.Drawing.Point(20,730)
$tab1.Controls.Add($btnCreate)

# Optional scaffold toggle + button
$chkScaffold = New-Object System.Windows.Forms.CheckBox
$chkScaffold.Text = "Scaffold default structure"
$chkScaffold.AutoSize = $true
$chkScaffold.Location = New-Object System.Drawing.Point(260,735)
$tab1.Controls.Add($chkScaffold)

$btnScaffold = New-Object System.Windows.Forms.Button
$btnScaffold.Text = "Scaffold Project Folders"
$btnScaffold.Size = New-Object System.Drawing.Size(180,32)
$btnScaffold.Location = New-Object System.Drawing.Point(420,730)
$tab1.Controls.Add($btnScaffold)

$btnReveal = New-Object System.Windows.Forms.Button
$btnReveal.Text = "Reveal in Explorer"
$btnReveal.Size = New-Object System.Drawing.Size(160,32)
$btnReveal.Location = New-Object System.Drawing.Point(610,730)
$tab1.Controls.Add($btnReveal)

$btnNotepad = New-Object System.Windows.Forms.Button
$btnNotepad.Text = "Open in Notepad"
$btnNotepad.Size = New-Object System.Drawing.Size(160,32)
$btnNotepad.Location = New-Object System.Drawing.Point(780,730)
$tab1.Controls.Add($btnNotepad)

$btnVSCode = New-Object System.Windows.Forms.Button
$btnVSCode.Text = "Open in VS Code"
$btnVSCode.Size = New-Object System.Drawing.Size(160,32)
$btnVSCode.Location = New-Object System.Drawing.Point(950,730)
$tab1.Controls.Add($btnVSCode)

# ——— Log
$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = "Last 20 actions:"
$lblLog.AutoSize = $true
$lblLog.Location = New-Object System.Drawing.Point(20,774)
$tab1.Controls.Add($lblLog)

$script:lstLog = New-Object System.Windows.Forms.ListBox
$script:lstLog.Size = New-Object System.Drawing.Size(1000,150)
$script:lstLog.Location = New-Object System.Drawing.Point(20,796)
$tab1.Controls.Add($script:lstLog)
if (Test-Path $sessionLog) { Get-Content -Path $sessionLog -Tail 20 | ForEach-Object { $script:lstLog.Items.Insert(0, $_) } }

# ——— Status bar
$status = New-Object System.Windows.Forms.StatusStrip
$sbText = New-Object System.Windows.Forms.ToolStripStatusLabel
$sbText.Text = "Ready."
[void]$status.Items.Add($sbText)
$status.Dock = 'Bottom'
$form.Controls.Add($status)
function Set-Status([string]$msg) { $sbText.Text = $msg }

Enable-TabAutoScroll $tab1

# ——— Template autofill
$cmbTpl.Add_SelectedIndexChanged({
    switch ($cmbTpl.SelectedItem) {
        'Blank (UTF-8)'            { $txtContent.Text = "" }
        'README (txt)' {
          $txtContent.Text = "Title`r`n=====`r`n`r`nSummary:`r`nCreated: $(Get-Date)"
          if (-not $txtName.Text -or -not $txtName.Text.EndsWith('.txt')) { $txtName.Text = 'README.txt' }
        }
        'PowerShell Script (.ps1)' { $txtContent.Text = "# PowerShell`r`nparam()`r`n`r`nWrite-Host 'Hello from BB-2'"; if (-not $txtName.Text.EndsWith('.ps1')) { $txtName.Text = 'NewScript.ps1' } }
        'JSON stub (.json)'        { $txtContent.Text = "{`r`n  `"name`": `"sample`",`r`n  `"created`": `"$(Get-Date -Format s)`"`r`n}"; if (-not $txtName.Text.EndsWith('.json')) { $txtName.Text = 'data.json' } }
    }
})

# ——— Helpers
function Write-BBLog { param([string]$m) try { $script:lstLog.Items.Insert(0, ("[{0}] {1}" -f (Get-Date -Format 'HH:mm:ss'), $m)) } catch {} }
function Ensure-Dir { param([string]$p) if (-not (Test-Path $p)) { New-Item -ItemType Directory -Path $p -Force | Out-Null } }
function Backup-File-Fallback {
    param([string]$Path)
    try {
        $name = [IO.Path]::GetFileName($Path)
        $ts   = Get-Date -Format "yyyyMMdd_HHmmss"
        $dir  = Join-Path $PSScriptRoot "Backups"
        Ensure-Dir $dir
        $dst  = Join-Path $dir "$name.create.bak.$ts"
        Copy-Item -Path $Path -Destination $dst -Force
        return $dst
    } catch { return $null }
}

function Backup-InitialCreate {
  param([string]$NewFilePath)
  try {
    if (-not (Test-Path -LiteralPath $NewFilePath)) { return $null }

    $backRoot = if ($BackupsDir) {
      $BackupsDir
    } elseif ($RootForState) {
      Join-Path $RootForState 'Backups'
    } else {
      Join-Path $PSScriptRoot 'Backups'
    }

    $createsDir = Join-Path $backRoot 'Creates'
    if (-not (Test-Path -LiteralPath $createsDir)) {
      New-Item -ItemType Directory -Path $createsDir -Force | Out-Null
    }

    $name = [IO.Path]::GetFileName($NewFilePath)
    $ts   = Get-Date -Format "yyyyMMdd_HHmmss"
    $dst  = Join-Path $createsDir "$name.create.init.$ts.bak"
    Copy-Item -LiteralPath $NewFilePath -Destination $dst -Force
    return $dst
  } catch { return $null }
}

function Resolve-CodeExe {
  try {
    foreach ($guess in @(
      "$env:LOCALAPPDATA\Programs\Microsoft VS Code\Code.exe",
      "$env:ProgramFiles\Microsoft VS Code\Code.exe",
      "$env:ProgramFiles(x86)\Microsoft VS Code\Code.exe"
    )) {
      if (Test-Path -LiteralPath $guess) { return $guess }
    }
  } catch {}
  try {
    $cmd = (Get-Command code -ErrorAction SilentlyContinue)
    if ($cmd -and $cmd.Path) { return $cmd.Path }  # fallback to code(.cmd)
  } catch {}
  return $null
}

function Sanitize-FileName {
  param([string]$Name)
  if ([string]::IsNullOrWhiteSpace($Name)) { return 'NewFile.txt' }

  $invalid = [IO.Path]::GetInvalidFileNameChars()
  $chars   = $Name.ToCharArray()
  for ($i=0; $i -lt $chars.Length; $i++) {
    if ($invalid -contains $chars[$i]) { $chars[$i] = '_' }
  }
  $safe = (-join $chars).TrimEnd('.',' ')
  if (-not $safe) { $safe = 'NewFile.txt' }

  $base = [IO.Path]::GetFileNameWithoutExtension($safe)
  $ext  = [IO.Path]::GetExtension($safe)
  $reserved = @('CON','PRN','AUX','NUL','COM1','COM2','COM3','COM4','COM5','COM6','COM7','COM8','COM9','LPT1','LPT2','LPT3','LPT4','LPT5','LPT6','LPT7','LPT8','LPT9')
  if ($reserved -contains $base.ToUpperInvariant()) {
    $safe = "${base}_$ext"
  }
  return $safe
}


# ——— Scaffolder
$btnScaffold.Add_Click({
    try {
        $folder = $txtFolder.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($folder)) { [Windows.Forms.MessageBox]::Show("Choose a base folder first.","Scaffold"); return }
        Ensure-Dir $folder
        $subs = @('src','docs','tests','backups','scripts')
        foreach ($s in $subs) { Ensure-Dir (Join-Path $folder $s) }
        Write-BBLog "Scaffolded: $(($subs -join ', '))"
        Set-Status "Scaffold created."
        [Windows.Forms.MessageBox]::Show("Scaffold ready under:`r`n$folder`r`n`r`n($($subs -join ', '))","Scaffold") | Out-Null
    } catch {
        Set-Status "Scaffold error: $($_.Exception.Message)"
        Write-BBLog ("ERROR (scaffold): {0}" -f $_.Exception.Message)
    }
})


# ——— Create (atomic) with auto-dir & backup
$btnCreate.Add_Click({
    $tmp            = $null
    $writeSucceeded = $false

    try {
        # capture UI values safely
        $folder = $txtFolder.Text
        if ($folder) { $folder = $folder.Trim() } else { $folder = '' }
        $name   = $txtName.Text
        if ($name)   { $name   = $name.Trim() }   else { $name   = '' }

        if ([string]::IsNullOrWhiteSpace($folder)) { [Windows.Forms.MessageBox]::Show("Choose a valid folder.","Input Error") | Out-Null; return }
        if ([string]::IsNullOrWhiteSpace($name))   { [Windows.Forms.MessageBox]::Show("Enter a file name.","Input Error")   | Out-Null; return }

        # ensure target directory exists (+ optional scaffold)
        Ensure-Dir $folder
        if ($chkScaffold -and $chkScaffold.Checked) {
            foreach ($s in @('src','docs','tests','backups','scripts')) { Ensure-Dir (Join-Path $folder $s) }
        }

        # sanitize filename if needed
        $origName = $name
        $invalid  = [IO.Path]::GetInvalidFileNameChars()
        $needsFix = ($name.IndexOfAny($invalid) -ge 0) -or
                    ($name.TrimEnd('.',' ') -ne $name) -or
                    (@('CON','PRN','AUX','NUL','COM1','COM2','COM3','COM4','COM5','COM6','COM7','COM8','COM9','LPT1','LPT2','LPT3','LPT4','LPT5','LPT6','LPT7','LPT8','LPT9') -contains ([IO.Path]::GetFileNameWithoutExtension($name).ToUpperInvariant()))
        if ($needsFix) {
          $name = Sanitize-FileName -Name $name
          $txtName.Text = $name
          Write-BBLog ("Note: filename adjusted from '{0}' to '{1}' for safety." -f $origName,$name)
          [Windows.Forms.MessageBox]::Show("Filename contained illegal characters/reserved words.`r`nAdjusted to:`r`n$name","Filename adjusted") | Out-Null
        }

        $path = Join-Path $folder $name
        if ($path.Length -gt 240) {
          Write-BBLog ("Warning: long path ({0} chars). Proceeding." -f $path.Length)
        }

        # backup on overwrite
        $preExisted = Test-Path -LiteralPath $path
        if ($preExisted) {
            if (Get-Command Backup-File -ErrorAction SilentlyContinue) {
                [void](Backup-File -Path $path)
            } else {
                [void](Backup-File-Fallback -Path $path)
            }
        }

        # atomic-ish write in same folder
        $rand    = [IO.Path]::GetRandomFileName()
        $tmpName = '.' + ([IO.Path]::GetFileName($name)) + '.' + $rand + '.tmp'
        $tmp     = Join-Path $folder $tmpName

        @($txtContent.Text) | Set-Content -LiteralPath $tmp -Encoding UTF8 -ErrorAction Stop
        Move-Item -LiteralPath $tmp -Destination $path -Force -ErrorAction Stop

        # backup on initial create
        if (-not $preExisted) {
            $initBak = Backup-InitialCreate -NewFilePath $path
            if ($initBak) { Write-BBLog ("Initial backup → {0}" -f $initBak) }
        }

        $writeSucceeded = $true

        if ($writeSucceeded) {
            Write-BBLog ("Created/Updated → {0}" -f $path)
            Set-Status "Wrote: $path"
            [Windows.Forms.MessageBox]::Show("Saved:`r`n$path","Success") | Out-Null
            $settings.LastPath = $folder
        }
    }
    catch {
        # tidy temp if it got created but not moved
        try {
            if ($tmp -and (Test-Path -LiteralPath $tmp)) {
                Remove-Item -LiteralPath $tmp -Force
            }
        } catch {}

        $msg = $_.Exception.Message
        Set-Status "Error: $msg"
        Write-BBLog ("ERROR (create): {0}" -f $msg)
        [Windows.Forms.MessageBox]::Show("Error:`r`n$msg","Failure") | Out-Null
    }
})




# ——— Reveal / Editors
$btnReveal.Add_Click({
  try {
    $folder = $txtFolder.Text.Trim()
    $name   = $txtName.Text.Trim()
    if (-not $folder) { return }
    $path = if ($name) { Join-Path $folder $name } else { $null }

    if ($path -and (Test-Path -LiteralPath $path)) {
      Start-Process -FilePath explorer.exe -ArgumentList "/select,`"$path`""
    } else {
      Start-Process -FilePath explorer.exe -ArgumentList "`"$folder`""
    }
    Write-BBLog "Explorer opened."; Set-Status "Explorer opened."
  } catch {}
})

$btnNotepad.Add_Click({
  try {
    $path = Join-Path $txtFolder.Text.Trim() $txtName.Text.Trim()
    if (Test-Path -LiteralPath $path) {
      Start-Process -FilePath notepad.exe -ArgumentList "`"$path`"" | Out-Null
      Write-BBLog "Opened in Notepad."; Set-Status "Opened in Notepad."
    } else {
      [Windows.Forms.MessageBox]::Show("File does not exist yet.","Info") | Out-Null
    }
  } catch {}
})

$btnVSCode.Add_Click({
  try {
    $path = Join-Path $txtFolder.Text.Trim() $txtName.Text.Trim()
    if (-not (Test-Path -LiteralPath $path)) {
      [Windows.Forms.MessageBox]::Show("File does not exist yet.","Info") | Out-Null
      return
    }
    $code = Resolve-CodeExe
    if ($code -and (Test-Path -LiteralPath $code)) {
      Start-Process -FilePath $code -ArgumentList "`"$path`"" | Out-Null
      Write-BBLog "Opened in VS Code."; Set-Status "Opened in VS Code."
    } else {
      Start-Process -FilePath notepad.exe -ArgumentList "`"$path`"" | Out-Null
      Write-BBLog "VS Code not found; opened in Notepad."; Set-Status "VS Code not found; Notepad opened."
    }
  } catch {}
})

# ——— Persist LastPath on close
$form.Add_FormClosing({
    try {
        $settings.LastPath = $txtFolder.Text
        $settings | ConvertTo-Json | Set-Content $settingsPath -Encoding UTF8
    } catch {}
})

# -------------------------------------------------------------------
# 5) ALIGNMENT TWEAK (do this after the controls are created):
#    Content + Last 20 actions boxes’ right edges align with the VS Code button.
#    VS Code button: X=950, Width=160 → right edge = 1110; left margin is 20 → desired width ~1090.
# -------------------------------------------------------------------
$txtContent.Size     = New-Object System.Drawing.Size(1090,560)
$script:lstLog.Size  = New-Object System.Drawing.Size(1090,150)

# =========================
# TAB 1 — Log Upgrade Patch
# =========================
# Drop this block **after** Tab 1 has created $lblLog and $script:lstLog.
# Idempotent: safe to run more than once in a session.

# --- Global session log bootstrap (shared by all tabs) ----------------
if (-not (Get-Variable -Scope Global -Name BB2_LogFolder -ErrorAction SilentlyContinue)) {
  $Global:BB2_LogFolder = Join-Path $env:APPDATA 'BellBoy-BB-2\logs'
  if (-not (Test-Path -LiteralPath $Global:BB2_LogFolder)) {
    New-Item -ItemType Directory -Path $Global:BB2_LogFolder -Force | Out-Null
  }
}
if (-not (Get-Variable -Scope Global -Name BB2_SessionLogPath -ErrorAction SilentlyContinue)) {
  $ts = (Get-Date).ToString('yyyyMMdd-HHmmss')
  $Global:BB2_SessionLogPath = Join-Path $Global:BB2_LogFolder ("Session-{0}.log" -f $ts)
  # seed file
  "[{0}] Session started." -f (Get-Date).ToString('s') | Out-File -FilePath $Global:BB2_SessionLogPath -Encoding UTF8 -Append
}

# convenience local alias to satisfy existing code that referenced $sessionLog
$sessionLog = $Global:BB2_SessionLogPath

# --- UI: extend to 50 items + context menu + buttons ------------------
try { if ($lblLog) { $lblLog.Text = 'Last 50 actions:' } } catch {}
try {
  # listbox tuning
  $script:lstLog.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended
  $script:lstLog.HorizontalScrollbar = $true

  # right-click menu
  $logMenu = New-Object System.Windows.Forms.ContextMenuStrip
  $miCopySel  = $logMenu.Items.Add('Copy Selected')
  $miCopyAll  = $logMenu.Items.Add('Copy Visible')
  $miOpenFull = $logMenu.Items.Add('Open Full Session Log')
  $script:lstLog.ContextMenuStrip = $logMenu

  $copySelected = {
    try {
      $lines = @()
      foreach ($it in $script:lstLog.SelectedItems) { $lines += [string]$it }
      if ($lines.Count -eq 0) { return }
      [Windows.Forms.Clipboard]::SetText(($lines -join [Environment]::NewLine))
    } catch {}
  }
  $copyAllVisible = {
    try {
      $lines = @()
      foreach ($it in $script:lstLog.Items) { $lines += [string]$it }
      if ($lines.Count -eq 0) { return }
      [Windows.Forms.Clipboard]::SetText(($lines -join [Environment]::NewLine))
    } catch {}
  }
  $openFullLog = {
    try {
      if (Test-Path -LiteralPath $sessionLog) {
        Start-Process notepad.exe "`"$sessionLog`""
      }
    } catch {}
  }

  $miCopySel.Add_Click($copySelected)
  $miCopyAll.Add_Click($copyAllVisible)
  $miOpenFull.Add_Click($openFullLog)

  # Ctrl+C copies selected rows when the list has focus
  $script:lstLog.Add_KeyDown({
    if ($_.Control -and $_.KeyCode -eq 'C') { & $copySelected }
  })

  # buttons row under the log
  $btnCopyLog = New-Object System.Windows.Forms.Button
  $btnCopyLog.Text = 'Copy Log'
  $btnCopyLog.Size = New-Object System.Drawing.Size(110,28)
  $btnCopyLog.Location = New-Object System.Drawing.Point(20, 950)
  $btnCopyLog.Add_Click($copyAllVisible)
  $tab1.Controls.Add($btnCopyLog)

  $btnSaveLog = New-Object System.Windows.Forms.Button
  $btnSaveLog.Text = 'Save Log'
  $btnSaveLog.Size = New-Object System.Drawing.Size(110,28)
  $btnSaveLog.Location = New-Object System.Drawing.Point(140, 950)
  $btnSaveLog.Add_Click({
    try {
      $dlg = New-Object System.Windows.Forms.SaveFileDialog
      $dlg.Title = 'Save current log view'
      $dlg.Filter = 'Text file (*.txt)|*.txt|All files (*.*)|*.*'
      $dlg.FileName = 'BB2-Log-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.txt'
      if ($dlg.ShowDialog() -eq 'OK') {
        $lines = @()
        foreach ($it in $script:lstLog.Items) { $lines += [string]$it }
        $enc = New-Object System.Text.UTF8Encoding($true)
        [IO.File]::WriteAllText($dlg.FileName, ($lines -join [Environment]::NewLine), $enc)
      }
    } catch {}
  })
  $tab1.Controls.Add($btnSaveLog)

  $btnOpenLog = New-Object System.Windows.Forms.Button
  $btnOpenLog.Text = 'Open Log'
  $btnOpenLog.Size = New-Object System.Drawing.Size(110,28)
  $btnOpenLog.Location = New-Object System.Drawing.Point(260, 950)
  $btnOpenLog.Add_Click({
    try {
      if (Test-Path -LiteralPath $sessionLog) {
        Start-Process -FilePath explorer.exe -ArgumentList "/select,`"$sessionLog`""
      } else {
        Start-Process -FilePath explorer.exe -ArgumentList "`"$Global:BB2_LogFolder`""
      }
    } catch {}
  })
  $tab1.Controls.Add($btnOpenLog)

  $btnFullLog = New-Object System.Windows.Forms.Button
  $btnFullLog.Text = 'Full Session Log'
  $btnFullLog.Size = New-Object System.Drawing.Size(140,28)
  $btnFullLog.Location = New-Object System.Drawing.Point(380, 950)
  $btnFullLog.Add_Click($openFullLog)
  $tab1.Controls.Add($btnFullLog)
} catch {}

# --- logging helpers (override simple Write-BBLog) --------------------
function Add-RecentLogLine {
  param([string]$line)
  try {
    $prefix = '[{0}] ' -f (Get-Date -Format 'HH:mm:ss')
    $script:lstLog.Items.Insert(0, ($prefix + $line))
    # trim to 50
    while ($script:lstLog.Items.Count -gt 50) {
      $script:lstLog.Items.RemoveAt($script:lstLog.Items.Count - 1)
    }
  } catch {}
}
function Append-SessionLog {
  param([string]$line)
  try {
    $enc = New-Object System.Text.UTF8Encoding($true)
    $full = '[{0}] {1}{2}' -f (Get-Date).ToString('s'), $line, [Environment]::NewLine
    [IO.File]::AppendAllText($sessionLog, $full, $enc)
  } catch {}
}

# Replace/upgrade existing Write-BBLog (safe re-define)
Remove-Item Function:\Write-BBLog -ErrorAction SilentlyContinue
function Write-BBLog {
  param([string]$m)
  if ([string]::IsNullOrWhiteSpace($m)) { return }
  Add-RecentLogLine -line $m
  Append-SessionLog  -line $m
}

# preload last 50 from prior session file if the simple tail was used
try {
  if (Test-Path -LiteralPath $sessionLog) {
    $lines = Get-Content -LiteralPath $sessionLog -Tail 50
    foreach ($ln in $lines) { $script:lstLog.Items.Insert(0, $ln) }
    while ($script:lstLog.Items.Count -gt 50) { $script:lstLog.Items.RemoveAt($script:lstLog.Items.Count - 1) }
  }
} catch {}

# quick status ping so we see the upgrade took
Write-BBLog 'Log upgraded: 50-item view, context menu, buttons, and session log.'

# ===== End Tab 1 =====


# ===================================================================
# TAB 2 — Append / Inject  (PS 5.1-safe, full rewrite)
# ===================================================================

# --- Tiny helpers (safe if already defined elsewhere) ---
if (-not (Get-Command Ensure-Dir -ErrorAction SilentlyContinue)) {
    function Ensure-Dir { param([string]$p) if (-not (Test-Path -LiteralPath $p)) { New-Item -ItemType Directory -Path $p -Force | Out-Null } }
}
if (-not (Get-Command Backup-File-Fallback -ErrorAction SilentlyContinue)) {
    function Backup-File-Fallback {
        param([string]$Path)
        try {
            if (-not (Test-Path -LiteralPath $Path)) { return $null }
            $name = [IO.Path]::GetFileName($Path)
            $ts   = Get-Date -Format "yyyyMMdd_HHmmss"

            # Robust backups root chain: $BackupsDir → $RootForState\Backups → %APPDATA%\BellBoy-BB-2\Backups → $PSScriptRoot\Backups
            $appData = [Environment]::GetFolderPath('ApplicationData')
            $dir = if ($BackupsDir) {
                $BackupsDir
            } elseif ($RootForState) {
                Join-Path $RootForState 'Backups'
            } elseif ($appData) {
                Join-Path $appData 'BellBoy-BB-2\Backups'
            } else {
                Join-Path $PSScriptRoot 'Backups'
            }

            if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            $dst  = Join-Path $dir "$name.append.bak.$ts"
            Copy-Item -LiteralPath $Path -Destination $dst -Force
            return $dst
        } catch { return $null }
    }
}
if (-not (Get-Command Write-BBLog -ErrorAction SilentlyContinue)) { function Write-BBLog { param([string]$m) } }
if (-not (Get-Command Set-Status -ErrorAction SilentlyContinue)) { function Set-Status { param([string]$m) } }

# --- Make a text box view-only (no typing, no tab focus) ---
function Set-Viewer {
    param([System.Windows.Forms.Control]$c)
    if ($c -is [System.Windows.Forms.TextBoxBase]) {
        $c.ReadOnly         = $true
        $c.TabStop          = $false
        $c.ShortcutsEnabled = $false
        $c.DetectUrls       = $false
        $c.WordWrap         = $false
        $c.HideSelection    = $false
        $c.BorderStyle      = 'FixedSingle'
        $c.Font             = New-Object Drawing.Font('Consolas', 10)
        $c.Add_KeyDown({ $_.SuppressKeyPress = $true })
        $c.Add_KeyPress({ $_.Handled = $true })
    }
}

# --- Unified diff (line-based), safe formatting (no fragile one-liners) ---
function Get-UnifiedDiffText {
    param(
        [string[]]$Old,
        [string[]]$New,
        [int]$Context = 3
    )
    if (-not $Old) { $Old = @() }
    if (-not $New) { $New = @() }
    $m = $Old.Count
    $n = $New.Count

    # DP table for LCS
    $dp = New-Object 'int[,]' ($m+1), ($n+1)
    for ($i = $m - 1; $i -ge 0; $i--) {
        for ($j = $n - 1; $j -ge 0; $j--) {
            if ($Old[$i] -eq $New[$j]) {
                $dp[$i,$j] = 1 + $dp[$i+1,$j+1]
            } else {
                $a = $dp[$i+1,$j]
                $b = $dp[$i,$j+1]
                if ($a -ge $b) { $dp[$i,$j] = $a } else { $dp[$i,$j] = $b }
            }
        }
    }

    # Reconstruct edit script as lines prefixed with ' ', '+', '-'
    $i = 0; $j = 0
    $ops = New-Object System.Collections.Generic.List[string]
    while ($i -lt $m -or $j -lt $n) {
        if ($i -lt $m -and $j -lt $n -and $Old[$i] -eq $New[$j]) {
            $ops.Add(" " + $Old[$i])
            $i++; $j++; continue
        }
        if ($j -lt $n -and ($i -ge $m -or $dp[$i,$j+1] -ge $dp[$i+1,$j])) {
            $ops.Add("+" + $New[$j])
            $j++; continue
        }
        if ($i -lt $m) {
            $ops.Add("-" + $Old[$i])
            $i++; continue
        }
    }

    # Split into hunks by runs of unchanged lines
    $hunks = @()
    $cur = @()
    $eqRun = 0
    foreach ($l in $ops) {
        if ($l.StartsWith(" ")) {
            $eqRun++
            $cur += $l
            if ($eqRun -gt (2 * $Context)) {
                # Split hunk: keep last Context lines as tail, flush the rest
                $tail = @()
                for ($k=0; $k -lt $Context; $k++) { $tail = @($cur[-(1+$k)]) + $tail }
                $cut = $cur.Count - $Context
                if ($cut -gt 0) { $hunks += ,($cur[0..($cut-1)]) }
                $cur = @($tail)
                $eqRun = [Math]::Min($eqRun, $Context)
            }
        } else {
            $eqRun = 0
            $cur += $l
        }
    }
    if ($cur.Count -gt 0) { $hunks += ,$cur }

    # Emit with @@ headers (line numbers are approximate but consistent)
    $out = New-Object System.Collections.Generic.List[string]
    $aPos = 0; $bPos = 0
    foreach ($h in $hunks) {
        $aStart = $aPos; $bStart = $bPos
        foreach ($l in $h) {
            if     ($l.StartsWith(" ")) { $aPos++; $bPos++ }
            elseif ($l.StartsWith("+")) { $bPos++ }
            elseif ($l.StartsWith("-")) { $aPos++ }
        }
        $aLen = $aPos - $aStart
        $bLen = $bPos - $bStart
        $out.Add(("@@ -{0},{1} +{2},{3} @@" -f ($aStart+1), [Math]::Max(0,$aLen), ($bStart+1), [Math]::Max(0,$bLen)))
        $out.AddRange($h)
    }
    return ($out -join [Environment]::NewLine)
}

# --- Layout ---
$tab2 = New-Object System.Windows.Forms.TabPage
$tab2.Text = "Tab 2 — Append / Inject"
$tabs.TabPages.Add($tab2)

# Row 1: Target picker
$lblTgt = New-Object System.Windows.Forms.Label
$lblTgt.Text = "Target File:"
$lblTgt.AutoSize = $true
$lblTgt.Location = New-Object System.Drawing.Point(20,20)
$tab2.Controls.Add($lblTgt)

$txtTgt = New-Object System.Windows.Forms.TextBox
$txtTgt.Size = New-Object System.Drawing.Size(820,24)
$txtTgt.Location = New-Object System.Drawing.Point(110,18)
$txtTgt.Text = $settings.LastFile
$tab2.Controls.Add($txtTgt)

$btnPick = New-Object System.Windows.Forms.Button
$btnPick.Text = "Browse"
$btnPick.Size = New-Object System.Drawing.Size(90,26)
$btnPick.Location = New-Object System.Drawing.Point(940,16)
$tab2.Controls.Add($btnPick)

$ofd = New-Object System.Windows.Forms.OpenFileDialog
$btnPick.Add_Click({
    if ($ofd.ShowDialog() -eq 'OK') {
        $txtTgt.Text = $ofd.FileName
        try { $rtbOriginal.Text = (Get-Content -Raw -LiteralPath $ofd.FileName) } catch { $rtbOriginal.Text = "" }
        $settings.LastFile = $ofd.FileName
    }
})

# Row 2: Mode + line + hint + DryRun
$lblMode = New-Object System.Windows.Forms.Label
$lblMode.Text = "Insert Mode:"
$lblMode.AutoSize = $true
$lblMode.Location = New-Object System.Drawing.Point(20,54)
$tab2.Controls.Add($lblMode)

$cmbMode = New-Object System.Windows.Forms.ComboBox
$cmbMode.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void]$cmbMode.Items.AddRange(@('Append to end','Insert at line ...'))
$cmbMode.SelectedIndex = 0
$cmbMode.Size = New-Object System.Drawing.Size(200,24)
$cmbMode.Location = New-Object System.Drawing.Point(110,52)
$tab2.Controls.Add($cmbMode)

$lblLine = New-Object System.Windows.Forms.Label
$lblLine.Text = "Line:"
$lblLine.AutoSize = $true
$lblLine.Location = New-Object System.Drawing.Point(330,54)
$tab2.Controls.Add($lblLine)

$numLine = New-Object System.Windows.Forms.NumericUpDown
$numLine.Minimum = 1
$numLine.Maximum = 100000
$numLine.Value = 1
$numLine.Size = New-Object System.Drawing.Size(100,24)
$numLine.Location = New-Object System.Drawing.Point(370,52)
$tab2.Controls.Add($numLine)

$lblHint = New-Object System.Windows.Forms.Label
$lblHint.Text = "Tip: Click Original to set line automatically."
$lblHint.AutoSize = $true
$lblHint.Location = New-Object System.Drawing.Point(500,54)
$tab2.Controls.Add($lblHint)

$chkDryRun = New-Object System.Windows.Forms.CheckBox
$chkDryRun.Text = "Dry-Run (no writes)"
$chkDryRun.AutoSize = $true
$chkDryRun.Location = New-Object System.Drawing.Point(780,52)
$tab2.Controls.Add($chkDryRun)

# Row 3: Panes
$lblOrig = New-Object System.Windows.Forms.Label
$lblOrig.Text = "Original Preview"
$lblOrig.AutoSize = $true
$lblOrig.Location = New-Object System.Drawing.Point(20,88)
$tab2.Controls.Add($lblOrig)

$rtbOriginal = New-Object System.Windows.Forms.RichTextBox
$rtbOriginal.Font = New-Object System.Drawing.Font("Consolas", 10)
$rtbOriginal.WordWrap = $false
$rtbOriginal.ScrollBars = 'Both'
$rtbOriginal.Size = New-Object System.Drawing.Size(500,420)
$rtbOriginal.Location = New-Object System.Drawing.Point(20,110)
$tab2.Controls.Add($rtbOriginal)
Set-Viewer $rtbOriginal

# Click → set line number
$rtbOriginal.Add_MouseUp({
    $index = $rtbOriginal.SelectionStart
    $line  = $rtbOriginal.GetLineFromCharIndex($index) + 1
    $numLine.Value = [decimal]$line
})

$lblAppend = New-Object System.Windows.Forms.Label
$lblAppend.Text = "Append / Inject Text"
$lblAppend.AutoSize = $true
$lblAppend.Location = New-Object System.Drawing.Point(540,88)
$tab2.Controls.Add($lblAppend)

$rtbAppend = New-Object System.Windows.Forms.RichTextBox
$rtbAppend.Font = New-Object System.Drawing.Font("Consolas", 10)
$rtbAppend.WordWrap = $false
$rtbAppend.ScrollBars = 'Both'
$rtbAppend.Size = New-Object System.Drawing.Size(500,200)
$rtbAppend.Location = New-Object System.Drawing.Point(540,110)
$rtbAppend.Text = "# Your appended block goes here.`r`n# $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
$tab2.Controls.Add($rtbAppend)

$btnPreview = New-Object System.Windows.Forms.Button
$btnPreview.Text = "Generate Preview + Diff"
$btnPreview.Size = New-Object System.Drawing.Size(220,30)
$btnPreview.Location = New-Object System.Drawing.Point(540,320)
$tab2.Controls.Add($btnPreview)

$lblFinal = New-Object System.Windows.Forms.Label
$lblFinal.Text = "Final Preview"
$lblFinal.AutoSize = $true
$lblFinal.Location = New-Object System.Drawing.Point(540,360)
$tab2.Controls.Add($lblFinal)

$rtbFinal = New-Object System.Windows.Forms.RichTextBox
$rtbFinal.Font = New-Object System.Drawing.Font("Consolas", 10)
$rtbFinal.WordWrap = $false
$rtbFinal.ScrollBars = 'Both'
$rtbFinal.Size = New-Object System.Drawing.Size(500,170)
$rtbFinal.Location = New-Object System.Drawing.Point(540,382)
$tab2.Controls.Add($rtbFinal)
Set-Viewer $rtbFinal

$lblDiff = New-Object System.Windows.Forms.Label
$lblDiff.Text = "Unified Diff (+ added, - removed)"
$lblDiff.AutoSize = $true
$lblDiff.Location = New-Object System.Drawing.Point(20,560)
$tab2.Controls.Add($lblDiff)

$rtbDiff = New-Object System.Windows.Forms.RichTextBox
$rtbDiff.Font = New-Object System.Drawing.Font("Consolas", 10)
$rtbDiff.WordWrap = $false
$rtbDiff.ScrollBars = 'Both'
$rtbDiff.Size = New-Object System.Drawing.Size(1020,240)
$rtbDiff.Location = New-Object System.Drawing.Point(20,582)
$tab2.Controls.Add($rtbDiff)
Set-Viewer $rtbDiff

# Row 4: Actions
$btnDoAppend = New-Object System.Windows.Forms.Button
$btnDoAppend.Text = "Append / Inject (Backups Enabled)"
$btnDoAppend.Size = New-Object System.Drawing.Size(260,34)
$btnDoAppend.Location = New-Object System.Drawing.Point(20,834)
$tab2.Controls.Add($btnDoAppend)

$btnOpenTgt = New-Object System.Windows.Forms.Button
$btnOpenTgt.Text = "Open Target in Notepad"
$btnOpenTgt.Size = New-Object System.Drawing.Size(200,34)
$btnOpenTgt.Location = New-Object System.Drawing.Point(300,834)
$tab2.Controls.Add($btnOpenTgt)

$btnExplorerTgt = New-Object System.Windows.Forms.Button
$btnExplorerTgt.Text = "Reveal Target in Explorer"
$btnExplorerTgt.Size = New-Object System.Drawing.Size(200,34)
$btnExplorerTgt.Location = New-Object System.Drawing.Point(520,834)
$tab2.Controls.Add($btnExplorerTgt)

$btnUndoAppend = New-Object System.Windows.Forms.Button
$btnUndoAppend.Text = "Undo Last Write"
$btnUndoAppend.Size = New-Object System.Drawing.Size(160,34)
$btnUndoAppend.Location = New-Object System.Drawing.Point(740,834)
$tab2.Controls.Add($btnUndoAppend)
# Keep your existing:  $btnUndoAppend.Add_Click({ Restore-LastWrite })

Enable-TabAutoScroll $tab2

# --- SIMPLE & SAFE PREVIEW + DIFF (no multidim arrays) ---
function Build-Preview {
    param(
        [string]$Original,
        [string]$Chunk,
        [string]$Mode,
        [int]   $Line,
        [int]   $Context = 3
    )

    # Split lines (PS 5.1-safe)
    $origLines  = @(); if ($Original) { $origLines  = $Original -split "`r?`n" }
    $chunkLines = @(); if ($Chunk)    { $chunkLines = $Chunk    -split "`r?`n" }

    # Determine insertion index
    $idx =
        if ($Mode -eq 'Insert at line ...') {
            [Math]::Max(0, [Math]::Min($origLines.Count, [int]$Line - 1))
        } else {
            # Append to end
            $origLines.Count
        }

    # Build new text
    $head     = if ($idx -gt 0) { $origLines[0..($idx-1)] } else { @() }
    $tail     = if ($idx -lt $origLines.Count) { $origLines[$idx..($origLines.Count-1)] } else { @() }
    $newLines = @($head + $chunkLines + $tail)
    $finalText = ($newLines -join [Environment]::NewLine)

    # ----- Unified diff (context-before, +added, context-after) -----
    $ctxBeforeStart = [Math]::Max(0, $idx - $Context)
    $ctxBeforeEnd   = $idx - 1
    $ctxAfterStart  = [Math]::Min($newLines.Count, $idx + $chunkLines.Count)
    $ctxAfterEnd    = [Math]::Min($newLines.Count - 1, $ctxAfterStart + $Context - 1)

    $aStart = $idx            # old file “a”: insertion happens *before* this line (1-based in header)
    $aLen   = 0               # we don’t delete anything
    $bStart = $idx + 1        # new file “b”: added lines start here (1-based)
    $bLen   = $chunkLines.Count

    $diff = New-Object System.Collections.Generic.List[string]
    $diff.Add(("@@ -{0},{1} +{2},{3} @@" -f ($aStart+1), $aLen, $bStart, $bLen))

    # context before
    if ($ctxBeforeEnd -ge $ctxBeforeStart) {
        for ($i=$ctxBeforeStart; $i -le $ctxBeforeEnd; $i++) { $diff.Add(" " + $origLines[$i]) }
    }
    # added block
    foreach ($l in $chunkLines) { $diff.Add("+ " + $l) }
    # context after
    if ($ctxAfterEnd -ge $ctxAfterStart) {
        for ($i=$ctxAfterStart; $i -le $ctxAfterEnd; $i++) { $diff.Add(" " + $newLines[$i]) }
    }

    [PSCustomObject]@{
        Final = $finalText
        Diff  = ($diff -join [Environment]::NewLine)
    }
}

# Render diff with per-line colors (no patchy artifacts)
# Render diff with per-line colors (prevents patchy artifacts)
function Render-Diff {
    param([string]$Text)
    $rtbDiff.SuspendLayout()
    $rtbDiff.Clear()
    if ($null -eq $Text) { $rtbDiff.ResumeLayout(); return }

    $lines = $Text -split "`r?`n"
    foreach ($line in $lines) {
        if     ($line.StartsWith('@@')) { $rtbDiff.SelectionColor = [Drawing.Color]::SteelBlue }
        elseif ($line.StartsWith('+'))  { $rtbDiff.SelectionColor = [Drawing.Color]::ForestGreen }
        elseif ($line.StartsWith('-'))  { $rtbDiff.SelectionColor = [Drawing.Color]::Firebrick }
        else                            { $rtbDiff.SelectionColor = [Drawing.Color]::Black }
        $rtbDiff.AppendText($line + [Environment]::NewLine)
    }
    $rtbDiff.SelectionColor = [Drawing.Color]::Black
    $rtbDiff.ResumeLayout()
}


# --- Buttons wiring ---
$btnPreview.Add_Click({
    try {
        $target = $txtTgt.Text.Trim()
        $orig   = if (Test-Path -LiteralPath $target) { Get-Content -Raw -LiteralPath $target } else { "" }
        $mode   = $cmbMode.SelectedItem
        $line   = [int]$numLine.Value
        $chunk  = $rtbAppend.Text

        $p = Build-Preview -Original $orig -Chunk $chunk -Mode $mode -Line $line
        $rtbFinal.Text = $p.Final
        $rtbDiff.Text  = ""          # optional clear
        Render-Diff $p.Diff
        Set-Status "Preview generated."
        Write-BBLog "Preview generated for $target"
    } catch {
        [Windows.Forms.MessageBox]::Show("Preview error:`r`n$($_.Exception.Message)","Error") | Out-Null
    }
})

$btnDoAppend.Add_Click({
    try {
        $target = $txtTgt.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($target)) {
            [Windows.Forms.MessageBox]::Show("Choose a target file.","Input Error") | Out-Null
            return
        }

        $orig  = if (Test-Path -LiteralPath $target) { Get-Content -Raw -LiteralPath $target } else { "" }
        $mode  = $cmbMode.SelectedItem
        $line  = [int]$numLine.Value
        $chunk = $rtbAppend.Text

        # Always build preview/diff first
        $p = Build-Preview -Original $orig -Chunk $chunk -Mode $mode -Line $line
        $rtbFinal.Text = $p.Final
        $rtbDiff.Text  = ""
        Render-Diff $p.Diff

        if ($chkDryRun.Checked) {
            Write-BBLog "Dry-Run (no write) for $target ($mode, line $line)"
            Set-Status "Dry-Run only — no changes written."
            return
        }

        # Pre-append backup for Undo
        if (Test-Path -LiteralPath $target) {
            $bk = if (Get-Command Backup-File -ErrorAction SilentlyContinue) { Backup-File -Path $target } else { Backup-File-Fallback -Path $target }
            $script:LastBackupFile = $bk
            $script:LastTargetPath = $target
        }

        # Atomic write
        $enc = if (Test-Path -LiteralPath $target) { Get-TextEncoding -Path $target } else { 'UTF8' }
        $tmp = "$target.bbtmp"
        $p.Final | Set-Content -LiteralPath $tmp -Encoding $enc
        Move-Item -LiteralPath $tmp -Destination $target -Force

        # Refresh original view
        if (Test-Path -LiteralPath $target) { $rtbOriginal.Text = (Get-Content -Raw -LiteralPath $target) }

        Write-BBLog "Appended to $target ($mode, line $line)"
        Set-Status "Appended to $target"
        [Windows.Forms.MessageBox]::Show("Append complete.`r`n$target","Success") | Out-Null
    } catch {
        Write-BBLog ("ERROR (append): {0}" -f $_.Exception.Message)
        [Windows.Forms.MessageBox]::Show("Append error:`r`n$($_.Exception.Message)","Error") | Out-Null
    }
})

$btnOpenTgt.Add_Click({ try { if (Test-Path -LiteralPath $txtTgt.Text) { Start-Process notepad.exe "`"$($txtTgt.Text)`"" | Out-Null } } catch { } })
$btnExplorerTgt.Add_Click({
    try {
        if (Test-Path -LiteralPath $txtTgt.Text) { Start-Process explorer.exe "/select,`"$($txtTgt.Text)`"" }
        elseif ($txtTgt.Text) { $dir = Split-Path -Path $txtTgt.Text -Parent; if ($dir) { Start-Process explorer.exe "`"$dir`"" } }
    } catch { }
})

# --- Undo/Restore helpers (PS 5.1-safe) ---

function Get-BackupsRoot {
    if ($BackupsDir) { return $BackupsDir }
    if ($RootForState) { return (Join-Path $RootForState 'Backups') }
    # ultimate fallback under AppData
    $app = [Environment]::GetFolderPath('ApplicationData')
    return (Join-Path $app 'BellBoy-BB-2\Backups')
}

function Find-LatestBackupFor {
    param([Parameter(Mandatory)][string]$TargetPath)
    $root = Get-BackupsRoot
    if (-not (Test-Path -LiteralPath $root)) { return $null }
    $name = [IO.Path]::GetFileName($TargetPath)

    $candidates = Get-ChildItem -LiteralPath $root -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object {
            # match common patterns we create (prefix or contains filename)
            $_.Name -like "$name*" -or $_.Name -like "*$name*"
        } | Sort-Object LastWriteTime -Descending

    if ($candidates) { return $candidates[0].FullName }
    return $null
}

function Restore-LastWrite {
    try {
        # Prefer the last write’s explicit vars (set during Append/Tab 1 create)
        $target = $script:LastTargetPath
        $bak    = $script:LastBackupFile

        # If we don’t have them, try to infer from the UI field
        if (-not $target -and $txtTgt -and $txtTgt.Text) {
            $tp = $txtTgt.Text.Trim()
            if ($tp) { $target = $tp }
        }
        if (-not $target -or [string]::IsNullOrWhiteSpace($target)) {
            [Windows.Forms.MessageBox]::Show("Nothing to undo: target path unknown.","Undo") | Out-Null
            Set-Status "Undo: no known target."
            return
        }

        if (-not $bak -or -not (Test-Path -LiteralPath $bak)) {
            $bak = Find-LatestBackupFor -TargetPath $target
        }
        if (-not $bak -or -not (Test-Path -LiteralPath $bak)) {
            [Windows.Forms.MessageBox]::Show("No backup found for:`r`n$target","Undo") | Out-Null
            Set-Status "Undo: no backup found."
            Write-BBLog "Undo failed — no backup found for $target"
            return
        }

        # Atomic restore: copy backup to temp, then move into place
        $tmp = "$target.bbrestore.tmp"
        Copy-Item -LiteralPath $bak -Destination $tmp -Force
        Move-Item -LiteralPath $tmp -Destination $target -Force

        # Refresh previews if Tab 2 is visible
        try { if (Test-Path -LiteralPath $target) { $rtbOriginal.Text = (Get-Content -Raw -LiteralPath $target) } } catch { }

        Write-BBLog "Undo restored from backup → $bak → $target"
        Set-Status "Undo complete."
        [Windows.Forms.MessageBox]::Show("Restored from backup:`r`n$bak","Undo complete") | Out-Null
    } catch {
        Write-BBLog ("ERROR (undo): {0}" -f $_.Exception.Message)
        Set-Status "Undo error."
        [Windows.Forms.MessageBox]::Show("Undo error:`r`n$($_.Exception.Message)","Error") | Out-Null
    }
}

function Get-TextEncoding {
    param([Parameter(Mandatory)][string]$Path)
    try {
        $fs = [IO.File]::OpenRead($Path)
        try {
            $bom = New-Object byte[] 4
            $read = $fs.Read($bom,0,4)
        } finally { $fs.Dispose() }
        if ($read -ge 2 -and $bom[0] -eq 0xFF -and $bom[1] -eq 0xFE) { return 'Unicode' }            # UTF-16 LE
        if ($read -ge 2 -and $bom[0] -eq 0xFE -and $bom[1] -eq 0xFF) { return 'BigEndianUnicode' }   # UTF-16 BE
        if ($read -ge 3 -and $bom[0] -eq 0xEF -and $bom[1] -eq 0xBB -and $bom[2] -eq 0xBF) { return 'UTF8' } # UTF-8 BOM
        return 'Default'  # system ANSI if no BOM
    } catch { return 'UTF8' }
}

# ================================
# SHARED PAYLOAD VALIDATOR (TAB 2)
# Drop this once near your other helper functions.
# ================================
if (-not (Get-Command Test-BB2Payload -ErrorAction SilentlyContinue)) {
  function Test-BB2Payload {
    param(
      [string]$Text,
      [ValidateSet('auto','json','ps')] [string]$Kind = 'auto'
    )

    $result = [ordered]@{
      Kind    = $Kind
      IsValid = $true
      Message = ''
      Line    = $null
      Column  = $null
    }

    $t = if ($Text) { $Text.Trim() } else { '' }
    if ([string]::IsNullOrWhiteSpace($t)) {
      $result.Message = 'Payload is empty.'
      return [pscustomobject]$result
    }

    if ($Kind -eq 'auto') {
      $first = $t.Substring(0,1)
      if ($first -eq '{' -or $first -eq '[') {
        $Kind = 'json'
      } else {
        $Kind = 'ps'
      }
    }
    $result.Kind = $Kind

    switch ($Kind) {
      'json' {
        try {
          $null = $t | ConvertFrom-Json -ErrorAction Stop
          $result.Message = 'JSON looks OK.'
        } catch {
          $result.IsValid = $false
          $result.Message = $_.Exception.Message
        }
      }
      'ps' {
        $errs = @()
        $null = [System.Management.Automation.PSParser]::Tokenize($t, [ref]$errs)
        if ($errs.Count -gt 0) {
          $first = $errs[0]
          $result.IsValid = $false
          $result.Message = $first.Message
          $result.Line    = $first.Extent.StartLineNumber
          $result.Column  = $first.Extent.StartColumnNumber
        } else {
          $result.Message = 'PowerShell syntax looks OK.'
        }
      }
    }

    return [pscustomobject]$result
  }
}

# ==================================================
# TAB 2 – VALIDATE BUTTON (Append / Inject payload)
# Put this in the Tab 2 section, AFTER you create the
# "Append / Inject Text" TextBox.
# ==================================================

# IMPORTANT:
# Replace `$txtTab2Payload` on the right-hand side below
# with the actual variable name of your "Append / Inject Text"
# TextBox (the one shown in the screenshot).
#
# Example, if your control is $txtT2InsertText, do:
#   $script:Tab2PayloadBox = $txtT2InsertText
#
# Just set this alias once near the Tab 2 UI build.
$script:Tab2PayloadBox = $rtbAppend # <-- CHANGE right side to your real textbox

$btnTab2Validate = New-Object System.Windows.Forms.Button
$btnTab2Validate.Text = 'Validate Payload'
$btnTab2Validate.Size = New-Object System.Drawing.Size(130,28)
# Position: adjust if it overlaps something; this is a safe default.
$btnTab2Validate.Location = New-Object System.Drawing.Point(20, 260)
$tab2.Controls.Add($btnTab2Validate)

$btnTab2Validate.Add_Click({
  try {
    $payload = if ($script:Tab2PayloadBox) { $script:Tab2PayloadBox.Text } else { '' }
    $val     = Test-BB2Payload -Text $payload -Kind 'auto'

    if ($val.IsValid) {
      Write-BBLog ("Tab 2 validator OK ({0}): {1}" -f $val.Kind, $val.Message)
      [Windows.Forms.MessageBox]::Show(
        "Payload looks OK.`r`n`r`nKind: $($val.Kind)`r`n$($val.Message)",
        "Tab 2 – Validator"
      ) | Out-Null
    } else {
      $loc = if ($val.Line -and $val.Column) {
        "Line $($val.Line), Column $($val.Column)"
      } else {
        "Location unknown"
      }

      Write-BBLog ("Tab 2 validator ERROR ({0}): {1} [{2}]" -f $val.Kind, $val.Message, $loc)
      [Windows.Forms.MessageBox]::Show(
        "Validator found problems in the payload.`r`n`r`nKind: $($val.Kind)`r`nMessage: $($val.Message)`r`n$loc",
        "Tab 2 – Validator",
        [Windows.Forms.MessageBoxButtons]::OK,
        [Windows.Forms.MessageBoxIcon]::Warning
      ) | Out-Null
    }
  } catch {
    Write-BBLog ("Tab 2 validator threw: {0}" -f $_.Exception.Message)
    [Windows.Forms.MessageBox]::Show(
      "Validator error:`r`n$($_.Exception.Message)",
      "Tab 2 – Validator Error"
    ) | Out-Null
  }
})

# Ensure the button calls it (keep this where you wire buttons)
$btnUndoAppend.Add_Click({ Restore-LastWrite })


# ===== End Tab 2 =====


# ===================================================================
# TAB 3 — Code Inject  (recalibrated)
# ===================================================================

# --- tiny guards (won't re-declare if you already have them) ---
if (-not (Get-Command Set-Viewer -ErrorAction SilentlyContinue)) {
  function Set-Viewer {
    param([System.Windows.Forms.Control]$c)
    if ($c -is [System.Windows.Forms.TextBoxBase]) {
      $c.ReadOnly         = $true
      $c.TabStop          = $false
      $c.ShortcutsEnabled = $false
      $c.DetectUrls       = $false
      $c.WordWrap         = $false
      $c.HideSelection    = $false
      $c.BorderStyle      = 'FixedSingle'
      $c.Font             = New-Object Drawing.Font('Consolas', 10)
      $c.Add_KeyDown({ $_.SuppressKeyPress = $true })
      $c.Add_KeyPress({ $_.Handled = $true })
    }
  }
}
if (-not (Get-Command Render-Diff -ErrorAction SilentlyContinue)) {
  function Render-Diff {
    param([string]$Text)
    $rtbCIDiff.SuspendLayout()
    $rtbCIDiff.Clear()
    if ($null -eq $Text) { $rtbCIDiff.ResumeLayout(); return }
    foreach ($line in ($Text -split "`r?`n")) {
      if     ($line.StartsWith('@@')) { $rtbCIDiff.SelectionColor = [Drawing.Color]::SteelBlue }
      elseif ($line.StartsWith('+'))  { $rtbCIDiff.SelectionColor = [Drawing.Color]::ForestGreen }
      elseif ($line.StartsWith('-'))  { $rtbCIDiff.SelectionColor = [Drawing.Color]::Firebrick }
      else                            { $rtbCIDiff.SelectionColor = [Drawing.Color]::Black }
      $rtbCIDiff.AppendText($line + [Environment]::NewLine)
    }
    $rtbCIDiff.SelectionColor = [Drawing.Color]::Black
    $rtbCIDiff.ResumeLayout()
  }
}

# =========[ SPOT 1 — Add this helper ONCE, near your other tiny guards ]========
# (Place ABOVE "# --- layout ---" for Tab 3, right after Set-Viewer/Render-Diff.)
if (-not (Get-Command Get-TextEncoding -ErrorAction SilentlyContinue)) {
  function Get-TextEncoding {
    param([Parameter(Mandatory)][string]$Path)
    try {
      $fs = [IO.File]::OpenRead($Path)
      try {
        $bom = New-Object byte[] 4
        $read = $fs.Read($bom,0,4)
      } finally { $fs.Dispose() }
      if ($read -ge 2 -and $bom[0] -eq 0xFF -and $bom[1] -eq 0xFE) { return 'Unicode' }           # UTF-16 LE
      if ($read -ge 2 -and $bom[0] -eq 0xFE -and $bom[1] -eq 0xFF) { return 'BigEndianUnicode' }  # UTF-16 BE
      if ($read -ge 3 -and $bom[0] -eq 0xEF -and $bom[1] -eq 0xBB -and $bom[2] -eq 0xBF) { return 'UTF8' } # UTF-8 BOM
      return 'Default'  # ANSI if no BOM
    } catch { return 'UTF8' }
  }
}

# --- 1) Helper (add once near other tiny guards) -------------------
if (-not (Get-Command Resolve-CodeExe -ErrorAction SilentlyContinue)) {
  function Resolve-CodeExe {
    try {
      $cmd = Get-Command code -ErrorAction SilentlyContinue
      if ($cmd -and $cmd.Source) { return $cmd.Source }
    } catch {}

    $candidates = @(
      "$env:LOCALAPPDATA\Programs\Microsoft VS Code\Code.exe",
      "$env:ProgramFiles\Microsoft VS Code\Code.exe",
      "$env:ProgramFiles(x86)\Microsoft VS Code\Code.exe",
      "$env:LOCALAPPDATA\Programs\Microsoft VS Code Insiders\Code - Insiders.exe",
      "$env:ProgramFiles\Microsoft VS Code Insiders\Code - Insiders.exe",
      "$env:ProgramFiles(x86)\Microsoft VS Code Insiders\Code - Insiders.exe"
    )
    foreach ($p in $candidates) { if ($p -and (Test-Path -LiteralPath $p)) { return $p } }
    return $null
  }
}

# Wipe ANY old versions that may have multiple parameter sets


function Write-BB2Tab3TextAtomic {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Content
    )

    $folder = Split-Path -LiteralPath $Path -Parent
    if (-not (Test-Path -LiteralPath $folder)) {
        throw "Target folder does not exist: $folder"
    }

    # Decide encoding based on existing file (if any)
    $enc = $null
    try {
        if (Test-Path -LiteralPath $Path) {
            switch (Get-TextEncoding -Path $Path) {
                'Unicode'          { $enc = [System.Text.Encoding]::Unicode }
                'BigEndianUnicode' { $enc = [System.Text.Encoding]::BigEndianUnicode }
                'Default'          { $enc = [System.Text.Encoding]::Default }
                'UTF8'             { $enc = New-Object System.Text.UTF8Encoding($false) } # UTF-8, no BOM
                default            { $enc = New-Object System.Text.UTF8Encoding($false) }
            }
        } else {
            $enc = New-Object System.Text.UTF8Encoding($false)
        }
    } catch {
        $enc = New-Object System.Text.UTF8Encoding($false)
    }

    $rand    = [System.IO.Path]::GetRandomFileName()
    $tmpName = '.' + [System.IO.Path]::GetFileName($Path) + '.' + $rand + '.tmp'
    $tmp     = Join-Path $folder $tmpName

    try {
        # Write temp file
        [System.IO.File]::WriteAllText($tmp, $Content, $enc)

        # Atomic-ish replace: if target exists, Replace; else Move
        if (Test-Path -LiteralPath $Path) {
            [System.IO.File]::Replace($tmp, $Path, $null)
        } else {
            [System.IO.File]::Move($tmp, $Path)
        }
    }
    finally {
        if ($tmp -and (Test-Path -LiteralPath $tmp)) {
            Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
        }
    }
}

# ===================================================================
# TAB 3 — add this helper NEAR YOUR OTHER TINY GUARDS (Set-Viewer, Render-Diff, Get-TextEncoding)
# Make sure you only have ONE Get-TextEncoding in the file.
# ===================================================================

function Invoke-BB2Tab3AtomicWrite {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Content
    )

    $folder = Split-Path -LiteralPath $Path -Parent
    if (-not (Test-Path -LiteralPath $folder)) {
        throw "Target folder does not exist: $folder"
    }

    # Guard: read-only target
    if (Test-Path -LiteralPath $Path) {
        try {
            $fi = Get-Item -LiteralPath $Path -ErrorAction Stop
            if ($fi.Attributes -band [System.IO.FileAttributes]::ReadOnly) {
                throw "Target file is marked Read-Only. Clear the attribute first, then retry."
            }
        } catch {
            # If we can't inspect attributes, fall through and let the write error speak
        }
    }

    # Decide encoding based on existing file (if any), using Get-TextEncoding
    $enc = $null
    try {
        if (Test-Path -LiteralPath $Path) {
            $encName = Get-TextEncoding -Path $Path
            switch ($encName) {
                'Unicode'          { $enc = [System.Text.Encoding]::Unicode }
                'BigEndianUnicode' { $enc = [System.Text.Encoding]::BigEndianUnicode }
                'Default'          { $enc = [System.Text.Encoding]::Default }
                'UTF8'             { $enc = New-Object System.Text.UTF8Encoding($false) } # UTF-8 no BOM
                default            { $enc = New-Object System.Text.UTF8Encoding($false) }
            }
        } else {
            $enc = New-Object System.Text.UTF8Encoding($false)
        }
    } catch {
        $enc = New-Object System.Text.UTF8Encoding($false)
    }

    $rand    = [System.IO.Path]::GetRandomFileName()
    $tmpName = '.' + ([System.IO.Path]::GetFileName($Path)) + '.' + $rand + '.tmp'
    $tmp     = Join-Path -Path $folder -ChildPath $tmpName

    try {
        # Write the temp file with the chosen encoding
        [System.IO.File]::WriteAllText($tmp, $Content, $enc)

        # Atomic-ish replace: Replace when target exists, else Move
        if (Test-Path -LiteralPath $Path) {
            [System.IO.File]::Replace($tmp, $Path, $null)
        } else {
            [System.IO.File]::Move($tmp, $Path)
        }
    } catch {
        $ex  = $_.Exception
        $msg = $ex.Message
        $hr  = $ex.HResult

        # Common Windows sharing-violation HRESULTs → friendlier message
        if ($hr -eq 0x80070020 -or $hr -eq 0x80070021) {
            throw "Target file appears to be locked by another process. Close any editors or tools holding it open, then retry.`r`n($msg)"
        }

        throw $msg
    }
    finally {
        if ($tmp -and (Test-Path -LiteralPath $tmp)) {
            Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
        }
    }
}

# --- layout --------------------------------------------------------
$tab3 = New-Object System.Windows.Forms.TabPage
$tab3.Text = "Tab 3 — Code Inject"
$tabs.TabPages.Add($tab3)

# Row 1: target
$lblCIFile = New-Object System.Windows.Forms.Label
$lblCIFile.Text = "Target File:"
$lblCIFile.AutoSize = $true
$lblCIFile.Location = New-Object System.Drawing.Point(20,20)
$tab3.Controls.Add($lblCIFile)

$txtCIFile = New-Object System.Windows.Forms.TextBox
$txtCIFile.Size = New-Object System.Drawing.Size(820,24)
$txtCIFile.Location = New-Object System.Drawing.Point(110,18)
$txtCIFile.Text = $settings.LastFile
$tab3.Controls.Add($txtCIFile)

$btnCIBrowse = New-Object System.Windows.Forms.Button
$btnCIBrowse.Text = "Browse"
$btnCIBrowse.Size = New-Object System.Drawing.Size(90,26)
$btnCIBrowse.Location = New-Object System.Drawing.Point(940,16)
$tab3.Controls.Add($btnCIBrowse)
$ofd3 = New-Object System.Windows.Forms.OpenFileDialog
$btnCIBrowse.Add_Click({
  if ($ofd3.ShowDialog() -eq 'OK') {
    $txtCIFile.Text = $ofd3.FileName
    try { $rtbCIOriginal.Text = (Get-Content -Raw -LiteralPath $ofd3.FileName) } catch { $rtbCIOriginal.Text = "" }
    $settings.LastFile = $ofd3.FileName
  }
})

# Row 2: mode + options
$lblCIMode = New-Object System.Windows.Forms.Label
$lblCIMode.Text = "Mode:"
$lblCIMode.AutoSize = $true
$lblCIMode.Location = New-Object System.Drawing.Point(20,54)
$tab3.Controls.Add($lblCIMode)

$cmbCIMode = New-Object System.Windows.Forms.ComboBox
$cmbCIMode.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void]$cmbCIMode.Items.AddRange(@(
  'Append to end',
  'After Pattern',
  'Before Pattern',
  'Replace Pattern',
  'Inside Region (Begin/End)'
))
$cmbCIMode.SelectedIndex = 0
$cmbCIMode.Size = New-Object System.Drawing.Size(220,24)
$cmbCIMode.Location = New-Object System.Drawing.Point(70,52)
$tab3.Controls.Add($cmbCIMode)

$chkCIRegex = New-Object System.Windows.Forms.CheckBox
$chkCIRegex.Text = "Regex"
$chkCIRegex.Checked = $true
$chkCIRegex.AutoSize = $true
$chkCIRegex.Location = New-Object System.Drawing.Point(310,54)
$tab3.Controls.Add($chkCIRegex)

$chkCICase = New-Object System.Windows.Forms.CheckBox
$chkCICase.Text = "Case-insensitive"
$chkCICase.Checked = $true
$chkCICase.AutoSize = $true
$chkCICase.Location = New-Object System.Drawing.Point(380,54)
$tab3.Controls.Add($chkCICase)

$chkCIDry = New-Object System.Windows.Forms.CheckBox
$chkCIDry.Text = "Dry-Run (no writes)"
$chkCIDry.AutoSize = $true
$chkCIDry.Location = New-Object System.Drawing.Point(520,54)
$tab3.Controls.Add($chkCIDry)

$lblCIPat = New-Object System.Windows.Forms.Label
$lblCIPat.Text = "Pattern:"
$lblCIPat.AutoSize = $true
$lblCIPat.Location = New-Object System.Drawing.Point(20,86)
$tab3.Controls.Add($lblCIPat)

$txtCIPat = New-Object System.Windows.Forms.TextBox
$txtCIPat.Size = New-Object System.Drawing.Size(900,24)
$txtCIPat.Location = New-Object System.Drawing.Point(80,84)
$tab3.Controls.Add($txtCIPat)

$lblCIBegin = New-Object System.Windows.Forms.Label
$lblCIBegin.Text = "Begin:"
$lblCIBegin.AutoSize = $true
$lblCIBegin.Location = New-Object System.Drawing.Point(20,116)
$tab3.Controls.Add($lblCIBegin)

$txtCIBegin = New-Object System.Windows.Forms.TextBox
$txtCIBegin.Size = New-Object System.Drawing.Size(430,24)
$txtCIBegin.Location = New-Object System.Drawing.Point(70,114)
$tab3.Controls.Add($txtCIBegin)

$lblCIEnd = New-Object System.Windows.Forms.Label
$lblCIEnd.Text = "End:"
$lblCIEnd.AutoSize = $true
$lblCIEnd.Location = New-Object System.Drawing.Point(520,116)
$tab3.Controls.Add($lblCIEnd)

$txtCIEnd = New-Object System.Windows.Forms.TextBox
$txtCIEnd.Size = New-Object System.Drawing.Size(460,24)
$txtCIEnd.Location = New-Object System.Drawing.Point(560,114)
$tab3.Controls.Add($txtCIEnd)

# Row 3: editors
$lblCIOrig = New-Object System.Windows.Forms.Label
$lblCIOrig.Text = "Original (preview)"
$lblCIOrig.AutoSize = $true
$lblCIOrig.Location = New-Object System.Drawing.Point(20,148)
$tab3.Controls.Add($lblCIOrig)

$rtbCIOriginal = New-Object System.Windows.Forms.RichTextBox
$rtbCIOriginal.Font = New-Object System.Drawing.Font("Consolas", 10)
$rtbCIOriginal.WordWrap = $false
$rtbCIOriginal.ScrollBars = 'Both'
$rtbCIOriginal.Size = New-Object System.Drawing.Size(500,420)
$rtbCIOriginal.Location = New-Object System.Drawing.Point(20,170)
$tab3.Controls.Add($rtbCIOriginal)
Set-Viewer $rtbCIOriginal

$lblCISnip = New-Object System.Windows.Forms.Label
$lblCISnip.Text = "Snippet to inject"
$lblCISnip.AutoSize = $true
$lblCISnip.Location = New-Object System.Drawing.Point(540,148)
$tab3.Controls.Add($lblCISnip)

$rtbCISnippet = New-Object System.Windows.Forms.RichTextBox
$rtbCISnippet.Font = New-Object System.Drawing.Font("Consolas", 10)
$rtbCISnippet.WordWrap = $false
$rtbCISnippet.ScrollBars = 'Both'
$rtbCISnippet.Size = New-Object System.Drawing.Size(500,220)
$rtbCISnippet.Location = New-Object System.Drawing.Point(540,170)
$rtbCISnippet.Text = "# injected by BB-2 on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
$tab3.Controls.Add($rtbCISnippet)
$rtbCISnippet.ShortcutsEnabled = $true

$cmsSnippet  = New-Object System.Windows.Forms.ContextMenuStrip
$miCut       = $cmsSnippet.Items.Add('Cut')
$miCopy      = $cmsSnippet.Items.Add('Copy')
$miPaste     = $cmsSnippet.Items.Add('Paste')
$miDelete    = $cmsSnippet.Items.Add('Delete')
$cmsSnippet.Items.Add('-') | Out-Null
$miSelectAll = $cmsSnippet.Items.Add('Select All')
$rtbCISnippet.ContextMenuStrip = $cmsSnippet

# Enable/disable based on selection & clipboard
$cmsSnippet.add_Opening({
    $hasSel = ($rtbCISnippet.SelectionLength -gt 0)
    $miCut.Enabled    = $hasSel
    $miCopy.Enabled   = $hasSel
    $miDelete.Enabled = $hasSel
    $miPaste.Enabled  = [System.Windows.Forms.Clipboard]::ContainsText()
})

# Actions
$miCut.Add_Click({ if ($rtbCISnippet.SelectionLength -gt 0) { $rtbCISnippet.Cut() } })
$miCopy.Add_Click({ if ($rtbCISnippet.SelectionLength -gt 0) { $rtbCISnippet.Copy() } })
$miPaste.Add_Click({ $rtbCISnippet.Paste() })
$miDelete.Add_Click({ if ($rtbCISnippet.SelectionLength -gt 0) { $rtbCISnippet.SelectedText = '' } })
$miSelectAll.Add_Click({ $rtbCISnippet.SelectAll() })


$btnCIPreview = New-Object System.Windows.Forms.Button
$btnCIPreview.Text = "Generate Preview + Diff"
$btnCIPreview.Size = New-Object System.Drawing.Size(220,30)
$btnCIPreview.Location = New-Object System.Drawing.Point(540,400)
$tab3.Controls.Add($btnCIPreview)

$lblCIFinal = New-Object System.Windows.Forms.Label
$lblCIFinal.Text = "Final Preview"
$lblCIFinal.AutoSize = $true
$lblCIFinal.Location = New-Object System.Drawing.Point(540,438)
$tab3.Controls.Add($lblCIFinal)

$rtbCIFinal = New-Object System.Windows.Forms.RichTextBox
$rtbCIFinal.Font = New-Object System.Drawing.Font("Consolas", 10)
$rtbCIFinal.WordWrap = $false
$rtbCIFinal.ScrollBars = 'Both'
$rtbCIFinal.Size = New-Object System.Drawing.Size(500,152)
$rtbCIFinal.Location = New-Object System.Drawing.Point(540,460)
$tab3.Controls.Add($rtbCIFinal)
Set-Viewer $rtbCIFinal

$lblCIDiff = New-Object System.Windows.Forms.Label
$lblCIDiff.Text = "Unified Diff (+ added, - removed)"
$lblCIDiff.AutoSize = $true
$lblCIDiff.Location = New-Object System.Drawing.Point(20,592)
$tab3.Controls.Add($lblCIDiff)

$rtbCIDiff = New-Object System.Windows.Forms.RichTextBox
$rtbCIDiff.Font = New-Object System.Drawing.Font("Consolas", 10)
$rtbCIDiff.WordWrap = $false
$rtbCIDiff.ScrollBars = 'Both'
$rtbCIDiff.Size = New-Object System.Drawing.Size(1020,240)
$rtbCIDiff.Location = New-Object System.Drawing.Point(20,614)
$tab3.Controls.Add($rtbCIDiff)
Set-Viewer $rtbCIDiff

# Row 4: actions
$btnCIInject = New-Object System.Windows.Forms.Button
$btnCIInject.Text = "Inject (Backups Enabled)"
$btnCIInject.Size = New-Object System.Drawing.Size(220,34)
$btnCIInject.Location = New-Object System.Drawing.Point(20,866)
$tab3.Controls.Add($btnCIInject)

$btnCIReveal = New-Object System.Windows.Forms.Button
$btnCIReveal.Text = "Reveal Target in Explorer"
$btnCIReveal.Size = New-Object System.Drawing.Size(200,34)
$btnCIReveal.Location = New-Object System.Drawing.Point(260,866)
$tab3.Controls.Add($btnCIReveal)

$btnCIOpen = New-Object System.Windows.Forms.Button
$btnCIOpen.Text = "Open Target in Notepad"
$btnCIOpen.Size = New-Object System.Drawing.Size(200,34)
$btnCIOpen.Location = New-Object System.Drawing.Point(480,866)
$tab3.Controls.Add($btnCIOpen)

$btnCIUndo = New-Object System.Windows.Forms.Button
$btnCIUndo.Text = "Undo Last Write"
$btnCIUndo.Size = New-Object System.Drawing.Size(160,34)
$btnCIUndo.Location = New-Object System.Drawing.Point(700,866)
$tab3.Controls.Add($btnCIUndo)
$btnCIUndo.Add_Click({ Restore-LastWrite })

$btnCICode = New-Object System.Windows.Forms.Button
$btnCICode.Text = "Open in VS Code"
$btnCICode.Size = New-Object System.Drawing.Size(160,34)
$btnCICode.Location = New-Object System.Drawing.Point(870,866)
$tab3.Controls.Add($btnCICode)

$btnCICode.Add_Click({
  try {
    $path = $txtCIFile.Text.Trim()
    if (-not (Test-Path -LiteralPath $path)) {
      [Windows.Forms.MessageBox]::Show("File does not exist yet.","Info") | Out-Null
      return
    }
    $code = Resolve-CodeExe
    if ($code) {
      Start-Process -FilePath $code -ArgumentList "`"$path`""
      Write-BBLog "Opened in VS Code."; Set-Status "Opened in VS Code."
    } else {
      [Windows.Forms.MessageBox]::Show("VS Code not found. Falling back to Notepad.","Info") | Out-Null
      Start-Process -FilePath notepad.exe -ArgumentList "`"$path`"" | Out-Null
      Write-BBLog "VS Code not found; opened in Notepad."
    }
  } catch {}
})

Enable-TabAutoScroll $tab3

# ---------------- core preview/insert engine -----------------------
function Get-LineEnding([string]$text) {
  if ($text -match "`r`n") { "`r`n" } elseif ($text -match "`n") { "`n" } else { "`r`n" }
}
function CI-FindSpan {
  param(
    [string]$Text,[string]$Pattern,[string]$Begin,[string]$End,
    [string]$Mode,[bool]$UseRegex,[bool]$CaseInsensitive
  )
  $RegexType = [System.Text.RegularExpressions.Regex]
  $opts = [System.Text.RegularExpressions.RegexOptions]::None
  if ($CaseInsensitive) { $opts = $opts -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase }
  if (-not $UseRegex) {
    $Pattern = $RegexType::Escape($Pattern)
    $Begin   = $RegexType::Escape($Begin)
    $End     = $RegexType::Escape($End)
  }

  switch ($Mode) {
    'After Pattern'  {
      $m = $RegexType::Match($Text, $Pattern, $opts)
      if (-not $m.Success) { return @{ Found=$false } }
      return @{ Found=$true; InsertAt=$m.Index + $m.Length }
    }
    'Before Pattern' {
      $m = $RegexType::Match($Text, $Pattern, $opts)
      if (-not $m.Success) { return @{ Found=$false } }
      return @{ Found=$true; InsertAt=$m.Index }
    }
    'Replace Pattern' {
      $m = $RegexType::Match($Text, $Pattern, $opts)
      if (-not $m.Success) { return @{ Found=$false } }
      return @{ Found=$true; ReplaceStart=$m.Index; ReplaceLen=$m.Length }
    }
    'Inside Region (Begin/End)' {
      $mb = $RegexType::Match($Text, $Begin, $opts)
      if (-not $mb.Success) { return @{ Found=$false } }

      $rxEnd = New-Object System.Text.RegularExpressions.Regex($End, $opts)
      $me    = $rxEnd.Match($Text, $mb.Index + $mb.Length)
      if (-not $me.Success) { return @{ Found=$false } }

      # *** NEW: snap RegionStart to the start of the NEXT LINE after BEGIN ***
      $regionStart = $mb.Index + $mb.Length
      # Look for CRLF first, then LF, but never move past the END marker.
      $crlf = $Text.IndexOf("`r`n", $regionStart)
      $lf   = $Text.IndexOf("`n",   $regionStart)
      if     ($crlf -ge 0 -and $crlf -lt $me.Index) { $regionStart = $crlf + 2 }
      elseif ($lf   -ge 0 -and $lf   -lt $me.Index) { $regionStart = $lf   + 1 }

      return @{ Found=$true; RegionStart=$regionStart; RegionEnd=$me.Index }
    }
    default {
      return @{ Found=$true; InsertAt=$Text.Length } # Append to end
    }
  }
}

# IMPORTANT: we avoid the old DP-table diff that threw index errors.
# We generate a simple unified-diff showing inserted block + a little context.
function CI-BuildPreview {
  param(
    [string]$Original,[string]$Snippet,[string]$Mode,
    [string]$Pattern,[string]$Begin,[string]$End,
    [bool]$UseRegex,[bool]$CaseInsensitive,
    [int]$Context = 3
  )
  $origText = if ($Original) { $Original } else { "" }
  $newline  = Get-LineEnding $origText
  $snipText = [string]($Snippet -replace "`r?`n", $newline)
  if ($snipText.Length -gt 0 -and (-not $snipText.EndsWith($newline))) { $snipText += $newline }

  $span = CI-FindSpan -Text $origText -Pattern $Pattern -Begin $Begin -End $End `
                      -Mode $Mode -UseRegex $UseRegex -CaseInsensitive $CaseInsensitive
  if (-not $span.Found) {
    return [PSCustomObject]@{ Final=$origText; Diff="No match found."; Note="No match found." }
  }

  switch ($Mode) {
    'Replace Pattern' {
      $pre  = $origText.Substring(0, $span.ReplaceStart)
      $post = $origText.Substring($span.ReplaceStart + $span.ReplaceLen)
      $final= $pre + $snipText + $post
      $insLine = (($pre -split "`r?`n").Count) + 1
      $added   = $snipText -split "`r?`n"; if ($added[-1] -eq '') { $added = $added[0..($added.Count-2)] }
    }
    'Inside Region (Begin/End)' {
      $pre   = $origText.Substring(0, $span.RegionStart)
      $post  = $origText.Substring($span.RegionEnd)
      $final = $pre + $snipText + $post
      $insLine = (($pre -split "`r?`n").Count) + 1
      $added   = $snipText -split "`r?`n"; if ($added[-1] -eq '') { $added = $added[0..($added.Count-2)] }
    }
    'After Pattern' {
      $pre  = $origText.Substring(0, $span.InsertAt)
      $post = $origText.Substring($span.InsertAt)
      $final= $pre + $snipText + $post
      $insLine = (($pre -split "`r?`n").Count) + 1
      $added   = $snipText -split "`r?`n"; if ($added[-1] -eq '') { $added = $added[0..($added.Count-2)] }
    }
    'Before Pattern' {
      $pre  = $origText.Substring(0, $span.InsertAt)
      $post = $origText.Substring($span.InsertAt)
      $final= $pre + $snipText + $post
      $insLine = (($pre -split "`r?`n").Count) + 1
      $added   = $snipText -split "`r?`n"; if ($added[-1] -eq '') { $added = $added[0..($added.Count-2)] }
    }
    default {
      $final= $origText + $snipText
      $insLine = (($origText -split "`r?`n").Count) + 1
      $added   = $snipText -split "`r?`n"; if ($added[-1] -eq '') { $added = $added[0..($added.Count-2)] }
    }
  }

  $oldLines = if ($origText) { $origText -split "`r?`n" } else { @() }
  $newLines = if ($final)    { $final    -split "`r?`n" } else { @() }

  # Build minimal unified diff header + context
  $ctxBeforeStart = [Math]::Max(0, $insLine - $Context - 1)
  $ctxBeforeEnd   = [Math]::Max(-1, $insLine - 2)
  $ctxAfterStart  = [Math]::Min($newLines.Count, $insLine - 1 + $added.Count)
  $ctxAfterEnd    = [Math]::Min($newLines.Count - 1, $ctxAfterStart + $Context - 1)

  $aStart = $insLine - 1;  $aLen = 0
  $bStart = $insLine;      $bLen = $added.Count

  $diff = New-Object System.Collections.Generic.List[string]
  $diff.Add(("@@ -{0},{1} +{2},{3} @@" -f ($aStart+1), $aLen, $bStart, $bLen))
  if ($ctxBeforeEnd -ge $ctxBeforeStart) { for ($i=$ctxBeforeStart; $i -le $ctxBeforeEnd; $i++) { $diff.Add(" " + $newLines[$i]) } }
  foreach ($l in $added) { $diff.Add("+ " + $l) }
  if ($ctxAfterEnd -ge $ctxAfterStart)  { for ($i=$ctxAfterStart;  $i -le $ctxAfterEnd;  $i++) { $diff.Add(" " + $newLines[$i]) } }

  [PSCustomObject]@{ Final=$final; Diff=($diff -join [Environment]::NewLine) }
}

# Tab 3: render unified diff into $rtbCIDiff (write first, then colorize) — PS 5.1-safe
function Render-Diff3 {
    param([string]$Text)

    if (-not $rtbCIDiff -or $rtbCIDiff.IsDisposed) { return }

    if ($null -eq $Text) { $Text = "" }
    $rtbCIDiff.Text = $Text

    $rtbCIDiff.SuspendLayout()
    $rtbCIDiff.SelectAll(); $rtbCIDiff.SelectionColor = [System.Drawing.Color]::Black
    $rtbCIDiff.DeselectAll()

    $pos = 0
    foreach ($line in $rtbCIDiff.Lines) {
        $len = $line.Length + [Environment]::NewLine.Length
        if     ($line.StartsWith('@@')) { $rtbCIDiff.Select($pos,$line.Length); $rtbCIDiff.SelectionColor = [System.Drawing.Color]::SteelBlue }
        elseif ($line.StartsWith('+'))  { $rtbCIDiff.Select($pos,$line.Length); $rtbCIDiff.SelectionColor = [System.Drawing.Color]::ForestGreen }
        elseif ($line.StartsWith('-'))  { $rtbCIDiff.Select($pos,$line.Length); $rtbCIDiff.SelectionColor = [System.Drawing.Color]::Firebrick  }
        $pos += $len
    }
    $rtbCIDiff.DeselectAll()
    $rtbCIDiff.ResumeLayout()
}




# ------------------------ button wiring ----------------------------
$btnCIPreview.Add_Click({
  try {
    $path = $txtCIFile.Text.Trim()
    $orig = if (Test-Path -LiteralPath $path) { Get-Content -Raw -LiteralPath $path } else { "" }
    $p = CI-BuildPreview -Original $orig -Snippet $rtbCISnippet.Text -Mode $cmbCIMode.SelectedItem `
         -Pattern $txtCIPat.Text -Begin $txtCIBegin.Text -End $txtCIEnd.Text `
         -UseRegex ([bool]$chkCIRegex.Checked) -CaseInsensitive ([bool]$chkCICase.Checked)
    $rtbCIFinal.Text = $p.Final
    Render-Diff3 $p.Diff
    Set-Status "Tab3: Preview generated."
    Write-BBLog "Tab3 Preview for $path"
  } catch {
    [Windows.Forms.MessageBox]::Show("Preview error:`r`n$($_.Exception.Message)","Error") | Out-Null
  }
})

# -------------------------------------------------------------------
# TAB 3 SAFE WRITER (UNIQUE NAME, PURE .NET, PS 5.1-SAFE)
# Drop this in place of the old "Write-BB2Tab3TextAtomic" helper block.
# -------------------------------------------------------------------
function Invoke-BB2Tab3AtomicWrite {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Content
    )

    $folder = Split-Path -LiteralPath $Path -Parent
    if (-not (Test-Path -LiteralPath $folder)) {
        throw "Target folder does not exist: $folder"
    }

    # Decide encoding based on existing file (if any), using Get-TextEncoding
    $enc = $null
    try {
        if (Test-Path -LiteralPath $Path) {
            switch (Get-TextEncoding -Path $Path) {
                'Unicode'          { $enc = [System.Text.Encoding]::Unicode }
                'BigEndianUnicode' { $enc = [System.Text.Encoding]::BigEndianUnicode }
                'Default'          { $enc = [System.Text.Encoding]::Default }
                'UTF8'             { $enc = New-Object System.Text.UTF8Encoding($false) } # UTF-8 no BOM
                default            { $enc = New-Object System.Text.UTF8Encoding($false) }
            }
        } else {
            $enc = New-Object System.Text.UTF8Encoding($false)
        }
    } catch {
        $enc = New-Object System.Text.UTF8Encoding($false)
    }

    $rand    = [IO.Path]::GetRandomFileName()
    $tmpName = '.' + ([IO.Path]::GetFileName($Path)) + '.' + $rand + '.tmp'
    $tmp     = Join-Path -Path $folder -ChildPath $tmpName

    try {
        # Write the temp file with the chosen encoding
        [System.IO.File]::WriteAllText($tmp, $Content, $enc)

        # Atomic-ish replace: Replace when target exists, else Move
        if (Test-Path -LiteralPath $Path) {
            [System.IO.File]::Replace($tmp, $Path, $null)
        } else {
            [System.IO.File]::Move($tmp, $Path)
        }
    }
    finally {
        if ($tmp -and (Test-Path -LiteralPath $tmp)) {
            Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
        }
    }
}

# --- Tab 3 atomic writer (simple, no parameter sets) -----------------
function Write-BB2Tab3TextAtomic {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Content
    )

    # Honor read-only: do NOT touch files marked ReadOnly
    if ([System.IO.File]::Exists($Path)) {
        $attrs = [System.IO.File]::GetAttributes($Path)
        if ($attrs -band [System.IO.FileAttributes]::ReadOnly) {
            throw "Target file is marked Read-Only. Clear the attribute first, then retry."
        }
    }

    # Preserve original encoding if the file already exists
    $enc = if (Test-Path -LiteralPath $Path) {
        Get-TextEncoding -Path $Path
    } else {
        'UTF8'
    }

    $tmp = "$Path.bbtmp"

    # Atomic write: temp → move over
    $Content | Set-Content -LiteralPath $tmp -Encoding $enc
    Move-Item -LiteralPath $tmp -Destination $Path -Force
}

# Shim to override any old advanced version with multiple parameter sets.
# Anything still calling Invoke-BB2Tab3AtomicWrite will now be routed
# through the simple writer above (no CmdletBinding, no parameter sets).
function Invoke-BB2Tab3AtomicWrite {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Content
    )

    Write-BB2Tab3TextAtomic -Path $Path -Content $Content
}

# -------------------------------------------------------------------
# TAB 3 – INJECT HANDLER
# Replace your existing $btnCIInject.Add_Click({ ... }) block with THIS.
# -------------------------------------------------------------------
$btnCIInject.Add_Click({
    try {
        $path = $txtCIFile.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($path)) {
            [Windows.Forms.MessageBox]::Show("Choose a target file.","Input Error") | Out-Null
            return
        }

        $modeLabel = [string]$cmbCIMode.SelectedItem

        # Read original for preview/ diff (ErrorAction Stop so our catch actually fires)
        $orig = if (Test-Path -LiteralPath $path) {
            Get-Content -Raw -LiteralPath $path -ErrorAction Stop
        } else {
            ""
        }

        # Build preview so Dry-Run and Apply share the same engine
        $p = CI-BuildPreview -Original $orig -Snippet $rtbCISnippet.Text -Mode $cmbCIMode.SelectedItem `
             -Pattern $txtCIPat.Text -Begin $txtCIBegin.Text -End $txtCIEnd.Text `
             -UseRegex ([bool]$chkCIRegex.Checked) -CaseInsensitive ([bool]$chkCICase.Checked)

        # Dry-Run: update preview + diff only, no writes
        if ($chkCIDry.Checked) {
            $rtbCIFinal.Text = $p.Final
            Render-Diff3 $p.Diff
            Write-BBLog ("Tab3 Dry-Run for {0} ({1})" -f $path,$modeLabel)
            Set-Status "Tab3: Dry-Run only — no changes written."
            return
        }

        # 1) Archive injected snippet for audit (pure .NET)
        try {
            $app = [Environment]::GetFolderPath('ApplicationData')
            if ($app) {
                $injectRoot = [System.IO.Path]::Combine($app,'BellBoy-BB-2','Backups','Injects')
                [System.IO.Directory]::CreateDirectory($injectRoot) | Out-Null

                $stem = [System.IO.Path]::GetFileName($path)
                $ts   = Get-Date -Format 'yyyyMMdd_HHmmss'
                $slug = ([string]$modeLabel) -replace '\s+', '_'
                $snipFile = [System.IO.Path]::Combine($injectRoot, ("{0}_inject_{1}_{2}.txt" -f $stem,$ts,$slug))

                [System.IO.File]::WriteAllText($snipFile, $rtbCISnippet.Text, [System.Text.Encoding]::UTF8)
                Write-BBLog ("Inject snippet saved → {0}" -f $snipFile)
            }
        } catch {
            Write-BBLog ("Tab3 snippet-archive error: {0}" -f $_.Exception.Message)
        }

        # 2) Backup original (doctrine: never prune) – local to Tab 3, pure .NET
        if ([System.IO.File]::Exists($path)) {
            try {
                $name = [System.IO.Path]::GetFileName($path)
                $ts   = Get-Date -Format "yyyyMMdd_HHmmss"
                $app  = [Environment]::GetFolderPath('ApplicationData')

                $root = if ($BackupsDir)        { $BackupsDir }
                        elseif ($RootForState) { [System.IO.Path]::Combine($RootForState,'Backups') }
                        elseif ($app)         { [System.IO.Path]::Combine($app,'BellBoy-BB-2','Backups') }
                        else                  { [System.IO.Path]::Combine($PSScriptRoot,'Backups') }

                [System.IO.Directory]::CreateDirectory($root) | Out-Null

                $bk = [System.IO.Path]::Combine($root, ("{0}.inject.bak.{1}" -f $name,$ts))
                [System.IO.File]::Copy($path, $bk, $true)

                Write-BBLog ("Backup created → {0}" -f $bk)
                $script:LastBackupFile = $bk
                $script:LastTargetPath = $path
            } catch {
                Write-BBLog ("ERROR (Tab3 backup): {0}" -f $_.Exception.Message)
            }
        }

        Write-BBLog "Tab3: prune disabled (doctrine) — keeping all backups."

        if ($path.Length -gt 240) {
            Write-BBLog ("Tab3: warning — long path ({0} chars). Proceeding." -f $path.Length)
        }

        # 3) Safe write via helper (handles read-only/locked cases)
        Invoke-BB2Tab3AtomicWrite -Path $path -Content $p.Final

        # 4) Refresh original/preview/diff from disk (ErrorAction Stop so we catch locks)
        try {
            if (Test-Path -LiteralPath $path) {
                $rtbCIOriginal.Text = Get-Content -Raw -LiteralPath $path -ErrorAction Stop
            }
        } catch {
            Write-BBLog ("Tab3 post-write refresh error: {0}" -f $_.Exception.Message)
        }

        $rtbCIFinal.Text = $p.Final
        Render-Diff3 $p.Diff

        Write-BBLog ("Tab3 Injected into {0} ({1})" -f $path,$modeLabel)
        Set-Status "Tab3: Inject complete."
        [Windows.Forms.MessageBox]::Show("Injection complete.`r`n$path","Success") | Out-Null
    }
    catch {
        $err = $_

        $cmdName = $null
        if ($err.InvocationInfo -and $err.InvocationInfo.MyCommand) {
            $cmdName = $err.InvocationInfo.MyCommand.Name
        }

        $fqid = $null
        if ($err.FullyQualifiedErrorId) {
            $fqid = $err.FullyQualifiedErrorId
        }

        $msg   = $err.Exception.Message
        $etype = $err.Exception.GetType().FullName
        $diag  = "Type: {0}; Cmd: {1}; FQID: {2}" -f $etype, $cmdName, $fqid

        Write-BBLog ("ERROR (Tab3 inject outer): {0} | {1}" -f $msg, $diag)
        Set-Status ("Tab 3 error: " + $msg)

        [Windows.Forms.MessageBox]::Show(
            "Injection error:`r`n$msg`r`n`r`n$diag",
            "Tab 3 – Injection Error"
        ) | Out-Null
    }
})



$btnCIReveal.Add_Click({
  try {
    if (Test-Path -LiteralPath $txtCIFile.Text) { Start-Process explorer.exe "/select,`"$($txtCIFile.Text)`"" }
    elseif ($txtCIFile.Text) { $dir = Split-Path -Path $txtCIFile.Text -Parent; if ($dir) { Start-Process explorer.exe "`"$dir`"" } }
  } catch { }
})
$btnCIOpen.Add_Click({ try { if (Test-Path -LiteralPath $txtCIFile.Text) { Start-Process notepad.exe "`"$($txtCIFile.Text)`"" | Out-Null } } catch { } })

# === Tab 3: Persist last-used settings (PS 5.1-safe) ===
# Requires: $cmbCIMode, $chkCIRegex, $chkCICase, $txtCIPat, $txtCIBegin, $txtCIEnd
# Uses:     $settings (your existing config object), $configPath (save target)



# === Tab 3: settings bootstrap + autosave (no FormClosing needed) ===

# 0) Config path & loader (reuse your existing $configPath if you have one)
if (-not $configPath) {
    $appData   = Join-Path $env:APPDATA 'BellBoy-BB-2'
    if (-not (Test-Path $appData)) { New-Item -ItemType Directory -Path $appData -Force | Out-Null }
    $configPath = Join-Path $appData 'config.json'
}

function Load-Settings {
    if (Test-Path -LiteralPath $configPath) {
        try   { return (Get-Content -Raw -LiteralPath $configPath | ConvertFrom-Json) }
        catch { return ([pscustomobject]@{}) }
    }
    else { return ([pscustomobject]@{}) }
}

function Save-Settings {
    param([Parameter(Mandatory)][object]$Data)
    try { $Data | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $configPath -Encoding UTF8 }
    catch { Write-Verbose "Save-Settings failed: $($_.Exception.Message)" }
}

# 1) Ensure $settings exists and has a Tab3 bag (PSCustomObject-safe)
if (-not $settings) { $settings = Load-Settings }
$hasTab3 = $settings.PSObject.Properties.Name -contains 'Tab3'
if (-not $hasTab3) {
    $defaults = [pscustomobject]@{
        Mode            = 'Append to end'
        Regex           = $true
        CaseInsensitive = $true
        Pattern         = ''
        Begin           = ''
        End             = ''
    }
    $settings | Add-Member -NotePropertyName Tab3 -NotePropertyValue $defaults
    Save-Settings -Data $settings
}

# 2) Load values from settings into the UI
function Load-Tab3Prefs {
    $t = $settings.Tab3
    $wantMode = if ($t.Mode) { [string]$t.Mode } else { 'Append to end' }
    if ($cmbCIMode.Items -contains $wantMode) { $cmbCIMode.SelectedItem = $wantMode } else { $cmbCIMode.SelectedIndex = 0 }
    $chkCIRegex.Checked  = [bool]  $t.Regex
    $chkCICase.Checked   = [bool]  $t.CaseInsensitive
    $txtCIPat.Text       = [string]$t.Pattern
    $txtCIBegin.Text     = [string]$t.Begin
    $txtCIEnd.Text       = [string]$t.End
}
Load-Tab3Prefs

# 3) Keep $settings.Tab3 updated + autosave on each change
function Update-Tab3AndSave {
    $settings.Tab3.Mode            = [string]$cmbCIMode.SelectedItem
    $settings.Tab3.Regex           = [bool]  $chkCIRegex.Checked
    $settings.Tab3.CaseInsensitive = [bool]  $chkCICase.Checked
    $settings.Tab3.Pattern         = [string]$txtCIPat.Text
    $settings.Tab3.Begin           = [string]$txtCIBegin.Text
    $settings.Tab3.End             = [string]$txtCIEnd.Text
    Save-Settings -Data $settings
}

$cmbCIMode.Add_SelectedIndexChanged({ Update-Tab3AndSave })
$chkCIRegex.Add_CheckedChanged(   { Update-Tab3AndSave })
$chkCICase.Add_CheckedChanged(    { Update-Tab3AndSave })
$txtCIPat.Add_TextChanged(        { Update-Tab3AndSave })
$txtCIBegin.Add_TextChanged(      { Update-Tab3AndSave })
$txtCIEnd.Add_TextChanged(        { Update-Tab3AndSave })
# === End Tab 3 autosave ===



# MODE: Patch existing file (Tab 3 only)
# TERMINAL: Windows PowerShell 5.1
# WORKDIR: D:\Trishula-Infra\Bell Boy\BB-2
# SAVE AS: BellBoy-BB2-Tab1_10-HeavyTestRun.ps1

# ================================
# SHARED PAYLOAD VALIDATOR (for TAB 3)
# Drop this once near your other helper functions (top of file or helpers region).
# Idempotent: only defines if not present.
# ================================
if (-not (Get-Command Test-BB2Payload -ErrorAction SilentlyContinue)) {
  function Test-BB2Payload {
    param(
      [string]$Text,
      [ValidateSet('auto','json','ps')] [string]$Kind = 'auto'
    )

    $result = [ordered]@{
      Kind    = $Kind
      IsValid = $true
      Message = ''
      Line    = $null
      Column  = $null
    }

    $t = if ($Text) { $Text.Trim() } else { '' }
    if ([string]::IsNullOrWhiteSpace($t)) {
      $result.Message = 'Payload is empty.'
      return [pscustomobject]$result
    }

    if ($Kind -eq 'auto') {
      $first = $t.Substring(0,1)
      if ($first -eq '{' -or $first -eq '[') {
        $Kind = 'json'
      } else {
        $Kind = 'ps'
      }
    }
    $result.Kind = $Kind

    switch ($Kind) {
      'json' {
        try {
          $null = $t | ConvertFrom-Json -ErrorAction Stop
          $result.Message = 'JSON looks OK.'
        } catch {
          $result.IsValid = $false
          $result.Message = $_.Exception.Message
        }
      }
      'ps' {
        $errs = @()
        $null = [System.Management.Automation.PSParser]::Tokenize($t, [ref]$errs)
        if ($errs.Count -gt 0) {
          $first = $errs[0]
          $result.IsValid = $false
          $result.Message = $first.Message
          $result.Line    = $first.Extent.StartLineNumber
          $result.Column  = $first.Extent.StartColumnNumber
        } else {
          $result.Message = 'PowerShell syntax looks OK.'
        }
      }
    }

    return [pscustomobject]$result
  }
}

# ==================================================
# TAB 3 – VALIDATE BUTTON (Code Inject payload)
# Place this in the Tab 3 section, AFTER the Tab 3 UI
# and the payload TextBox are created.
# ==================================================

# IMPORTANT:
# 1) Set this alias ONCE to your actual Tab 3 payload TextBox.
#    Change the right-hand side to whatever you really use.
#    Examples you might have in your code:
#      $txtTab3Payload
#      $txtT3InsertText
#      $txtInjectText
#
#    Pick the real one and assign it here, e.g.:
#      $script:Tab3PayloadBox = $txtInjectText
#
# 2) If you already have a good place for the button, just tweak Location.
# -------------------------------------------------
$script:Tab3PayloadBox = $rtbCISnippet   # <-- CHANGE right side to your real Tab 3 payload TextBox

$btnTab3Validate = New-Object System.Windows.Forms.Button
$btnTab3Validate.Text = 'Validate Payload'
$btnTab3Validate.Size = New-Object System.Drawing.Size(130,28)
# Adjust coordinates so it sits nicely near your Tab 3 payload box
$btnTab3Validate.Location = New-Object System.Drawing.Point(20, 260)
$tab3.Controls.Add($btnTab3Validate)

$btnTab3Validate.Add_Click({
  try {
    $payload = if ($script:Tab3PayloadBox) { $script:Tab3PayloadBox.Text } else { '' }
    $val     = Test-BB2Payload -Text $payload -Kind 'auto'

    if ($val.IsValid) {
      Write-BBLog ("Tab 3 validator OK ({0}): {1}" -f $val.Kind, $val.Message)
      [Windows.Forms.MessageBox]::Show(
        "Payload looks OK.`r`n`r`nKind: $($val.Kind)`r`n$($val.Message)",
        "Tab 3 – Validator"
      ) | Out-Null
    }
    else {
      $loc = if ($val.Line -and $val.Column) {
        "Line $($val.Line), Column $($val.Column)"
      } else {
        "Location unknown"
      }

      Write-BBLog ("Tab 3 validator ERROR ({0}): {1} [{2}]" -f $val.Kind, $val.Message, $loc)
      [Windows.Forms.MessageBox]::Show(
        "Validator found problems in the payload.`r`n`r`nKind: $($val.Kind)`r`nMessage: $($val.Message)`r`n$loc",
        "Tab 3 – Validator",
        [Windows.Forms.MessageBoxButtons]::OK,
        [Windows.Forms.MessageBoxIcon]::Warning
      ) | Out-Null
    }
  } catch {
    Write-BBLog ("Tab 3 validator threw: {0}" -f $_.Exception.Message)
    [Windows.Forms.MessageBox]::Show(
      "Validator error:`r`n$($_.Exception.Message)",
      "Tab 3 – Validator Error"
    ) | Out-Null
  }
})

# Load values into UI now
Load-Tab3Prefs
#$cfg.Tab3 = $settings.Tab3


# ===== End Tab 3 =====



[void]$form.ShowDialog()
