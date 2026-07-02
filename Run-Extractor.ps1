<#
.SYNOPSIS
  Interactive, double-click-friendly front end for Extract-Attachments.ps1.

.DESCRIPTION
  Guides the user through configuring and running an extraction:
    - choose & validate the save path (with an auto-detected default)
    - choose & validate one or more target Outlook folders
    - accurately estimate what a real run will extract (size + max aware)
    - run a safe single-email test before processing everything
    - remember your settings so the next run doesn't need reconfiguring
  No parameters: just double-click Run-Extractor.bat (which calls this script).
#>

$ErrorActionPreference = "Stop"
$ScriptVersion = "2.0.0"
$scriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$worker      = Join-Path $scriptDir "Extract-Attachments.ps1"

if (-not (Test-Path $worker)) {
    Write-Host "ERROR: Extract-Attachments.ps1 not found next to this launcher." -ForegroundColor Red
    Write-Host "Keep both files in the same folder." -ForegroundColor Yellow
    Read-Host "Press Enter to exit"; exit 1
}

# ---- Connect to Outlook ----
function Connect-Outlook {
    for ($r = 1; $r -le 3; $r++) {
        try {
            $script:outlook = New-Object -ComObject Outlook.Application
            $script:ns = $outlook.GetNamespace("MAPI")
            $null = $ns.GetDefaultFolder(6)
            return $true
        } catch {
            Write-Host "Attempt $r/3 - cannot reach Outlook: $($_.Exception.Message)" -ForegroundColor Yellow
            Start-Sleep -Seconds (3 * $r)
        }
    }
    return $false
}

Write-Host "Connecting to Outlook..." -ForegroundColor Cyan
if (-not (Connect-Outlook)) {
    Write-Host ""
    Write-Host "ERROR: Could not connect to Outlook." -ForegroundColor Red
    Write-Host "  - Make sure the CLASSIC Outlook desktop app is open and signed in." -ForegroundColor Yellow
    Write-Host "  - 'New Outlook' and Outlook on the web are NOT supported." -ForegroundColor Yellow
    Read-Host "Press Enter to exit"; exit 1
}
Write-Host "Connected." -ForegroundColor Green

# ---- Auto-detect default save path (mirrors Extract-Attachments.ps1) ----
function Get-DefaultSavePath {
    $candidates = @($env:OneDriveCommercial, $env:OneDrive, "$env:USERPROFILE\OneDrive")
    Get-ChildItem "$env:USERPROFILE" -Directory -Filter "OneDrive*" -ErrorAction SilentlyContinue |
        ForEach-Object { $candidates += $_.FullName }
    foreach ($p in $candidates) {
        if ($p -and (Test-Path $p)) { return (Join-Path $p "Downloads") }
    }
    return (Join-Path $env:USERPROFILE "Downloads")
}

# ---- Auto-detect SharePoint web link base + OneDrive synced root from registry ----
function Get-DefaultSharePoint {
    $result = @{ WebRoot = ""; LocalRoot = "" }
    $accounts = Get-ChildItem 'HKCU:\Software\Microsoft\OneDrive\Accounts' -ErrorAction SilentlyContinue |
        Where-Object { $_.PSChildName -like 'Business*' }
    foreach ($acct in $accounts) {
        $p = Get-ItemProperty $acct.PSPath -ErrorAction SilentlyContinue
        if ($p -and $p.UserFolder -and $p.ServiceEndpointUri -and (Test-Path $p.UserFolder)) {
            $result.LocalRoot = $p.UserFolder
            $result.WebRoot   = ($p.ServiceEndpointUri -replace '/_api$', '') + '/Documents'
            break
        }
    }
    return $result
}

# ---- Remember settings between runs (per-user config file) ----
# Persist the user's choices locally so the next launch doesn't need
# reconfiguring. Stored per Windows user, outside the app folder.
$configDir  = Join-Path $env:USERPROFILE ".outlook-attachment-extractor"
$configPath = Join-Path $configDir "config.json"

function Save-Config {
    # Persist the durable settings from $cfg. FolderObjs are live COM objects and
    # are intentionally NOT saved - they are rebuilt from the folder names on load.
    try {
        if (-not (Test-Path $configDir)) {
            New-Item -ItemType Directory -Path $configDir -Force -ErrorAction Stop | Out-Null
        }
        $toSave = [ordered]@{
            OutputPath        = $cfg.OutputPath
            Folder            = $cfg.Folder
            MinSizeMB         = $cfg.MinSizeMB
            OlderThanDays     = $cfg.OlderThanDays
            MaxItems          = $cfg.MaxItems
            SharePointWebRoot = $cfg.SharePointWebRoot
            OneDriveLocalRoot = $cfg.OneDriveLocalRoot
            _savedAt          = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        }
        $toSave | ConvertTo-Json | Out-File -FilePath $configPath -Encoding UTF8 -ErrorAction Stop
    } catch {
        Write-Host "NOTE: couldn't save settings to $configPath ($($_.Exception.Message))" -ForegroundColor DarkYellow
    }
}

function Load-Config {
    # Overlay any saved settings onto $cfg, re-validating so stale values (a
    # deleted folder or a missing save path) fall back to the detected defaults.
    if (-not (Test-Path $configPath)) { return }
    try {
        $saved = Get-Content $configPath -Raw -ErrorAction Stop | ConvertFrom-Json
    } catch {
        Write-Host "NOTE: saved settings unreadable ($configPath); using defaults." -ForegroundColor DarkYellow
        return
    }
    if ($null -ne $saved.MinSizeMB)     { $cfg.MinSizeMB     = [double]$saved.MinSizeMB }
    if ($null -ne $saved.OlderThanDays) { $cfg.OlderThanDays = [int]$saved.OlderThanDays }
    if ($null -ne $saved.MaxItems)      { $cfg.MaxItems      = [int]$saved.MaxItems }
    if ($saved.SharePointWebRoot) { $cfg.SharePointWebRoot = [string]$saved.SharePointWebRoot }
    if ($saved.OneDriveLocalRoot) { $cfg.OneDriveLocalRoot = [string]$saved.OneDriveLocalRoot }
    if ($saved.OutputPath) {
        if (Test-Path $saved.OutputPath) {
            $cfg.OutputPath = [string]$saved.OutputPath
        } else {
            Write-Host ("NOTE: saved save path no longer exists ({0}); using {1}." -f $saved.OutputPath, $cfg.OutputPath) -ForegroundColor DarkYellow
        }
    }
    if ($saved.Folder) {
        $resolved = Resolve-Folders ([string]$saved.Folder)
        if ($resolved.List.Count -gt 0) {
            if ($resolved.Bad.Count -eq 0) {
                $cfg.Folder = [string]$saved.Folder
            } else {
                $cfg.Folder = ($resolved.List | ForEach-Object { $_.Name }) -join ','
                Write-Host ("NOTE: saved folder(s) not found: {0}; using '{1}'." -f ($resolved.Bad -join ', '), $cfg.Folder) -ForegroundColor DarkYellow
            }
            $cfg.FolderObjs = @($resolved.List | ForEach-Object { $_.Obj })
        } else {
            Write-Host ("NOTE: saved folder(s) '{0}' not found; using '{1}'." -f $saved.Folder, $cfg.Folder) -ForegroundColor DarkYellow
        }
    }
    Write-Host "Restored your saved settings from $configPath" -ForegroundColor Green
}

# ---- Resolve an Outlook folder by name (inbox/sent/drafts/all/custom) ----
function Resolve-Folder($name) {
    switch ($name.ToLower()) {
        "inbox"  { return @{ Name="Inbox";      Obj=$ns.GetDefaultFolder(6) } }
        "sent"   { return @{ Name="Sent Items"; Obj=$ns.GetDefaultFolder(5) } }
        "drafts" { return @{ Name="Drafts";     Obj=$ns.GetDefaultFolder(16) } }
        "all"    { return @{ Name="all";        Obj=$null } }   # handled by worker
        default {
            $root = $ns.GetDefaultFolder(6).Store.GetRootFolder()
            foreach ($f in $root.Folders) { if ($f.Name -eq $name) { return @{ Name=$f.Name; Obj=$f } } }
            return $null
        }
    }
}

# ---- Resolve a comma-separated list of folders into live COM objects ----
# Accepts numbers-already-mapped names and keywords (inbox/sent/drafts/all).
# "all" expands to Inbox + Sent Items. Duplicates (by name) are removed.
# Returns @{ List = @(@{Name;Obj}...); Bad = @(unresolved names) }.
function Resolve-Folders($spec) {
    $list = @(); $bad = @(); $seen = @{}
    foreach ($tok in ($spec -split ',')) {
        $name = $tok.Trim()
        if ($name -eq "") { continue }
        if ($name.ToLower() -eq "all") {
            $expand = @(
                @{ Name = "Inbox";      Obj = $ns.GetDefaultFolder(6) },
                @{ Name = "Sent Items"; Obj = $ns.GetDefaultFolder(5) }
            )
        } else {
            $r = Resolve-Folder $name
            if (-not $r -or -not $r.Obj) { $bad += $name; continue }
            $expand = @(@{ Name = $r.Name; Obj = $r.Obj })
        }
        foreach ($f in $expand) {
            $key = $f.Name.ToLower()
            if (-not $seen.ContainsKey($key)) { $seen[$key] = $true; $list += $f }
        }
    }
    return @{ List = $list; Bad = $bad }
}

# ---- Estimate what a real run would actually extract ----
# Mirrors the worker exactly: per folder, oldest-first, restricted by the age
# filter, scan up to MaxItems and count emails that truly have a qualifying
# attachment (>= MinSizeMB, not an inline image), summing their size. This
# reflects the size filter AND the per-run MaxItems cap - not just folder+date.
function Estimate-Run($folderObjs) {
    $dateMatch = 0; $emails = 0; $atts = 0; $sizeMB = 0.0; $capped = $false
    foreach ($fo in $folderObjs) {
        if (-not $fo) { continue }
        try {
            $items = $fo.Items
            $items.Sort("[ReceivedTime]", $false)
            if ($cfg.OlderThanDays -gt 0) {
                $cut = (Get-Date).AddDays(-$cfg.OlderThanDays).ToString("MM/dd/yyyy")
                $filtered = $items.Restrict("[ReceivedTime] < '$cut'")
            } else {
                $filtered = $items
            }
            $dateMatch += $filtered.Count
            if ($filtered.Count -gt $cfg.MaxItems) { $capped = $true }
            $processed = 0
            $item = $filtered.GetFirst()
            while ($item -ne $null -and $processed -lt $cfg.MaxItems) {
                if ($item.Attachments.Count -gt 0) {
                    $hit = $false
                    for ($a = 1; $a -le $item.Attachments.Count; $a++) {
                        $att = $item.Attachments.Item($a)
                        $mb = [Math]::Round($att.Size / 1MB, 2)
                        if ($mb -ge $cfg.MinSizeMB -and $att.Type -ne 6) {
                            $hit = $true; $atts++; $sizeMB += $mb
                        }
                    }
                    if ($hit) { $emails++ }
                }
                $processed++
                $item = $filtered.GetNext()
            }
        } catch {
            Write-Host "NOTE: couldn't scan a folder for the estimate ($($_.Exception.Message))" -ForegroundColor DarkYellow
        }
    }
    return [ordered]@{
        DateMatch   = $dateMatch
        Emails      = $emails
        Attachments = $atts
        SizeMB      = [Math]::Round($sizeMB, 1)
        Capped      = $capped
    }
}

# ---- Recompute + cache the estimate (call after folder/filter changes) ----
function Refresh-Estimate {
    Write-Host ("Estimating (scanning up to {0} email(s) per folder)..." -f $cfg.MaxItems) -ForegroundColor DarkGray
    try   { $script:estimate = Estimate-Run $cfg.FolderObjs }
    catch { $script:estimate = $null; Write-Host "NOTE: estimate failed ($($_.Exception.Message))" -ForegroundColor DarkYellow }
}

# ---- Validate / prompt for a writable save path ----
function Set-SavePath {
    while ($true) {
        $p = Read-Host "Enter save path (blank = keep '$($cfg.OutputPath)')"
        if ([string]::IsNullOrWhiteSpace($p)) { return }
        try {
            if (-not (Test-Path $p)) {
                $mk = Read-Host "Path doesn't exist. Create it? [Y/N]"
                if ($mk -match '^(y|yes)$') { New-Item -ItemType Directory -Path $p -Force | Out-Null }
                else { continue }
            }
            # writability check
            $probe = Join-Path $p (".writetest_{0}.tmp" -f (Get-Random))
            "ok" | Out-File -FilePath $probe -Encoding ASCII; Remove-Item $probe -Force
            $cfg.OutputPath = (Resolve-Path $p).Path
            Write-Host "Save path set: $($cfg.OutputPath)" -ForegroundColor Green
            return
        } catch {
            Write-Host "Not usable: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# ---- Validate / prompt for one or MORE target folders ----
function Set-Folder {
    $root = $ns.GetDefaultFolder(6).Store.GetRootFolder()
    $topFolders = @($root.Folders | ForEach-Object { $_.Name })
    Write-Host "Available top-level folders:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $topFolders.Count; $i++) {
        Write-Host ("  {0}) {1}" -f ($i + 1), $topFolders[$i])
    }
    Write-Host "Pick one or MORE, comma-separated - by number or name." -ForegroundColor DarkGray
    Write-Host "Keywords: inbox, sent, drafts, all.   e.g.  1,sent   or   inbox,Projects" -ForegroundColor DarkGray
    while ($true) {
        $reply = Read-Host "Enter folder(s) (blank = keep '$($cfg.Folder)')"
        if ([string]::IsNullOrWhiteSpace($reply)) { return }
        # Map numeric tokens to the listed folder names; keep keywords/names as-is.
        $tokens = @(); $ok = $true
        foreach ($tok in ($reply -split ',')) {
            $t = $tok.Trim()
            if ($t -eq "") { continue }
            if ($t -match '^\d+$') {
                $idx = [int]$t
                if ($idx -ge 1 -and $idx -le $topFolders.Count) {
                    $tokens += $topFolders[$idx - 1]
                } else {
                    Write-Host "Number '$t' is out of range (1-$($topFolders.Count))." -ForegroundColor Red
                    $ok = $false; break
                }
            } else {
                $tokens += $t
            }
        }
        if (-not $ok) { continue }
        if ($tokens.Count -eq 0) { Write-Host "Nothing entered. Try again." -ForegroundColor Red; continue }
        $spec = ($tokens -join ',')
        $resolved = Resolve-Folders $spec
        if ($resolved.Bad.Count -gt 0) {
            Write-Host ("Folder(s) not found: {0}. Try again." -f ($resolved.Bad -join ', ')) -ForegroundColor Red
            continue
        }
        if ($resolved.List.Count -eq 0) {
            Write-Host "No valid folder chosen. Try again." -ForegroundColor Red
            continue
        }
        $cfg.Folder     = $spec
        $cfg.FolderObjs = @($resolved.List | ForEach-Object { $_.Obj })
        $names = ($resolved.List | ForEach-Object { $_.Name }) -join ', '
        Write-Host ("Folder(s) set: {0}" -f $names) -ForegroundColor Green
        return
    }
}

function Set-SharePoint {
    Write-Host "SharePoint web links power the 'Open in browser' link in each email." -ForegroundColor Cyan
    Write-Host "They only work when your save path is inside a OneDrive-synced folder." -ForegroundColor DarkGray
    Write-Host (" Current web base : {0}" -f $(if ($cfg.SharePointWebRoot) { $cfg.SharePointWebRoot } else { '(none - local links only)' }))
    Write-Host (" OneDrive root    : {0}" -f $(if ($cfg.OneDriveLocalRoot) { $cfg.OneDriveLocalRoot } else { '(none)' }))
    Write-Host " a) Auto-detect from OneDrive   o) Disable (local links only)   e) Enter manually"
    $pick = Read-Host "Choose [a/o/e] (blank = keep)"
    switch ($pick.ToLower()) {
        "a" {
            $d = Get-DefaultSharePoint
            if ($d.WebRoot) { $cfg.SharePointWebRoot = $d.WebRoot; $cfg.OneDriveLocalRoot = $d.LocalRoot; Write-Host "Detected: $($d.WebRoot)" -ForegroundColor Green }
            else { Write-Host "No OneDrive Business account detected." -ForegroundColor Yellow }
        }
        "o" { $cfg.SharePointWebRoot = ""; $cfg.OneDriveLocalRoot = ""; Write-Host "Web links disabled." -ForegroundColor Green }
        "e" {
            $w = Read-Host "SharePoint web base (e.g. https://contoso-my.sharepoint.com/personal/you_contoso_com/Documents)"
            if ($w -notmatch '^https://.*sharepoint\.com/.+') { Write-Host "That doesn't look like a SharePoint URL. Keeping previous." -ForegroundColor Red; return }
            $l = Read-Host "Local OneDrive root that maps to it (e.g. C:\Users\you\OneDrive - Contoso)"
            if (-not (Test-Path $l)) { Write-Host "Local root not found. Keeping previous." -ForegroundColor Red; return }
            $cfg.SharePointWebRoot = $w.TrimEnd('/'); $cfg.OneDriveLocalRoot = (Resolve-Path $l).Path
            Write-Host "Set." -ForegroundColor Green
        }
        default { }
    }
    if ($cfg.SharePointWebRoot -and $cfg.OutputPath -and -not $cfg.OutputPath.StartsWith($cfg.OneDriveLocalRoot, [StringComparison]::OrdinalIgnoreCase)) {
        Write-Host "NOTE: your save path is not under the OneDrive root, so web links won't be generated for it." -ForegroundColor Yellow
    }
}

function Set-Filters {
    $s = Read-Host "Min attachment size in MB (blank = $($cfg.MinSizeMB))"
    if ($s -match '^\d+(\.\d+)?$') { $cfg.MinSizeMB = [double]$s }
    $d = Read-Host "Only emails older than N days, 0 = all (blank = $($cfg.OlderThanDays))"
    if ($d -match '^\d+$') { $cfg.OlderThanDays = [int]$d }
    $m = Read-Host "Max emails per run (blank = $($cfg.MaxItems))"
    if ($m -match '^\d+$') { $cfg.MaxItems = [int]$m }
    Write-Host "Filters updated." -ForegroundColor Green
}

# ---- Invoke the worker ----
function Run-Worker([bool]$DryRun, [bool]$Remove, [int]$MaxOverride) {
    $max = if ($MaxOverride -gt 0) { $MaxOverride } else { $cfg.MaxItems }
    & $worker -Folder $cfg.Folder -MinSizeMB $cfg.MinSizeMB -MaxItems $max `
        -OlderThanDays $cfg.OlderThanDays -DryRun $DryRun -RemoveAfterExtract $Remove `
        -OutputPath $cfg.OutputPath -SharePointWebRoot $cfg.SharePointWebRoot `
        -OneDriveLocalRoot $cfg.OneDriveLocalRoot
}

# ---- Initial config ----
$spDefault = Get-DefaultSharePoint
$cfg = [ordered]@{
    OutputPath        = Get-DefaultSavePath
    Folder            = "inbox"
    FolderObjs        = @($ns.GetDefaultFolder(6))
    MinSizeMB         = 1
    OlderThanDays     = 90
    MaxItems          = 100
    SharePointWebRoot = $spDefault.WebRoot
    OneDriveLocalRoot = $spDefault.LocalRoot
}

# Restore previously saved settings (if any); seed the file on first run so the
# detected defaults are remembered next time too.
Load-Config
if (-not (Test-Path $configPath)) { Save-Config }
Write-Host "Your settings are remembered in $configPath" -ForegroundColor DarkGray

function Show-Status {
    Write-Host ""
    $spText = if ($cfg.SharePointWebRoot) { $cfg.SharePointWebRoot } else { "(none - local links only)" }
    Write-Host "============== Outlook Attachment Extractor  v$ScriptVersion ==============" -ForegroundColor Cyan
    Write-Host (" Save path    : {0}" -f $cfg.OutputPath)
    Write-Host (" Email folder : {0}" -f $cfg.Folder)
    Write-Host (" SharePoint   : {0}" -f $spText)
    Write-Host (" Filters      : >= {0} MB | older than {1} days | max {2}/run per folder" -f $cfg.MinSizeMB, $cfg.OlderThanDays, $cfg.MaxItems)
    $e = $script:estimate
    if ($e) {
        Write-Host (" Estimate     : {0} attachment(s) in {1} email(s), ~{2} MB will be extracted" -f $e.Attachments, $e.Emails, $e.SizeMB) -ForegroundColor Yellow
        $note = if ($e.Capped) { "; raise 'max' to cover the rest" } else { "" }
        Write-Host ("                (>= {0} MB filter; {1} email(s) match the age filter; up to {2} scanned per folder{3})" -f $cfg.MinSizeMB, $e.DateMatch, $cfg.MaxItems, $note) -ForegroundColor DarkYellow
    } else {
        Write-Host " Estimate     : (unavailable - change folder/filters to recompute)" -ForegroundColor DarkYellow
    }
    Write-Host "----------------------------------------------------------------------"
    Write-Host " 1) Change save path"
    Write-Host " 2) Change email folder(s)"
    Write-Host " 3) Change SharePoint web link base ('Open in browser' links)"
    Write-Host " 4) Change filters (size / age / max)"
    Write-Host " 5) Preview (dry run - exact attachment list, no changes)"
    Write-Host " 6) TEST: extract just ONE email (keeps it in the email)"
    Write-Host " 7) Extract ALL (save to disk, keep attachments in email)"
    Write-Host " 8) Extract ALL and REMOVE from email (frees mailbox - permanent)"
    Write-Host " 9) Quit"
    Write-Host "======================================================================" -ForegroundColor Cyan
}

Refresh-Estimate

while ($true) {
    Show-Status
    $choice = Read-Host "Select [1-9]"
    switch ($choice) {
        "1" { Set-SavePath;   Save-Config }
        "2" { Set-Folder;     Save-Config; Refresh-Estimate }
        "3" { Set-SharePoint; Save-Config }
        "4" { Set-Filters;    Save-Config; Refresh-Estimate }
        "5" { Run-Worker $true $false 0; Read-Host "Press Enter to continue" }
        "6" {
            Write-Host "Extracting ONE email as a test (no removal)..." -ForegroundColor Cyan
            Run-Worker $false $false 1
            Write-Host "Test done. Check your save path to confirm the file looks right." -ForegroundColor Green
            Read-Host "Press Enter to continue"
        }
        "7" {
            $c = Read-Host "Extract all matching attachments to '$($cfg.OutputPath)'? [Y/N]"
            if ($c -match '^(y|yes)$') { Run-Worker $false $false 0 }
            Read-Host "Press Enter to continue"
        }
        "8" {
            Write-Host "WARNING: This PERMANENTLY removes attachments from the emails after saving." -ForegroundColor Red
            $c = Read-Host "Type EXACTLY 'REMOVE' to proceed"
            if ($c -ceq "REMOVE") { Run-Worker $false $true 0 }
            else { Write-Host "Cancelled." -ForegroundColor Yellow }
            Read-Host "Press Enter to continue"
        }
        # Use 'exit' here, not 'break': in PowerShell a 'break' inside a switch
        # only leaves the switch, so the menu loop would redraw instead of
        # quitting. Exiting the script lets the Run-Extractor.bat window close.
        "9" { Write-Host "Bye."; exit 0 }
        default { Write-Host "Enter a number 1-9." -ForegroundColor Yellow }
    }
}
