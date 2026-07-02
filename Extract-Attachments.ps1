<#
.SYNOPSIS
  Extract large attachments from Outlook to a local folder (e.g. OneDrive),
  preserving email-attachment links.

.DESCRIPTION
  Scans Outlook folders for emails with large attachments, saves them directly
  to a configurable output directory (no extra sub-folders are created),
  creates an index CSV mapping, inserts a reference link into the email body,
  and optionally removes the attachment from the email to free mailbox space.

  Requires:
    - Windows 10/11 with Microsoft Outlook desktop app (classic, not "New Outlook")
    - PowerShell 5.1+

.PARAMETER Folder
  Outlook folder(s) to scan. Accepts a single value or a comma-separated list:
  inbox, sent, drafts, all, and/or any custom top-level folder name(s), e.g.
  "inbox,sent,Projects". "all" expands to Inbox + Sent Items. Default: inbox

.PARAMETER MinSizeMB
  Minimum attachment size in MB to extract. Default: 1

.PARAMETER MaxItems
  Maximum number of emails to process per run, to prevent timeouts. Default: 100

.PARAMETER OlderThanDays
  Only process emails older than N days. Default: 90

.PARAMETER DryRun
  Preview mode - list what would be extracted without making any changes. Default: true
  ALWAYS do a dry run first before extracting.

.PARAMETER RemoveAfterExtract
  Remove attachment from the email after saving to disk. This frees mailbox/server
  space but is PERMANENT - the attachment cannot be recovered from the email.
  Default: false

.PARAMETER OutputPath
  Base directory for extracted attachments. Defaults to your OneDrive Downloads folder.
  Falls back to ~/Downloads if OneDrive is not configured.

.EXAMPLE
  # Preview large attachments in Inbox (safe - no changes made)
  .\Extract-Attachments.ps1 -DryRun $true

.EXAMPLE
  # Extract attachments > 5MB from Inbox, older than 180 days
  .\Extract-Attachments.ps1 -MinSizeMB 5 -OlderThanDays 180 -DryRun $false

.EXAMPLE
  # Extract from Sent Items folder
  .\Extract-Attachments.ps1 -Folder sent -DryRun $false

.EXAMPLE
  # Extract from several folders in one run (comma-separated)
  .\Extract-Attachments.ps1 -Folder "inbox,sent,Projects" -DryRun $false

.EXAMPLE
  # Extract AND remove attachments to free mailbox space
  .\Extract-Attachments.ps1 -DryRun $false -RemoveAfterExtract $true

.EXAMPLE
  # Extract to a custom output directory
  .\Extract-Attachments.ps1 -DryRun $false -OutputPath "D:\EmailBackups"

.EXAMPLE
  # Scan all folders (Inbox + Sent Items)
  .\Extract-Attachments.ps1 -Folder all -DryRun $true
#>

param(
    [string]$Folder = "inbox",
    [double]$MinSizeMB = 1,
    [int]$MaxItems = 100,
    [int]$OlderThanDays = 90,
    [bool]$DryRun = $true,
    [bool]$RemoveAfterExtract = $false,
    [string]$OutputPath = "",
    [string]$SharePointWebRoot = "",
    [string]$OneDriveLocalRoot = "",
    [switch]$Version
)

$ScriptVersion = "2.0.0"
if ($Version) { Write-Host "Outlook Attachment Extractor v$ScriptVersion"; exit 0 }

$ErrorActionPreference = "Continue"

# === Determine output directory ===
if ($OutputPath -eq "") {
    # Try OneDrive paths (common enterprise and personal patterns)
    $oneDrivePaths = @(
        $env:OneDriveCommercial,
        $env:OneDrive,
        "$env:USERPROFILE\OneDrive"
    )
    # Also check for OneDrive folders with org names
    $oneDriveFolders = Get-ChildItem "$env:USERPROFILE" -Directory -Filter "OneDrive*" -ErrorAction SilentlyContinue
    foreach ($od in $oneDriveFolders) {
        $oneDrivePaths += $od.FullName
    }

    $OutputPath = ""
    foreach ($p in $oneDrivePaths) {
        if ($p -and (Test-Path $p)) {
            $OutputPath = Join-Path $p "Downloads"
            break
        }
    }
    if ($OutputPath -eq "") {
        $OutputPath = Join-Path $env:USERPROFILE "Downloads"
    }
}

$indexFile = Join-Path $OutputPath "_attachment_index.csv"

Write-Host "Outlook Attachment Extractor v$ScriptVersion" -ForegroundColor DarkGray
Write-Host "Output directory: $OutputPath" -ForegroundColor Cyan

# === Determine SharePoint web link base (for "Open in browser" links) ===
# Auto-detect from the OneDrive Business registry so the tenant URL is always
# correct and survives tenant renames. Both values can be overridden via params.
if ($SharePointWebRoot -eq "" -or $OneDriveLocalRoot -eq "") {
    $odAccounts = Get-ChildItem 'HKCU:\Software\Microsoft\OneDrive\Accounts' -ErrorAction SilentlyContinue |
        Where-Object { $_.PSChildName -like 'Business*' }
    foreach ($acct in $odAccounts) {
        $props = Get-ItemProperty $acct.PSPath -ErrorAction SilentlyContinue
        if (-not $props) { continue }
        if ($OneDriveLocalRoot -eq "" -and $props.UserFolder -and (Test-Path $props.UserFolder)) {
            # Only adopt this account if our OutputPath lives under its synced folder
            if ($OutputPath.StartsWith($props.UserFolder, [StringComparison]::OrdinalIgnoreCase)) {
                $OneDriveLocalRoot = $props.UserFolder
                if ($SharePointWebRoot -eq "" -and $props.ServiceEndpointUri) {
                    $SharePointWebRoot = ($props.ServiceEndpointUri -replace '/_api$', '') + '/Documents'
                }
            }
        }
    }
}

if ($SharePointWebRoot -ne "" -and $OneDriveLocalRoot -ne "") {
    Write-Host "Web links via: $SharePointWebRoot" -ForegroundColor Cyan
} else {
    Write-Host "Web links: (none - output is not under a OneDrive-synced folder)" -ForegroundColor DarkGray
}
Write-Host ""

# Convert a local OneDrive path to its SharePoint web URL (URL-encoded). Returns
# $null when the path is outside the synced root or no web root is configured.
function Get-OneDriveWebUrl($localPath) {
    if ($SharePointWebRoot -eq "" -or $OneDriveLocalRoot -eq "") { return $null }
    if (-not $localPath.StartsWith($OneDriveLocalRoot, [StringComparison]::OrdinalIgnoreCase)) { return $null }
    $relative = $localPath.Substring($OneDriveLocalRoot.Length)
    $segments = $relative -split '\\' | ForEach-Object { [Uri]::EscapeDataString($_) }
    return $SharePointWebRoot + ($segments -join '/')
}

# === File extension to category mapping ===
$categoryMap = @{
    # Work documents
    '.pptx' = 'Work'; '.ppt' = 'Work'; '.docx' = 'Work'; '.doc' = 'Work'
    '.pdf'  = 'Work'; '.msg' = 'Work'; '.eml' = 'Work'; '.rtf' = 'Work'
    '.one'  = 'Work'; '.vsdx' = 'Work'; '.vsd' = 'Work'
    # Data & Analytics
    '.xlsx' = 'Analytics & Data'; '.xls' = 'Analytics & Data'; '.csv' = 'Analytics & Data'
    '.json' = 'Analytics & Data'; '.xml' = 'Analytics & Data'; '.parquet' = 'Analytics & Data'
    '.accdb' = 'Analytics & Data'; '.mdb' = 'Analytics & Data'
    # Media
    '.png'  = 'Media'; '.jpg' = 'Media'; '.jpeg' = 'Media'; '.gif' = 'Media'
    '.bmp'  = 'Media'; '.svg' = 'Media'; '.mp4' = 'Media'; '.mp3' = 'Media'
    '.mov'  = 'Media'; '.avi' = 'Media'; '.wmv' = 'Media'; '.wav' = 'Media'
    '.tiff' = 'Media'; '.tif' = 'Media'; '.webp' = 'Media'
    # Installers & Tools
    '.zip'  = 'Installers & Tools'; '.rar' = 'Installers & Tools'; '.7z' = 'Installers & Tools'
    '.msi'  = 'Installers & Tools'; '.exe' = 'Installers & Tools'; '.iso' = 'Installers & Tools'
    '.tar'  = 'Installers & Tools'; '.gz' = 'Installers & Tools'
}

function Get-FileCategory($filename) {
    $ext = [IO.Path]::GetExtension($filename).ToLower()
    if ($categoryMap.ContainsKey($ext)) {
        return $categoryMap[$ext]
    }
    return 'Work'  # default category
}

# === Connect to Outlook (with retry) ===
$maxRetries = 3
$outlook = $null
for ($r = 1; $r -le $maxRetries; $r++) {
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $ns = $outlook.GetNamespace("MAPI")
        $null = $ns.GetDefaultFolder(6)  # test connection
        Write-Host "Connected to Outlook" -ForegroundColor Green
        break
    } catch {
        Write-Host "Attempt $r/$maxRetries - COM error: $($_.Exception.Message)" -ForegroundColor Yellow
        if ($r -lt $maxRetries) {
            Start-Sleep -Seconds (5 * $r)
        } else {
            Write-Host ""
            Write-Host "ERROR: Cannot connect to Outlook after $maxRetries attempts." -ForegroundColor Red
            Write-Host ""
            Write-Host "Troubleshooting:" -ForegroundColor Yellow
            Write-Host "  1. Make sure Outlook desktop app (classic) is running"
            Write-Host "  2. 'New Outlook' is NOT supported - switch to classic Outlook"
            Write-Host "  3. Try closing and reopening Outlook, then run this script again"
            Write-Host "  4. Run: Get-Process OUTLOOK  (should show a running process)"
            exit 1
        }
    }
}

# === Get target folder ===
function Get-OutlookFolder($folderName) {
    switch ($folderName.ToLower()) {
        "inbox"  { return $ns.GetDefaultFolder(6) }
        "sent"   { return $ns.GetDefaultFolder(5) }
        "drafts" { return $ns.GetDefaultFolder(16) }
        default {
            $root = $ns.GetDefaultFolder(6).Store.GetRootFolder()
            foreach ($f in $root.Folders) {
                if ($f.Name -eq $folderName) { return $f }
            }
            Write-Host "ERROR: Folder '$folderName' not found." -ForegroundColor Red
            Write-Host "Available top-level folders:" -ForegroundColor Yellow
            foreach ($f in $root.Folders) {
                Write-Host "  - $($f.Name)"
            }
            exit 1
        }
    }
}

# Build the list of folders to scan. -Folder accepts a single value or a
# comma-separated list (e.g. "inbox,sent,Projects"). "all" expands to Inbox +
# Sent Items. Keyword folders are normalized to their canonical names so a
# folder referenced two ways (e.g. "sent" and "all") is scanned only once.
$folders = @()
$seen = @{}
foreach ($token in ($Folder -split ',')) {
    $name = $token.Trim()
    if ($name -eq "") { continue }
    switch ($name.ToLower()) {
        "all" {
            $expand = @(
                @{ Name = "Inbox";      Obj = $ns.GetDefaultFolder(6) },
                @{ Name = "Sent Items"; Obj = $ns.GetDefaultFolder(5) }
            )
        }
        "inbox"  { $expand = @(@{ Name = "Inbox";      Obj = $ns.GetDefaultFolder(6) }) }
        "sent"   { $expand = @(@{ Name = "Sent Items"; Obj = $ns.GetDefaultFolder(5) }) }
        "drafts" { $expand = @(@{ Name = "Drafts";     Obj = $ns.GetDefaultFolder(16) }) }
        default  { $expand = @(@{ Name = $name;        Obj = (Get-OutlookFolder $name) }) }
    }
    foreach ($f in $expand) {
        $key = $f.Name.ToLower()
        if (-not $seen.ContainsKey($key)) { $seen[$key] = $true; $folders += $f }
    }
}
if ($folders.Count -eq 0) {
    Write-Host "ERROR: No valid folder specified in -Folder '$Folder'." -ForegroundColor Red
    exit 1
}

# === Ensure output directory exists ===
if (-not $DryRun) {
    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }
    if (-not (Test-Path $indexFile)) {
        "EmailEntryID,Subject,Sender,ReceivedDate,Folder,AttachmentName,OriginalSizeMB,SavedPath,Category,ExtractedDate" | Out-File -FilePath $indexFile -Encoding UTF8
    }
}

# === Helper: Sanitize filename ===
function Sanitize-Filename($name) {
    $invalid = [IO.Path]::GetInvalidFileNameChars() -join ''
    $sanitized = $name -replace "[$([regex]::Escape($invalid))]", '_'
    if ($sanitized.Length -gt 80) {
        $ext = [IO.Path]::GetExtension($sanitized)
        $sanitized = $sanitized.Substring(0, 76) + $ext
    }
    return $sanitized
}

# === Helper: Get short sender name ===
function Get-ShortSender($senderName) {
    $parts = $senderName -split '[,\s]+'
    if ($parts.Count -ge 2) {
        return ($parts[0] + "_" + $parts[1]).Substring(0, [Math]::Min(20, ($parts[0] + "_" + $parts[1]).Length))
    }
    return $senderName.Substring(0, [Math]::Min(20, $senderName.Length))
}

# === Helper: Never overwrite an existing file ===
# With a flat output folder, two different emails can produce the same
# "date_sender_filename". Append " (2)", " (3)", ... on collision so no
# extracted attachment is ever silently lost.
function Get-UniquePath($path) {
    if (-not (Test-Path -LiteralPath $path)) { return $path }
    $dir  = [IO.Path]::GetDirectoryName($path)
    $base = [IO.Path]::GetFileNameWithoutExtension($path)
    $ext  = [IO.Path]::GetExtension($path)
    $n = 2
    do {
        $candidate = Join-Path $dir ("{0} ({1}){2}" -f $base, $n, $ext)
        $n++
    } while (Test-Path -LiteralPath $candidate)
    return $candidate
}

# === Main Processing ===
$cutoffDate = (Get-Date).AddDays(-$OlderThanDays)
$totalExtracted = 0
$totalSizeMB = 0
$totalEmails = 0
$results = @()

foreach ($folderInfo in $folders) {
    $currentFolder = $folderInfo.Obj
    $folderName = $folderInfo.Name
    Write-Host "`n=== Processing: $folderName ===" -ForegroundColor Cyan

    $items = $currentFolder.Items
    $items.Sort("[ReceivedTime]", $false)  # oldest first

    if ($OlderThanDays -le 0) {
        $filtered = $items
        $itemCount = $items.Count
        Write-Host "All items (no date filter): $itemCount"
    } else {
        $dateFilter = $cutoffDate.ToString("MM/dd/yyyy")
        $filtered = $items.Restrict("[ReceivedTime] < '$dateFilter'")
        $itemCount = $filtered.Count
        Write-Host "Items older than $OlderThanDays days: $itemCount"
    }

    $processed = 0
    $item = $filtered.GetFirst()

    while ($item -ne $null -and $processed -lt $MaxItems) {
        try {
            if ($item.Attachments.Count -gt 0) {
                $hasLargeAttachment = $false
                $savedPaths = @()

                for ($a = 1; $a -le $item.Attachments.Count; $a++) {
                    $att = $item.Attachments.Item($a)
                    $sizeMB = [Math]::Round($att.Size / 1MB, 2)

                    # Skip inline images (OLE type 6) and small attachments
                    if ($sizeMB -ge $MinSizeMB -and $att.Type -ne 6) {
                        $hasLargeAttachment = $true
                        $totalExtracted++
                        $totalSizeMB += $sizeMB

                        $shortSender = Get-ShortSender $item.SenderName
                        $safeFilename = Sanitize-Filename $att.FileName
                        $saveName = "{0}_{1}_{2}" -f $item.ReceivedTime.ToString("yyyy-MM-dd"), $shortSender, $safeFilename
                        $category = Get-FileCategory $att.FileName
                        # Save directly into the chosen output folder (no category /
                        # month sub-folders). Category is kept as index metadata only.
                        $savePath = Join-Path $OutputPath $saveName

                        $result = [PSCustomObject]@{
                            Folder     = $folderName
                            Date       = $item.ReceivedTime.ToString("yyyy-MM-dd")
                            Sender     = $item.SenderName
                            Subject    = $item.Subject.Substring(0, [Math]::Min(60, $item.Subject.Length))
                            Attachment = $att.FileName
                            SizeMB     = $sizeMB
                            Category   = $category
                            SavePath   = $savePath
                        }
                        $results += $result

                        if ($DryRun) {
                            Write-Host ("  [{0}] {1,7:N1} MB  {2,-30} | {3}" -f `
                                $item.ReceivedTime.ToString("yyyy-MM-dd"), $sizeMB, `
                                $att.FileName.Substring(0, [Math]::Min(30, $att.FileName.Length)), `
                                $item.Subject.Substring(0, [Math]::Min(50, $item.Subject.Length)))
                        } else {
                            # Never overwrite an existing file (flat folder means
                            # possible name clashes across months / senders).
                            $savePath = Get-UniquePath $savePath
                            $result.SavePath = $savePath

                            $att.SaveAsFile($savePath)
                            Write-Host ("  SAVED: {0} ({1:N1} MB)" -f ([IO.Path]::GetFileName($savePath)), $sizeMB) -ForegroundColor Green

                            # Append to index CSV
                            $csvLine = '"{0}","{1}","{2}","{3}","{4}","{5}",{6},"{7}","{8}","{9}"' -f `
                                $item.EntryID, `
                                ($item.Subject -replace '"','""'), `
                                ($item.SenderName -replace '"','""'), `
                                $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm"), `
                                $folderName, `
                                ($att.FileName -replace '"','""'), `
                                $sizeMB, `
                                ($savePath -replace '"','""'), `
                                $category, `
                                (Get-Date).ToString("yyyy-MM-dd HH:mm")
                            $csvLine | Out-File -FilePath $indexFile -Encoding UTF8 -Append

                            $savedPaths += $savePath
                        }
                    }
                }

                if ($hasLargeAttachment -and -not $DryRun) {
                    # Insert file paths as links at the top of the email body
                    if ($savedPaths.Count -gt 0) {
                        if ($item.BodyFormat -eq 2) {
                            # HTML format
                            $htmlBlock = "<div style='background:#f0f7ff;border:1px solid #b3d4fc;padding:8px 12px;margin-bottom:10px;font-family:Calibri,sans-serif;font-size:11px;'>"
                            $htmlBlock += "<b>&#128206; Attachment extracted to:</b><br>"
                            foreach ($sp in $savedPaths) {
                                $fileUri = "file:///" + (($sp -replace '\\', '/') -replace ' ', '%20')
                                $fileName = [IO.Path]::GetFileName($sp)
                                $webUrl = Get-OneDriveWebUrl $sp
                                $htmlBlock += "<a href=`"$fileUri`">&#128193; $fileName</a>"
                                if ($webUrl) {
                                    $htmlBlock += " &nbsp;|&nbsp; <a href=`"$webUrl`">&#127760; Open in browser</a>"
                                }
                                $htmlBlock += "<br>"
                            }
                            if (Get-OneDriveWebUrl $savedPaths[0]) {
                                $htmlBlock += "<span style='color:#888;font-size:10px;'>&#128193; = local (this PC) &nbsp; &#127760; = OneDrive web (mobile / forwarding)</span>"
                            }
                            $htmlBlock += "</div>"

                            try {
                                if ($item.HTMLBody -match '<body[^>]*>') {
                                    $item.HTMLBody = $item.HTMLBody -replace '(<body[^>]*>)', "`$1$htmlBlock"
                                } else {
                                    $item.HTMLBody = $htmlBlock + $item.HTMLBody
                                }
                            } catch {
                                Write-Host "  WARNING: Could not insert link into email body: $($_.Exception.Message)" -ForegroundColor Yellow
                            }
                        } else {
                            # Plain text
                            $pathLines = "Attachment extracted to:`n"
                            foreach ($sp in $savedPaths) {
                                $pathLines += "  Local: $sp`n"
                                $webUrl = Get-OneDriveWebUrl $sp
                                if ($webUrl) { $pathLines += "  Web:   $webUrl`n" }
                            }
                            try {
                                $item.Body = $pathLines + "`n" + $item.Body
                            } catch {
                                Write-Host "  WARNING: Could not insert link into email body: $($_.Exception.Message)" -ForegroundColor Yellow
                            }
                        }
                    }

                    # Remove attachments if requested (iterate in reverse to avoid index shift)
                    if ($RemoveAfterExtract) {
                        for ($a = $item.Attachments.Count; $a -ge 1; $a--) {
                            $att = $item.Attachments.Item($a)
                            $sizeMB = [Math]::Round($att.Size / 1MB, 2)
                            if ($sizeMB -ge $MinSizeMB -and $att.Type -ne 6) {
                                $attName = $att.FileName
                                $att.Delete()
                                Write-Host "    REMOVED from email: $attName" -ForegroundColor Yellow
                            }
                        }
                    }

                    $item.Save()
                }

                if ($hasLargeAttachment) { $totalEmails++ }
            }
        } catch {
            Write-Host "  ERROR processing item: $($_.Exception.Message)" -ForegroundColor Red
        }

        $processed++
        $item = $filtered.GetNext()
    }
}

# === Summary ===
Write-Host ""
Write-Host "=============================" -ForegroundColor Cyan
if ($DryRun) {
    Write-Host "  DRY RUN SUMMARY (no changes made)" -ForegroundColor Yellow
} else {
    Write-Host "  EXTRACTION COMPLETE" -ForegroundColor Green
}
Write-Host "=============================" -ForegroundColor Cyan
Write-Host "Emails with large attachments: $totalEmails"
Write-Host "Attachments found:             $totalExtracted"
Write-Host ("Total size:                    {0:N1} MB ({1:N2} GB)" -f $totalSizeMB, ($totalSizeMB / 1024))
Write-Host ""
if (-not $DryRun) {
    Write-Host "Saved to:  $OutputPath" -ForegroundColor Green
    Write-Host "Index CSV: $indexFile" -ForegroundColor Green
    if ($RemoveAfterExtract) {
        Write-Host ""
        Write-Host "WARNING: Attachments were REMOVED from emails (permanent)." -ForegroundColor Yellow
    } else {
        Write-Host ""
        Write-Host "Attachments are still in emails. To free mailbox space, re-run with:" -ForegroundColor Cyan
        Write-Host "  .\Extract-Attachments.ps1 -DryRun `$false -RemoveAfterExtract `$true"
    }
} else {
    Write-Host ""
    Write-Host "To extract these attachments, re-run with:" -ForegroundColor Cyan
    Write-Host "  .\Extract-Attachments.ps1 -DryRun `$false"
}

# === Output results table for large runs ===
if ($results.Count -gt 0 -and $results.Count -le 50) {
    Write-Host ""
    $results | Format-Table -Property Date, @{L='Size(MB)';E={$_.SizeMB};F='N1'}, Category, Attachment, @{L='Subject';E={$_.Subject}} -AutoSize
}
