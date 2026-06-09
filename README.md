🚀 Work in Progress
This repository is a work in progress, and I am actively looking for feedback and contributors to help refine it!

If you see something that can be optimized, a bug that needs squashing, or want to co-develop a feature, please feel free to:

Open an Issue to discuss your ideas.
Submit a Pull Request with your improvements.
Let's collaborate and make this tool awesome!

# Outlook Attachment Extractor

> **⚠️ Windows Only** — Requires Windows 10/11 with Microsoft Outlook desktop app (classic).

Extract large attachments from your Outlook mailbox to a local folder (e.g. OneDrive), organized by file type and date. Free up mailbox space without losing access to your files.

## What It Does

```
Outlook Mailbox                          Your OneDrive / Downloads
┌──────────────────────┐                 ┌──────────────────────────────────────┐
│ Inbox (50 GB)        │                 │ Downloads/                           │
│  📧 + 📎 15MB report │──extract──────→│   Work/Email Attachments/2025-03/    │
│  📧 + 📎 8MB deck   │                 │     2025-03-15_John_Q1_Report.xlsx   │
│  📧 + 📎 3MB data   │                 │   Analytics & Data/Email Attach.../  │
│  ...                 │                 │     2025-03-10_Jane_Sales_Data.csv   │
│                      │  ←─link─────── │                                      │
│  📧 "Extracted to:   │                 │ Email Attachments/                   │
│      C:\Users\...\   │                 │   _attachment_index.csv              │
│      Q1_Report.xlsx" │                 └──────────────────────────────────────┘
└──────────────────────┘
```

**Key features:**
- 📂 **Auto-categorizes** files into Work, Analytics & Data, Media, Installers & Tools
- 🔗 **Inserts a clickable link** into the email body pointing to the saved file
- 📋 **Creates an index CSV** mapping every extracted attachment back to its source email
- 🧹 **Optionally removes** attachments from emails to free server-side mailbox space
- 🔍 **Dry-run mode** — preview everything before making any changes
- ⏱️ **Batch processing** — processes oldest emails first, with configurable limits

## Requirements

| Requirement | Details |
|---|---|
| **OS** | Windows 10 or Windows 11 |
| **Outlook** | Microsoft Outlook desktop app (**classic** — the "New Outlook" app is NOT supported) |
| **PowerShell** | 5.1+ (included with Windows 10/11) |
| **Mailbox** | Exchange Online, Exchange Server, or Outlook.com account configured in Outlook |

> **Note:** This tool uses the Outlook COM API, which is only available on Windows with the classic Outlook desktop application. It does not work on macOS, Linux, Outlook Web (OWA), or the "New Outlook" app.

## Installation

1. **Clone this repository** (or download the ZIP):
   ```powershell
   git clone https://github.com/tamv123/outlook-attachment-extractor.git
   cd outlook-attachment-extractor
   ```

2. **That's it.** No dependencies to install — it's a single PowerShell script.

## Quick Start

### Step 1: Preview (Dry Run) — Always do this first

```powershell
# See what would be extracted from your Inbox
# Default: attachments > 1 MB, from emails > 90 days old
.\Extract-Attachments.ps1
```

### Step 2: Extract Attachments

```powershell
# Save attachments to OneDrive/Downloads (keeps attachments in email)
.\Extract-Attachments.ps1 -DryRun $false
```

### Step 3 (Optional): Remove from Email to Free Space

```powershell
# Extract AND remove attachments from emails (⚠️ permanent!)
.\Extract-Attachments.ps1 -DryRun $false -RemoveAfterExtract $true
```

## Usage

```
.\Extract-Attachments.ps1 [parameters]
```

### Parameters

| Parameter | Default | Description |
|---|---|---|
| `-Folder` | `inbox` | Outlook folder: `inbox`, `sent`, `drafts`, `all`, or any custom folder name |
| `-MinSizeMB` | `1` | Minimum attachment size in MB to extract |
| `-MaxItems` | `100` | Max emails to process per run (prevents timeouts) |
| `-OlderThanDays` | `90` | Only process emails older than N days |
| `-DryRun` | `$true` | Preview mode — no files saved, no emails modified |
| `-RemoveAfterExtract` | `$false` | Remove attachment from email after saving to disk (**permanent!**) |
| `-OutputPath` | *(auto-detected)* | Output directory. Auto-detects OneDrive, falls back to `~/Downloads` |

### Examples

```powershell
# Preview all folders
.\Extract-Attachments.ps1 -Folder all -DryRun $true

# Extract large files (> 5 MB) from emails older than 6 months
.\Extract-Attachments.ps1 -MinSizeMB 5 -OlderThanDays 180 -DryRun $false

# Extract from Sent Items
.\Extract-Attachments.ps1 -Folder sent -DryRun $false

# Extract to a custom directory
.\Extract-Attachments.ps1 -DryRun $false -OutputPath "D:\EmailBackups"

# Process a custom folder (e.g., "Projects" or "Archive")
.\Extract-Attachments.ps1 -Folder "Projects" -DryRun $false

# Process up to 500 emails in one run
.\Extract-Attachments.ps1 -MaxItems 500 -DryRun $false
```

## Output Structure

Extracted files are organized by category and month:

```
{OutputPath}/
├── Work/
│   └── Email Attachments/
│       ├── 2025-01/
│       │   ├── 2025-01-15_John_Smith_Q4_Report.pptx
│       │   └── 2025-01-20_Jane_Doe_Project_Plan.docx
│       └── 2025-02/
│           └── ...
├── Analytics & Data/
│   └── Email Attachments/
│       └── 2025-01/
│           └── 2025-01-10_Bob_Sales_Data.xlsx
├── Media/
│   └── Email Attachments/
│       └── ...
├── Installers & Tools/
│   └── Email Attachments/
│       └── ...
└── Email Attachments/
    └── _attachment_index.csv          ← master index of all extractions
```

### File Naming

Files are named with date, sender, and original filename for easy identification:

```
{date}_{sender}_{original_filename}
Example: 2025-03-15_John_Smith_Q1_Media_Report.xlsx
```

### Category Mapping

| Category | File Extensions |
|---|---|
| **Work** | `.pptx` `.ppt` `.docx` `.doc` `.pdf` `.msg` `.eml` `.rtf` `.one` `.vsdx` |
| **Analytics & Data** | `.xlsx` `.xls` `.csv` `.json` `.xml` `.parquet` `.accdb` |
| **Media** | `.png` `.jpg` `.jpeg` `.gif` `.bmp` `.svg` `.mp4` `.mp3` `.mov` `.wav` |
| **Installers & Tools** | `.zip` `.rar` `.7z` `.msi` `.exe` `.iso` `.tar` `.gz` |

Unrecognized extensions default to **Work**.

### Index CSV

The `_attachment_index.csv` file tracks every extraction:

| Column | Description |
|---|---|
| `EmailEntryID` | Outlook unique email identifier |
| `Subject` | Email subject line |
| `Sender` | Sender name |
| `ReceivedDate` | When the email was received |
| `Folder` | Source folder (Inbox, Sent Items, etc.) |
| `AttachmentName` | Original filename |
| `OriginalSizeMB` | File size in MB |
| `SavedPath` | Full path to the saved file |
| `Category` | Auto-assigned category |
| `ExtractedDate` | When extraction was performed |

The index is **append-only** — safe to run the script multiple times without duplicating entries.

## How It Works

1. **Connects** to Outlook via the COM API (with automatic retry)
2. **Scans** the target folder for emails older than the specified threshold
3. **Filters** attachments by minimum size (skips inline images)
4. **Saves** each qualifying attachment to the categorized folder structure
5. **Logs** the extraction to the index CSV
6. **Inserts** a clickable `file:///` link at the top of the email body
7. *(Optional)* **Removes** the attachment from the email and saves the modified email

### What happens to the email?

After extraction, a blue info box is inserted at the top of the email body:

> 📎 **Attachment extracted to:**
> `C:\Users\you\OneDrive\Downloads\Work\Email Attachments\2025-03\2025-03-15_Report.pptx`

The original email text is preserved. If you click the link, the file opens directly.

## Recommended Workflow

For large mailboxes, work in phases:

### Phase 1: Quick Assessment
```powershell
# See total size of extractable attachments across all folders
.\Extract-Attachments.ps1 -Folder all -DryRun $true
```

### Phase 2: Extract (Keep in Email)
```powershell
# Start with Inbox, oldest and largest first
.\Extract-Attachments.ps1 -Folder inbox -MinSizeMB 5 -DryRun $false

# Then Sent Items
.\Extract-Attachments.ps1 -Folder sent -MinSizeMB 5 -DryRun $false
```

### Phase 3: Verify & Remove
```powershell
# Verify files exist in your output folder, then free mailbox space
.\Extract-Attachments.ps1 -Folder inbox -MinSizeMB 5 -DryRun $false -RemoveAfterExtract $true
```

### Phase 4: Ongoing Maintenance
Run monthly to keep your mailbox lean:
```powershell
.\Extract-Attachments.ps1 -OlderThanDays 30 -DryRun $false
```

## Troubleshooting

### "Cannot connect to Outlook"
- Make sure the **classic** Outlook desktop app is running (not "New Outlook")
- Try closing and reopening Outlook
- Check that Outlook is not in a "needs password" or "disconnected" state

### "Folder not found"
- The script will list available folders. Use the exact folder name.
- Custom folders must be top-level (under your mailbox root).

### Script won't run (Execution Policy)
```powershell
# Option 1: Run with bypass (recommended for one-time use)
powershell.exe -ExecutionPolicy Bypass -File .\Extract-Attachments.ps1

# Option 2: Set policy for current user (permanent)
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Not all emails are processed
- Outlook in **Cached Mode** only processes locally cached emails.
- To process all server-side emails, change Outlook sync settings:
  **File → Account Settings → Account Settings → Change → Mail to keep offline → All**
- Full sync may take 30–60 minutes depending on mailbox size.

### "New Outlook" vs Classic Outlook
This tool requires the **classic** Outlook desktop app. To check:
- Classic: has the **File** menu with **Account Settings**
- New: has a simplified ribbon with a toggle "Try the new Outlook"

If you're on New Outlook, you can switch back via the toggle in the top-right corner.

## Limitations

- **Windows only** — the Outlook COM API is a Windows-specific technology
- **Classic Outlook only** — "New Outlook" and Outlook Web (OWA) are not supported
- **Cached Mode** — only processes locally cached emails (change sync settings for full access)
- **IRM/S-MIME** — cannot process encrypted or rights-managed emails
- **Shared mailboxes** — may not work with read-only shared folders
- **No undo for removal** — once `-RemoveAfterExtract $true` is used and the email is saved, the attachment is permanently removed from the Exchange server

## License

MIT License. See [LICENSE](LICENSE) for details.
