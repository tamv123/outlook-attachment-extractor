@echo off
REM Portable launcher for the Outlook Attachment Extractor.
REM Double-click this file. It runs the interactive menu against your own
REM Outlook and saves to your own OneDrive/Downloads. No install required.
REM Keep this .bat next to Run-Extractor.ps1 and Extract-Attachments.ps1.

cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Run-Extractor.ps1"

if errorlevel 1 (
  echo.
  echo The extractor exited with an error. See the messages above.
  pause
)
