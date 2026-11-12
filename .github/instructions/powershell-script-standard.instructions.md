---
description: "Use when creating, editing, reviewing, or standardizing PowerShell scripts (.ps1): Bentley banner comment header, comment-based help, begin/process/end lifecycle, .LastRun.log transcript logging, ProjectWise login."
applyTo: "**/*.ps1"
---

# PowerShell Script Standard

[docs/SCRIPT_STANDARD.md](../../docs/SCRIPT_STANDARD.md) is the authority for metadata, help, login, and lifecycle rules. [Get-PWDatasourceHealthSnapshot.ps1](../../DatasourceUtilities/Get-PWDatasourceHealthSnapshot.ps1) is a complete example. Apply the standard to every new script and to any script you are asked to update, without waiting to be asked again.

## Banner header

- Copy the banner from an existing retained script. Change only the text after `| |#  `; the ASCII art and the `| |#` column must stay identical on every line.
- Fill every field. Dates use `MM-dd-yyyy`. Use `Unknown` when history cannot establish a value; never invent authors or dates.
- When modifying a script, update `Modified By`, `Modified Date`, and `Script Version`.
- `Description/Notes` states the purpose, the ProjectWise cmdlets used, and anything destructive.

## Help placement

- Leave one blank line between the banner and `<# .SYNOPSIS`; without it PowerShell does not detect the help.
- Include `.SYNOPSIS`, `.DESCRIPTION`, a `.PARAMETER` for every parameter, a safe `.EXAMPLE`, and `.NOTES`.

## Lifecycle and logging

`[CmdletBinding()]` and `param`, then `begin`, `process`, `end`. Define helper functions in `begin`. The `begin` block opens with:

```powershell
$InformationPreference = 'Continue'
# $VerbosePreference = 'Continue' # Uncomment this line if you want to see verbose messages.
# Log all script output to a file for easy reference later if needed.
[string] $lastRunLogFilePath = "$PSCommandPath.$(Get-Date -Format 'yyyyMMdd_HHmmss').LastRun.log"
Start-Transcript -Path $lastRunLogFilePath -ErrorAction Stop

# Display the time that this script started running.
[DateTime] $startTime = Get-Date
Write-Host "Starting script at '$($startTime.ToString('u'))'."
```

- `process` wraps its work in `try`/`catch`; the `catch` logs out sessions the script opened, stops the transcript, and rethrows.
- `end` reports elapsed time and calls `Stop-Transcript` in a `finally`.
- Keep the existing execution model for sources in [docs/LIFECYCLE_EXCEPTIONS.md](../../docs/LIFECYCLE_EXCEPTIONS.md), modules, and profiles.

## Validate

Run `& .\.github\scripts\Test-ScriptStandard.ps1 -Path <script or folder>` and fix findings until it reports none. For batches, use its output as the work list instead of tracking files by hand.
