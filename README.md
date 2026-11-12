# ProjectWise PowerShell Examples

A library of example PowerShell scripts for administering Bentley ProjectWise with the `pwps_dab` module. The scripts cover reporting, auditing, bulk updates, and housekeeping. Each top-level folder holds the scripts for one area of ProjectWise and has a `Readme.txt` that describes them.

> **Disclaimer:** These scripts are provided "as is" as examples only. Bentley does not maintain or support them. Test them thoroughly in a non-production datasource before you use them. The full disclaimer is in the header of each script.

---

## Contents

- [Prerequisites](#prerequisites)
- [Getting Started](#getting-started)
- [Repository Layout](#repository-layout)
- [Script Catalog](#script-catalog)
- [Script Conventions](#script-conventions)
- [Logging](#logging)
- [Testing and Validation](#testing-and-validation)
- [Reference Documentation](#reference-documentation)
- [Contributing](#contributing)

---

## Prerequisites

| Requirement | Notes |
|---|---|
| Windows PowerShell 5.1 (x64) | Required by every script. |
| ProjectWise Explorer | Must be installed on the workstation. |
| [`pwps_dab`](https://www.powershellgallery.com/packages/pwps_dab) | Required by all scripts except those in `General/`. |
| [`ImportExcel`](https://www.powershellgallery.com/packages/ImportExcel) | Required by scripts that read or write `.xlsx` files. |
| Rendition service / Plot IPS Files add-in | Needed only for `Renditions/`. |

Install the modules from the PowerShell Gallery:

```powershell
Install-Module -Name pwps_dab -Scope CurrentUser
Install-Module -Name ImportExcel -Scope CurrentUser
```

You can check a workstation with [General/Test-PWWorkstationPrereqs.ps1](General/Test-PWWorkstationPrereqs.ps1).

---

## Getting Started

1. Clone the repository.
2. Open **Windows PowerShell 5.1 (x64)**.
3. Read the help for the script you want to run:

   ```powershell
   Get-Help .\Reports\Get-PWInactiveUsersReport.ps1 -Full
   ```

4. Run the script. If no ProjectWise session is active, you are prompted to log in through Bentley IMS:

   ```powershell
   .\Reports\Get-PWInactiveUsersReport.ps1
   ```

Scripts that change data (for example `Sync-*`, `Set-*`, `Update-*`, `Disable-*`, `New-*`, `Submit-*`) state this in their help. Try them on a test datasource first.

---

## Repository Layout

```
pwps_examples/
├── .github/
│   ├── instructions/        # Copilot instructions for the script standard
│   └── scripts/             # Test-ScriptStandard.ps1 validator
├── <Area>/                  # One folder per ProjectWise area (see catalog)
│   ├── *.ps1
│   └── Readme.txt
├── docs/                    # Script standard, cmdlet references, API specs
└── Tests/                   # Smoke tests against a live datasource
```

---

## Script Catalog

### AccessControl
| Script | Description |
|---|---|
| [Export-PWMemberFolderAccess.ps1](AccessControl/Export-PWMemberFolderAccess.ps1) | Lists the folders where a user, group, or user list has explicit access. |
| [Find-PWFoldersWithExplicitSecurity.ps1](AccessControl/Find-PWFoldersWithExplicitSecurity.ps1) | Lists folders that have explicit (non-inherited) access control. |

### AuditTrail
| Script | Description |
|---|---|
| [Export-PWFolderAuditTrail.ps1](AuditTrail/Export-PWFolderAuditTrail.ps1) | Exports the audit trail of a folder and its documents to Excel. |
| [Get-PWUserActivitySummary.ps1](AuditTrail/Get-PWUserActivitySummary.ps1) | Summarizes audit trail activity per user. |

### ConnectedProjects
| Script | Description |
|---|---|
| [Get-PWWorkAreaConnectionReport.ps1](ConnectedProjects/Get-PWWorkAreaConnectionReport.ps1) | Reports which work areas are connected to a cloud project. |
| [New-PWWorkAreasFromSpreadsheet.ps1](ConnectedProjects/New-PWWorkAreasFromSpreadsheet.ps1) | Creates work areas in bulk from an Excel file, with checks first. |
| [iTwinMembershipManagementTool/](ConnectedProjects/iTwinMembershipManagementTool) | WPF tool (`ManageUsersConnectedProjects.ps1`) for managing iTwin memberships, role assignments, role copies, membership reports, and bulk actions. |

### DatasourceUtilities
| Script | Description |
|---|---|
| [Compare-PWDatasourceUsersAndGroups.ps1](DatasourceUtilities/Compare-PWDatasourceUsersAndGroups.ps1) | Compares users, groups, and user lists between two datasources. |
| [Get-PWDatasourceHealthSnapshot.ps1](DatasourceUtilities/Get-PWDatasourceHealthSnapshot.ps1) | Captures a health snapshot of the current datasource. |

### Documents
| Script | Description |
|---|---|
| [Find-PWDocumentsWithoutFiles.ps1](Documents/Find-PWDocumentsWithoutFiles.ps1) | Reports documents that have no file or an empty file. |
| [Get-PWStaleCheckoutsReport.ps1](Documents/Get-PWStaleCheckoutsReport.ps1) | Lists documents that have been checked out for too long. |

### EnvironmentAttributes
| Script | Description |
|---|---|
| [Export-PWDocumentAttributes.ps1](EnvironmentAttributes/Export-PWDocumentAttributes.ps1) | Exports document attributes below a folder to Excel. |
| [Update-PWDocumentAttributesFromExcel.ps1](EnvironmentAttributes/Update-PWDocumentAttributesFromExcel.ps1) | Updates document attributes from an Excel workbook. |

### Environments
| Script | Description |
|---|---|
| [Export-PWEnvironmentDefinitions.ps1](Environments/Export-PWEnvironmentDefinitions.ps1) | Documents environment definitions. |
| [Get-PWEnvironmentUsageReport.ps1](Environments/Get-PWEnvironmentUsageReport.ps1) | Reports how many folders use each environment. |

### FileManipulation
| Script | Description |
|---|---|
| [Export-PWFolderToLocal.ps1](FileManipulation/Export-PWFolderToLocal.ps1) | Exports a folder tree to a local folder. |
| [Get-PWLargestFilesReport.ps1](FileManipulation/Get-PWLargestFilesReport.ps1) | Reports the largest files and folder sizes below a folder. |

### Folder
| Script | Description |
|---|---|
| [Get-PWEmptyFoldersReport.ps1](Folder/Get-PWEmptyFoldersReport.ps1) | Lists empty leaf folders below a folder. |
| [New-PWFolderStructureFromCsv.ps1](Folder/New-PWFolderStructureFromCsv.ps1) | Creates a folder structure from a CSV template. |

### General
| Script | Description |
|---|---|
| [Get-PWModuleUpdateStatus.ps1](General/Get-PWModuleUpdateStatus.ps1) | Reports whether newer versions of the ProjectWise PowerShell modules are available. |
| [Test-PWWorkstationPrereqs.ps1](General/Test-PWWorkstationPrereqs.ps1) | Checks a workstation for ProjectWise PowerShell prerequisites. |

### Groups&UserLists
| Script | Description |
|---|---|
| [Export-PWGroupAndUserListMembership.ps1](Groups&UserLists/Export-PWGroupAndUserListMembership.ps1) | Exports group and user list membership to Excel. |
| [Sync-PWGroupMembershipFromCsv.ps1](Groups&UserLists/Sync-PWGroupMembershipFromCsv.ps1) | Synchronizes group membership with a CSV file. |

### Licensing
| Script | Description |
|---|---|
| [Get-PWConnectedUsersReport.ps1](Licensing/Get-PWConnectedUsersReport.ps1) | Lists connected users and logs the count over time. |
| [Get-PWMonthlyLoginReport.ps1](Licensing/Get-PWMonthlyLoginReport.ps1) | Exports login activity for a month or a year. |

### ManagedWorkspace
| Script | Description |
|---|---|
| [Compare-PWFolderWorkspaces.ps1](ManagedWorkspace/Compare-PWFolderWorkspaces.ps1) | Compares the Managed Workspace CSBs of two folders. |
| [Get-PWWorkAreaWorkspaceReport.ps1](ManagedWorkspace/Get-PWWorkAreaWorkspaceReport.ps1) | Reports the Managed Workspace CSBs of every work area. |

### Renditions
| Script | Description |
|---|---|
| [Get-PWRenditionProfileAssignments.ps1](Renditions/Get-PWRenditionProfileAssignments.ps1) | Reports rendition profiles and the work areas that use them. |
| [Submit-PWRenditionRequests.ps1](Renditions/Submit-PWRenditionRequests.ps1) | Submits the documents in a folder to a rendition profile. |

### Reports
| Script | Description |
|---|---|
| [Get-PWInactiveUsersReport.ps1](Reports/Get-PWInactiveUsersReport.ps1) | Lists users who have not logged in recently. |
| [Get-PWWorkAreaLastActivityReport.ps1](Reports/Get-PWWorkAreaLastActivityReport.ps1) | Finds dormant work areas by their last file update. |

### SavedSearches
| Script | Description |
|---|---|
| [Export-PWSavedSearchList.ps1](SavedSearches/Export-PWSavedSearchList.ps1) | Exports the list of saved searches to Excel. |
| [Export-PWSavedSearchResults.ps1](SavedSearches/Export-PWSavedSearchResults.ps1) | Runs a saved search and exports the results to Excel. |

### Scheduled Tasks
| Script | Description |
|---|---|
| [Get-PWScheduledTaskStatus.ps1](Scheduled%20Tasks/Get-PWScheduledTaskStatus.ps1) | Shows the run status of ProjectWise scheduled tasks. |
| [Register-PWScriptScheduledTask.ps1](Scheduled%20Tasks/Register-PWScriptScheduledTask.ps1) | Schedules a script from this library to run daily or weekly. |

### States
| Script | Description |
|---|---|
| [Get-PWDocumentStateReport.ps1](States/Get-PWDocumentStateReport.ps1) | Reports how many documents are in each workflow state. |
| [Set-PWDocumentStatesFromCsv.ps1](States/Set-PWDocumentStatesFromCsv.ps1) | Changes document states from a CSV list. |

### StorageAreas
| Script | Description |
|---|---|
| [Find-PWFoldersByStorageArea.ps1](StorageAreas/Find-PWFoldersByStorageArea.ps1) | Lists folders by storage area and flags mixed storage in a folder tree. |
| [Get-PWStorageAreaUsage.ps1](StorageAreas/Get-PWStorageAreaUsage.ps1) | Reports storage areas with document counts and disk usage. |

### Tests
| Script | Description |
|---|---|
| [Test-PWReadOnlySmokeTest.ps1](Tests/Test-PWReadOnlySmokeTest.ps1) | Smoke-tests read-only `pwps_dab` cmdlets against the current datasource. |
| [Test-PWDocumentRoundTrip.ps1](Tests/Test-PWDocumentRoundTrip.ps1) | Tests the full document life cycle (create, check out, check in, delete) in a test folder. |

### User
| Script | Description |
|---|---|
| [Export-PWUsersReport.ps1](User/Export-PWUsersReport.ps1) | Exports users to Excel. |
| [Disable-PWUsersFromCsv.ps1](User/Disable-PWUsersFromCsv.ps1) | Disables the users listed in a CSV file. |

### WorkFlows
| Script | Description |
|---|---|
| [Export-PWWorkflowDiagram.ps1](WorkFlows/Export-PWWorkflowDiagram.ps1) | Draws workflows as Mermaid state diagrams. |
| [Find-PWUnusedWorkflows.ps1](WorkFlows/Find-PWUnusedWorkflows.ps1) | Finds workflows that no folder or document uses. |

---

## Script Conventions

All scripts follow [docs/SCRIPT_STANDARD.md](docs/SCRIPT_STANDARD.md). In summary:

- **Header banner:** Each script starts with the Bentley banner comment. It records the name, purpose, author, dates, version, the PowerShell, ProjectWise, and module versions, and the disclaimer.
- **Comment-based help:** `.SYNOPSIS`, `.DESCRIPTION`, a `.PARAMETER` entry for every parameter, safe `.EXAMPLE` entries, and `.NOTES` covering the required modules, inputs and outputs, login behavior, and any destructive effects.
- **Lifecycle:** `[CmdletBinding()]` with `param`, `begin`, `process`, and `end` blocks.
  - `begin` checks dependencies, starts logging, and records the start time.
  - `process` does the work inside `try`/`catch`.
  - `end` closes sessions, reports the elapsed time, and stops the transcript.
- **Login:** Scripts reuse the current session (`Get-PWCurrentDatasource`) or prompt with `New-PWLogin -BentleyIMS`. Scripts that use several datasources select each one explicitly.
- **No secrets in source:** Passwords, tokens, endpoints, and datasource names are never hard-coded, and they must not appear in transcripts or verbose output.

---

## Logging

Each run writes a timestamped transcript next to the script:

```
<ScriptName>.ps1.yyyyMMdd_HHmmss.LastRun.log
```

The `.gitignore` excludes `*.log`, `*.xlsx`, and `*.csv`, so logs and generated reports are not committed.

---

## Testing and Validation

**Standard check:** Before you commit, check new or changed scripts against the standard:

```powershell
& .\.github\scripts\Test-ScriptStandard.ps1 -Path .\Reports\Get-PWInactiveUsersReport.ps1
& .\.github\scripts\Test-ScriptStandard.ps1 -Path .\Reports
```

Fix any findings and run the check again until none remain.

**Smoke tests:** The scripts in [Tests/](Tests) run against a live datasource:

- `Test-PWReadOnlySmokeTest.ps1` is read-only and safe to run on any datasource.
- `Test-PWDocumentRoundTrip.ps1` creates and deletes documents. Point it **only** at a dedicated test folder.

---

## Reference Documentation

The [docs/](docs) folder contains offline reference material:

| File / Folder | Contents |
|---|---|
| [SCRIPT_STANDARD.md](docs/SCRIPT_STANDARD.md) | The script standard that all scripts follow. |
| [CmdletReferenceOutputter.ps1](docs/CmdletReferenceOutputter.ps1) | Writes a cmdlet help reference file for each selected module. |
| `*-Reference.txt` | Cmdlet references for `pwps`, `pwps_dab`, `pwps_cloud`, `pwps_masterdoc`, `pwps_pwdm`, `ImportExcel`, `Microsoft.Entra`, `Pester`, `PSSQLite`, and `PSWriteHTML`. |
| [PerformanceConsiderations.txt](docs/PerformanceConsiderations.txt) | Performance notes for writing scripts. |
| [iTwinAPI_Reference/](docs/iTwinAPI_Reference) | OpenAPI specifications for iTwin platform services. |
| [pwps_cloud/](docs/pwps_cloud) | HTML reference for the `pwps_cloud` module. |

---

## Contributing

1. Add the script to the folder for its area. Create a new folder with a `Readme.txt` if no folder fits.
2. Copy the banner and structure from an existing script. Change only the metadata text.
3. Write complete comment-based help, and describe any destructive behavior in `.NOTES`.
4. Run `Test-ScriptStandard.ps1` until it reports no findings.
5. Add the script to the folder `Readme.txt` and to the [Script Catalog](#script-catalog) above.
