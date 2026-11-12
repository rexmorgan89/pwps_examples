# Script Standard

## Metadata and Help

Retained scripts use the Bentley comment header represented by `PW_CopyFolders_Between_Datasources.ps1` in the adjacent `pwps_bentley` repository. Populate the following values from source history where available:

- Name, Type, Purpose, Author, Creation Date.
- Modified By, Modified Date, Script Version.
- PowerShell Version, ProjectWise Version, PWPS Module Version, PWPS_DAB Module Version.
- Requirements.

Do not invent dates, authors, or version requirements. Use `Unknown` when source history cannot establish a value.

The header is followed by comment-based help containing `.SYNOPSIS`, `.DESCRIPTION`, `.PARAMETER` for every public parameter, a safe `.EXAMPLE`, and `.NOTES`. Notes must identify required modules, inputs/outputs, ProjectWise cmdlets, login behavior, and destructive effects.

## Login Rules

For a script that requires one datasource, use the active ProjectWise datasource or prompt through Bentley IMS:

```powershell
$currentDatasource = Get-PWCurrentDatasource -ErrorAction SilentlyContinue
if ([string]::IsNullOrWhiteSpace([string] $currentDatasource)) {
    $null = New-PWLogin -BentleyIMS -Verbose -ErrorAction Stop
    $currentDatasource = Get-PWCurrentDatasource -ErrorAction SilentlyContinue

    if ([string]::IsNullOrWhiteSpace([string] $currentDatasource)) {
        throw 'ProjectWise login did not establish an active datasource connection.'
    }
}
```

For a script that requires multiple datasources, select each datasource explicitly and then log in through Bentley IMS:

```powershell
$sourceDatasource = Get-PWDatasource |
    Select-Object 'InternalName', 'HostName' |
    Out-GridView -Title 'Select a Source Datasource' -OutputMode Single

if ($null -eq $sourceDatasource) {
    throw 'A source datasource must be selected.'
}

[string] $sourceDatasourceName = '{0}:{1}' -f $sourceDatasource.HostName, $sourceDatasource.InternalName
$null = New-PWLogin $sourceDatasourceName -BentleyIMS -ErrorAction Stop
```

Never store a password, token, endpoint, or datasource name in source. Do not write secrets to transcripts or verbose output.

## Lifecycle Rules

For non-GUI scripts that do not consume pipeline input, use `[CmdletBinding()]`, then `param`, `begin`, `process`, and `end`.

- `begin`: validate/import dependencies, establish safe logging, and record the start time.
- `process`: perform the script's main work.
- `end`: clean up ProjectWise sessions that the script owns, report duration, and stop any started transcript.

Use `try`/`finally` whenever a failure could leave a ProjectWise session or transcript active. Do not add an outer lifecycle wrapper to `.psm1` modules, PowerShell profiles, GUI forms, or existing pipeline/lifecycle implementations; preserve their established execution model.
