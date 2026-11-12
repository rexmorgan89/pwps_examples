#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Get-PWDatasourceHealthSnapshot
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Captures a point-in-time summary of a datasource: object counts, statistics and storage areas.
####    ###########    #######         ######    #         ####         #    ######          ######    ########    ######| |#  Author                   : GitHub Copilot
####    ##########-    ####               ###                ##         #    ####              ####    ########    ######| |#  Creation Date            : 10-27-2026
#                     #####     ######    ###      #####     ###    #####    ###     ######     ###    ########    ######| |#  Modified By              : GitHub Copilot
#                    #####    #########    ##    ########     ##    #####    ##     ########    ###    ########    ######| |#  Modified Date            : 10-27-2026
####    ##########      ##                 ##    #########    ##    #####    ##                  ##    ########    ######| |#  Script Version           : 1.0
####    ############    ##                 ##    #########    ##    #####    ##                  ##    ########    ######| |#  PowerShell Version       : Windows PowerShell 5.1 x64
####    ############    ##    ###############    #########    ##    #####    ##     ###############    #######     ######| |#  ProjectWise Version      : Compatible ProjectWise Explorer installation required
####    ###########     ###     ######    ###    #########    ##    #####    ###     #######    ###     #####      ######| |#  PWPS Module Version      : Not required
####                   ####               ###    #########    ##    #####    ####              #####               ######| |#  PWPS_DAB Module Version  : Repository reference: 24.0.2
####                  #######           #####    #########    ##    #####    ######           ########        #    ######| |#
##############################################################################################################    #######| |#  Requirements:
###################################################################################################     ###       #######| |#  pwps_dab, ImportExcel and an administrator login.
##################################################################################################              #########| |#
#######################################################################################################      ############| |#
#########################################################################################################################| |#
#####################################################################################################################################################################################################################
##
##  Description/Notes:
##  Records counts of work areas, environments, workflows, states, storage areas and connected users plus datasource statistics.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWStatistics, Get-PWRichProjects, Get-PWEnvironments, Get-PWWorkflows, Get-PWStates, Get-PWStorageAreaList, Get-PWConnectedUserList, Undo-PWLogin.
##  Read-only; does not change ProjectWise data.
#####################################################################################################################################################################################################################
##
##   Disclaimer:   The information is provided "as is," for example purposes only, and shall not be maintained or supported by Bentley.
##   Bentley makes no warranty as to the effectiveness of this content in a specific situation or for a specific user or specific users.
##   The use of this or similar content in any kind of production environment is entirely at the risk of the user.
##   Bentley suggests extensive testing before deployment. Bentley expressly disclaims express or implied warranties of any kind,
##   including without limitation warranties of merchantability, non-infringement, or fitness for a particular purpose.
##   Bentley does not warrant the accuracy or completeness of any information, text, links, or other items contained within
##   the script/module; that the script/module is error-free. Bentley further expressly disclaims any warranty to authorized users or to any third party.
##
#####################################################################################################################################################################################################################

<#
.SYNOPSIS
Captures a health snapshot of the current ProjectWise datasource.

.DESCRIPTION
Writes an Excel workbook with:
- Summary: counts of work areas, environments, workflows, states, storage areas and
  connected users, with the time of the snapshot.
- Statistics: the output of Get-PWStatistics.
- StorageAreas: each storage area with its host and path.
Run it before and after upgrades or migrations, or on a schedule, and compare snapshots.

.PARAMETER RefreshStatistics
Passes -ForceRefresh to Get-PWStatistics. This can slow the datasource; use outside
working hours.

.PARAMETER OutputPath
Full path of the Excel workbook to create. Defaults to a timestamped file beside the script.

.EXAMPLE
& '.\DatasourceUtilities\Get-PWDatasourceHealthSnapshot.ps1'
Captures a snapshot using cached statistics.

.INPUTS
None.

.OUTPUTS
An Excel workbook and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab and ImportExcel. Uses the active ProjectWise session or prompts for a
Bentley IMS login, and logs out only if this script opened the session. A failing
section is reported as a warning and the remaining sections still run.
#>

[CmdletBinding()]
param(
    [switch] $RefreshStatistics,

    [string] $OutputPath = (Join-Path $PSScriptRoot ('DatasourceHealth_{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd_HHmmss')))
)

begin {
    $InformationPreference = 'Continue'
    # $VerbosePreference = 'Continue' # Uncomment this line if you want to see verbose messages.
    # Log all script output to a file for easy reference later if needed.
    [string] $lastRunLogFilePath = "$PSCommandPath.$(Get-Date -Format 'yyyyMMdd_HHmmss').LastRun.log"
    Start-Transcript -Path $lastRunLogFilePath -ErrorAction Stop

    # Display the time that this script started running.
    [DateTime] $startTime = Get-Date
    Write-Host "Starting script at '$($startTime.ToString('u'))'."

    [bool] $loggedInByScript = $false

    function ConvertTo-PlainObject {
        param([Parameter(ValueFromPipeline)] $InputObject)
        process {
            if ($InputObject -is [System.Data.DataRow]) {
                $row = [ordered] @{}
                foreach ($column in $InputObject.Table.Columns) { $row[$column.ColumnName] = $InputObject[$column] }
                [pscustomobject] $row
            }
            elseif ($null -ne $InputObject) { $InputObject }
        }
    }

    function Measure-Section {
        param([string] $Name, [scriptblock] $Script)
        try { @(& $Script).Count }
        catch {
            Write-Warning "Could not count ${Name}: $($_.Exception.Message)"
            $null
        }
    }
}

process {
    try {
        Import-Module pwps_dab -DisableNameChecking -ErrorAction Stop
        Import-Module ImportExcel -ErrorAction Stop

        $currentDatasource = Get-PWCurrentDatasource -ErrorAction SilentlyContinue
        if ([string]::IsNullOrWhiteSpace([string] $currentDatasource)) {
            $null = New-PWLogin -BentleyIMS -Verbose -ErrorAction Stop
            $loggedInByScript = $true
            $currentDatasource = Get-PWCurrentDatasource -ErrorAction SilentlyContinue

            if ([string]::IsNullOrWhiteSpace([string] $currentDatasource)) {
                throw 'ProjectWise login did not establish an active datasource connection.'
            }
        }
        Write-Host "Connected to datasource '${currentDatasource}'."

        $storageAreas = @(Get-PWStorageAreaList -ErrorAction Continue)
        $summary = [ordered] @{
            Datasource     = [string] $currentDatasource
            SnapshotTime   = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            WorkAreas      = Measure-Section 'work areas' { Get-PWRichProjects -ErrorAction Stop }
            Environments   = Measure-Section 'environments' { Get-PWEnvironments -ErrorAction Stop }
            Workflows      = Measure-Section 'workflows' { Get-PWWorkflows -ErrorAction Stop }
            States         = Measure-Section 'states' { Get-PWStates -ErrorAction Stop }
            StorageAreas   = $storageAreas.Count
            ConnectedUsers = Measure-Section 'connected users' { Get-PWConnectedUserList -ErrorAction Stop }
        }
        $summary.GetEnumerator() | Select-Object @{ Name = 'Item'; Expression = { $_.Key } }, @{ Name = 'Value'; Expression = { $_.Value } } |
            Export-Excel -Path $OutputPath -WorksheetName 'Summary' -AutoSize -BoldTopRow
        $summary.GetEnumerator() | ForEach-Object { Write-Host ('{0,-15} {1}' -f $_.Key, $_.Value) }

        try {
            $statistics = @(Get-PWStatistics -ForceRefresh:$RefreshStatistics -IgnoreLastLoginCount -ErrorAction Stop | ConvertTo-PlainObject)
            if ($statistics.Count -gt 0) {
                $statistics | Export-Excel -Path $OutputPath -WorksheetName 'Statistics' -AutoSize -FreezeTopRow -BoldTopRow
            }
        }
        catch {
            Write-Warning "Could not read datasource statistics: $($_.Exception.Message)"
        }

        if ($storageAreas.Count -gt 0) {
            $storageAreas | Select-Object Name, ID, Description, Host, Path, Protocol |
                Export-Excel -Path $OutputPath -WorksheetName 'StorageAreas' -AutoSize -FreezeTopRow -BoldTopRow
        }
        Write-Host "Snapshot saved to '${OutputPath}'."
    }
    catch {
        if ($loggedInByScript) { $null = Undo-PWLogin -ErrorAction SilentlyContinue }
        $null = Stop-Transcript
        throw
    }
}

end {
    try {
        if ($loggedInByScript) { $null = Undo-PWLogin -ErrorAction SilentlyContinue }

        # Display the time that this script finished running, and how long it took to run.
        [DateTime] $finishTime = Get-Date
        [TimeSpan] $elapsedTime = $finishTime - $startTime
        Write-Host "Finished script at '$($finishTime.ToString('u'))'. Took '$elapsedTime' to run."
    }
    finally {
        Stop-Transcript
    }
}
