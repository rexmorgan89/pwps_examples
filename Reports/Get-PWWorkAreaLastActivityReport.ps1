#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Get-PWWorkAreaLastActivityReport
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Reports the most recent file update in each work area to find dormant projects.
####    ###########    #######         ######    #         ####         #    ######          ######    ########    ######| |#  Author                   : GitHub Copilot
####    ##########-    ####               ###                ##         #    ####              ####    ########    ######| |#  Creation Date            : 10-27-2026
#                     #####     ######    ###      #####     ###    #####    ###     ######     ###    ########    ######| |#  Modified By              : GitHub Copilot
#                    #####    #########    ##    ########     ##    #####    ##     ########    ###    ########    ######| |#  Modified Date            : 10-27-2026
####    ##########      ##                 ##    #########    ##    #####    ##                  ##    ########    ######| |#  Script Version           : 1.0
####    ############    ##                 ##    #########    ##    #####    ##                  ##    ########    ######| |#  PowerShell Version       : Windows PowerShell 5.1 x64
####    ############    ##    ###############    #########    ##    #####    ##     ###############    #######     ######| |#  ProjectWise Version      : Datasource with dms_doc.o_fupdatetime
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
##  Reads the latest file update per folder with one SQL query, rolls folders up to their work area and flags work areas idle for -DormantDays.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWRichProjects, Get-PWFolders, Select-PWSQLDataTable, Undo-PWLogin.
##  Read-only; does not archive or change work areas.
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
Finds dormant work areas by their last file update.

.DESCRIPTION
For every work area below -FolderPath, reports the most recent file update date of any
document in the work area, the document count and the number of days since the last
update. Work areas with no update in -DormantDays days, or with no documents at all,
are flagged as dormant. Use the report to plan archiving.

.PARAMETER FolderPath
Optional folder path to limit the report. Reports every work area when empty.

.PARAMETER DormantDays
Days without a file update before a work area is flagged. Defaults to 365.

.PARAMETER OutputPath
Full path of the Excel workbook to create. Defaults to a timestamped file beside the script.

.EXAMPLE
& '.\Reports\Get-PWWorkAreaLastActivityReport.ps1' -FolderPath 'Projects' -DormantDays 730
Flags work areas below 'Projects' with no file updates in two years.

.INPUTS
None.

.OUTPUTS
An Excel workbook and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab and ImportExcel. Uses the active ProjectWise session or prompts for a
Bentley IMS login, and logs out only if this script opened the session. Activity that
does not change files (viewing, attribute edits) is not counted. Nested work areas are
counted under the deepest work area that contains each folder.
#>

[CmdletBinding()]
param(
    [string] $FolderPath,

    [ValidateRange(1, 36500)]
    [int] $DormantDays = 365,

    [string] $OutputPath = (Join-Path $PSScriptRoot ('WorkAreaLastActivity_{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd_HHmmss')))
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

        $searchParameters = @{ PopulatePaths = $true; ErrorAction = 'Stop' }
        $folderParameters = @{ ErrorAction = 'Stop' }
        if (-not [string]::IsNullOrWhiteSpace($FolderPath)) {
            $searchParameters['FolderPath'] = $FolderPath
            $folderParameters['FolderPath'] = $FolderPath
        }
        $workAreas = @(Get-PWRichProjects @searchParameters | Sort-Object { $_.FullPath.Length } -Descending)
        if ($workAreas.Count -eq 0) {
            Write-Warning 'No work areas were found.'
            return
        }
        $folders = @(Get-PWFolders @folderParameters)

        $activity = @{}
        $sql = 'SELECT o_projectno, MAX(o_fupdatetime) AS LastUpdate, COUNT(*) AS DocumentCount FROM dms_doc GROUP BY o_projectno'
        foreach ($row in (Select-PWSQLDataTable -SQLSelectStatement $sql -ErrorAction Stop).Rows) {
            $activity[[int] $row.o_projectno] = $row
        }

        $totals = @{}
        foreach ($workArea in $workAreas) { $totals[$workArea.FullPath] = @{ LastUpdate = $null; Documents = 0 } }
        foreach ($folder in $folders) {
            $row = $activity[[int] $folder.ProjectID]
            if ($null -eq $row) { continue }
            # Work areas are sorted deepest first, so the first match is the closest owner.
            $owner = $workAreas | Where-Object { $folder.FullPath -eq $_.FullPath -or $folder.FullPath.StartsWith($_.FullPath + '\') } | Select-Object -First 1
            if ($null -eq $owner) { continue }
            $total = $totals[$owner.FullPath]
            $total.Documents += [int] $row.DocumentCount
            $lastUpdate = $row.LastUpdate -as [DateTime]
            if ($null -ne $lastUpdate -and ($null -eq $total.LastUpdate -or $lastUpdate -gt $total.LastUpdate)) { $total.LastUpdate = $lastUpdate }
        }

        [DateTime] $now = Get-Date
        $report = foreach ($workArea in $workAreas) {
            $total = $totals[$workArea.FullPath]
            $idleDays = if ($total.LastUpdate) { [int] ($now - $total.LastUpdate).TotalDays } else { $null }
            [pscustomobject] @{
                WorkArea       = $workArea.FullPath
                Documents      = $total.Documents
                LastFileUpdate = $total.LastUpdate
                DaysIdle       = $idleDays
                Dormant        = ($null -eq $idleDays) -or ($idleDays -ge $DormantDays)
            }
        }

        $report | Sort-Object -Property @{ Expression = 'Dormant'; Descending = $true }, @{ Expression = 'DaysIdle'; Descending = $true } |
            Export-Excel -Path $OutputPath -WorksheetName 'WorkAreas' -AutoSize -FreezeTopRow -BoldTopRow
        $dormant = @($report | Where-Object Dormant).Count
        Write-Host "Work areas: $($report.Count); dormant for ${DormantDays}+ day(s) or empty: ${dormant}. Report saved to '${OutputPath}'."
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
