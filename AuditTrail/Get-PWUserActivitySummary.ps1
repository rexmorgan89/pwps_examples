#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Get-PWUserActivitySummary
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Summarizes audit trail actions per user over a number of days and exports them to Excel.
####    ###########    #######         ######    #         ####         #    ######          ######    ########    ######| |#  Author                   : GitHub Copilot
####    ##########-    ####               ###                ##         #    ####              ####    ########    ######| |#  Creation Date            : 10-27-2026
#                     #####     ######    ###      #####     ###    #####    ###     ######     ###    ########    ######| |#  Modified By              : GitHub Copilot
#                    #####    #########    ##    ########     ##    #####    ##     ########    ###    ########    ######| |#  Modified Date            : 10-27-2026
####    ##########      ##                 ##    #########    ##    #####    ##                  ##    ########    ######| |#  Script Version           : 1.0
####    ############    ##                 ##    #########    ##    #####    ##                  ##    ########    ######| |#  PowerShell Version       : Windows PowerShell 5.1 x64
####    ############    ##    ###############    #########    ##    #####    ##     ###############    #######     ######| |#  ProjectWise Version      : SQL Server datasource with the dms_audt table
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
##  Counts audit trail records per user and action for the last N days with a single SQL query and exports two worksheets.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Select-PWSQLDataTable, Undo-PWLogin.
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
Summarizes ProjectWise audit trail activity per user.

.DESCRIPTION
Queries the dms_audt table for records created in the last -Days days, grouped by user,
object type and action. Writes two worksheets to an Excel workbook:
- Detail: one row per user and action, with the count and most recent time.
- PerUser: total actions and most recent action per user.
Document action codes (1000-1038) are translated to readable names; other codes are
shown as numbers.

.PARAMETER Days
Number of days to look back. Defaults to 30.

.PARAMETER OutputPath
Full path of the Excel workbook to create. Defaults to a timestamped file beside the script.

.EXAMPLE
& '.\AuditTrail\Get-PWUserActivitySummary.ps1' -Days 7
Summarizes the last week of activity.

.INPUTS
None.

.OUTPUTS
An Excel workbook and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab and ImportExcel. Uses the active ProjectWise session or prompts for a
Bentley IMS login, and logs out only if this script opened the session. The date filter
uses SQL Server syntax; adjust it for Oracle datasources. Only records still in dms_audt
are counted; archived or trimmed records are not included.
#>

[CmdletBinding()]
param(
    [ValidateRange(1, 3650)]
    [int] $Days = 30,

    [string] $OutputPath = (Join-Path $PSScriptRoot ('UserActivitySummary_{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd_HHmmss')))
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

    $documentActions = @{
        1000 = 'Unknown Action'; 1001 = 'Created'; 1002 = 'General Properties Modified'; 1003 = 'Attributes Modified'
        1004 = 'File Added'; 1005 = 'File Removed'; 1006 = 'File Replaced'; 1007 = 'Checked In'; 1008 = 'Viewed'
        1009 = 'Checked Out'; 1010 = 'Copied Out'; 1011 = 'Given Out'; 1012 = 'State Changed'; 1013 = 'Final Status Set'
        1014 = 'Final Status Removed'; 1015 = 'New Version Created'; 1016 = 'Moved'; 1017 = 'Copied'
        1018 = 'Permissions Changed'; 1019 = 'Redline Created or Accessed'; 1020 = 'Deleted'; 1021 = 'Exported'
        1022 = 'Freed'; 1023 = 'Extracted'; 1024 = 'Distributed'; 1025 = 'Sent'; 1026 = 'Commented'; 1027 = 'Imported'
        1028 = 'ACL Assigned'; 1029 = 'ACL Modified'; 1030 = 'ACL Removed'; 1031 = 'Revit Action'; 1032 = 'Packed'
        1033 = 'Unpacked'; 1034 = 'WRE Execution Started'; 1035 = 'WRE Execution Succeeded'; 1036 = 'WRE Execution Failed'
        1037 = 'Restored'; 1038 = 'Action Unassigned'
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

        [string] $since = (Get-Date).Date.AddDays(-$Days).ToString('yyyy-MM-dd HH:mm:ss')
        [string] $sql = @"
SELECT o_username, o_objtype, o_action, COUNT(*) AS ActionCount, MAX(o_acttime) AS LastActionTime
FROM dms_audt
WHERE o_acttime >= '$since'
GROUP BY o_username, o_objtype, o_action
"@
        $table = Select-PWSQLDataTable -SQLSelectStatement $sql -ErrorAction Stop

        $detail = @(foreach ($row in $table.Rows) {
                [int] $action = $row.o_action
                [pscustomobject] @{
                    UserName       = [string] $row.o_username
                    ObjectType     = [int] $row.o_objtype
                    ActionCode     = $action
                    Action         = if ($documentActions.ContainsKey($action)) { "Document $($documentActions[$action])" } else { "Action $action" }
                    ActionCount    = [int] $row.ActionCount
                    LastActionTime = $row.LastActionTime
                }
            })

        if ($detail.Count -eq 0) {
            Write-Warning "No audit trail records were found in the last ${Days} day(s)."
            return
        }

        $perUser = $detail | Group-Object -Property UserName | ForEach-Object {
            [pscustomobject] @{
                UserName       = $_.Name
                TotalActions   = ($_.Group | Measure-Object -Property ActionCount -Sum).Sum
                LastActionTime = ($_.Group | Sort-Object -Property LastActionTime -Descending | Select-Object -First 1).LastActionTime
            }
        } | Sort-Object -Property TotalActions -Descending

        $detail | Sort-Object UserName, Action | Export-Excel -Path $OutputPath -WorksheetName 'Detail' -AutoSize -FreezeTopRow -BoldTopRow
        $perUser | Export-Excel -Path $OutputPath -WorksheetName 'PerUser' -AutoSize -FreezeTopRow -BoldTopRow
        Write-Host "Summarized $($perUser.Count) user(s). Report saved to '${OutputPath}'."
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
