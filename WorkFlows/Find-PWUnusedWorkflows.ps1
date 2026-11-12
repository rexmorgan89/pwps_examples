#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Find-PWUnusedWorkflows
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Counts the folders and documents that use each workflow and lists unused workflows.
####    ###########    #######         ######    #         ####         #    ######          ######    ########    ######| |#  Author                   : GitHub Copilot
####    ##########-    ####               ###                ##         #    ####              ####    ########    ######| |#  Creation Date            : 10-27-2026
#                     #####     ######    ###      #####     ###    #####    ###     ######     ###    ########    ######| |#  Modified By              : GitHub Copilot
#                    #####    #########    ##    ########     ##    #####    ##     ########    ###    ########    ######| |#  Modified Date            : 10-27-2026
####    ##########      ##                 ##    #########    ##    #####    ##                  ##    ########    ######| |#  Script Version           : 1.0
####    ############    ##                 ##    #########    ##    #####    ##                  ##    ########    ######| |#  PowerShell Version       : Windows PowerShell 5.1 x64
####    ############    ##    ###############    #########    ##    #####    ##     ###############    #######     ######| |#  ProjectWise Version      : Datasource with dms_proj and dms_doc workflow columns
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
##  Counts folders (dms_proj) and documents (dms_doc) per workflow ID with two SQL queries and flags workflows that nothing uses.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWWorkflows, Select-PWSQLDataTable, Undo-PWLogin.
##  Read-only; does not delete workflows.
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
Finds workflows that no folder or document uses.

.DESCRIPTION
Lists every workflow with the number of folders it is assigned to and the number of
documents that carry it. Workflows with zero folders and zero documents are flagged as
unused and are candidates for clean-up. Workflows still carried by documents but no
longer assigned to any folder are flagged as orphaned.

.PARAMETER OutputPath
Full path of the Excel workbook to create. Defaults to a timestamped file beside the script.

.EXAMPLE
& '.\WorkFlows\Find-PWUnusedWorkflows.ps1'
Writes the workflow usage report.

.INPUTS
None.

.OUTPUTS
An Excel workbook and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab and ImportExcel. Uses the active ProjectWise session or prompts for a
Bentley IMS login, and logs out only if this script opened the session. Workflows can also
be referenced by work area templates or integrations; check before removing them.
#>

[CmdletBinding()]
param(
    [string] $OutputPath = (Join-Path $PSScriptRoot ('WorkflowUsage_{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd_HHmmss')))
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

    function Get-CountByWorkflow {
        param([string] $Table)
        $counts = @{}
        $sql = "SELECT o_workflowno, COUNT(*) AS ItemCount FROM $Table GROUP BY o_workflowno"
        foreach ($row in (Select-PWSQLDataTable -SQLSelectStatement $sql -ErrorAction Stop).Rows) {
            if ($row.o_workflowno -isnot [DBNull]) { $counts[[int] $row.o_workflowno] = [int] $row.ItemCount }
        }
        $counts
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

        $workflows = @(Get-PWWorkflows -ErrorAction Stop)
        $folderCounts = Get-CountByWorkflow -Table 'dms_proj'
        $documentCounts = Get-CountByWorkflow -Table 'dms_doc'

        $report = foreach ($workflow in $workflows) {
            [int] $id = $workflow.ID
            [int] $folders = $folderCounts[$id]
            [int] $documents = $documentCounts[$id]
            [pscustomobject] @{
                Workflow    = $workflow.Name
                ID          = $id
                Description = $workflow.Description
                Folders     = $folders
                Documents   = $documents
                Status      = if ($folders -eq 0 -and $documents -eq 0) { 'Unused' } elseif ($folders -eq 0) { 'Orphaned' } else { 'In use' }
            }
        }
        $report = @($report)

        $report | Sort-Object Status, Workflow |
            Export-Excel -Path $OutputPath -WorksheetName 'Workflows' -AutoSize -FreezeTopRow -BoldTopRow
        $report | Group-Object Status | ForEach-Object { Write-Host "$($_.Name): $($_.Count)" }
        Write-Host "Report saved to '${OutputPath}'."
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
