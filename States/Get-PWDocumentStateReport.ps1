#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Get-PWDocumentStateReport
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Counts documents per workflow and state, overall and per folder.
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
###################################################################################################     ###       #######| |#  pwps_dab, ImportExcel and read access to the target folder.
##################################################################################################              #########| |#
#######################################################################################################      ############| |#
#########################################################################################################################| |#
#####################################################################################################################################################################################################################
##
##  Description/Notes:
##  Searches documents below a folder and counts them by workflow and state, with a per-folder breakdown and a document list.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWDocumentsBySearch, Undo-PWLogin.
##  Read-only; does not change states.
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
Reports how many documents are in each workflow state.

.DESCRIPTION
Writes an Excel workbook with:
- ByState: document count per workflow and state.
- ByFolder: document count per folder, workflow and state.
- Documents: every document with its state (with -IncludeDocumentList).
The Documents sheet can be filtered and saved as a CSV for Set-PWDocumentStatesFromCsv.ps1.

.PARAMETER FolderPath
ProjectWise folder path to search, without the leading 'Documents' folder.

.PARAMETER IncludeDocumentList
Adds a Documents worksheet with FolderPath, DocumentName and State columns.

.PARAMETER OutputPath
Full path of the Excel workbook to create. Defaults to a timestamped file beside the script.

.EXAMPLE
& '.\States\Get-PWDocumentStateReport.ps1' -FolderPath 'Projects\11431' -IncludeDocumentList
Reports document states in one project, with a document list.

.INPUTS
None.

.OUTPUTS
An Excel workbook and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab and ImportExcel. Uses the active ProjectWise session or prompts for a
Bentley IMS login, and logs out only if this script opened the session. Only the active
version of each document is counted. Documents not in a workflow show '(none)'.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $FolderPath,

    [switch] $IncludeDocumentList,

    [string] $OutputPath = (Join-Path $PSScriptRoot ('DocumentStates_{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd_HHmmss')))
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

        $documents = @(Get-PWDocumentsBySearch -FolderPath $FolderPath -ErrorAction Stop | ForEach-Object {
                [pscustomobject] @{
                    FolderPath   = $_.FolderPath
                    DocumentName = $_.Name
                    FileName     = $_.FileName
                    Workflow     = if ($_.Workflow) { $_.Workflow } else { '(none)' }
                    State        = if ($_.WorkflowState) { $_.WorkflowState } else { '(none)' }
                    StateChanged = $_.StatusChangeDate
                }
            })
        if ($documents.Count -eq 0) {
            Write-Warning "No documents were found below '${FolderPath}'."
            return
        }

        $documents | Group-Object Workflow, State | ForEach-Object {
            [pscustomobject] @{ Workflow = $_.Group[0].Workflow; State = $_.Group[0].State; Documents = $_.Count }
        } | Sort-Object Workflow, State |
            Export-Excel -Path $OutputPath -WorksheetName 'ByState' -AutoSize -FreezeTopRow -BoldTopRow

        $documents | Group-Object FolderPath, Workflow, State | ForEach-Object {
            [pscustomobject] @{ FolderPath = $_.Group[0].FolderPath; Workflow = $_.Group[0].Workflow; State = $_.Group[0].State; Documents = $_.Count }
        } | Sort-Object FolderPath, Workflow, State |
            Export-Excel -Path $OutputPath -WorksheetName 'ByFolder' -AutoSize -FreezeTopRow -BoldTopRow

        if ($IncludeDocumentList) {
            $documents | Sort-Object FolderPath, DocumentName |
                Export-Excel -Path $OutputPath -WorksheetName 'Documents' -AutoSize -FreezeTopRow -BoldTopRow
        }
        Write-Host "Counted $($documents.Count) document(s). Report saved to '${OutputPath}'."
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
