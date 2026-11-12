#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Get-PWStaleCheckoutsReport
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Lists documents below a folder that have been checked out longer than a set number of days.
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
##  Finds documents that are out to a user and whose check-out date is older than -OlderThanDays, and exports them to Excel.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWDocumentsBySearch, Undo-PWLogin.
##  Read-only; does not free or check in documents.
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
Lists documents that have been checked out for too long.

.DESCRIPTION
Searches every document below -FolderPath and reports those that are out to a user
(checked out or exported) with a check-out date older than -OlderThanDays days.
The report lists the user, the check-out date, the days out and the document path so
that administrators can follow up before freeing documents.

.PARAMETER FolderPath
ProjectWise folder path to search, without the leading 'Documents' folder.

.PARAMETER OlderThanDays
Minimum number of days a document must have been out. Defaults to 14.

.PARAMETER OutputPath
Full path of the Excel workbook to create. Defaults to a timestamped file beside the script.

.EXAMPLE
& '.\Documents\Get-PWStaleCheckoutsReport.ps1' -FolderPath 'Projects' -OlderThanDays 30
Lists documents below 'Projects' that have been out for 30 days or more.

.INPUTS
None.

.OUTPUTS
An Excel workbook and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab and ImportExcel. Uses the active ProjectWise session or prompts for a
Bentley IMS login, and logs out only if this script opened the session. Searching a
large folder tree can take a long time. To free documents after review, use
CheckIn-PWDocumentsOrFree -Free on the documents you have confirmed.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $FolderPath,

    [ValidateRange(0, 3650)]
    [int] $OlderThanDays = 14,

    [string] $OutputPath = (Join-Path $PSScriptRoot ('StaleCheckouts_{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd_HHmmss')))
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

        [DateTime] $cutoff = (Get-Date).AddDays(-$OlderThanDays)
        $documents = @(Get-PWDocumentsBySearch -FolderPath $FolderPath -ErrorAction Stop)
        Write-Host "Checked $($documents.Count) document(s) below '${FolderPath}'."

        $stale = @(foreach ($document in $documents) {
                if ([string]::IsNullOrWhiteSpace([string] $document.DocumentOutToName)) { continue }
                $checkOutDate = $document.DocumentCheckOutDate -as [DateTime]
                if ($null -eq $checkOutDate -or $checkOutDate -gt $cutoff) { continue }

                [pscustomobject] @{
                    OutTo          = $document.DocumentOutToName
                    CheckOutDate   = $checkOutDate
                    DaysOut        = [int] ((Get-Date) - $checkOutDate).TotalDays
                    Status         = $document.DocumentStatus
                    Name           = $document.Name
                    FileName       = $document.FileName
                    Version        = $document.Version
                    FolderPath     = $document.FolderPath
                    WorkflowState  = $document.WorkflowState
                    DocumentGUID   = $document.DocumentGUIDString
                }
            })

        if ($stale.Count -eq 0) {
            Write-Host "No documents have been out for ${OlderThanDays} day(s) or more."
            return
        }

        $stale | Sort-Object -Property DaysOut -Descending |
            Export-Excel -Path $OutputPath -WorksheetName 'StaleCheckouts' -AutoSize -FreezeTopRow -BoldTopRow
        Write-Host "Found $($stale.Count) stale check-out(s). Report saved to '${OutputPath}'."
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
