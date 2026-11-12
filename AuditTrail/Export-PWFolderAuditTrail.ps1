#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Export-PWFolderAuditTrail
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Exports folder and document audit trail records below a ProjectWise folder to Excel.
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
###################################################################################################     ###       #######| |#  pwps_dab and read access to the target folder.
##################################################################################################              #########| |#
#######################################################################################################      ############| |#
#########################################################################################################################| |#
#####################################################################################################################################################################################################################
##
##  Description/Notes:
##  Exports the audit trail of a folder, its documents and optionally its subfolders to an Excel workbook.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWFolders, Get-PWFolderAuditTrailRecords, Undo-PWLogin.
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
Exports the audit trail of a ProjectWise folder and its documents to Excel.

.DESCRIPTION
Uses Get-PWFolderAuditTrailRecords to collect audit trail records for the folder and
the documents it contains, and writes them to an Excel workbook. Use -IncludeSubFolders
to include every folder and document below the target folder.

.PARAMETER FolderPath
ProjectWise folder path, without the leading 'Documents' folder.

.PARAMETER IncludeSubFolders
Includes audit trail records for all subfolders and their documents.

.PARAMETER OutputPath
Full path of the Excel workbook to create. Defaults to a timestamped file beside the script.

.EXAMPLE
& '.\AuditTrail\Export-PWFolderAuditTrail.ps1' -FolderPath 'Projects\11431'
Exports audit trail records for the folder and the documents directly in it.

.EXAMPLE
& '.\AuditTrail\Export-PWFolderAuditTrail.ps1' -FolderPath 'Projects\11431' -IncludeSubFolders
Exports audit trail records for the whole folder tree.

.INPUTS
None.

.OUTPUTS
An Excel workbook and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab. Uses the active ProjectWise session or prompts for a Bentley IMS
login, and logs out only if this script opened the session. Large folder trees can
produce very large workbooks; start without -IncludeSubFolders.
References: docs/pwps_dab-Reference.md entry for Get-PWFolderAuditTrailRecords.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $FolderPath,

    [switch] $IncludeSubFolders,

    [string] $OutputPath = (Join-Path $PSScriptRoot ('FolderAuditTrail_{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd_HHmmss')))
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

        $folder = Get-PWFolders -FolderPath $FolderPath -JustOne -ErrorAction Stop
        if ($null -eq $folder) {
            throw "Folder '${FolderPath}' was not found."
        }

        $null = Get-PWFolderAuditTrailRecords -InputFolder $folder -IncludeDocuments `
            -IncludeSubFolders:$IncludeSubFolders -OutputFilePathName $OutputPath -OutputFormat 'Excel' -ErrorAction Stop

        if (Test-Path -LiteralPath $OutputPath) {
            Write-Host "Audit trail exported to '${OutputPath}'."
        }
        else {
            Write-Warning "No workbook was created. The folder may have no audit trail records."
        }
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
