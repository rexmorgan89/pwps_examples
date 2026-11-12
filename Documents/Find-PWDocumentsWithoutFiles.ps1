#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Find-PWDocumentsWithoutFiles
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Reports abstract documents and documents with empty files below a ProjectWise folder.
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
##  Finds abstract documents (no file attached) and documents whose file size is zero, and exports them to Excel.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWDocumentsBySearch, Undo-PWLogin.
##  Read-only; does not delete or change documents.
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
Reports documents that have no file or an empty file.

.DESCRIPTION
Searches every document below -FolderPath and reports:
- Abstract documents, which have no file attached.
- Documents with a file name but a file size of zero.
These are often left behind by failed imports or interrupted check-ins.

.PARAMETER FolderPath
ProjectWise folder path to search, without the leading 'Documents' folder.

.PARAMETER ExcludeAbstract
Leaves abstract documents out of the report, for datasources that use them on purpose.

.PARAMETER OutputPath
Full path of the Excel workbook to create. Defaults to a timestamped file beside the script.

.EXAMPLE
& '.\Documents\Find-PWDocumentsWithoutFiles.ps1' -FolderPath 'Projects\11431'
Reports abstract and empty documents in one project.

.INPUTS
None.

.OUTPUTS
An Excel workbook and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab and ImportExcel. Uses the active ProjectWise session or prompts for a
Bentley IMS login, and logs out only if this script opened the session. A zero file size
reflects the ProjectWise record; it does not check the storage area file itself.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $FolderPath,

    [switch] $ExcludeAbstract,

    [string] $OutputPath = (Join-Path $PSScriptRoot ('DocumentsWithoutFiles_{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd_HHmmss')))
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

        $documents = @(Get-PWDocumentsBySearch -FolderPath $FolderPath -ErrorAction Stop)
        Write-Host "Checked $($documents.Count) document(s) below '${FolderPath}'."

        $results = @(foreach ($document in $documents) {
                $reason = $null
                if ($document.IsAbstract -or [string]::IsNullOrWhiteSpace([string] $document.FileName)) {
                    if (-not $ExcludeAbstract) { $reason = 'Abstract (no file)' }
                }
                elseif (($document.FileSize -as [long]) -eq 0) {
                    $reason = 'Empty file'
                }
                if ($null -eq $reason) { continue }

                [pscustomobject] @{
                    Reason       = $reason
                    Name         = $document.Name
                    FileName     = $document.FileName
                    Version      = $document.Version
                    FolderPath   = $document.FolderPath
                    Creator      = $document.DocumentCreatorName
                    CreateDate   = $document.CreateDate
                    Updater      = $document.DocumentUpdaterName
                    UpdateDate   = $document.DocumentUpdateDate
                    DocumentGUID = $document.DocumentGUIDString
                }
            })

        if ($results.Count -eq 0) {
            Write-Host 'No documents without files were found.'
            return
        }

        $results | Sort-Object Reason, FolderPath, Name |
            Export-Excel -Path $OutputPath -WorksheetName 'DocumentsWithoutFiles' -AutoSize -FreezeTopRow -BoldTopRow
        $results | Group-Object Reason | ForEach-Object { Write-Host "$($_.Name): $($_.Count)" }
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
