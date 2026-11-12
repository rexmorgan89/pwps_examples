#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Export-PWDocumentAttributes
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Exports document properties and environment attribute values below a folder to Excel.
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
##  Writes one row per document with its GUID, path, file name and every environment attribute value, ready for review or bulk editing.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWDocumentsBySearch, Undo-PWLogin.
##  Read-only. Pair with Update-PWDocumentAttributesFromExcel.ps1 to write edited values back.
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
Exports document attributes below a ProjectWise folder to Excel.

.DESCRIPTION
Searches the documents in -FolderPath (and its subfolders unless -ThisFolderOnly is set)
and writes their properties and environment attribute values to an Excel workbook.
The DocumentGUID column identifies each document, so the workbook can be edited and
passed to Update-PWDocumentAttributesFromExcel.ps1.

.PARAMETER FolderPath
ProjectWise folder path, without the leading 'Documents' folder.

.PARAMETER ThisFolderOnly
Exports only documents directly in -FolderPath.

.PARAMETER OutputPath
Full path of the Excel workbook to create. Defaults to a timestamped file beside the script.

.EXAMPLE
& '.\EnvironmentAttributes\Export-PWDocumentAttributes.ps1' -FolderPath 'Projects\11431\03_Design' -ThisFolderOnly
Exports attributes of the documents in one folder.

.INPUTS
None.

.OUTPUTS
An Excel workbook and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab and ImportExcel. Uses the active ProjectWise session or prompts for a
Bentley IMS login, and logs out only if this script opened the session. When folders use
different environments, columns that do not apply to a document are left blank.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $FolderPath,

    [switch] $ThisFolderOnly,

    [string] $OutputPath = (Join-Path $PSScriptRoot ('DocumentAttributes_{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd_HHmmss')))
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

        $documents = @(Get-PWDocumentsBySearch -FolderPath $FolderPath -GetAttributes -JustThisFolder:$ThisFolderOnly -ErrorAction Stop)
        if ($documents.Count -eq 0) {
            Write-Warning "No documents were found in '${FolderPath}'."
            return
        }

        $rows = foreach ($document in $documents) {
            $row = [ordered] @{
                DocumentGUID = $document.DocumentGUIDString
                FolderPath   = $document.FolderPath
                Name         = $document.Name
                FileName     = $document.FileName
                Description  = $document.Description
                Version      = $document.Version
            }
            if ($document.Attributes -is [System.Collections.IDictionary]) {
                foreach ($key in $document.Attributes.Keys) {
                    if (-not $row.Contains([string] $key)) { $row[[string] $key] = $document.Attributes[$key] }
                }
            }
            [pscustomobject] $row
        }

        $columns = @($rows | ForEach-Object { $_.PSObject.Properties.Name } | Select-Object -Unique)
        $rows | Select-Object -Property $columns |
            Export-Excel -Path $OutputPath -WorksheetName 'Attributes' -AutoSize -FreezeTopRow -BoldTopRow
        Write-Host "Exported $($documents.Count) document(s) with $($columns.Count - 6) attribute column(s) to '${OutputPath}'."
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
