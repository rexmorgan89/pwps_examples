#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Export-PWSavedSearchResults
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Runs a saved search and exports the documents it returns, with optional attributes, to Excel.
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
###################################################################################################     ###       #######| |#  pwps_dab and ImportExcel.
##################################################################################################              #########| |#
#######################################################################################################      ############| |#
#########################################################################################################################| |#
#####################################################################################################################################################################################################################
##
##  Description/Notes:
##  Runs a global, work area or personal saved search and writes one row per returned document to Excel.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWDocumentsBySearch, Undo-PWLogin.
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
Runs a ProjectWise saved search and exports the results to Excel.

.DESCRIPTION
Runs the saved search named -SearchName and writes the documents it returns to an Excel
workbook. With -IncludeAttributes, each environment attribute becomes its own column.
Useful for scheduled document registers or transmittal checks without opening
ProjectWise Explorer.

.PARAMETER SearchName
Name of the saved search to run.

.PARAMETER WorkAreaPath
Folder path of the work area that owns the search. Leave empty for global searches.

.PARAMETER Personal
Runs the current user's personal saved search with this name.

.PARAMETER IncludeAttributes
Adds environment attribute values as extra columns.

.PARAMETER OutputPath
Full path of the Excel workbook to create. Defaults to a timestamped file beside the script.

.EXAMPLE
& '.\SavedSearches\Export-PWSavedSearchResults.ps1' -SearchName 'Issued for Construction'
Exports the results of a global saved search.

.EXAMPLE
& '.\SavedSearches\Export-PWSavedSearchResults.ps1' -SearchName 'My Drawings' -Personal -IncludeAttributes
Exports a personal saved search, including attribute columns.

.INPUTS
None.

.OUTPUTS
An Excel workbook and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab and ImportExcel. Uses the active ProjectWise session or prompts for a
Bentley IMS login, and logs out only if this script opened the session. Results reflect
the access rights of the user running the script.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $SearchName,

    [string] $WorkAreaPath,

    [switch] $Personal,

    [switch] $IncludeAttributes,

    [string] $OutputPath = (Join-Path $PSScriptRoot ('SavedSearchResults_{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd_HHmmss')))
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

        $searchParameters = @{
            SearchName    = $SearchName
            Personal      = $Personal
            GetAttributes = $IncludeAttributes
            ErrorAction   = 'Stop'
        }
        if (-not [string]::IsNullOrWhiteSpace($WorkAreaPath)) { $searchParameters['PathOfSearchProject'] = $WorkAreaPath }
        $documents = @(Get-PWDocumentsBySearch @searchParameters)

        if ($documents.Count -eq 0) {
            Write-Warning "Saved search '${SearchName}' returned no documents, or was not found."
            return
        }

        $rows = foreach ($document in $documents) {
            $row = [ordered] @{
                Name          = $document.Name
                FileName      = $document.FileName
                Description   = $document.Description
                Version       = $document.Version
                FolderPath    = $document.FolderPath
                Workflow      = $document.Workflow
                WorkflowState = $document.WorkflowState
                OutTo         = $document.DocumentOutToName
                UpdateDate    = $document.DocumentUpdateDate
                FileSize      = $document.FileSize
                WebLink       = $document.ProjectWiseWebLink
            }
            if ($IncludeAttributes -and $document.Attributes -is [System.Collections.IDictionary]) {
                foreach ($key in $document.Attributes.Keys) {
                    if (-not $row.Contains([string] $key)) { $row[[string] $key] = $document.Attributes[$key] }
                }
            }
            [pscustomobject] $row
        }

        # Attribute columns can differ per environment; collect them all so Excel gets every column.
        $columns = @($rows | ForEach-Object { $_.PSObject.Properties.Name } | Select-Object -Unique)
        $rows | Select-Object -Property $columns |
            Export-Excel -Path $OutputPath -WorksheetName 'Results' -AutoSize -FreezeTopRow -BoldTopRow
        Write-Host "Exported $($documents.Count) document(s) to '${OutputPath}'."
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
