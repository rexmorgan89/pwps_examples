#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Export-PWSavedSearchList
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Exports global, work area and personal saved searches to Excel.
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
##  Lists global saved searches, the saved searches of each work area below a path and optionally the current user's personal searches.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWSavedSearches, Get-PWRichProjects, Undo-PWLogin.
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
Exports the list of ProjectWise saved searches to Excel.

.DESCRIPTION
Writes up to three worksheets:
- Global: saved searches at the datasource level.
- WorkAreas: saved searches defined in each work area below -WorkAreaPath (when set).
- Personal: the current user's personal saved searches (with -IncludePersonal).
Use the list to document searches before a migration or to find duplicates.

.PARAMETER WorkAreaPath
Optional folder path. Saved searches of every work area below this path are listed.

.PARAMETER IncludePersonal
Also lists the personal saved searches of the user running the script.

.PARAMETER OutputPath
Full path of the Excel workbook to create. Defaults to a timestamped file beside the script.

.EXAMPLE
& '.\SavedSearches\Export-PWSavedSearchList.ps1'
Lists global saved searches.

.EXAMPLE
& '.\SavedSearches\Export-PWSavedSearchList.ps1' -WorkAreaPath 'Projects' -IncludePersonal
Lists global, work area and personal saved searches.

.INPUTS
None.

.OUTPUTS
An Excel workbook and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab and ImportExcel. Uses the active ProjectWise session or prompts for a
Bentley IMS login, and logs out only if this script opened the session. Personal searches
of other users are not visible to this script.
#>

[CmdletBinding()]
param(
    [string] $WorkAreaPath,

    [switch] $IncludePersonal,

    [string] $OutputPath
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

        if ([string]::IsNullOrWhiteSpace($OutputPath)) {
            $OutputPath = Join-Path $PSScriptRoot ('SavedSearches_{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
        }
        
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

        [int] $total = 0
        $global = @(Get-PWSavedSearches -ErrorAction Stop)
        if ($global.Count -gt 0) {
            $global | Export-Excel -Path $OutputPath -WorksheetName 'Global' -AutoSize -FreezeTopRow -BoldTopRow
        }
        Write-Host "Global saved searches: $($global.Count)"
        $total += $global.Count

        if (-not [string]::IsNullOrWhiteSpace($WorkAreaPath)) {
            $workAreaSearches = @(foreach ($workArea in @(Get-PWRichProjects -FolderPath $WorkAreaPath -PopulatePaths -ErrorAction Stop)) {
                    foreach ($search in @(Get-PWSavedSearches -InputFolder $workArea -ErrorAction Continue)) {
                        $search | Select-Object -Property @{ Name = 'WorkArea'; Expression = { $workArea.FullPath } }, *
                    }
                })
            if ($workAreaSearches.Count -gt 0) {
                $workAreaSearches | Export-Excel -Path $OutputPath -WorksheetName 'WorkAreas' -AutoSize -FreezeTopRow -BoldTopRow
            }
            Write-Host "Work area saved searches: $($workAreaSearches.Count)"
            $total += $workAreaSearches.Count
        }

        if ($IncludePersonal) {
            $personal = @(Get-PWSavedSearches -Personal -ErrorAction Stop)
            if ($personal.Count -gt 0) {
                $personal | Export-Excel -Path $OutputPath -WorksheetName 'Personal' -AutoSize -FreezeTopRow -BoldTopRow
            }
            Write-Host "Personal saved searches: $($personal.Count)"
            $total += $personal.Count
        }

        if ($total -eq 0) {
            Write-Warning 'No saved searches were found; no workbook was created.'
        }
        else {
            Write-Host "Report saved to '${OutputPath}'."
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
