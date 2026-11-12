#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Get-PWEmptyFoldersReport
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Lists folders below a path that contain no documents and no subfolders.
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
##  Lists empty leaf folders (no documents, no subfolders) with their creator and dates, for review before clean-up.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWEmptyLeafFolders, Undo-PWLogin.
##  Read-only; does not delete folders.
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
Lists empty leaf folders below a ProjectWise folder.

.DESCRIPTION
Uses Get-PWEmptyLeafFolders to find folders with no documents and no subfolders, and
writes them to Excel with their creator, creation date and last update date.
Folders created as part of a standard structure are often empty on purpose; use
-OlderThanDays to focus on folders that have stayed empty for a while.

.PARAMETER FolderPath
ProjectWise folder path to search, without the leading 'Documents' folder.

.PARAMETER OlderThanDays
Only lists folders last updated more than this many days ago. Defaults to 0 (all).

.PARAMETER OutputPath
Full path of the Excel workbook to create. Defaults to a timestamped file beside the script.

.EXAMPLE
& '.\Folder\Get-PWEmptyFoldersReport.ps1' -FolderPath 'Projects' -OlderThanDays 365
Lists folders below 'Projects' that have been empty and unchanged for a year.

.INPUTS
None.

.OUTPUTS
An Excel workbook and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab and ImportExcel. Uses the active ProjectWise session or prompts for a
Bentley IMS login, and logs out only if this script opened the session. To remove empty
folders after review, see Remove-PWEmptyFolders in the pwps_dab reference.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $FolderPath,

    [ValidateRange(0, 36500)]
    [int] $OlderThanDays = 0,

    [string] $OutputPath = (Join-Path $PSScriptRoot ('EmptyFolders_{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd_HHmmss')))
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
        $folders = @(Get-PWEmptyLeafFolders -FolderPath $FolderPath -FillPaths -ErrorAction Stop |
                Where-Object { $OlderThanDays -eq 0 -or ($_.UpdateDateTime -as [DateTime]) -lt $cutoff })

        if ($folders.Count -eq 0) {
            Write-Host 'No empty folders matched.'
            return
        }

        $folders | Sort-Object -Property FullPath |
            Select-Object FullPath, Name, Description, FolderCreatorName, CreateDateTime, UpdateDateTime, IsRichProject |
            Export-Excel -Path $OutputPath -WorksheetName 'EmptyFolders' -AutoSize -FreezeTopRow -BoldTopRow
        Write-Host "Found $($folders.Count) empty folder(s). Report saved to '${OutputPath}'."
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
