#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Find-PWFoldersByStorageArea
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Lists the folders below a path that use a given storage area, or a different one from their parent.
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
##  Lists folders below a path with their storage area and flags folders whose storage area differs from their parent.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWFolders, Undo-PWLogin.
##  Read-only; does not move folders or files.
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
Lists folders by storage area and flags mixed storage in a folder tree.

.DESCRIPTION
Reads every folder below -FolderPath and reports its storage area. Each folder whose
storage area differs from its parent folder is flagged, which helps find work areas
that are split across storage areas before a migration.
Use -StorageAreaName to list only the folders that use one storage area.

.PARAMETER FolderPath
ProjectWise folder path to search, without the leading 'Documents' folder.

.PARAMETER StorageAreaName
Optional storage area name. When set, only folders using this storage area are listed.

.PARAMETER OutputPath
Full path of the Excel workbook to create. Defaults to a timestamped file beside the script.

.EXAMPLE
& '.\StorageAreas\Find-PWFoldersByStorageArea.ps1' -FolderPath 'Projects'
Lists every folder below 'Projects' and flags storage area changes.

.EXAMPLE
& '.\StorageAreas\Find-PWFoldersByStorageArea.ps1' -FolderPath 'Projects' -StorageAreaName 'Old Storage'
Lists only folders that still use 'Old Storage'.

.INPUTS
None.

.OUTPUTS
An Excel workbook and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab and ImportExcel. Uses the active ProjectWise session or prompts for a
Bentley IMS login, and logs out only if this script opened the session. Folder storage
does not always match the storage of the documents in it; see
Update-PWDocumentStorageToMatchFolder in the pwps_dab reference to align them.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $FolderPath,

    [string] $StorageAreaName,

    [string] $OutputPath = (Join-Path $PSScriptRoot ('FoldersByStorageArea_{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd_HHmmss')))
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

        $folders = @(Get-PWFolders -FolderPath $FolderPath -PopulatePaths -ErrorAction Stop)
        if ($folders.Count -eq 0) {
            throw "Folder '${FolderPath}' was not found."
        }
        $storageById = @{}
        foreach ($folder in $folders) { $storageById[[int] $folder.ProjectID] = [string] $folder.Storage }

        $results = foreach ($folder in $folders) {
            $parentStorage = $storageById[[int] $folder.ParentID]
            [pscustomobject] @{
                FullPath          = $folder.FullPath
                StorageArea       = $folder.Storage
                ParentStorage     = $parentStorage
                DiffersFromParent = ($null -ne $parentStorage -and $parentStorage -ne [string] $folder.Storage)
                IsWorkArea        = $folder.IsRichProject
                Environment       = $folder.Environment
            }
        }
        if (-not [string]::IsNullOrWhiteSpace($StorageAreaName)) {
            $results = @($results | Where-Object { $_.StorageArea -eq $StorageAreaName })
        }
        $results = @($results)

        if ($results.Count -eq 0) {
            Write-Host 'No matching folders were found.'
            return
        }

        $results | Sort-Object -Property FullPath |
            Export-Excel -Path $OutputPath -WorksheetName 'Folders' -AutoSize -FreezeTopRow -BoldTopRow
        $results | Group-Object -Property StorageArea | Sort-Object -Property Count -Descending |
            Select-Object @{ Name = 'StorageArea'; Expression = { $_.Name } }, @{ Name = 'FolderCount'; Expression = { $_.Count } } |
            Export-Excel -Path $OutputPath -WorksheetName 'Summary' -AutoSize -BoldTopRow
        $mixed = @($results | Where-Object DiffersFromParent)
        Write-Host "Listed $($results.Count) folder(s); $($mixed.Count) use a different storage area from their parent."
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
