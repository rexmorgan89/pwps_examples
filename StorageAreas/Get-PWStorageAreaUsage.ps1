#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Get-PWStorageAreaUsage
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Lists storage areas with their paths, document counts and disk space.
####    ###########    #######         ######    #         ####         #    ######          ######    ########    ######| |#  Author                   : GitHub Copilot
####    ##########-    ####               ###                ##         #    ####              ####    ########    ######| |#  Creation Date            : 10-27-2026
#                     #####     ######    ###      #####     ###    #####    ###     ######     ###    ########    ######| |#  Modified By              : GitHub Copilot
#                    #####    #########    ##    ########     ##    #####    ##     ########    ###    ########    ######| |#  Modified Date            : 10-27-2026
####    ##########      ##                 ##    #########    ##    #####    ##                  ##    ########    ######| |#  Script Version           : 1.0
####    ############    ##                 ##    #########    ##    #####    ##                  ##    ########    ######| |#  PowerShell Version       : Windows PowerShell 5.1 x64
####    ############    ##    ###############    #########    ##    #####    ##     ###############    #######     ######| |#  ProjectWise Version      : Datasource with dms_doc and dms_proj storage columns
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
##  Lists every storage area with its host and path, and counts the documents and folders that use each one.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWStorageAreaList, Get-PWStorageDiskUsage, Select-PWSQLDataTable, Undo-PWLogin.
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
Reports storage areas with document counts and disk usage.

.DESCRIPTION
Lists every storage area in the datasource with its host and path, then counts the
documents (dms_doc) and folders (dms_proj) that point at each storage area. When
-IncludeDiskUsage is used, a second worksheet holds Get-PWStorageDiskUsage results.
Use the report to plan storage migrations or to find unused storage areas.

.PARAMETER IncludeDiskUsage
Adds a DiskUsage worksheet from Get-PWStorageDiskUsage. The storage hosts must be
reachable from this workstation.

.PARAMETER OutputPath
Full path of the Excel workbook to create. Defaults to a timestamped file beside the script.

.EXAMPLE
& '.\StorageAreas\Get-PWStorageAreaUsage.ps1'
Lists storage areas with document and folder counts.

.EXAMPLE
& '.\StorageAreas\Get-PWStorageAreaUsage.ps1' -IncludeDiskUsage
Also reports free disk space for each storage area.

.INPUTS
None.

.OUTPUTS
An Excel workbook and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab and ImportExcel. Uses the active ProjectWise session or prompts for a
Bentley IMS login, and logs out only if this script opened the session. Counts include
every document version. The SQL uses the o_storno column of dms_doc and dms_proj.
#>

[CmdletBinding()]
param(
    [switch] $IncludeDiskUsage,

    [string] $OutputPath = (Join-Path $PSScriptRoot ('StorageAreaUsage_{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd_HHmmss')))
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

        $storageAreas = @(Get-PWStorageAreaList -ErrorAction Stop)
        $documentCounts = @{}
        foreach ($row in (Select-PWSQLDataTable -SQLSelectStatement 'SELECT o_storno, COUNT(*) AS ItemCount FROM dms_doc GROUP BY o_storno' -ErrorAction Stop).Rows) {
            $documentCounts[[int] $row.o_storno] = [int] $row.ItemCount
        }
        $folderCounts = @{}
        foreach ($row in (Select-PWSQLDataTable -SQLSelectStatement 'SELECT o_storno, COUNT(*) AS ItemCount FROM dms_proj GROUP BY o_storno' -ErrorAction Stop).Rows) {
            $folderCounts[[int] $row.o_storno] = [int] $row.ItemCount
        }

        $usage = foreach ($storage in $storageAreas) {
            [int] $id = $storage.ID
            [pscustomobject] @{
                StorageArea   = $storage.Name
                ID            = $id
                Description   = $storage.Description
                Host          = $storage.Host
                Path          = $storage.Path
                DocumentCount = [int] $documentCounts[$id]
                FolderCount   = [int] $folderCounts[$id]
            }
        }

        $usage | Sort-Object -Property DocumentCount -Descending |
            Export-Excel -Path $OutputPath -WorksheetName 'Usage' -AutoSize -FreezeTopRow -BoldTopRow
        $unused = @($usage | Where-Object { $_.DocumentCount -eq 0 -and $_.FolderCount -eq 0 })
        Write-Host "Listed $($storageAreas.Count) storage area(s); $($unused.Count) are not used by any folder or document."

        if ($IncludeDiskUsage) {
            $diskUsage = @(Get-PWStorageDiskUsage -InputStorageAreas $storageAreas -ErrorAction Continue)
            if ($diskUsage.Count -gt 0) {
                $diskUsage | Export-Excel -Path $OutputPath -WorksheetName 'DiskUsage' -AutoSize -FreezeTopRow -BoldTopRow
            }
        }
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
