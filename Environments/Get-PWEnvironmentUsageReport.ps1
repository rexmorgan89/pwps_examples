#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Get-PWEnvironmentUsageReport
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Counts the folders that use each environment and lists environments that no folder uses.
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
###################################################################################################     ###       #######| |#  pwps_dab, ImportExcel and an administrator login.
##################################################################################################              #########| |#
#######################################################################################################      ############| |#
#########################################################################################################################| |#
#####################################################################################################################################################################################################################
##
##  Description/Notes:
##  Reads all folders (or those below a path), groups them by environment and compares the result with the list of environments.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWEnvironments, Get-PWFolders, Undo-PWLogin.
##  Read-only; does not change environments or folders.
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
Reports how many folders use each environment.

.DESCRIPTION
Writes an Excel workbook with:
- Usage: each environment with its folder count and an example folder path.
- Folders: every folder and its environment (with -IncludeFolderList).
Environments with a count of zero are candidates for clean-up. When -FolderPath is set,
"unused" only means unused below that path.

.PARAMETER FolderPath
Optional folder path to limit the report. Reads every folder when empty.

.PARAMETER IncludeFolderList
Adds a Folders worksheet listing every folder with its environment.

.PARAMETER OutputPath
Full path of the Excel workbook to create. Defaults to a timestamped file beside the script.

.EXAMPLE
& '.\Environments\Get-PWEnvironmentUsageReport.ps1'
Reports environment use across the whole datasource.

.INPUTS
None.

.OUTPUTS
An Excel workbook and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab and ImportExcel. Uses the active ProjectWise session or prompts for a
Bentley IMS login, and logs out only if this script opened the session. Reading every
folder in a large datasource can take a long time.
#>

[CmdletBinding()]
param(
    [string] $FolderPath,

    [switch] $IncludeFolderList,

    [string] $OutputPath = (Join-Path $PSScriptRoot ('EnvironmentUsage_{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd_HHmmss')))
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

        $environmentNames = @(Get-PWEnvironments -ErrorAction Stop | ForEach-Object { [string] $_.Name })
        $folderParameters = @{ ErrorAction = 'Stop' }
        if (-not [string]::IsNullOrWhiteSpace($FolderPath)) { $folderParameters['FolderPath'] = $FolderPath }
        $folders = @(Get-PWFolders @folderParameters)
        Write-Host "Read $($folders.Count) folder(s) and $($environmentNames.Count) environment(s)."

        $byEnvironment = @{}
        foreach ($group in ($folders | Group-Object -Property { [string] $_.Environment })) { $byEnvironment[$group.Name] = $group.Group }

        $usage = foreach ($name in $environmentNames) {
            $used = if ($byEnvironment.ContainsKey($name)) { @($byEnvironment[$name]) } else { @() }
            [pscustomobject] @{
                Environment   = $name
                FolderCount   = $used.Count
                ExampleFolder = if ($used.Count -gt 0) { $used[0].FullPath } else { '' }
            }
        }
        if ($byEnvironment.ContainsKey('')) {
            $noEnvironment = @($byEnvironment[''])
            $usage = @($usage) + [pscustomobject] @{ Environment = '(none)'; FolderCount = $noEnvironment.Count; ExampleFolder = $noEnvironment[0].FullPath }
        }

        $usage | Sort-Object -Property FolderCount -Descending |
            Export-Excel -Path $OutputPath -WorksheetName 'Usage' -AutoSize -FreezeTopRow -BoldTopRow
        if ($IncludeFolderList) {
            $folders | Sort-Object FullPath | Select-Object FullPath, Environment, IsRichProject |
                Export-Excel -Path $OutputPath -WorksheetName 'Folders' -AutoSize -FreezeTopRow -BoldTopRow
        }
        $unused = @($usage | Where-Object { $_.FolderCount -eq 0 })
        Write-Host "$($unused.Count) environment(s) are not used by any folder read."
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
