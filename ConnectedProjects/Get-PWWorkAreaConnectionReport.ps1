#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Get-PWWorkAreaConnectionReport
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Lists work areas and whether each one is connected to a cloud project (iTwin).
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
##  Lists every work area below a path with its type and connected project ID, and summarizes connected versus not connected.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWRichProjects, Undo-PWLogin.
##  Read-only; does not connect or disconnect work areas.
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
Reports which work areas are connected to a cloud project.

.DESCRIPTION
Lists every work area (rich project) below -FolderPath, or in the whole datasource,
with its work area type and connected project (iTwin) ID. A Summary worksheet counts
connected and not connected work areas. Use it to find work areas that still need
to be connected, for example before enabling cloud services.

.PARAMETER FolderPath
Optional folder path to limit the report. Reports the whole datasource when empty.

.PARAMETER OutputPath
Full path of the Excel workbook to create. Defaults to a timestamped file beside the script.

.EXAMPLE
& '.\ConnectedProjects\Get-PWWorkAreaConnectionReport.ps1' -FolderPath 'Projects'
Reports the connection status of every work area below 'Projects'.

.INPUTS
None.

.OUTPUTS
An Excel workbook and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab and ImportExcel. Uses the active ProjectWise session or prompts for a
Bentley IMS login, and logs out only if this script opened the session. The report uses
the connected project ID stored in ProjectWise; it does not check that the cloud project
still exists.
#>

[CmdletBinding()]
param(
    [string] $FolderPath,

    [string] $OutputPath = (Join-Path $PSScriptRoot ('WorkAreaConnections_{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd_HHmmss')))
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

        $searchParameters = @{ PopulatePaths = $true; ErrorAction = 'Stop' }
        if (-not [string]::IsNullOrWhiteSpace($FolderPath)) { $searchParameters['FolderPath'] = $FolderPath }
        $workAreas = @(Get-PWRichProjects @searchParameters)

        if ($workAreas.Count -eq 0) {
            Write-Warning 'No work areas were found.'
            return
        }

        $rows = foreach ($workArea in $workAreas) {
            [string] $connectedId = $workArea.ConnectedProjectId
            [pscustomobject] @{
                FullPath           = $workArea.FullPath
                Name               = $workArea.Name
                Description        = $workArea.Description
                WorkAreaType       = $workArea.ProjectType
                IsConnected        = -not [string]::IsNullOrWhiteSpace($connectedId) -and $connectedId -ne [guid]::Empty.ToString()
                ConnectedProjectId = $connectedId
                Storage            = $workArea.Storage
                Created            = $workArea.CreateDateTime
            }
        }

        $rows | Sort-Object -Property FullPath |
            Export-Excel -Path $OutputPath -WorksheetName 'WorkAreas' -AutoSize -FreezeTopRow -BoldTopRow
        $connected = @($rows | Where-Object IsConnected).Count
        @(
            [pscustomobject] @{ Status = 'Connected'; WorkAreas = $connected }
            [pscustomobject] @{ Status = 'Not connected'; WorkAreas = $rows.Count - $connected }
        ) | Export-Excel -Path $OutputPath -WorksheetName 'Summary' -AutoSize -BoldTopRow
        Write-Host "Work areas: $($rows.Count); connected: ${connected}. Report saved to '${OutputPath}'."
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
