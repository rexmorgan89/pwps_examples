#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Get-PWWorkAreaWorkspaceReport
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Lists the Managed Workspace CSBs that apply to each work area and flags work areas with none.
####    ###########    #######         ######    #         ####         #    ######          ######    ########    ######| |#  Author                   : GitHub Copilot
####    ##########-    ####               ###                ##         #    ####              ####    ########    ######| |#  Creation Date            : 10-27-2026
#                     #####     ######    ###      #####     ###    #####    ###     ######     ###    ########    ######| |#  Modified By              : GitHub Copilot
#                    #####    #########    ##    ########     ##    #####    ##     ########    ###    ########    ######| |#  Modified Date            : 10-27-2026
####    ##########      ##                 ##    #########    ##    #####    ##                  ##    ########    ######| |#  Script Version           : 1.0
####    ############    ##                 ##    #########    ##    #####    ##                  ##    ########    ######| |#  PowerShell Version       : Windows PowerShell 5.1 x64
####    ############    ##    ###############    #########    ##    #####    ##     ###############    #######     ######| |#  ProjectWise Version      : Datasource with Managed Workspace configured
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
##  For each work area below a path, lists the assigned and inherited CSBs by level, so missing or unexpected workspaces stand out.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWRichProjects, Get-PWManagedWorkspaces, Undo-PWLogin.
##  Read-only; does not change workspace assignments.
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
Reports the Managed Workspace CSBs of every work area.

.DESCRIPTION
Writes an Excel workbook with:
- WorkAreas: one row per work area with its CSB count and CSB names per level.
- Blocks: one row per work area and CSB, including where the CSB is assigned.
Work areas with no CSBs are listed first, because they usually open with the wrong or a
default workspace.

.PARAMETER FolderPath
Optional folder path to limit the report. Reports every work area when empty.

.PARAMETER OutputPath
Full path of the Excel workbook to create. Defaults to a timestamped file beside the script.

.EXAMPLE
& '.\ManagedWorkspace\Get-PWWorkAreaWorkspaceReport.ps1' -FolderPath 'Projects'
Reports the CSBs of every work area below 'Projects'.

.INPUTS
None.

.OUTPUTS
An Excel workbook and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab and ImportExcel. Uses the active ProjectWise session or prompts for a
Bentley IMS login, and logs out only if this script opened the session. Datasource-level
CSBs apply to every work area and are included in each row.
#>

[CmdletBinding()]
param(
    [string] $FolderPath,

    [string] $OutputPath = (Join-Path $PSScriptRoot ('WorkAreaWorkspaces_{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd_HHmmss')))
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

        $blocks = [System.Collections.Generic.List[object]]::new()
        $summary = foreach ($workArea in $workAreas) {
            $assigned = @(Get-PWManagedWorkspaces -InputFolders $workArea -IncludeInherited -ErrorAction Continue)
            foreach ($block in $assigned) {
                $blocks.Add([pscustomobject] @{
                        WorkArea   = $workArea.FullPath
                        Level      = $block.Level
                        CSB        = $block.Name
                        AssignedAt = $block.FolderPath
                    })
            }
            $row = [ordered] @{ WorkArea = $workArea.FullPath; CSBCount = $assigned.Count }
            foreach ($level in ($assigned | Group-Object -Property Level)) {
                $row[[string] $level.Name] = ($level.Group.Name | Sort-Object) -join '; '
            }
            [pscustomobject] $row
        }
        $summary = @($summary)

        $columns = @($summary | ForEach-Object { $_.PSObject.Properties.Name } | Select-Object -Unique)
        $summary | Sort-Object CSBCount, WorkArea | Select-Object -Property $columns |
            Export-Excel -Path $OutputPath -WorksheetName 'WorkAreas' -AutoSize -FreezeTopRow -BoldTopRow
        if ($blocks.Count -gt 0) {
            $blocks | Export-Excel -Path $OutputPath -WorksheetName 'Blocks' -AutoSize -FreezeTopRow -BoldTopRow
        }
        $none = @($summary | Where-Object { $_.CSBCount -eq 0 }).Count
        Write-Host "Work areas: $($summary.Count); without any CSB: ${none}. Report saved to '${OutputPath}'."
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
