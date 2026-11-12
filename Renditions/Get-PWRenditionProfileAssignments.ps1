#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Get-PWRenditionProfileAssignments
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Lists rendition profiles and the default profile of each work area.
####    ###########    #######         ######    #         ####         #    ######          ######    ########    ######| |#  Author                   : GitHub Copilot
####    ##########-    ####               ###                ##         #    ####              ####    ########    ######| |#  Creation Date            : 10-27-2026
#                     #####     ######    ###      #####     ###    #####    ###     ######     ###    ########    ######| |#  Modified By              : GitHub Copilot
#                    #####    #########    ##    ########     ##    #####    ##     ########    ###    ########    ######| |#  Modified Date            : 11-12-2026
####    ##########      ##                 ##    #########    ##    #####    ##                  ##    ########    ######| |#  Script Version           : 1.1
####    ############    ##                 ##    #########    ##    #####    ##                  ##    ########    ######| |#  PowerShell Version       : Windows PowerShell 5.1 x64
####    ############    ##    ###############    #########    ##    #####    ##     ###############    #######     ######| |#  ProjectWise Version      : Datasource with rendition profiles configured
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
##  Exports all rendition profiles and the default rendition profile of every work area below a path.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWRenditionProfiles, Get-PWRichProjects, Get-PWFolderRenditionProfile, Undo-PWLogin.
##  Read-only; does not change rendition settings.
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
Reports rendition profiles and which work areas use them.

.DESCRIPTION
Writes an Excel workbook with:
- Profiles: every rendition profile with its ID and description.
- WorkAreas: the default rendition profile of each work area below -FolderPath.
Use it before retiring a profile or to check that new work areas got the right default.

.PARAMETER FolderPath
Optional folder path to limit the work area list. Lists every work area when empty.

.PARAMETER OutputPath
Full path of the Excel workbook to create. Defaults to a timestamped file beside the script.

.EXAMPLE
& '.\Renditions\Get-PWRenditionProfileAssignments.ps1' -FolderPath 'Projects'
Lists profiles and the default profile of every work area below 'Projects'.

.INPUTS
None.

.OUTPUTS
An Excel workbook and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab and ImportExcel. Uses the active ProjectWise session or prompts for a
Bentley IMS login, and logs out only if this script opened the session. Work areas with
no default profile are listed with an empty profile column.
#>

[CmdletBinding()]
param(
    [string] $FolderPath,

    [string] $OutputPath = (Join-Path $PSScriptRoot ('RenditionProfiles_{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd_HHmmss')))
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

        $profiles = @(Get-PWRenditionProfiles -ErrorAction Stop)
        if ($profiles.Count -gt 0) {
            $profiles | Select-Object ID, Name, Description | Sort-Object Name |
                Export-Excel -Path $OutputPath -WorksheetName 'Profiles' -AutoSize -FreezeTopRow -BoldTopRow
        }
        Write-Host "Rendition profiles: $($profiles.Count)"

        $searchParameters = @{ PopulatePaths = $true; ErrorAction = 'Stop' }
        if (-not [string]::IsNullOrWhiteSpace($FolderPath)) { $searchParameters['FolderPath'] = $FolderPath }
        $workAreas = @(Get-PWRichProjects @searchParameters)

        $assignments = foreach ($workArea in $workAreas) {
            $assigned = @(Get-PWFolderRenditionProfile -InputFolder $workArea -ErrorAction Continue)
            [pscustomobject] @{
                WorkArea = $workArea.FullPath
                Profile  = ($assigned | ForEach-Object { if ($_.PSObject.Properties['Name']) { $_.Name } else { [string] $_ } }) -join '; '
            }
        }
        $assignments = @($assignments)
        if ($assignments.Count -gt 0) {
            $assignments | Sort-Object WorkArea |
                Export-Excel -Path $OutputPath -WorksheetName 'WorkAreas' -AutoSize -FreezeTopRow -BoldTopRow
            $assignments | Group-Object Profile | Sort-Object Count -Descending | ForEach-Object {
                Write-Host ('{0,5}  {1}' -f $_.Count, $(if ($_.Name) { $_.Name } else { '(no default profile)' }))
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
