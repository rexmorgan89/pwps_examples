#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Compare-PWDatasourceUsersAndGroups
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Compares users, groups and user lists between two ProjectWise datasources.
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
###################################################################################################     ###       #######| |#  pwps_dab, ImportExcel and an administrator login to both datasources.
##################################################################################################              #########| |#
#######################################################################################################      ############| |#
#########################################################################################################################| |#
#####################################################################################################################################################################################################################
##
##  Description/Notes:
##  Logs in to a source and a target datasource in turn, reads their users, groups and user lists, and reports the differences.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, Get-PWDatasource, New-PWLogin, Get-PWUsersByMatch, Get-PWGroupNames, Get-PWUserListNames, Undo-PWLogin.
##  Read-only; does not create or change users or groups.
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
Compares users, groups and user lists between two datasources.

.DESCRIPTION
Prompts for a source and a target datasource, logs in to each with Bentley IMS, and
reports items that exist in only one of them. For users found in both, it also flags
differences in email address and disabled state. Use it to check a migration or to keep
a test datasource in step with production.

.PARAMETER OutputPath
Full path of the Excel workbook to create. Defaults to a timestamped file beside the script.

.EXAMPLE
& '.\DatasourceUtilities\Compare-PWDatasourceUsersAndGroups.ps1'
Prompts for both datasources and writes the comparison workbook.

.INPUTS
None.

.OUTPUTS
An Excel workbook and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab and ImportExcel. Refuses to run when a ProjectWise session is already
active, because it logs in and out of two datasources. Names are compared without regard
to case. Group and user list membership is not compared.
#>

[CmdletBinding()]
param(
    [string] $OutputPath = (Join-Path $PSScriptRoot ('DatasourceUserCompare_{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd_HHmmss')))
)

begin {
    Import-Module pwps_dab -DisableNameChecking -ErrorAction Stop
    Import-Module ImportExcel -ErrorAction Stop
    if (-not [string]::IsNullOrWhiteSpace([string] (Get-PWCurrentDatasource -ErrorAction SilentlyContinue))) {
        throw 'A ProjectWise session is already active. Use a fresh PowerShell session or log out before running this script.'
    }

    $InformationPreference = 'Continue'
    # $VerbosePreference = 'Continue' # Uncomment this line if you want to see verbose messages.
    # Log all script output to a file for easy reference later if needed.
    [string] $lastRunLogFilePath = "$PSCommandPath.$(Get-Date -Format 'yyyyMMdd_HHmmss').LastRun.log"
    Start-Transcript -Path $lastRunLogFilePath -ErrorAction Stop

    # Display the time that this script started running.
    [DateTime] $startTime = Get-Date
    Write-Host "Starting script at '$($startTime.ToString('u'))'."

    function Get-DatasourceInventory {
        param([string] $Title)

        $datasource = Get-PWDatasource | Select-Object 'InternalName', 'HostName' |
            Out-GridView -Title $Title -OutputMode Single
        if ($null -eq $datasource) { throw 'A datasource must be selected.' }

        [string] $datasourceName = '{0}:{1}' -f $datasource.HostName, $datasource.InternalName
        $null = New-PWLogin $datasourceName -BentleyIMS -ErrorAction Stop
        try {
            Write-Host "Reading users, groups and user lists from '${datasourceName}'."
            [pscustomobject] @{
                Name      = $datasourceName
                Users     = @(Get-PWUsersByMatch -ErrorAction Stop | Select-Object UserName, Email, IsDisabled, Type)
                Groups    = @(Get-PWGroupNames -ErrorAction Stop | ForEach-Object { [string] $_ })
                UserLists = @(Get-PWUserListNames -ErrorAction Stop | ForEach-Object { [string] $_ })
            }
        }
        finally {
            $null = Undo-PWLogin -ErrorAction SilentlyContinue
        }
    }

    function Compare-NameList {
        param([string] $Type, [string[]] $Source, [string[]] $Target, [string] $SourceName, [string] $TargetName)
        foreach ($difference in @(Compare-Object -ReferenceObject @($Source) -DifferenceObject @($Target))) {
            [pscustomobject] @{
                Type   = $Type
                Name   = $difference.InputObject
                Status = if ($difference.SideIndicator -eq '<=') { "Only in ${SourceName}" } else { "Only in ${TargetName}" }
                Detail = ''
            }
        }
    }
}

process {
    try {
        $source = Get-DatasourceInventory -Title 'Select the SOURCE datasource'
        $target = Get-DatasourceInventory -Title 'Select the TARGET datasource'

        $differences = [System.Collections.Generic.List[object]]::new()
        foreach ($item in @(Compare-NameList 'User' $source.Users.UserName $target.Users.UserName $source.Name $target.Name)) { $differences.Add($item) }
        foreach ($item in @(Compare-NameList 'Group' $source.Groups $target.Groups $source.Name $target.Name)) { $differences.Add($item) }
        foreach ($item in @(Compare-NameList 'UserList' $source.UserLists $target.UserLists $source.Name $target.Name)) { $differences.Add($item) }

        $targetUsers = @{}
        foreach ($user in $target.Users) { $targetUsers[$user.UserName.ToUpperInvariant()] = $user }
        foreach ($user in $source.Users) {
            $match = $targetUsers[$user.UserName.ToUpperInvariant()]
            if ($null -eq $match) { continue }
            $details = @()
            if ([string] $user.Email -ne [string] $match.Email) { $details += "Email: '$($user.Email)' vs '$($match.Email)'" }
            if ([bool] $user.IsDisabled -ne [bool] $match.IsDisabled) { $details += "Disabled: $($user.IsDisabled) vs $($match.IsDisabled)" }
            if ($details.Count -gt 0) {
                $differences.Add([pscustomobject] @{ Type = 'User'; Name = $user.UserName; Status = 'Different'; Detail = $details -join '; ' })
            }
        }

        if ($differences.Count -eq 0) {
            Write-Host 'No differences were found.'
            return
        }
        $differences | Sort-Object Type, Status, Name |
            Export-Excel -Path $OutputPath -WorksheetName 'Differences' -AutoSize -FreezeTopRow -BoldTopRow
        $differences | Group-Object Type, Status | ForEach-Object { Write-Host "$($_.Name): $($_.Count)" }
        Write-Host "Report saved to '${OutputPath}'."
    }
    catch {
        $null = Undo-PWLogin -ErrorAction SilentlyContinue
        $null = Stop-Transcript
        throw
    }
}

end {
    try {
        # Display the time that this script finished running, and how long it took to run.
        [DateTime] $finishTime = Get-Date
        [TimeSpan] $elapsedTime = $finishTime - $startTime
        Write-Host "Finished script at '$($finishTime.ToString('u'))'. Took '$elapsedTime' to run."
    }
    finally {
        Stop-Transcript
    }
}
