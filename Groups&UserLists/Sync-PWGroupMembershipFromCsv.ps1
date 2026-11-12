#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Sync-PWGroupMembershipFromCsv
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Adds users to groups from a CSV and can remove users that are not listed.
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
###################################################################################################     ###       #######| |#  pwps_dab and an administrator login.
##################################################################################################              #########| |#
#######################################################################################################      ############| |#
#########################################################################################################################| |#
#####################################################################################################################################################################################################################
##
##  Description/Notes:
##  Compares each group's current users with GroupName,UserName rows from a CSV, adds missing users and, with -RemoveUnlisted, removes extra users.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWUsersInGroup, Add-PWUserToGroup, Remove-PWUserFromGroup, Undo-PWLogin.
##  CHANGES group membership. Supports -WhatIf; writes a results CSV beside the script.
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
Synchronizes ProjectWise group membership with a CSV.

.DESCRIPTION
Reads GroupName and UserName columns from a CSV. For each group in the file:
- Users in the CSV but not in the group are added.
- With -RemoveUnlisted, users in the group but not in the CSV are removed.
Groups that do not appear in the CSV are not touched. Run with -WhatIf first.

.PARAMETER CsvPath
CSV file with GroupName and UserName columns.

.PARAMETER RemoveUnlisted
Removes users from each listed group when they are not in the CSV for that group.

.PARAMETER ResultPath
CSV file for the per-change results. Defaults to a timestamped file beside the script.

.EXAMPLE
& '.\Groups&UserLists\Sync-PWGroupMembershipFromCsv.ps1' -CsvPath 'C:\Temp\Membership.csv' -WhatIf
Shows the users that would be added.

.EXAMPLE
& '.\Groups&UserLists\Sync-PWGroupMembershipFromCsv.ps1' -CsvPath 'C:\Temp\Membership.csv' -RemoveUnlisted
Adds missing users and removes users that are not in the CSV, after confirmation.

.INPUTS
None.

.OUTPUTS
A results CSV and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab. Uses the active ProjectWise session or prompts for a Bentley IMS
login, and logs out only if this script opened the session. User names are compared
without regard to case. Removing an administrator from an administrator group can lock
out administration; review -WhatIf output carefully.
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string] $CsvPath,

    [switch] $RemoveUnlisted,

    [string] $ResultPath = (Join-Path $PSScriptRoot ('GroupSyncResults_{0}.csv' -f (Get-Date -Format 'yyyyMMdd_HHmmss')))
)

begin {
    $InformationPreference = 'Continue'
    # $VerbosePreference = 'Continue' # Uncomment this line if you want to see verbose messages.
    # Log all script output to a file for easy reference later if needed.
    [string] $lastRunLogFilePath = "$PSCommandPath.$(Get-Date -Format 'yyyyMMdd_HHmmss').LastRun.log"
    Start-Transcript -Path $lastRunLogFilePath -ErrorAction Stop -WhatIf:$false

    # Display the time that this script started running.
    [DateTime] $startTime = Get-Date
    Write-Host "Starting script at '$($startTime.ToString('u'))'."

    [bool] $loggedInByScript = $false
}

process {
    try {
        Import-Module pwps_dab -DisableNameChecking -ErrorAction Stop

        $rows = @(Import-Csv -LiteralPath $CsvPath |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_.GroupName) -and -not [string]::IsNullOrWhiteSpace($_.UserName) })
        if ($rows.Count -eq 0) { throw "CSV '${CsvPath}' has no GroupName,UserName rows." }

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

        $results = foreach ($group in ($rows | Group-Object -Property { $_.GroupName.Trim() })) {
            [string] $groupName = $group.Name
            try {
                $current = @(Get-PWUsersInGroup -GroupName $groupName -ErrorAction Stop | ForEach-Object { [string] $_.UserName })
            }
            catch {
                [pscustomobject] @{ GroupName = $groupName; UserName = ''; Action = 'Read'; Result = 'Failed'; Message = $_.Exception.Message }
                continue
            }
            $desired = @($group.Group | ForEach-Object { $_.UserName.Trim() } | Sort-Object -Unique)

            $changes = @($desired | Where-Object { $current -notcontains $_ } | ForEach-Object { @{ User = $_; Action = 'Add' } })
            if ($RemoveUnlisted) {
                $changes += @($current | Where-Object { $desired -notcontains $_ } | ForEach-Object { @{ User = $_; Action = 'Remove' } })
            }

            foreach ($change in $changes) {
                [string] $verb = if ($change.Action -eq 'Add') { "Add to group '${groupName}'" } else { "Remove from group '${groupName}'" }
                if (-not $PSCmdlet.ShouldProcess($change.User, $verb)) {
                    [pscustomobject] @{ GroupName = $groupName; UserName = $change.User; Action = $change.Action; Result = 'Skipped'; Message = 'WhatIf or declined' }
                    continue
                }
                try {
                    if ($change.Action -eq 'Add') {
                        $null = Add-PWUserToGroup -GroupName $groupName -UserName $change.User -ErrorAction Stop
                    }
                    else {
                        $null = Remove-PWUserFromGroup -GroupName $groupName -UserName $change.User -ErrorAction Stop
                    }
                    [pscustomobject] @{ GroupName = $groupName; UserName = $change.User; Action = $change.Action; Result = 'Done'; Message = '' }
                }
                catch {
                    [pscustomobject] @{ GroupName = $groupName; UserName = $change.User; Action = $change.Action; Result = 'Failed'; Message = $_.Exception.Message }
                }
            }
        }

        $results = @($results)
        if ($results.Count -eq 0) {
            Write-Host 'Group membership already matches the CSV.'
            return
        }
        $results | Export-Csv -LiteralPath $ResultPath -NoTypeInformation -Encoding UTF8 -WhatIf:$false
        $results | Group-Object Action, Result | ForEach-Object { Write-Host "$($_.Name): $($_.Count)" }
        Write-Host "Results saved to '${ResultPath}'."
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
