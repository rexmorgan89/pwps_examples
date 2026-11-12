#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Export-PWGroupAndUserListMembership
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Exports every group and user list with its members to Excel.
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
##  Lists the members of each group and user list (optionally filtered by name) and a summary of member counts, including empty ones.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWGroupNames, Get-PWUserListNames, Get-PWMembers, Undo-PWLogin.
##  Read-only; does not change membership.
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
Exports group and user list membership to Excel.

.DESCRIPTION
Writes an Excel workbook with:
- Summary: each group and user list with its member count (empty ones included).
- Groups: one row per group member.
- UserLists: one row per user list member.
Use it for access reviews, audits and before restructuring groups.

.PARAMETER NameLike
Optional wildcard filter on group and user list names, for example 'PRJ_*'.

.PARAMETER OutputPath
Full path of the Excel workbook to create. Defaults to a timestamped file beside the script.

.EXAMPLE
& '.\Groups&UserLists\Export-PWGroupAndUserListMembership.ps1'
Exports all groups and user lists.

.EXAMPLE
& '.\Groups&UserLists\Export-PWGroupAndUserListMembership.ps1' -NameLike 'PRJ_11431*'
Exports only groups and user lists for one project.

.INPUTS
None.

.OUTPUTS
An Excel workbook and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab and ImportExcel. Uses the active ProjectWise session or prompts for a
Bentley IMS login, and logs out only if this script opened the session. User lists can
contain groups and other user lists; members are listed as returned, not expanded.
#>

[CmdletBinding()]
param(
    [string] $NameLike = '*',

    [string] $OutputPath = (Join-Path $PSScriptRoot ('GroupAndUserListMembership_{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd_HHmmss')))
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

        $summary = [System.Collections.Generic.List[object]]::new()
        foreach ($set in @(
                @{ Type = 'Group'; Sheet = 'Groups'; Names = @(Get-PWGroupNames -ErrorAction Stop) }
                @{ Type = 'UserList'; Sheet = 'UserLists'; Names = @(Get-PWUserListNames -ErrorAction Stop) }
            )) {
            $names = @($set.Names | ForEach-Object { [string] $_ } | Where-Object { $_ -like $NameLike } | Sort-Object)
            $members = foreach ($name in $names) {
                $found = @(Get-PWMembers -Type $set.Type -Name $name -ErrorAction Continue)
                $summary.Add([pscustomobject] @{ Type = $set.Type; Name = $name; Members = $found.Count })
                $found | Select-Object -Property @{ Name = $set.Type; Expression = { $name } }, *
            }
            $members = @($members)
            if ($members.Count -gt 0) {
                $columns = @($members | ForEach-Object { $_.PSObject.Properties.Name } | Select-Object -Unique)
                $members | Select-Object -Property $columns |
                    Export-Excel -Path $OutputPath -WorksheetName $set.Sheet -AutoSize -FreezeTopRow -BoldTopRow
            }
            Write-Host "$($set.Type): $($names.Count) checked, $($members.Count) membership row(s)."
        }

        if ($summary.Count -eq 0) {
            Write-Warning "No groups or user lists match '${NameLike}'."
            return
        }
        $summary | Export-Excel -Path $OutputPath -WorksheetName 'Summary' -AutoSize -FreezeTopRow -BoldTopRow -MoveToStart
        $empty = @($summary | Where-Object { $_.Members -eq 0 }).Count
        Write-Host "${empty} group(s) or user list(s) have no members. Report saved to '${OutputPath}'."
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
