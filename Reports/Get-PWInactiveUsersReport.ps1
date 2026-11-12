#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Get-PWInactiveUsersReport
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Lists enabled users who have not logged in for a set number of days or have never logged in.
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
##  Uses Get-PWUserByLastLogin to list users whose last login is older than -Days or who never logged in, for licence and clean-up reviews.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWUserByLastLogin, Undo-PWLogin.
##  Read-only; does not disable users. See User\Disable-PWUsersFromCsv.ps1 for that step.
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
Lists users who have not logged in recently.

.DESCRIPTION
Writes an Excel workbook of users whose last login is more than -Days days ago, or who
have never logged in. Disabled users and logical (service) accounts are left out unless
requested. Users created within the last -Days days are also left out, because they may
not have had a chance to log in yet.

The UserName column can be copied into a CSV for User\Disable-PWUsersFromCsv.ps1.

.PARAMETER Days
Number of days without a login. Defaults to 90.

.PARAMETER IncludeDisabled
Includes users who are already disabled.

.PARAMETER IncludeIdentityInfo
Adds identity and identity provider columns. Slower.

.PARAMETER OutputPath
Full path of the Excel workbook to create. Defaults to a timestamped file beside the script.

.EXAMPLE
& '.\Reports\Get-PWInactiveUsersReport.ps1' -Days 180
Lists enabled users with no login in the last 180 days.

.INPUTS
None.

.OUTPUTS
An Excel workbook and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab and ImportExcel. Uses the active ProjectWise session or prompts for a
Bentley IMS login, and logs out only if this script opened the session.
#>

[CmdletBinding()]
param(
    [ValidateRange(1, 3650)]
    [int] $Days = 90,

    [switch] $IncludeDisabled,

    [switch] $IncludeIdentityInfo,

    [string] $OutputPath = (Join-Path $PSScriptRoot ('InactiveUsers_{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd_HHmmss')))
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

        [DateTime] $cutoff = (Get-Date).AddDays(-$Days)
        $users = @(Get-PWUserByLastLogin -DaysAgo $Days -IdentityInfo:$IncludeIdentityInfo -ErrorAction Stop | Where-Object {
                ($IncludeDisabled -or -not $_.IsDisabled) -and
                $_.Type -ne 'Logical' -and
                ($null -eq ($_.CreationDate -as [DateTime]) -or ($_.CreationDate -as [DateTime]) -lt $cutoff)
            })

        if ($users.Count -eq 0) {
            Write-Host "No users have been inactive for ${Days} day(s)."
            return
        }

        $users | Sort-Object -Property LastLogin | Select-Object -Property @(
            'UserName', 'Name', 'Email', 'Description', 'Type', 'SecProvider'
            @{ Name = 'LastLogin'; Expression = { if ($_.LastLogin) { $_.LastLogin } else { 'Never' } } }
            'CreationDate', 'IsDisabled'
            if ($IncludeIdentityInfo) { 'Identity'; 'IdentityProvider' }
        ) | Export-Excel -Path $OutputPath -WorksheetName 'InactiveUsers' -AutoSize -FreezeTopRow -BoldTopRow
        Write-Host "Found $($users.Count) inactive user(s). Report saved to '${OutputPath}'."
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
