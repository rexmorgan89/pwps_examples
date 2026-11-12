#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Disable-PWUsersFromCsv
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Disables the ProjectWise user accounts listed in a CSV.
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
##  Reads UserName values from a CSV, finds each exact user, skips accounts already disabled and disables the rest.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWUsersByMatch, Set-PWUserDisabled, Undo-PWLogin.
##  CHANGES user accounts; disabled users cannot log in. Supports -WhatIf; writes a results CSV beside the script.
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
Disables ProjectWise users listed in a CSV.

.DESCRIPTION
Reads a CSV with a UserName column (for example from Reports\Get-PWInactiveUsersReport.ps1),
finds each user by exact name and disables the account. Users that are already disabled
are skipped, and names that do not match exactly one user are reported as failed.
The account running the script is never disabled.

Disabling keeps the user, their history and their ownership; it only stops them logging in.

.PARAMETER CsvPath
CSV file with a UserName column.

.PARAMETER ResultPath
CSV file for the per-user results. Defaults to a timestamped file beside the script.

.EXAMPLE
& '.\User\Disable-PWUsersFromCsv.ps1' -CsvPath 'C:\Temp\UsersToDisable.csv' -WhatIf
Shows which users would be disabled.

.INPUTS
None.

.OUTPUTS
A results CSV and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab. Uses the active ProjectWise session or prompts for a Bentley IMS
login, and logs out only if this script opened the session. Re-enable a user with
Set-PWUserEnabled if needed.
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string] $CsvPath,

    [string] $ResultPath = (Join-Path $PSScriptRoot ('DisableUsersResults_{0}.csv' -f (Get-Date -Format 'yyyyMMdd_HHmmss')))
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

        $names = @(Import-Csv -LiteralPath $CsvPath | ForEach-Object { ([string] $_.UserName).Trim() } |
                Where-Object { $_ } | Sort-Object -Unique)
        if ($names.Count -eq 0) { throw "CSV '${CsvPath}' has no UserName values." }

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

        [string] $currentUser = [string] (Get-PWCurrentUser -ErrorAction SilentlyContinue).UserName

        $results = foreach ($name in $names) {
            try {
                $found = @(Get-PWUsersByMatch -UserName $name -ErrorAction Stop | Where-Object { $_.UserName -eq $name })
                if ($found.Count -ne 1) { throw "Expected 1 user named '${name}', found $($found.Count)." }
                $user = $found[0]

                if ($name -eq $currentUser) {
                    [pscustomobject] @{ UserName = $name; Result = 'Skipped'; Message = 'Account running this script' }
                }
                elseif ($user.IsDisabled) {
                    [pscustomobject] @{ UserName = $name; Result = 'Skipped'; Message = 'Already disabled' }
                }
                elseif (-not $PSCmdlet.ShouldProcess("${name} ($($user.Email))", 'Disable user')) {
                    [pscustomobject] @{ UserName = $name; Result = 'Skipped'; Message = 'WhatIf or declined' }
                }
                else {
                    $null = Set-PWUserDisabled -InputUser $user -ErrorAction Stop
                    [pscustomobject] @{ UserName = $name; Result = 'Disabled'; Message = '' }
                }
            }
            catch {
                [pscustomobject] @{ UserName = $name; Result = 'Failed'; Message = $_.Exception.Message }
            }
        }

        $results = @($results)
        $results | Export-Csv -LiteralPath $ResultPath -NoTypeInformation -Encoding UTF8 -WhatIf:$false
        $results | Group-Object Result | ForEach-Object { Write-Host "$($_.Name): $($_.Count)" }
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
