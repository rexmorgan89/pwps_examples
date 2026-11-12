#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Get-PWConnectedUsersReport
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Lists the users currently connected to the datasource and appends the count to a CSV log.
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
##  Lists connected users and appends a timestamped count to a CSV, so repeated runs from a scheduled task show peak usage.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWConnectedUserList, Undo-PWLogin.
##  Read-only; does not disconnect users.
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
Lists connected ProjectWise users and logs the count over time.

.DESCRIPTION
Gets the users currently connected to the datasource and:
- Writes the list to a CSV snapshot (-SnapshotPath).
- Appends one line with the date, time, datasource and count to a history CSV
  (-HistoryPath), so a scheduled task run every 15-30 minutes shows peak use.

.PARAMETER IncludeLastLogin
Adds each user's last login time. This makes the query much slower.

.PARAMETER SnapshotPath
CSV file for the current list of users. Defaults to a timestamped file beside the script.

.PARAMETER HistoryPath
CSV file that collects one count per run. Defaults to ConnectedUsersHistory.csv beside the script.

.EXAMPLE
& '.\Licensing\Get-PWConnectedUsersReport.ps1'
Lists connected users and adds the count to the history file.

.INPUTS
None.

.OUTPUTS
CSV files and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab. Uses the active ProjectWise session or prompts for a Bentley IMS
login, and logs out only if this script opened the session. For unattended runs, see
the Scheduled Tasks folder. The session of the account running this script is counted.
#>

[CmdletBinding()]
param(
    [switch] $IncludeLastLogin,

    [string] $SnapshotPath = (Join-Path $PSScriptRoot ('ConnectedUsers_{0}.csv' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))),

    [string] $HistoryPath = (Join-Path $PSScriptRoot 'ConnectedUsersHistory.csv')
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

        $users = @(Get-PWConnectedUserList -LastLogin:$IncludeLastLogin -ErrorAction Stop)
        if ($users.Count -gt 0) {
            $users | Export-Csv -LiteralPath $SnapshotPath -NoTypeInformation -Encoding UTF8
            Write-Host "Connected user list saved to '${SnapshotPath}'."
        }

        [pscustomobject] @{
            Timestamp      = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            Datasource     = [string] $currentDatasource
            ConnectedUsers = $users.Count
        } | Export-Csv -LiteralPath $HistoryPath -NoTypeInformation -Encoding UTF8 -Append
        Write-Host "Connected users: $($users.Count). Count appended to '${HistoryPath}'."
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
