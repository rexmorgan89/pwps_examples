#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Get-PWMonthlyLoginReport
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Exports the users who logged in to the datasource during a month or year.
####    ###########    #######         ######    #         ####         #    ######          ######    ########    ######| |#  Author                   : GitHub Copilot
####    ##########-    ####               ###                ##         #    ####              ####    ########    ######| |#  Creation Date            : 10-27-2026
#                     #####     ######    ###      #####     ###    #####    ###     ######     ###    ########    ######| |#  Modified By              : GitHub Copilot
#                    #####    #########    ##    ########     ##    #####    ##     ########    ###    ########    ######| |#  Modified Date            : 10-27-2026
####    ##########      ##                 ##    #########    ##    #####    ##                  ##    ########    ######| |#  Script Version           : 1.0
####    ############    ##                 ##    #########    ##    #####    ##                  ##    ########    ######| |#  PowerShell Version       : Windows PowerShell 5.1 x64
####    ############    ##    ###############    #########    ##    #####    ##     ###############    #######     ######| |#  ProjectWise Version      : Datasource with login auditing enabled
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
##  Wraps Get-PWMonthlyLogins to export login activity for the previous month by default, a chosen month, or a whole year.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWMonthlyLogins, Undo-PWLogin.
##  Read-only; does not change ProjectWise data.
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
Exports ProjectWise login activity for a month or a year.

.DESCRIPTION
Runs Get-PWMonthlyLogins for the requested period and saves the result to an Excel
workbook. Defaults to the previous calendar month, which suits a monthly scheduled task
for licence and usage reviews.

.PARAMETER Month
Month number (1-12). Defaults to the previous month.

.PARAMETER Year
Four-digit year. Defaults to the year of the previous month.

.PARAMETER WholeYear
Reports every month of -Year instead of a single month.

.PARAMETER OutputPath
Full path of the Excel workbook to create. Defaults to a file named for the period beside the script.

.EXAMPLE
& '.\Licensing\Get-PWMonthlyLoginReport.ps1'
Reports logins for the previous month.

.EXAMPLE
& '.\Licensing\Get-PWMonthlyLoginReport.ps1' -Year 2026 -WholeYear
Reports logins for every month of 2026.

.INPUTS
None.

.OUTPUTS
An Excel workbook, the objects returned by Get-PWMonthlyLogins, and a timestamped
.LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab. Uses the active ProjectWise session or prompts for a Bentley IMS
login, and logs out only if this script opened the session. Results depend on login
records still held in the audit trail; archived records are not counted.
#>

[CmdletBinding()]
param(
    [ValidateRange(1, 12)]
    [int] $Month = (Get-Date).AddMonths(-1).Month,

    [ValidateRange(2000, 2100)]
    [int] $Year = (Get-Date).AddMonths(-1).Year,

    [switch] $WholeYear,

    [string] $OutputPath
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
    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        [string] $period = if ($WholeYear) { "$Year" } else { '{0}-{1:00}' -f $Year, $Month }
        $OutputPath = Join-Path $PSScriptRoot "MonthlyLogins_${period}.xlsx"
    }
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

        Get-PWMonthlyLogins -Month $Month -Year $Year -WholeYear:$WholeYear -OutputFileName $OutputPath -ErrorAction Stop

        if (Test-Path -LiteralPath $OutputPath) {
            Write-Host "Login report saved to '${OutputPath}'."
        }
        else {
            Write-Warning 'No workbook was created. There may be no login records for this period.'
        }
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
