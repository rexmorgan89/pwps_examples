#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Test-PWReadOnlySmokeTest
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Runs common read-only pwps_dab cmdlets and reports pass/fail and timing for each.
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
###################################################################################################     ###       #######| |#  pwps_dab.
##################################################################################################              #########| |#
#######################################################################################################      ############| |#
#########################################################################################################################| |#
#####################################################################################################################################################################################################################
##
##  Description/Notes:
##  Calls a set of read-only cmdlets (environments, workflows, states, storage, groups, folders, documents) and reports counts and timings.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWEnvironments, Get-PWWorkflows, Get-PWStates, Get-PWStorageAreaList, Get-PWGroupNames, Get-PWFolders, Get-PWDocumentsBySearch, Undo-PWLogin.
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
Smoke-tests read-only pwps_dab cmdlets against the current datasource.

.DESCRIPTION
Runs each check, records how many objects it returned and how long it took, and prints
Pass or Fail. A failure does not stop later checks. Folder and document checks run only
when -SampleFolderPath is given.

Run it after a pwps_dab or ProjectWise upgrade, or to compare response times between
servers. No data is changed.

.PARAMETER SampleFolderPath
Optional folder used for the folder and document checks.

.EXAMPLE
& '.\Tests\Test-PWReadOnlySmokeTest.ps1' -SampleFolderPath 'Projects\11431'
Runs every check, including folder and document reads in one project.

.INPUTS
None.

.OUTPUTS
One result object per check and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab. Uses the active ProjectWise session or prompts for a Bentley IMS
login, and logs out only if this script opened the session. Some checks need
administrator rights and fail for normal users.
#>

[CmdletBinding()]
param(
    [string] $SampleFolderPath
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

        $checks = [ordered] @{
            'Get-PWEnvironments'    = { Get-PWEnvironments -ErrorAction Stop }
            'Get-PWWorkflows'       = { Get-PWWorkflows -ErrorAction Stop }
            'Get-PWStates'          = { Get-PWStates -ErrorAction Stop }
            'Get-PWStorageAreaList' = { Get-PWStorageAreaList -ErrorAction Stop }
            'Get-PWGroupNames'      = { Get-PWGroupNames -ErrorAction Stop }
        }
        if (-not [string]::IsNullOrWhiteSpace($SampleFolderPath)) {
            $checks['Get-PWFolders (sample)'] = { Get-PWFolders -FolderPath $SampleFolderPath -JustOne -ErrorAction Stop }
            $checks['Get-PWDocumentsBySearch (sample)'] = { Get-PWDocumentsBySearch -FolderPath $SampleFolderPath -JustThisFolder -ErrorAction Stop }
        }

        $results = foreach ($check in $checks.GetEnumerator()) {
            $timer = [System.Diagnostics.Stopwatch]::StartNew()
            try {
                $count = @(& $check.Value).Count
                [pscustomobject] @{ Check = $check.Key; Result = 'Pass'; Objects = $count; Seconds = [math]::Round($timer.Elapsed.TotalSeconds, 2); Message = '' }
            }
            catch {
                [pscustomobject] @{ Check = $check.Key; Result = 'Fail'; Objects = 0; Seconds = [math]::Round($timer.Elapsed.TotalSeconds, 2); Message = $_.Exception.Message }
            }
        }

        $results | Format-Table -AutoSize | Out-Host
        $failed = @($results | Where-Object Result -eq 'Fail').Count
        if ($failed -eq 0) { Write-Host 'All checks passed.' -ForegroundColor Green }
        else { Write-Host "${failed} check(s) failed." -ForegroundColor Red }
        $results
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
