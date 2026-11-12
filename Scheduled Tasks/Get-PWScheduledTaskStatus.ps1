#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Get-PWScheduledTaskStatus
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Reports the last and next run of the scheduled tasks in the \ProjectWise\ task folder.
####    ###########    #######         ######    #         ####         #    ######          ######    ########    ######| |#  Author                   : GitHub Copilot
####    ##########-    ####               ###                ##         #    ####              ####    ########    ######| |#  Creation Date            : 10-27-2026
#                     #####     ######    ###      #####     ###    #####    ###     ######     ###    ########    ######| |#  Modified By              : GitHub Copilot
#                    #####    #########    ##    ########     ##    #####    ##     ########    ###    ########    ######| |#  Modified Date            : 10-27-2026
####    ##########      ##                 ##    #########    ##    #####    ##                  ##    ########    ######| |#  Script Version           : 1.0
####    ############    ##                 ##    #########    ##    #####    ##                  ##    ########    ######| |#  PowerShell Version       : Windows PowerShell 5.1 x64
####    ############    ##    ###############    #########    ##    #####    ##     ###############    #######     ######| |#  ProjectWise Version      : Not required
####    ###########     ###     ######    ###    #########    ##    #####    ###     #######    ###     #####      ######| |#  PWPS Module Version      : Not required
####                   ####               ###    #########    ##    #####    ####              #####               ######| |#  PWPS_DAB Module Version  : Not required
####                  #######           #####    #########    ##    #####    ######           ########        #    ######| |#
##############################################################################################################    #######| |#  Requirements:
###################################################################################################     ###       #######| |#  ScheduledTasks module.
##################################################################################################              #########| |#
#######################################################################################################      ############| |#
#########################################################################################################################| |#
#####################################################################################################################################################################################################################
##
##  Description/Notes:
##  Lists each task in a task folder with its state, last run time, last result and next run time, and highlights failures.
##  No ProjectWise cmdlets are used.
##  Read-only; does not start, stop or change tasks.
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
Shows the run status of ProjectWise scheduled tasks.

.DESCRIPTION
Reads every task in -TaskPath (default \ProjectWise\, where
Register-PWScriptScheduledTask.ps1 creates tasks) and reports its state, last run time,
last result and next run time. A last result other than 0 is shown as Failed.
Use -FailedOnly in a monitoring job to report only problems.

.PARAMETER TaskPath
Task Scheduler folder to read. Defaults to \ProjectWise\.

.PARAMETER FailedOnly
Lists only tasks whose last run did not succeed.

.EXAMPLE
& '.\Scheduled Tasks\Get-PWScheduledTaskStatus.ps1'
Lists every task in \ProjectWise\.

.EXAMPLE
& '.\Scheduled Tasks\Get-PWScheduledTaskStatus.ps1' -FailedOnly
Lists only tasks whose last run failed.

.INPUTS
None.

.OUTPUTS
One status object per task and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires the ScheduledTasks module. Result 267011 means the task has not run yet.
The transcript of each scheduled script (its .LastRun.log) has the details of a failure.
#>

[CmdletBinding()]
param(
    [string] $TaskPath = '\ProjectWise\',

    [switch] $FailedOnly
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
}

process {
    try {
        Import-Module ScheduledTasks -ErrorAction Stop

        $tasks = @(Get-ScheduledTask -TaskPath $TaskPath -ErrorAction SilentlyContinue)
        if ($tasks.Count -eq 0) {
            Write-Warning "No scheduled tasks were found in '${TaskPath}'."
            return
        }

        $status = foreach ($task in $tasks) {
            $info = $task | Get-ScheduledTaskInfo
            [string] $result = switch ($info.LastTaskResult) {
                0 { 'Succeeded' }
                267011 { 'Never run' }
                267009 { 'Running' }
                default { 'Failed (0x{0:X})' -f $info.LastTaskResult }
            }
            [pscustomobject] @{
                Task        = $task.TaskName
                State       = [string] $task.State
                LastRunTime = $info.LastRunTime
                LastResult  = $result
                NextRunTime = $info.NextRunTime
                RunAs       = $task.Principal.UserId
                Action      = ($task.Actions | ForEach-Object { "$($_.Execute) $($_.Arguments)" }) -join '; '
            }
        }
        if ($FailedOnly) { $status = @($status | Where-Object { $_.LastResult -like 'Failed*' }) }
        $status = @($status)

        if ($status.Count -eq 0) {
            Write-Host 'No failed tasks.'
            return
        }
        $status | Select-Object Task, State, LastRunTime, LastResult, NextRunTime, RunAs | Format-Table -AutoSize | Out-Host
        $status
    }
    catch {
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
