#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Register-PWScriptScheduledTask
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Registers a Windows scheduled task that runs a script from this library in 64-bit Windows PowerShell.
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
###################################################################################################     ###       #######| |#  ScheduledTasks module and rights to register tasks (elevated session for other accounts).
##################################################################################################              #########| |#
#######################################################################################################      ############| |#
#########################################################################################################################| |#
#####################################################################################################################################################################################################################
##
##  Description/Notes:
##  Creates a daily or weekly task under \ProjectWise\ that runs a .ps1 file with optional arguments as a chosen account.
##  No ProjectWise cmdlets are used. The account password is prompted for and passed to Task Scheduler; it is not stored or logged.
##  CREATES or replaces a scheduled task. Supports -WhatIf.
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
Schedules a script from this library to run daily or weekly.

.DESCRIPTION
Registers a Windows scheduled task in the \ProjectWise\ task folder that runs
-ScriptPath with 64-bit Windows PowerShell 5.1. The task runs whether or not the user
is logged on, as the account entered at the credential prompt.

The scheduled script must be able to log in without a prompt. The scripts in this
library prompt for Bentley IMS login when no session exists; for unattended use, adapt
the script to log in with a credential file created by Export-Clixml under the task
account.

.PARAMETER ScriptPath
Full path of the .ps1 file to run.

.PARAMETER TaskName
Name of the task. Defaults to the script file name.

.PARAMETER Schedule
Daily or Weekly. Defaults to Daily.

.PARAMETER At
Time of day to run, for example '02:00'. Defaults to 02:00.

.PARAMETER DaysOfWeek
Days for a weekly schedule. Defaults to Monday.

.PARAMETER ScriptArguments
Optional arguments added after the script path, for example "-Days 30".

.EXAMPLE
& '.\Scheduled Tasks\Register-PWScriptScheduledTask.ps1' -ScriptPath 'C:\Scripts\Licensing\Get-PWConnectedUsersReport.ps1' -Schedule Weekly -DaysOfWeek Monday -At '06:00' -WhatIf
Shows the task that would be registered.

.INPUTS
None.

.OUTPUTS
The registered task and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires the ScheduledTasks module (Windows 8 / Server 2012 or later). Registering a
task for another account usually needs an elevated session. An existing task with the
same name in \ProjectWise\ is replaced. Use Get-PWScheduledTaskStatus.ps1 to check runs.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ (Test-Path -LiteralPath $_ -PathType Leaf) -and $_ -like '*.ps1' })]
    [string] $ScriptPath,

    [string] $TaskName,

    [ValidateSet('Daily', 'Weekly')]
    [string] $Schedule = 'Daily',

    [DateTime] $At = '02:00',

    [ValidateSet('Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday')]
    [string[]] $DaysOfWeek = @('Monday'),

    [string] $ScriptArguments
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
}

process {
    try {
        Import-Module ScheduledTasks -ErrorAction Stop

        [string] $resolvedScript = (Resolve-Path -LiteralPath $ScriptPath).ProviderPath
        if ([string]::IsNullOrWhiteSpace($TaskName)) { $TaskName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedScript) }
        [string] $powershell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
        [string] $arguments = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}" {1}' -f $resolvedScript, $ScriptArguments

        $action = New-ScheduledTaskAction -Execute $powershell -Argument $arguments.Trim() -WorkingDirectory (Split-Path -Parent $resolvedScript)
        $trigger = if ($Schedule -eq 'Daily') {
            New-ScheduledTaskTrigger -Daily -At $At
        }
        else {
            New-ScheduledTaskTrigger -Weekly -DaysOfWeek $DaysOfWeek -At $At
        }
        $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 12) -MultipleInstances IgnoreNew

        [string] $when = if ($Schedule -eq 'Daily') { "daily at $($At.ToString('HH:mm'))" } else { "every $($DaysOfWeek -join ', ') at $($At.ToString('HH:mm'))" }
        Write-Host "Task:    \ProjectWise\${TaskName}"
        Write-Host "Runs:    ${powershell} ${arguments}"
        Write-Host "When:    ${when}"

        if ($PSCmdlet.ShouldProcess("\ProjectWise\${TaskName}", 'Register scheduled task')) {
            $credential = Get-Credential -Message "Account to run '${TaskName}'" -UserName "${env:USERDOMAIN}\${env:USERNAME}"
            if ($null -eq $credential) { throw 'A run-as account is required.' }

            $task = Register-ScheduledTask -TaskName $TaskName -TaskPath '\ProjectWise\' -Action $action -Trigger $trigger -Settings $settings `
                -User $credential.UserName -Password $credential.GetNetworkCredential().Password -RunLevel Limited -Force -ErrorAction Stop
            Write-Host "Registered task '\ProjectWise\${TaskName}' to run as '$($credential.UserName)'."
            $task
        }
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
