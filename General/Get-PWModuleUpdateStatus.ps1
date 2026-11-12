#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Get-PWModuleUpdateStatus
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Compares installed module versions with the latest versions on the PowerShell Gallery.
####    ###########    #######         ######    #         ####         #    ######          ######    ########    ######| |#  Author                   : GitHub Copilot
####    ##########-    ####               ###                ##         #    ####              ####    ########    ######| |#  Creation Date            : 10-27-2026
#                     #####     ######    ###      #####     ###    #####    ###     ######     ###    ########    ######| |#  Modified By              : GitHub Copilot
#                    #####    #########    ##    ########     ##    #####    ##     ########    ###    ########    ######| |#  Modified Date            : 10-27-2026
####    ##########      ##                 ##    #########    ##    #####    ##                  ##    ########    ######| |#  Script Version           : 1.0
####    ############    ##                 ##    #########    ##    #####    ##                  ##    ########    ######| |#  PowerShell Version       : Windows PowerShell 5.1 x64
####    ############    ##    ###############    #########    ##    #####    ##     ###############    #######     ######| |#  ProjectWise Version      : Not required
####    ###########     ###     ######    ###    #########    ##    #####    ###     #######    ###     #####      ######| |#  PWPS Module Version      : Not required
####                   ####               ###    #########    ##    #####    ####              #####               ######| |#  PWPS_DAB Module Version  : Checked, not required to run
####                  #######           #####    #########    ##    #####    ######           ########        #    ######| |#
##############################################################################################################    #######| |#  Requirements:
###################################################################################################     ###       #######| |#  PowerShellGet and internet access to the PowerShell Gallery.
##################################################################################################              #########| |#
#######################################################################################################      ############| |#
#########################################################################################################################| |#
#####################################################################################################################################################################################################################
##
##  Description/Notes:
##  Lists the installed and latest gallery versions of pwps_dab, ImportExcel and any other named modules.
##  No ProjectWise cmdlets are used and no login is made.
##  Read-only; does not install or update modules.
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
Reports whether newer versions of the ProjectWise PowerShell modules are available.

.DESCRIPTION
For each module in -Name, shows the newest installed version, the latest version on the
PowerShell Gallery, and whether an update is available. It does not install anything;
test new pwps_dab versions against your scripts before upgrading production machines.

.PARAMETER Name
Module names to check. Defaults to pwps_dab and ImportExcel.

.EXAMPLE
& '.\General\Get-PWModuleUpdateStatus.ps1'
Checks pwps_dab and ImportExcel.

.EXAMPLE
& '.\General\Get-PWModuleUpdateStatus.ps1' -Name pwps_dab, ImportExcel, PSSQLite
Also checks PSSQLite.

.INPUTS
None.

.OUTPUTS
One status object per module and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires PowerShellGet (Find-Module) and access to https://www.powershellgallery.com.
Proxy servers may need configuration. To update a module after testing, run
Update-Module -Name <name> in an elevated session or Install-Module -Scope CurrentUser.
#>

[CmdletBinding()]
param(
    [string[]] $Name = @('pwps_dab', 'ImportExcel')
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
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

        $status = foreach ($moduleName in $Name) {
            $installed = Get-Module -ListAvailable -Name $moduleName | Sort-Object Version -Descending | Select-Object -First 1
            $latest = $null
            $note = ''
            try {
                $latest = Find-Module -Name $moduleName -Repository PSGallery -ErrorAction Stop
            }
            catch {
                $note = "Gallery lookup failed: $($_.Exception.Message)"
            }

            $updateAvailable = $null -ne $installed -and $null -ne $latest -and ([version] $latest.Version -gt $installed.Version)
            if ($null -eq $installed) { $note = (('Not installed. ' + $note).Trim()) }
            elseif ($updateAvailable) { $note = 'Update available.' }
            elseif ($null -ne $latest) { $note = 'Up to date.' }

            [pscustomobject] @{
                Module          = $moduleName
                Installed       = if ($installed) { [string] $installed.Version } else { '' }
                Latest          = if ($latest) { [string] $latest.Version } else { '' }
                UpdateAvailable = $updateAvailable
                Note            = $note
            }
        }

        $status | Format-Table -AutoSize | Out-Host
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
