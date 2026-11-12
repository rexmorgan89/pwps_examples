#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Test-PWWorkstationPrereqs
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Checks that a workstation meets the requirements for the scripts in this library.
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
###################################################################################################     ###       #######| |#  None. Reads local settings and the registry only.
##################################################################################################              #########| |#
#######################################################################################################      ############| |#
#########################################################################################################################| |#
#####################################################################################################################################################################################################################
##
##  Description/Notes:
##  Checks PowerShell edition and version, 64-bit process, .NET Framework 4.8, pwps_dab and ImportExcel versions, and installed Bentley clients.
##  No ProjectWise cmdlets are used and no login is made.
##  Read-only.
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
Checks a workstation for ProjectWise PowerShell prerequisites.

.DESCRIPTION
Runs a set of checks and prints a Pass, Warn or Fail line for each:
- Windows PowerShell 5.1 (Desktop edition).
- 64-bit PowerShell process.
- .NET Framework 4.8 or later.
- pwps_dab installed at or above -MinimumPwpsDabVersion.
- ImportExcel installed at or above -MinimumImportExcelVersion.
- ProjectWise Explorer and Bentley CONNECTION Client installed.
Run it on a new admin workstation or server before using the other scripts.

.PARAMETER MinimumPwpsDabVersion
Lowest acceptable pwps_dab version. Defaults to 24.0.2.

.PARAMETER MinimumImportExcelVersion
Lowest acceptable ImportExcel version. Defaults to 7.8.6.

.EXAMPLE
& '.\General\Test-PWWorkstationPrereqs.ps1'
Prints the result of every check.

.INPUTS
None.

.OUTPUTS
One result object per check and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires no modules and no ProjectWise login. Installed applications are detected from
the Windows uninstall registry keys by display name, so renamed products may show as
missing.
#>

[CmdletBinding()]
param(
    [version] $MinimumPwpsDabVersion = '24.0.2',

    [version] $MinimumImportExcelVersion = '7.8.6'
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

    function New-CheckResult {
        param([string] $Check, [ValidateSet('Pass', 'Warn', 'Fail')] [string] $Result, [string] $Detail)
        [pscustomobject] @{ Check = $Check; Result = $Result; Detail = $Detail }
    }

    function Test-ModuleVersion {
        param([string] $Name, [version] $Minimum)
        $installed = Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending | Select-Object -First 1
        if ($null -eq $installed) { return New-CheckResult $Name 'Fail' 'Not installed.' }
        if ($installed.Version -lt $Minimum) { return New-CheckResult $Name 'Warn' "Version $($installed.Version) is below ${Minimum}." }
        New-CheckResult $Name 'Pass' "Version $($installed.Version)."
    }
}

process {
    try {
        $results = [System.Collections.Generic.List[object]]::new()

        $edition = $PSVersionTable.PSEdition
        $psVersion = $PSVersionTable.PSVersion
        if ($edition -eq 'Desktop' -and $psVersion.Major -eq 5 -and $psVersion.Minor -ge 1) {
            $results.Add((New-CheckResult 'Windows PowerShell 5.1' 'Pass' "$edition $psVersion"))
        }
        else {
            $results.Add((New-CheckResult 'Windows PowerShell 5.1' 'Fail' "Running $edition $psVersion. Most scripts here need Windows PowerShell 5.1."))
        }

        if ([Environment]::Is64BitProcess) {
            $results.Add((New-CheckResult '64-bit process' 'Pass' 'x64'))
        }
        else {
            $results.Add((New-CheckResult '64-bit process' 'Fail' 'Running x86. Start the 64-bit Windows PowerShell.'))
        }

        $netRelease = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -Name Release -ErrorAction SilentlyContinue).Release
        if ($netRelease -ge 528040) {
            $results.Add((New-CheckResult '.NET Framework 4.8' 'Pass' "Release ${netRelease}"))
        }
        else {
            $results.Add((New-CheckResult '.NET Framework 4.8' 'Fail' "Release '${netRelease}' is older than 4.8."))
        }

        $results.Add((Test-ModuleVersion -Name 'pwps_dab' -Minimum $MinimumPwpsDabVersion))
        $results.Add((Test-ModuleVersion -Name 'ImportExcel' -Minimum $MinimumImportExcelVersion))

        $uninstallKeys = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
        $installed = @(Get-ItemProperty -Path $uninstallKeys -ErrorAction SilentlyContinue | Where-Object DisplayName)
        foreach ($product in @(
                @{ Check = 'ProjectWise Explorer'; Pattern = 'ProjectWise*Explorer*' }
                @{ Check = 'Bentley CONNECTION Client'; Pattern = '*CONNECTION Client*' }
            )) {
            $match = $installed | Where-Object { $_.DisplayName -like $product.Pattern } | Select-Object -First 1
            if ($match) {
                $results.Add((New-CheckResult $product.Check 'Pass' "$($match.DisplayName) $($match.DisplayVersion)"))
            }
            else {
                $results.Add((New-CheckResult $product.Check 'Warn' 'Not found in installed programs.'))
            }
        }

        $colors = @{ Pass = 'Green'; Warn = 'Yellow'; Fail = 'Red' }
        foreach ($result in $results) {
            Write-Host ('[{0}] {1}: {2}' -f $result.Result.ToUpperInvariant(), $result.Check, $result.Detail) -ForegroundColor $colors[$result.Result]
        }
        $results
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
