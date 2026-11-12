#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : CmdletReferenceOutputter
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Generate <module>-Reference.txt cmdlet help references for installed modules.
####    ###########    #######         ######    #         ####         #    ######          ######    ########    ######| |#  Author                   : Unknown
####    ##########-    ####               ###                ##         #    ####              ####    ########    ######| |#  Creation Date            : Unknown
#                     #####     ######    ###      #####     ###    #####    ###     ######     ###    ########    ######| |#  Modified By              : Andrew Frimpong
#                    #####    #########    ##    ########     ##    #####    ##     ########    ###    ########    ######| |#  Modified Date            : 11-06-2026
####    ##########      ##                 ##    #########    ##    #####    ##                  ##    ########    ######| |#  Script Version           : 1.0
####    ############    ##                 ##    #########    ##    #####    ##                  ##    ########    ######| |#  PowerShell Version       : 5.1
####    ############    ##    ###############    #########    ##    #####    ##     ###############    #######     ######| |#  ProjectWise Version      : Unknown
####    ###########     ###     ######    ###    #########    ##    #####    ###     #######    ###     #####      ######| |#  PWPS Module Version      : N/A
####                   ####               ###    #########    ##    #####    ####              #####               ######| |#  PWPS_DAB Module Version  : Not required
####                  #######           #####    #########    ##    #####    ######           ########        #    ######| |#
##############################################################################################################    #######| |#  Requirements:
###################################################################################################     ###       #######| |#  The modules to document must be installed; write access to the script folder (docs).
##################################################################################################              #########| |#
#######################################################################################################      ############| |#
#########################################################################################################################| |#
#####################################################################################################################################################################################################################
##
##  Description/Notes:
##  Picks modules by -ModuleName (newest installed version of each) or from an Out-GridView list of Get-Module -ListAvailable,
##  then writes the full Get-Help text of every command from Get-Command -Module to <module>-Reference.txt beside the script.
##  Overwrites existing reference files of the same name. Does not connect to ProjectWise.
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
Writes a full cmdlet help reference text file for each selected module.

.DESCRIPTION
For each module named in -ModuleName (newest installed version only), or each module picked in
Out-GridView when -ModuleName is omitted, collects the commands with Get-Command -Module and
writes the Get-Help -Full text of every command to <module name>-Reference.txt in the script
folder. Warns about requested modules that are not installed.

.PARAMETER ModuleName
Names of installed modules to document. Omit to pick modules interactively in Out-GridView.

.EXAMPLE
.\CmdletReferenceOutputter.ps1 -ModuleName pwps_dab
Writes pwps_dab-Reference.txt beside the script.

.EXAMPLE
.\CmdletReferenceOutputter.ps1
Opens Out-GridView to pick one or more installed modules to document.

.NOTES
Needs only the modules being documented. Outputs <module name>-Reference.txt (UTF-8) beside the
script, overwriting existing files, and a timestamped .LastRun.log transcript. No ProjectWise login.
#>

[CmdletBinding()]
param(
    # Omit to pick modules interactively in Out-GridView.
    [string[]] $ModuleName
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
        #$yourReferenceDocName = Read-Host -Prompt 'Enter the name of the reference document (without extension)' 
        if ($ModuleName) {
            # Keep only the newest installed version of each module.
            $referenceModules = Get-Module -ListAvailable -Name $ModuleName |
                Sort-Object -Property Version -Descending |
                Group-Object -Property Name |
                ForEach-Object { $_.Group[0] }

            $missing = $ModuleName | Where-Object { $_ -notin $referenceModules.Name }
            if ($missing) {
                Write-Warning "Module(s) not installed: $($missing -join ', ')"
            }
        }
        else {
            $referenceModules = (Get-Module -ListAvailable | Out-GridView -OutputMode Multiple -Title 'Select the module(s) to generate reference documentation for')
        }

        foreach ($referenceModule in $referenceModules) {
    
            $cmds = Get-Command -Module $referenceModule

            $output = foreach ($cmd in $cmds) {
                "## $($cmd.Name)`n"
                Get-Help $cmd.Name -Full | Out-String
                "`n--`n"
            }

            $output | Out-File -FilePath (Join-Path $PSScriptRoot "$($referenceModule.Name)-Reference.txt") -Encoding utf8 -Force
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
