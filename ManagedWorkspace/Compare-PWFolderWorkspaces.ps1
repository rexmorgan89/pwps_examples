#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Compare-PWFolderWorkspaces
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Compares the Managed Workspace configuration blocks that apply to two folders.
####    ###########    #######         ######    #         ####         #    ######          ######    ########    ######| |#  Author                   : GitHub Copilot
####    ##########-    ####               ###                ##         #    ####              ####    ########    ######| |#  Creation Date            : 10-27-2026
#                     #####     ######    ###      #####     ###    #####    ###     ######     ###    ########    ######| |#  Modified By              : GitHub Copilot
#                    #####    #########    ##    ########     ##    #####    ##     ########    ###    ########    ######| |#  Modified Date            : 11-12-2026
####    ##########      ##                 ##    #########    ##    #####    ##                  ##    ########    ######| |#  Script Version           : 1.1
####    ############    ##                 ##    #########    ##    #####    ##                  ##    ########    ######| |#  PowerShell Version       : Windows PowerShell 5.1 x64
####    ############    ##    ###############    #########    ##    #####    ##     ###############    #######     ######| |#  ProjectWise Version      : Datasource with Managed Workspace configured
####    ###########     ###     ######    ###    #########    ##    #####    ###     #######    ###     #####      ######| |#  PWPS Module Version      : Not required
####                   ####               ###    #########    ##    #####    ####              #####               ######| |#  PWPS_DAB Module Version  : Repository reference: 24.0.2
####                  #######           #####    #########    ##    #####    ######           ########        #    ######| |#
##############################################################################################################    #######| |#  Requirements:
###################################################################################################     ###       #######| |#  pwps_dab and read access to both folders.
##################################################################################################              #########| |#
#######################################################################################################      ############| |#
#########################################################################################################################| |#
#####################################################################################################################################################################################################################
##
##  Description/Notes:
##  Lists the CSBs assigned to or inherited by two folders and shows which apply to only one of them.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWFolders, Get-PWManagedWorkspaces, Undo-PWLogin.
##  Read-only; does not change workspace assignments.
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
Compares the Managed Workspace CSBs of two folders.

.DESCRIPTION
Gets the configuration settings blocks (CSBs) that apply to each folder, including
inherited ones, and lists every CSB with whether it applies to the first folder, the
second folder or both. Use it to answer "why does this project behave differently from
that one in MicroStation or OpenRoads?".

.PARAMETER FolderPath
ProjectWise folder path of the first folder.

.PARAMETER CompareFolderPath
ProjectWise folder path of the second folder.

.PARAMETER DifferencesOnly
Lists only CSBs that apply to one folder but not the other.

.EXAMPLE
& '.\ManagedWorkspace\Compare-PWFolderWorkspaces.ps1' -FolderPath 'Projects\11431' -CompareFolderPath 'Projects\11500' -DifferencesOnly
Shows the CSBs that apply to only one of the two projects.

.INPUTS
None.

.OUTPUTS
One comparison object per CSB and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab. Uses the active ProjectWise session or prompts for a Bentley IMS
login, and logs out only if this script opened the session. CSBs are matched by name and
level; variable values inside the CSBs are not compared.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $FolderPath,

    [Parameter(Mandatory)]
    [string] $CompareFolderPath,

    [switch] $DifferencesOnly
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

    function Get-FolderBlockKeys {
        param([string] $Path)
        $folder = Get-PWFolders -FolderPath $Path -JustOne -ErrorAction SilentlyContinue
        if ($null -eq $folder) { throw "Folder '${Path}' was not found." }
        $keys = @{}
        foreach ($block in @(Get-PWManagedWorkspaces -InputFolders $folder -IncludeInherited -ErrorAction Stop)) {
            $keys['{0}|{1}' -f $block.Level, $block.Name] = $block
        }
        $keys
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

        $first = Get-FolderBlockKeys -Path $FolderPath
        $second = Get-FolderBlockKeys -Path $CompareFolderPath

        $comparison = foreach ($key in (@($first.Keys) + @($second.Keys) | Sort-Object -Unique)) {
            $block = if ($first.ContainsKey($key)) { $first[$key] } else { $second[$key] }
            [string] $appliesTo = if ($first.ContainsKey($key) -and $second.ContainsKey($key)) { 'Both' }
            elseif ($first.ContainsKey($key)) { 'First only' }
            else { 'Second only' }
            if ($DifferencesOnly -and $appliesTo -eq 'Both') { continue }

            [pscustomobject] @{
                Level       = $block.Level
                CSB         = $block.Name
                AppliesTo   = $appliesTo
                AssignedAt  = $block.FolderPath
                Description = $block.Description
            }
        }
        $comparison = @($comparison)

        Write-Host "First:  '${FolderPath}' ($($first.Count) CSB(s))"
        Write-Host "Second: '${CompareFolderPath}' ($($second.Count) CSB(s))"
        if ($comparison.Count -eq 0) {
            Write-Host 'Both folders use the same CSBs.'
            return
        }
        $comparison | Sort-Object Level, CSB | Format-Table -AutoSize | Out-Host
        $comparison
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
