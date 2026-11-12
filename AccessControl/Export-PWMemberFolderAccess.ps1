#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Export-PWMemberFolderAccess
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Lists every folder whose access control includes a given user, group or user list.
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
###################################################################################################     ###       #######| |#  pwps_dab, ImportExcel and an administrator login.
##################################################################################################              #########| |#
#######################################################################################################      ############| |#
#########################################################################################################################| |#
#####################################################################################################################################################################################################################
##
##  Description/Notes:
##  Finds the folders where a user, group or user list is named in the folder access control, below a path or across the datasource.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWFolders, Get-PWFoldersByAccessMember, Undo-PWLogin.
##  Read-only; does not change security.
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
Lists the folders where a user, group or user list has explicit access.

.DESCRIPTION
Uses Get-PWFoldersByAccessMember to find folders whose access control names the
member. Searches below -FolderPath, or the whole datasource when -FolderPath is empty.
Use it before removing a group, retiring a user list or answering "where does this
account have access?".

.PARAMETER MemberType
Type of access control member: User, Group or UserList.

.PARAMETER MemberName
Name of the user, group or user list.

.PARAMETER FolderPath
Optional folder path to limit the search. Searches the whole datasource when empty.

.PARAMETER OutputPath
Full path of the Excel workbook to create. Defaults to a timestamped file beside the script.

.EXAMPLE
& '.\AccessControl\Export-PWMemberFolderAccess.ps1' -MemberType UserList -MemberName 'Read_Only' -FolderPath 'Projects'
Lists folders below 'Projects' that have the 'Read_Only' user list in their security.

.INPUTS
None.

.OUTPUTS
An Excel workbook and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab and ImportExcel. Uses the active ProjectWise session or prompts for a
Bentley IMS login, and logs out only if this script opened the session. Only explicit
entries are found; access a member gets through inheritance or group membership is not.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateSet('User', 'Group', 'UserList')]
    [string] $MemberType,

    [Parameter(Mandatory)]
    [string] $MemberName,

    [string] $FolderPath,

    [string] $OutputPath = (Join-Path $PSScriptRoot ('MemberFolderAccess_{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd_HHmmss')))
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
        Import-Module ImportExcel -ErrorAction Stop

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

        if ([string]::IsNullOrWhiteSpace($FolderPath)) {
            $folders = @(Get-PWFoldersByAccessMember -MemberType $MemberType -MemberName $MemberName -SearchEntireDatasource -ErrorAction Stop)
        }
        else {
            $searchFolders = @(Get-PWFolders -FolderPath $FolderPath -ErrorAction Stop)
            if ($searchFolders.Count -eq 0) { throw "Folder '${FolderPath}' was not found." }
            $folders = @(Get-PWFoldersByAccessMember -InputFolder $searchFolders -MemberType $MemberType -MemberName $MemberName -ErrorAction Stop)
        }

        if ($folders.Count -eq 0) {
            Write-Host "No folders name ${MemberType} '${MemberName}' in their access control."
            return
        }

        $folders | Sort-Object -Property FullPath | Select-Object -Property @(
            @{ Name = 'MemberType'; Expression = { $MemberType } }
            @{ Name = 'MemberName'; Expression = { $MemberName } }
            'FullPath', 'Name', 'IsRichProject', 'Environment', 'Workflow', 'Storage'
        ) | Export-Excel -Path $OutputPath -WorksheetName 'Folders' -AutoSize -FreezeTopRow -BoldTopRow
        Write-Host "Found $($folders.Count) folder(s). Report saved to '${OutputPath}'."
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
