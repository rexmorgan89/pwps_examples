#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Export-PWUsersReport
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Exports every ProjectWise user with type, identity, email, last login and status to Excel.
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
##  Exports all users with their account type, security provider, identity, email, last login and disabled flag, plus a summary by type.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWUsersByMatch, Undo-PWLogin.
##  Read-only; does not change users.
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
Exports ProjectWise users to Excel.

.DESCRIPTION
Writes an Excel workbook with:
- Users: one row per user with user name, full name, email, type, security provider,
  identity, identity provider, creation date, last login and disabled flag.
- Summary: user counts by type and status.
- MissingEmail: enabled users with no email address, which often cannot sign in with
  Bentley IMS.

.PARAMETER UserName
Optional user name patterns (for example 'jsmith*'). Exports every user when empty.

.PARAMETER OutputPath
Full path of the Excel workbook to create. Defaults to a timestamped file beside the script.

.EXAMPLE
& '.\User\Export-PWUsersReport.ps1'
Exports every user.

.INPUTS
None.

.OUTPUTS
An Excel workbook and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab and ImportExcel. Uses the active ProjectWise session or prompts for a
Bentley IMS login, and logs out only if this script opened the session.
#>

[CmdletBinding()]
param(
    [string[]] $UserName,

    [string] $OutputPath = (Join-Path $PSScriptRoot ('Users_{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd_HHmmss')))
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

        $userParameters = @{ LastLogin = $true; ErrorAction = 'Stop' }
        if ($UserName) { $userParameters['UserName'] = $UserName }
        $users = @(Get-PWUsersByMatch @userParameters |
                Select-Object UserName, Name, Email, Description, Type, SecProvider, Identity, IdentityProvider, CreationDate, LastLogin, IsDisabled)

        if ($users.Count -eq 0) {
            Write-Warning 'No users were found.'
            return
        }

        $users | Sort-Object UserName |
            Export-Excel -Path $OutputPath -WorksheetName 'Users' -AutoSize -FreezeTopRow -BoldTopRow
        $users | Group-Object Type, IsDisabled | ForEach-Object {
            [pscustomobject] @{ Type = $_.Group[0].Type; Disabled = $_.Group[0].IsDisabled; Users = $_.Count }
        } | Sort-Object Type, Disabled |
            Export-Excel -Path $OutputPath -WorksheetName 'Summary' -AutoSize -BoldTopRow

        $missingEmail = @($users | Where-Object { -not $_.IsDisabled -and $_.Type -ne 'Logical' -and [string]::IsNullOrWhiteSpace($_.Email) })
        if ($missingEmail.Count -gt 0) {
            $missingEmail | Export-Excel -Path $OutputPath -WorksheetName 'MissingEmail' -AutoSize -FreezeTopRow -BoldTopRow
        }
        Write-Host "Exported $($users.Count) user(s); $($missingEmail.Count) enabled user(s) have no email. Report saved to '${OutputPath}'."
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
