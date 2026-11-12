#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Export-PWEnvironmentDefinitions
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Documents environment columns in Excel and optionally exports full definitions for re-import.
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
##  Lists the columns of every environment (or selected ones) and can also run Export-PWEnvironments for use with Import-PWEnvironments.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWEnvironments, Get-PWEnvironmentColumns, Export-PWEnvironments, Undo-PWLogin.
##  Read-only; does not change environments.
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
Documents ProjectWise environment definitions.

.DESCRIPTION
Writes an Excel workbook with one row per environment column, for review or change
control. With -IncludeFullExport, also runs Export-PWEnvironments to a second file that
Import-PWEnvironments can read, for copying environments to another datasource.

.PARAMETER EnvironmentName
Optional environment names. All environments are documented when empty.

.PARAMETER IncludeFullExport
Also writes a full Export-PWEnvironments file beside the report.

.PARAMETER OutputPath
Full path of the Excel workbook to create. Defaults to a timestamped file beside the script.

.EXAMPLE
& '.\Environments\Export-PWEnvironmentDefinitions.ps1'
Documents the columns of every environment.

.EXAMPLE
& '.\Environments\Export-PWEnvironmentDefinitions.ps1' -EnvironmentName 'Civil' -IncludeFullExport
Documents one environment and exports its full definition.

.INPUTS
None.

.OUTPUTS
One or two Excel workbooks and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab and ImportExcel. Uses the active ProjectWise session or prompts for a
Bentley IMS login, and logs out only if this script opened the session.
#>

[CmdletBinding()]
param(
    [string[]] $EnvironmentName,

    [switch] $IncludeFullExport,

    [string] $OutputPath = (Join-Path $PSScriptRoot ('EnvironmentColumns_{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd_HHmmss')))
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

        $names = @(Get-PWEnvironments -ErrorAction Stop | ForEach-Object { [string] $_.Name })
        if ($EnvironmentName) {
            $unknown = @($EnvironmentName | Where-Object { $names -notcontains $_ })
            if ($unknown.Count -gt 0) { throw "Environment(s) not found: $($unknown -join ', ')" }
            $names = @($EnvironmentName)
        }

        $columns = foreach ($name in $names) {
            Get-PWEnvironmentColumns -EnvironmentName $name -ErrorAction Continue |
                Select-Object -Property @{ Name = 'Environment'; Expression = { $name } }, *
        }
        $columns = @($columns)
        if ($columns.Count -gt 0) {
            $headers = @($columns | ForEach-Object { $_.PSObject.Properties.Name } | Select-Object -Unique)
            $columns | Select-Object -Property $headers |
                Export-Excel -Path $OutputPath -WorksheetName 'Columns' -AutoSize -FreezeTopRow -BoldTopRow
            Write-Host "Documented $($columns.Count) column(s) in $($names.Count) environment(s) in '${OutputPath}'."
        }
        else {
            Write-Warning 'No environment columns were returned.'
        }

        if ($IncludeFullExport) {
            [string] $fullExportPath = [System.IO.Path]::ChangeExtension($OutputPath, $null).TrimEnd('.') + '_FullExport.xlsx'
            $exportParameters = @{ OutputFileName = $fullExportPath; ErrorAction = 'Stop' }
            if ($EnvironmentName) { $exportParameters['Environments'] = $names }
            $null = Export-PWEnvironments @exportParameters
            Write-Host "Full environment export saved to '${fullExportPath}'."
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
