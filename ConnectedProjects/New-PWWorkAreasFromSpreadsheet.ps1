#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : New-PWWorkAreasFromSpreadsheet
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Validates inputs and creates work areas in bulk from an Excel input file.
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
###################################################################################################     ###       #######| |#  pwps_dab, ImportExcel and rights to create work areas.
##################################################################################################              #########| |#
#######################################################################################################      ############| |#
#########################################################################################################################| |#
#####################################################################################################################################################################################################################
##
##  Description/Notes:
##  Checks the input file, parent folder, template and storage area, then calls New-PWRichProjectsFromSpreadsheet.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWFolders, Get-PWStorageAreaList, New-PWRichProjectsFromSpreadsheet, Undo-PWLogin.
##  CREATES work areas and folders. Supports -WhatIf; test in a non-production datasource first.
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
Creates work areas in bulk from an Excel input file, with checks first.

.DESCRIPTION
Before creating anything, confirms that:
- The input workbook exists and has at least one data row.
- The default parent folder exists.
- The default template folder exists.
- The default storage area exists.
Then calls New-PWRichProjectsFromSpreadsheet. Use -WhatIf to run the checks only.

The input workbook layout is defined by New-PWRichProjectsFromSpreadsheet; see its help
in docs/pwps_dab-Reference.md. ConnectedProjects\New-PWProjectInputFile0.3.ps1 creates
a starting workbook.

.PARAMETER InputFileName
Full path of the Excel input workbook.

.PARAMETER DefaultNewParentPath
Folder path under which work areas are created when a row does not give a parent.

.PARAMETER DefaultTemplateName
Folder path of the work area template to copy when a row does not give a template.

.PARAMETER DefaultStorageArea
Storage area used when a row does not give one.

.EXAMPLE
& '.\ConnectedProjects\New-PWWorkAreasFromSpreadsheet.ps1' -InputFileName 'C:\Temp\NewProjects.xlsx' -DefaultNewParentPath 'Projects' -DefaultTemplateName 'Templates\Standard Project' -DefaultStorageArea 'Storage01' -WhatIf
Runs every check without creating work areas.

.INPUTS
None.

.OUTPUTS
Output from New-PWRichProjectsFromSpreadsheet and a timestamped .LastRun.log transcript.

.NOTES
Requires pwps_dab and ImportExcel. Uses the active ProjectWise session or prompts for a
Bentley IMS login, and logs out only if this script opened the session. Creating work
areas cannot be undone by this script; remove unwanted work areas manually.
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string] $InputFileName,

    [Parameter(Mandatory)]
    [string] $DefaultNewParentPath,

    [Parameter(Mandatory)]
    [string] $DefaultTemplateName,

    [Parameter(Mandatory)]
    [string] $DefaultStorageArea
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

        $rowCount = @(Import-Excel -Path $InputFileName -ErrorAction Stop).Count
        if ($rowCount -eq 0) { throw "Input file '${InputFileName}' has no data rows." }
        Write-Host "Input file has ${rowCount} row(s)."

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

        if ($null -eq (Get-PWFolders -FolderPath $DefaultNewParentPath -JustOne -ErrorAction SilentlyContinue)) {
            throw "Parent folder '${DefaultNewParentPath}' was not found."
        }
        if ($null -eq (Get-PWFolders -FolderPath $DefaultTemplateName -JustOne -ErrorAction SilentlyContinue)) {
            throw "Template folder '${DefaultTemplateName}' was not found."
        }
        if (@(Get-PWStorageAreaList -StorageAreaNames $DefaultStorageArea -ErrorAction SilentlyContinue).Count -eq 0) {
            throw "Storage area '${DefaultStorageArea}' was not found."
        }
        Write-Host 'All checks passed.'

        if ($PSCmdlet.ShouldProcess("${rowCount} row(s) from '${InputFileName}' under '${DefaultNewParentPath}'", 'Create work areas')) {
            New-PWRichProjectsFromSpreadsheet -InputFileName $InputFileName -DefaultNewParentPath $DefaultNewParentPath `
                -DefaultTemplateName $DefaultTemplateName -DefaultStorageArea $DefaultStorageArea -ErrorAction Stop
            Write-Host 'Work area creation finished. Review the output above for rows that failed.'
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
