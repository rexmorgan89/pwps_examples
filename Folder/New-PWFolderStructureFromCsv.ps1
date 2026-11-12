#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : New-PWFolderStructureFromCsv
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Creates a standard folder structure under a parent folder from a CSV template.
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
###################################################################################################     ###       #######| |#  pwps_dab and rights to create folders under the parent folder.
##################################################################################################              #########| |#
#######################################################################################################      ############| |#
#########################################################################################################################| |#
#####################################################################################################################################################################################################################
##
##  Description/Notes:
##  Reads RelativePath (and optional Description) rows from a CSV and creates any missing folders, parents first.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWFolders, New-PWFolder, Undo-PWLogin.
##  CREATES folders. Existing folders are skipped, never changed. Supports -WhatIf.
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
Creates a folder structure from a CSV template.

.DESCRIPTION
Reads a CSV with a RelativePath column (for example '03_Design\Plan') and an optional
Description column, and creates each folder under -ParentFolderPath. Rows are processed
shallowest first so that parents exist before their children. Folders that already exist
are skipped.

New folders use -StorageArea, -Environment and -Workflow when given; otherwise they use
the storage area of the parent folder and no environment or workflow.

.PARAMETER ParentFolderPath
Existing ProjectWise folder under which the structure is created.

.PARAMETER CsvPath
CSV file with a RelativePath column and an optional Description column.

.PARAMETER StorageArea
Storage area for new folders. Defaults to the parent folder's storage area.

.PARAMETER Environment
Optional environment for new folders.

.PARAMETER Workflow
Optional workflow for new folders.

.EXAMPLE
& '.\Folder\New-PWFolderStructureFromCsv.ps1' -ParentFolderPath 'Projects\11431' -CsvPath '.\Folder\StandardStructure.csv' -WhatIf
Lists the folders that would be created.

.INPUTS
None.

.OUTPUTS
A results table and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab. Uses the active ProjectWise session or prompts for a Bentley IMS
login, and logs out only if this script opened the session. New folders inherit
security from their parent; this script does not set access control.
Example CSV:
    RelativePath,Description
    01_Admin,Administration
    03_Design,Design
    03_Design\Plan,Plan sheets
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string] $ParentFolderPath,

    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string] $CsvPath,

    [string] $StorageArea,

    [string] $Environment,

    [string] $Workflow
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

    [bool] $loggedInByScript = $false
}

process {
    try {
        Import-Module pwps_dab -DisableNameChecking -ErrorAction Stop

        $rows = @(Import-Csv -LiteralPath $CsvPath | Where-Object { -not [string]::IsNullOrWhiteSpace($_.RelativePath) })
        if ($rows.Count -eq 0) { throw "CSV '${CsvPath}' has no RelativePath values." }

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

        $parent = Get-PWFolders -FolderPath $ParentFolderPath -JustOne -ErrorAction SilentlyContinue
        if ($null -eq $parent) { throw "Parent folder '${ParentFolderPath}' was not found." }
        if ([string]::IsNullOrWhiteSpace($StorageArea)) { $StorageArea = $parent.Storage }

        $ordered = $rows | Sort-Object -Property { ($_.RelativePath.Trim('\') -split '\\').Count }, RelativePath
        $results = foreach ($row in $ordered) {
            [string] $fullPath = '{0}\{1}' -f $ParentFolderPath.TrimEnd('\'), $row.RelativePath.Trim().Trim('\')
            if ($null -ne (Get-PWFolders -FolderPath $fullPath -JustOne -ErrorAction SilentlyContinue)) {
                [pscustomobject] @{ Folder = $fullPath; Result = 'Exists'; Message = '' }
                continue
            }
            if (-not $PSCmdlet.ShouldProcess($fullPath, 'Create folder')) {
                [pscustomobject] @{ Folder = $fullPath; Result = 'Skipped'; Message = 'WhatIf or declined' }
                continue
            }
            try {
                $newFolder = @{ FolderPath = $fullPath; StorageArea = $StorageArea; ErrorAction = 'Stop' }
                if (-not [string]::IsNullOrWhiteSpace($Environment)) { $newFolder['Environment'] = $Environment }
                if (-not [string]::IsNullOrWhiteSpace($Workflow)) { $newFolder['Workflow'] = $Workflow }
                if (-not [string]::IsNullOrWhiteSpace($row.Description)) { $newFolder['Description'] = $row.Description }
                $null = New-PWFolder @newFolder
                [pscustomobject] @{ Folder = $fullPath; Result = 'Created'; Message = '' }
            }
            catch {
                [pscustomobject] @{ Folder = $fullPath; Result = 'Failed'; Message = $_.Exception.Message }
            }
        }

        $results | Format-Table -AutoSize | Out-Host
        $results | Group-Object Result | ForEach-Object { Write-Host "$($_.Name): $($_.Count)" }
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
