#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Find-PWFoldersWithExplicitSecurity
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Lists folders below a path that have their own access control instead of inheriting it.
####    ###########    #######         ######    #         ####         #    ######          ######    ########    ######| |#  Author                   : GitHub Copilot
####    ##########-    ####               ###                ##         #    ####              ####    ########    ######| |#  Creation Date            : 10-27-2026
#                     #####     ######    ###      #####     ###    #####    ###     ######     ###    ########    ######| |#  Modified By              : GitHub Copilot
#                    #####    #########    ##    ########     ##    #####    ##     ########    ###    ########    ######| |#  Modified Date            : 10-27-2026
####    ##########      ##                 ##    #########    ##    #####    ##                  ##    ########    ######| |#  Script Version           : 1.0
####    ############    ##                 ##    #########    ##    #####    ##                  ##    ########    ######| |#  PowerShell Version       : Windows PowerShell 5.1 x64
####    ############    ##    ###############    #########    ##    #####    ##     ###############    #######     ######| |#  ProjectWise Version      : Datasource with the dms_proj.o_aclno column
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
##  Flags folders with an access control list of their own (dms_proj.o_aclno > 0) and can list their access entries.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWFolders, Select-PWSQLDataTable, Get-PWFolderSecurity, Undo-PWLogin.
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
Lists folders that have explicit (non-inherited) access control.

.DESCRIPTION
Reads every folder below -FolderPath and flags those with their own access control
list. These folders stop inheriting security from their parent, so changes made higher
in the tree do not reach them. With -IncludeAccessEntries, a second worksheet lists
the access entries of each flagged folder.

Run this before using the inheritance reset scripts in this folder, to see what would
change.

.PARAMETER FolderPath
ProjectWise folder path to search, without the leading 'Documents' folder.

.PARAMETER IncludeAccessEntries
Adds an AccessEntries worksheet from Get-PWFolderSecurity. Slower on large trees.

.PARAMETER OutputPath
Full path of the Excel workbook to create. Defaults to a timestamped file beside the script.

.EXAMPLE
& '.\AccessControl\Find-PWFoldersWithExplicitSecurity.ps1' -FolderPath 'Projects\11431' -IncludeAccessEntries
Lists folders in one project with explicit security, and their access entries.

.INPUTS
None.

.OUTPUTS
An Excel workbook and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab and ImportExcel. Uses the active ProjectWise session or prompts for a
Bentley IMS login, and logs out only if this script opened the session. Document-level
security is not included.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $FolderPath,

    [switch] $IncludeAccessEntries,

    [string] $OutputPath = (Join-Path $PSScriptRoot ('FoldersWithExplicitSecurity_{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd_HHmmss')))
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

    function ConvertTo-PlainObject {
        param([Parameter(ValueFromPipeline)] $InputObject)
        process {
            if ($InputObject -is [System.Data.DataRow]) {
                $row = [ordered] @{}
                foreach ($column in $InputObject.Table.Columns) { $row[$column.ColumnName] = $InputObject[$column] }
                [pscustomobject] $row
            }
            elseif ($null -ne $InputObject) { $InputObject }
        }
    }
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

        $folders = @(Get-PWFolders -FolderPath $FolderPath -ErrorAction Stop)
        if ($folders.Count -eq 0) { throw "Folder '${FolderPath}' was not found." }

        $explicitIds = @{}
        foreach ($row in (Select-PWSQLDataTable -SQLSelectStatement 'SELECT o_projectno FROM dms_proj WHERE o_aclno > 0' -ErrorAction Stop).Rows) {
            $explicitIds[[int] $row.o_projectno] = $true
        }

        $flagged = @($folders | Where-Object { $explicitIds.ContainsKey([int] $_.ProjectID) } | Sort-Object -Property FullPath)
        Write-Host "Checked $($folders.Count) folder(s); $($flagged.Count) have explicit security."
        if ($flagged.Count -eq 0) { return }

        $flagged | Select-Object FullPath, Name, IsRichProject, Workflow, FolderDepth |
            Export-Excel -Path $OutputPath -WorksheetName 'Folders' -AutoSize -FreezeTopRow -BoldTopRow

        if ($IncludeAccessEntries) {
            $entries = foreach ($folder in $flagged) {
                Get-PWFolderSecurity -InputFolder $folder -ErrorAction Continue | ConvertTo-PlainObject |
                    Select-Object -Property @{ Name = 'FolderPath'; Expression = { $folder.FullPath } }, *
            }
            $entries = @($entries)
            if ($entries.Count -gt 0) {
                $columns = @($entries | ForEach-Object { $_.PSObject.Properties.Name } | Select-Object -Unique)
                $entries | Select-Object -Property $columns |
                    Export-Excel -Path $OutputPath -WorksheetName 'AccessEntries' -AutoSize -FreezeTopRow -BoldTopRow
            }
        }
        Write-Host "Report saved to '${OutputPath}'."
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
