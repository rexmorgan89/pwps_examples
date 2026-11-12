#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Get-PWLargestFilesReport
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Lists the largest files below a folder and the total size of each folder.
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
###################################################################################################     ###       #######| |#  pwps_dab, ImportExcel and read access to the target folder.
##################################################################################################              #########| |#
#######################################################################################################      ############| |#
#########################################################################################################################| |#
#####################################################################################################################################################################################################################
##
##  Description/Notes:
##  Searches the documents below a folder, lists the N largest files and sums file size per folder and per file extension.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWDocumentsBySearch, Undo-PWLogin.
##  Read-only; does not change ProjectWise data.
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
Reports the largest files and folder sizes below a ProjectWise folder.

.DESCRIPTION
Writes an Excel workbook with:
- LargestFiles: the -Top largest files, with size in MB.
- ByFolder: total file size and document count per folder.
- ByExtension: total file size and document count per file extension.
Use it to find storage growth, oversized files and folders worth archiving.

.PARAMETER FolderPath
ProjectWise folder path to search, without the leading 'Documents' folder.

.PARAMETER Top
Number of largest files to list. Defaults to 100.

.PARAMETER IncludeVersions
Includes earlier document versions in the totals.

.PARAMETER OutputPath
Full path of the Excel workbook to create. Defaults to a timestamped file beside the script.

.EXAMPLE
& '.\FileManipulation\Get-PWLargestFilesReport.ps1' -FolderPath 'Projects' -Top 50
Lists the 50 largest files below 'Projects'.

.INPUTS
None.

.OUTPUTS
An Excel workbook and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab and ImportExcel. Uses the active ProjectWise session or prompts for a
Bentley IMS login, and logs out only if this script opened the session. Sizes come from
the ProjectWise document records.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $FolderPath,

    [ValidateRange(1, 100000)]
    [int] $Top = 100,

    [switch] $IncludeVersions,

    [string] $OutputPath = (Join-Path $PSScriptRoot ('LargestFiles_{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd_HHmmss')))
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

        $documents = @(Get-PWDocumentsBySearch -FolderPath $FolderPath -GetVersionsToo:$IncludeVersions -ErrorAction Stop |
                Where-Object { -not [string]::IsNullOrWhiteSpace([string] $_.FileName) })
        if ($documents.Count -eq 0) {
            Write-Warning "No documents with files were found below '${FolderPath}'."
            return
        }

        $files = foreach ($document in $documents) {
            [long] $bytes = $document.FileSize -as [long]
            [pscustomobject] @{
                SizeMB     = [math]::Round($bytes / 1MB, 2)
                Bytes      = $bytes
                FileName   = $document.FileName
                Extension  = [System.IO.Path]::GetExtension([string] $document.FileName).ToLowerInvariant()
                Name       = $document.Name
                Version    = $document.Version
                FolderPath = $document.FolderPath
                UpdateDate = $document.FileUpdateDate
            }
        }

        $files | Sort-Object -Property Bytes -Descending | Select-Object -First $Top -Property * -ExcludeProperty Bytes |
            Export-Excel -Path $OutputPath -WorksheetName 'LargestFiles' -AutoSize -FreezeTopRow -BoldTopRow

        foreach ($grouping in @(@{ Sheet = 'ByFolder'; Property = 'FolderPath' }, @{ Sheet = 'ByExtension'; Property = 'Extension' })) {
            $files | Group-Object -Property $grouping.Property | ForEach-Object {
                [pscustomobject] @{
                    $grouping.Property = $_.Name
                    Documents          = $_.Count
                    TotalMB            = [math]::Round((($_.Group | Measure-Object -Property Bytes -Sum).Sum / 1MB), 2)
                }
            } | Sort-Object -Property TotalMB -Descending |
                Export-Excel -Path $OutputPath -WorksheetName $grouping.Sheet -AutoSize -FreezeTopRow -BoldTopRow
        }

        [double] $totalGB = [math]::Round((($files | Measure-Object -Property Bytes -Sum).Sum / 1GB), 2)
        Write-Host "Checked $($documents.Count) file(s), ${totalGB} GB in total. Report saved to '${OutputPath}'."
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
