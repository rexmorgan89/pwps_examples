#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Export-PWFolderToLocal
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Exports the files in a ProjectWise folder tree to a local folder with a metadata workbook.
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
###################################################################################################     ###       #######| |#  pwps_dab and read access to the source folder.
##################################################################################################              #########| |#
#######################################################################################################      ############| |#
#########################################################################################################################| |#
#####################################################################################################################################################################################################################
##
##  Description/Notes:
##  Wraps Export-PWDocuments to copy files out of a folder tree into a local folder and write a metadata workbook.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWFolders, Export-PWDocuments, Undo-PWLogin.
##  Creates local files. Documents are copied out, which ProjectWise records in the audit trail; their content is not changed.
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
Exports a ProjectWise folder tree to a local folder.

.DESCRIPTION
Copies the files of every document in -FolderPath (and its subfolders unless
-ThisFolderOnly is set) to -OutputFolder, and writes a metadata workbook describing them.
Use it for project hand-over, archive copies or offline review packages.

.PARAMETER FolderPath
ProjectWise folder path to export, without the leading 'Documents' folder.

.PARAMETER OutputFolder
Local folder to receive the files. It is created if needed and must be empty.

.PARAMETER ThisFolderOnly
Exports only files directly in -FolderPath.

.PARAMETER IncludeVersions
Also exports earlier versions; their file names get a .verXXXX suffix.

.EXAMPLE
& '.\FileManipulation\Export-PWFolderToLocal.ps1' -FolderPath 'Projects\11431\05_Issued' -OutputFolder 'D:\Handover\11431'
Exports the issued folder tree and its metadata.

.INPUTS
None.

.OUTPUTS
Local files, a metadata workbook and a timestamped .LastRun.log transcript.

.NOTES
Requires pwps_dab. Uses the active ProjectWise session or prompts for a Bentley IMS
login, and logs out only if this script opened the session. Check free disk space before
exporting large trees. The metadata workbook is saved as _metadata.xlsx in -OutputFolder.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $FolderPath,

    [Parameter(Mandatory)]
    [string] $OutputFolder,

    [switch] $ThisFolderOnly,

    [switch] $IncludeVersions
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

        if (Test-Path -LiteralPath $OutputFolder) {
            if (@(Get-ChildItem -LiteralPath $OutputFolder -Force).Count -gt 0) {
                throw "Output folder '${OutputFolder}' is not empty. Choose an empty or new folder."
            }
        }
        else {
            $null = New-Item -ItemType Directory -Path $OutputFolder -ErrorAction Stop
        }

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

        if ($null -eq (Get-PWFolders -FolderPath $FolderPath -JustOne -ErrorAction SilentlyContinue)) {
            throw "Folder '${FolderPath}' was not found."
        }

        [string] $metadataPath = Join-Path $OutputFolder '_metadata.xlsx'
        $null = Export-PWDocuments -ProjectWiseFolder $FolderPath -OutputFolder $OutputFolder -ExportMetadata `
            -ExportMetadataFile $metadataPath -ExportVersions:$IncludeVersions -JustOneFolder:$ThisFolderOnly -ErrorAction Stop

        $files = @(Get-ChildItem -LiteralPath $OutputFolder -Recurse -File | Where-Object { $_.FullName -ne $metadataPath })
        [double] $sizeMB = [math]::Round((($files | Measure-Object -Property Length -Sum).Sum / 1MB), 1)
        Write-Host "Exported $($files.Count) file(s), ${sizeMB} MB, to '${OutputFolder}'."
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
