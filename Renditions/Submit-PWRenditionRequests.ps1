#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Submit-PWRenditionRequests
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Submits the documents in a folder that match a file name pattern to a rendition profile.
####    ###########    #######         ######    #         ####         #    ######          ######    ########    ######| |#  Author                   : GitHub Copilot
####    ##########-    ####               ###                ##         #    ####              ####    ########    ######| |#  Creation Date            : 10-27-2026
#                     #####     ######    ###      #####     ###    #####    ###     ######     ###    ########    ######| |#  Modified By              : GitHub Copilot
#                    #####    #########    ##    ########     ##    #####    ##     ########    ###    ########    ######| |#  Modified Date            : 11-12-2026
####    ##########      ##                 ##    #########    ##    #####    ##                  ##    ########    ######| |#  Script Version           : 1.1
####    ############    ##                 ##    #########    ##    #####    ##                  ##    ########    ######| |#  PowerShell Version       : Windows PowerShell 5.1 x64
####    ############    ##    ###############    #########    ##    #####    ##     ###############    #######     ######| |#  ProjectWise Version      : Datasource with a working rendition service
####    ###########     ###     ######    ###    #########    ##    #####    ###     #######    ###     #####      ######| |#  PWPS Module Version      : Not required
####                   ####               ###    #########    ##    #####    ####              #####               ######| |#  PWPS_DAB Module Version  : Repository reference: 24.0.2
####                  #######           #####    #########    ##    #####    ######           ########        #    ######| |#
##############################################################################################################    #######| |#  Requirements:
###################################################################################################     ###       #######| |#  pwps_dab and a rendition profile suitable for the selected files.
##################################################################################################              #########| |#
#######################################################################################################      ############| |#
#########################################################################################################################| |#
#####################################################################################################################################################################################################################
##
##  Description/Notes:
##  Checks the rendition profile exists, finds matching documents and submits them.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWRenditionProfiles, Get-PWDocumentsBySearch, New-PWRenditionRequest, Undo-PWLogin.
##  Creates rendition requests, which can create or replace output files. Supports -WhatIf.
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
Submits documents in a folder to a rendition profile.

.DESCRIPTION
Confirms that -ProfileName exists, searches -FolderPath for documents whose file name
matches -FileName, and submits them to New-PWRenditionRequest. By default only the
folder itself is searched; use -IncludeSubFolders to search below it.
Submission does not wait for the renditions to finish.

.PARAMETER FolderPath
ProjectWise folder path, without the leading 'Documents' folder.

.PARAMETER ProfileName
Name of the rendition profile to use.

.PARAMETER FileName
File name pattern using ProjectWise wildcards (%), for example '%.ips'. Defaults to all files.

.PARAMETER IncludeSubFolders
Also submits matching documents in subfolders.

.PARAMETER SinglePlotRequest
Submits all documents as one combined request (-MakeSinglePlotRequest).

.EXAMPLE
& '.\Renditions\Submit-PWRenditionRequests.ps1' -FolderPath 'Projects\11431\03_Design\Plan\FinalPlan\Plot Sets' -FileName '%.ips' -ProfileName 'MNDOT-11X17-IPS-BW' -WhatIf
Lists the documents that would be submitted.

.INPUTS
None.

.OUTPUTS
Output from New-PWRenditionRequest and a timestamped .LastRun.log transcript.

.NOTES
Requires pwps_dab. Uses the active ProjectWise session or prompts for a Bentley IMS
login, and logs out only if this script opened the session. Check job status in the
rendition service or with Get-PWRenditionJobsStatus.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string] $FolderPath,

    [Parameter(Mandatory)]
    [string] $ProfileName,

    [string] $FileName = '%',

    [switch] $IncludeSubFolders,

    [switch] $SinglePlotRequest
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

        if (@(Get-PWRenditionProfiles -ProfileName $ProfileName -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $ProfileName }).Count -eq 0) {
            throw "Rendition profile '${ProfileName}' was not found."
        }

        $documents = @(Get-PWDocumentsBySearch -FolderPath $FolderPath -FileName $FileName -JustThisFolder:(-not $IncludeSubFolders) -ErrorAction Stop)
        if ($documents.Count -eq 0) {
            Write-Warning "No documents matching '${FileName}' were found in '${FolderPath}'."
            return
        }
        $documents | Select-Object FolderPath, Name, FileName | Format-Table -AutoSize | Out-Host

        if ($PSCmdlet.ShouldProcess("$($documents.Count) document(s) in '${FolderPath}'", "Submit to rendition profile '${ProfileName}'")) {
            New-PWRenditionRequest -InputDocuments $documents -ProfileName $ProfileName -MakeSinglePlotRequest:$SinglePlotRequest -ErrorAction Stop
            Write-Host "Submitted $($documents.Count) document(s) to '${ProfileName}'."
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
