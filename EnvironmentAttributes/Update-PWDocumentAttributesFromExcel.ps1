#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Update-PWDocumentAttributesFromExcel
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Updates selected environment attributes of documents from an edited Excel workbook.
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
###################################################################################################     ###       #######| |#  pwps_dab, ImportExcel and write access to the documents.
##################################################################################################              #########| |#
#######################################################################################################      ############| |#
#########################################################################################################################| |#
#####################################################################################################################################################################################################################
##
##  Description/Notes:
##  Reads a workbook with a DocumentGUID column, finds each document and writes only the named attribute columns.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWDocumentsByGUIDs, Update-PWDocumentAttributes, Undo-PWLogin.
##  CHANGES document attributes. Supports -WhatIf; writes a results CSV beside the script.
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
Updates document attributes from an Excel workbook.

.DESCRIPTION
Reads a workbook (for example one created by Export-PWDocumentAttributes.ps1 and then
edited), finds each document by its DocumentGUID and updates only the attribute columns
named in -AttributeName. Other columns are ignored, so read-only or system columns in the
workbook are never written.

Blank cells are skipped unless -IncludeBlankValues is set, so an empty cell does not
clear an existing value by accident.

.PARAMETER InputPath
Full path of the Excel workbook. It must contain a DocumentGUID column.

.PARAMETER AttributeName
One or more attribute column names to update.

.PARAMETER WorksheetName
Worksheet to read. Defaults to the first worksheet.

.PARAMETER IncludeBlankValues
Writes blank cells as empty values, clearing the attribute.

.PARAMETER ResultPath
CSV file for the per-document results. Defaults to a timestamped file beside the script.

.EXAMPLE
& '.\EnvironmentAttributes\Update-PWDocumentAttributesFromExcel.ps1' -InputPath 'C:\Temp\Attributes.xlsx' -AttributeName 'TITLE_1', 'SHEET_NUMBER' -WhatIf
Shows which documents would be updated without changing them.

.INPUTS
None.

.OUTPUTS
A results CSV and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab and ImportExcel. Uses the active ProjectWise session or prompts for a
Bentley IMS login, and logs out only if this script opened the session. Attribute names
are case-sensitive in some environments. Export the current values first so that you
can restore them if needed.
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string] $InputPath,

    [Parameter(Mandatory)]
    [string[]] $AttributeName,

    [string] $WorksheetName,

    [switch] $IncludeBlankValues,

    [string] $ResultPath = (Join-Path $PSScriptRoot ('AttributeUpdateResults_{0}.csv' -f (Get-Date -Format 'yyyyMMdd_HHmmss')))
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
        Import-Module ImportExcel -ErrorAction Stop

        $importParameters = @{ Path = $InputPath; ErrorAction = 'Stop' }
        if (-not [string]::IsNullOrWhiteSpace($WorksheetName)) { $importParameters['WorksheetName'] = $WorksheetName }
        $rows = @(Import-Excel @importParameters)
        if ($rows.Count -eq 0) { throw "Workbook '${InputPath}' has no data rows." }

        $headers = @($rows[0].PSObject.Properties.Name)
        if ($headers -notcontains 'DocumentGUID') { throw "Workbook '${InputPath}' has no DocumentGUID column." }
        $missing = @($AttributeName | Where-Object { $headers -notcontains $_ })
        if ($missing.Count -gt 0) { throw "Workbook is missing attribute column(s): $($missing -join ', ')" }

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

        $results = foreach ($row in $rows) {
            [string] $target = ''
            [string] $guid = $row.DocumentGUID
            if ([string]::IsNullOrWhiteSpace($guid)) { continue }

            $attributes = @{}
            foreach ($name in $AttributeName) {
                $value = $row.$name
                if ($IncludeBlankValues -or -not [string]::IsNullOrWhiteSpace([string] $value)) { $attributes[$name] = [string] $value }
            }
            if ($attributes.Count -eq 0) {
                [pscustomobject] @{ DocumentGUID = $guid; Document = ''; Result = 'Skipped'; Message = 'No values to write' }
                continue
            }

            try {
                $document = Get-PWDocumentsByGUIDs -DocumentGUIDs $guid -ErrorAction Stop | Select-Object -First 1
                if ($null -eq $document) { throw 'Document not found.' }

                [string] $target = '{0}\{1}' -f $document.FolderPath, $document.Name
                if (-not $PSCmdlet.ShouldProcess($target, "Set $($attributes.Keys -join ', ')")) {
                    [pscustomobject] @{ DocumentGUID = $guid; Document = $target; Result = 'Skipped'; Message = 'WhatIf or declined' }
                    continue
                }
                $updated = Update-PWDocumentAttributes -InputDocuments $document -Attributes $attributes -ReturnBoolean -ErrorAction Stop
                if ($updated -eq $false) { throw 'Update-PWDocumentAttributes returned false.' }
                [pscustomobject] @{ DocumentGUID = $guid; Document = $target; Result = 'Updated'; Message = '' }
            }
            catch {
                [pscustomobject] @{ DocumentGUID = $guid; Document = $target; Result = 'Failed'; Message = $_.Exception.Message }
            }
        }

        $results = @($results)
        $results | Export-Csv -LiteralPath $ResultPath -NoTypeInformation -Encoding UTF8 -WhatIf:$false
        $results | Group-Object Result | ForEach-Object { Write-Host "$($_.Name): $($_.Count)" }
        Write-Host "Results saved to '${ResultPath}'."
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
