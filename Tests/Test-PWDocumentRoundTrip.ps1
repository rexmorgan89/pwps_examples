#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Test-PWDocumentRoundTrip
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Creates, checks out, checks in and deletes a test document to prove core document operations work.
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
###################################################################################################     ###       #######| |#  pwps_dab and create/delete rights in a test folder.
##################################################################################################              #########| |#
#######################################################################################################      ############| |#
#########################################################################################################################| |#
#####################################################################################################################################################################################################################
##
##  Description/Notes:
##  Runs create, find, check-out, check-in and delete on one uniquely named test document and reports pass/fail and timing per step.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, New-PWDocument, Get-PWDocumentsBySearch, CheckOut-PWDocuments, CheckIn-PWDocumentsOrFree, Remove-PWDocuments, Undo-PWLogin.
##  CREATES and DELETES one document in the given test folder. Supports -WhatIf.
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
Tests the full document life cycle in a test folder.

.DESCRIPTION
Creates a small text file, then:
1. Creates a document from it in -TestFolderPath.
2. Finds the new document.
3. Checks it out.
4. Checks it back in.
5. Deletes it (skipped with -KeepDocument).
Each step reports Pass or Fail with its duration. A failing step stops the later steps,
but the delete step still runs so that no test document is left behind.

Use it after ProjectWise or pwps_dab upgrades, or to check storage area and integration
server connectivity.

.PARAMETER TestFolderPath
ProjectWise folder used for the test. Use a dedicated test folder.

.PARAMETER KeepDocument
Leaves the test document in place for inspection.

.EXAMPLE
& '.\Tests\Test-PWDocumentRoundTrip.ps1' -TestFolderPath 'Admin\PowerShell Tests'
Runs the full round trip.

.INPUTS
None.

.OUTPUTS
One result object per step and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab. Uses the active ProjectWise session or prompts for a Bentley IMS
login, and logs out only if this script opened the session. Remove-PWDocuments deletes
permanently unless the datasource uses soft delete. The local temp file is removed.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string] $TestFolderPath,

    [switch] $KeepDocument
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
    [string] $documentName = 'PWPS_RoundTrip_{0}' -f (Get-Date -Format 'yyyyMMdd_HHmmss')
    [string] $localFile = Join-Path $env:TEMP "${documentName}.txt"
    $results = [System.Collections.Generic.List[object]]::new()

    function Invoke-Step {
        param([string] $Name, [scriptblock] $Action)
        $timer = [System.Diagnostics.Stopwatch]::StartNew()
        try {
            $output = & $Action
            $results.Add([pscustomobject] @{ Step = $Name; Result = 'Pass'; Seconds = [math]::Round($timer.Elapsed.TotalSeconds, 2); Message = '' })
            $output
        }
        catch {
            $results.Add([pscustomobject] @{ Step = $Name; Result = 'Fail'; Seconds = [math]::Round($timer.Elapsed.TotalSeconds, 2); Message = $_.Exception.Message })
            throw
        }
    }
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

        if (-not $PSCmdlet.ShouldProcess("${TestFolderPath}\${documentName}", 'Create, check out, check in and delete test document')) {
            return
        }

        Set-Content -LiteralPath $localFile -Value "Round-trip test created $(Get-Date -Format u) by ${env:USERNAME}." -WhatIf:$false
        $document = $null
        try {
            $null = Invoke-Step 'Create document' {
                $null = New-PWDocument -FolderPath $TestFolderPath -FilePath $localFile -DocumentName $documentName -Description 'PowerShell round-trip test' -ErrorAction Stop
            }
            $document = Invoke-Step 'Find document' {
                $found = @(Get-PWDocumentsBySearch -FolderPath $TestFolderPath -DocumentName $documentName -JustThisFolder -ErrorAction Stop)
                if ($found.Count -ne 1) { throw "Expected 1 document named '${documentName}', found $($found.Count)." }
                $found[0]
            }
            $null = Invoke-Step 'Check out' { $null = CheckOut-PWDocuments -InputDocument $document -ErrorAction Stop }
            $null = Invoke-Step 'Check in' { $null = CheckIn-PWDocumentsOrFree -InputDocument $document -ErrorAction Stop }
        }
        catch {
            Write-Warning "Round trip stopped: $($_.Exception.Message)"
        }
        finally {
            if ($null -ne $document -and -not $KeepDocument) {
                try {
                    $null = Invoke-Step 'Delete document' {
                        $null = CheckIn-PWDocumentsOrFree -InputDocument $document -Free -ErrorAction SilentlyContinue
                        $null = Remove-PWDocuments -InputDocument $document -ErrorAction Stop
                    }
                }
                catch {
                    Write-Warning "Test document '${TestFolderPath}\${documentName}' could not be deleted; remove it manually."
                }
            }
            Remove-Item -LiteralPath $localFile -Force -ErrorAction SilentlyContinue -WhatIf:$false
        }

        $results | Format-Table -AutoSize | Out-Host
        $failed = @($results | Where-Object Result -eq 'Fail').Count
        if ($failed -eq 0) { Write-Host 'Round trip passed.' -ForegroundColor Green }
        else { Write-Host "Round trip failed at ${failed} step(s)." -ForegroundColor Red }
        $results
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
