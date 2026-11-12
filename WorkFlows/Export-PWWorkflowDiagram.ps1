#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : Export-PWWorkflowDiagram
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Writes a Mermaid state diagram for each ProjectWise workflow.
####    ###########    #######         ######    #         ####         #    ######          ######    ########    ######| |#  Author                   : GitHub Copilot
####    ##########-    ####               ###                ##         #    ####              ####    ########    ######| |#  Creation Date            : 10-27-2026
#                     #####     ######    ###      #####     ###    #####    ###     ######     ###    ########    ######| |#  Modified By              : GitHub Copilot
#                    #####    #########    ##    ########     ##    #####    ##     ########    ###    ########    ######| |#  Modified Date            : 11-12-2026
####    ##########      ##                 ##    #########    ##    #####    ##                  ##    ########    ######| |#  Script Version           : 1.1
####    ############    ##                 ##    #########    ##    #####    ##                  ##    ########    ######| |#  PowerShell Version       : Windows PowerShell 5.1 x64
####    ############    ##    ###############    #########    ##    #####    ##     ###############    #######     ######| |#  ProjectWise Version      : Compatible ProjectWise Explorer installation required
####    ###########     ###     ######    ###    #########    ##    #####    ###     #######    ###     #####      ######| |#  PWPS Module Version      : Not required
####                   ####               ###    #########    ##    #####    ####              #####               ######| |#  PWPS_DAB Module Version  : Repository reference: 24.0.2
####                  #######           #####    #########    ##    #####    ######           ########        #    ######| |#
##############################################################################################################    #######| |#  Requirements:
###################################################################################################     ###       #######| |#  pwps_dab.
##################################################################################################              #########| |#
#######################################################################################################      ############| |#
#########################################################################################################################| |#
#####################################################################################################################################################################################################################
##
##  Description/Notes:
##  Reads each workflow's state links (previous and next state IDs) and writes one Mermaid stateDiagram-v2 .mmd file per workflow.
##  ProjectWise cmdlets used: Get-PWCurrentDatasource, New-PWLogin, Get-PWWorkflows, Get-PWWorkflowStateLinks, Undo-PWLogin.
##  Read-only; writes local .mmd files only.
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
Draws ProjectWise workflows as Mermaid state diagrams.

.DESCRIPTION
For each workflow (or those named in -WorkflowName), reads its state links and writes a
Mermaid stateDiagram-v2 file to -OutputFolder. Each state points to its next state;
the first state comes from [*] and the last goes to [*].

Open the .mmd files in any Mermaid viewer or paste them into a Markdown mermaid code
block in workflow documentation.

.PARAMETER WorkflowName
Optional workflow names. Every workflow is drawn when empty.

.PARAMETER OutputFolder
Local folder for the .mmd files. Defaults to a timestamped folder beside the script.

.EXAMPLE
& '.\WorkFlows\Export-PWWorkflowDiagram.ps1' -WorkflowName 'Default Workflow'
Writes one diagram for 'Default Workflow'.

.INPUTS
None.

.OUTPUTS
Mermaid (.mmd) files and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires pwps_dab. Uses the active ProjectWise session or prompts for a Bentley IMS
login, and logs out only if this script opened the session. ProjectWise workflows are
linear (each state has one previous and one next state); access rules per state are not
drawn.
#>

[CmdletBinding()]
param(
    [string[]] $WorkflowName,

    [string] $OutputFolder = (Join-Path $PSScriptRoot ('WorkflowDiagrams_{0}' -f (Get-Date -Format 'yyyyMMdd_HHmmss')))
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

        $workflows = @(Get-PWWorkflows -ErrorAction Stop)
        if ($WorkflowName) {
            $workflows = @($workflows | Where-Object { $WorkflowName -contains $_.Name })
            $missing = @($WorkflowName | Where-Object { @($workflows.Name) -notcontains $_ })
            if ($missing.Count -gt 0) { Write-Warning "Workflow(s) not found: $($missing -join ', ')" }
        }
        if ($workflows.Count -eq 0) {
            Write-Warning 'No workflows to draw.'
            return
        }
        $null = New-Item -ItemType Directory -Path $OutputFolder -Force

        foreach ($workflow in $workflows) {
            $states = @(Get-PWWorkflowStateLinks -WorkflowName $workflow.Name -ErrorAction Continue)
            $lines = [System.Collections.Generic.List[string]]::new()
            $lines.Add('---')
            $lines.Add(('title: {0}' -f $workflow.Name))
            $lines.Add('---')
            $lines.Add('stateDiagram-v2')
            foreach ($state in $states) {
                $lines.Add(('    state "{0}" as s{1}' -f ([string] $state.Name).Replace('"', "'"), $state.ID))
            }
            foreach ($state in $states) {
                if ([int] $state.iPrevious -eq 0) { $lines.Add(('    [*] --> s{0}' -f $state.ID)) }
                if ([int] $state.iNext -eq 0) { $lines.Add(('    s{0} --> [*]' -f $state.ID)) }
                else { $lines.Add(('    s{0} --> s{1}' -f $state.ID, $state.iNext)) }
            }

            [string] $safeName = ($workflow.Name -replace '[\\/:*?"<>|]', '_').Trim()
            [string] $path = Join-Path $OutputFolder "${safeName}.mmd"
            Set-Content -LiteralPath $path -Value $lines -Encoding UTF8
            Write-Host ("{0,-40} {1} state(s)" -f $workflow.Name, $states.Count)
        }
        Write-Host "Diagrams saved to '${OutputFolder}'."
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
