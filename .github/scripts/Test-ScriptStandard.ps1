<#
.SYNOPSIS
Checks PowerShell scripts against the repository script standard.

.DESCRIPTION
Parses each .ps1 and .psm1 file under -Path and reports one finding per gap:
  Parse      The file does not parse.
  Header     Bentley banner header is missing, has empty or missing fields, or its '| |#'
             column is not aligned.
  Help       Script-level comment-based help is missing .SYNOPSIS, .DESCRIPTION, .EXAMPLE,
             .NOTES, or a .PARAMETER entry for a declared parameter.
  Lifecycle  Script lacks [CmdletBinding()] param, begin, process, or end blocks.
  Logging    begin does not call Start-Transcript, or end does not call Stop-Transcript.

Modules (.psm1) are checked for the header only. Sources listed in the lifecycle exception
list skip the Lifecycle and Logging rules. No output means every checked file complies.

.PARAMETER Path
Files or folders to check. Defaults to the repository root (two levels above this script).

.PARAMETER ExceptionListPath
Markdown table whose first column holds back-quoted repository-relative paths or wildcards
that are exempt from the lifecycle rules. Defaults to docs/LIFECYCLE_EXCEPTIONS.md when present.

.PARAMETER ExcludeDirectory
Folder names that are skipped anywhere in the tree. The default skips archives, tooling,
build output, and bundled copies of the pwps_dab module.

.PARAMETER ExcludeFile
File name wildcards that are skipped.

.EXAMPLE
& .\.github\scripts\Test-ScriptStandard.ps1 -Path .\ManagedWorkspace\ManagedWorkspaceChart.ps1
Check one script.

.EXAMPLE
& .\.github\scripts\Test-ScriptStandard.ps1 | Group-Object File | Sort-Object Count -Descending
List non-compliant files across the repository, worst first.

.OUTPUTS
PSCustomObject with File, Rule, and Detail for each finding.

.NOTES
Read-only. Parses files with the PowerShell AST; never runs them.
#>

[CmdletBinding()]
param(
    [string[]] $Path,
    [string] $ExceptionListPath,
    [string[]] $ExcludeDirectory = @('z_Archive', '.git', '.github', 'node_modules', 'bin', 'obj', 'pwps_dab'),
    [string[]] $ExcludeFile = @('*.Tests.ps1')
)

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
if (-not $Path) { $Path = @($repoRoot) }
if (-not $ExceptionListPath) {
    $defaultExceptionList = Join-Path $repoRoot 'docs\LIFECYCLE_EXCEPTIONS.md'
    if (Test-Path -LiteralPath $defaultExceptionList -PathType Leaf) { $ExceptionListPath = $defaultExceptionList }
}

$exceptionPatterns = @()
if ($ExceptionListPath) {
    $exceptionPatterns = @(Get-Content -LiteralPath $ExceptionListPath | ForEach-Object {
            if ($_ -match '^\|\s*`([^`]+)`') { $Matches[1].Replace('\', '/') }
        })
}

$headerFields = @('Name', 'Type', 'Purpose', 'Author', 'Creation Date', 'Modified By', 'Modified Date', 'Script Version',
    'PowerShell Version', 'ProjectWise Version', 'PWPS Module Version', 'PWPS_DAB Module Version')

$files = foreach ($item in $Path) {
    $resolved = Get-Item -LiteralPath $item -ErrorAction Stop
    # -Include is ignored with -LiteralPath in Windows PowerShell 5.1, so filter by extension.
    if ($resolved.PSIsContainer) { Get-ChildItem -LiteralPath $resolved.FullName -Recurse -File | Where-Object { $_.Extension -in '.ps1', '.psm1' } }
    else { $resolved }
}

$checked = 0
$findingCount = 0
foreach ($file in @($files | Sort-Object FullName -Unique)) {
    $relativePath = $file.FullName
    if ($relativePath.StartsWith($repoRoot, [StringComparison]::OrdinalIgnoreCase)) {
        $relativePath = $relativePath.Substring($repoRoot.Length).TrimStart('\', '/')
    }
    $relativePath = $relativePath.Replace('\', '/')
    $folders = @($relativePath.Split('/') | Select-Object -SkipLast 1)
    if (@($folders | Where-Object { $ExcludeDirectory -contains $_ }).Count -gt 0) { continue }
    if (@($ExcludeFile | Where-Object { $file.Name -like $_ }).Count -gt 0) { continue }

    $checked++
    $findings = [System.Collections.Generic.List[object]]::new()
    $addFinding = { param($Rule, $Detail) $findings.Add([pscustomobject]@{ File = $relativePath; Rule = $Rule; Detail = $Detail }) }

    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref] $tokens, [ref] $parseErrors)
    if ($parseErrors.Count -gt 0) {
        & $addFinding 'Parse' "$($parseErrors[0].Message) (line $($parseErrors[0].Extent.StartLineNumber))"
        $findings
        $findingCount += $findings.Count
        continue
    }

    #region Header

    $lines = @(Get-Content -LiteralPath $file.FullName)
    $text = $lines -join "`n"
    if ($lines.Count -eq 0 -or $lines[0] -notmatch '^#{20,}\s*$') {
        & $addFinding 'Header' 'File does not start with the Bentley banner header.'
    }
    else {
        foreach ($field in $headerFields) {
            $match = [regex]::Match($text, "(?m)^.*\| \|#\s+$([regex]::Escape($field))\s*:[ \t]*(.*)$")
            if (-not $match.Success) { & $addFinding 'Header' "Missing field '$field'." }
            elseif ([string]::IsNullOrWhiteSpace($match.Groups[1].Value)) { & $addFinding 'Header' "Field '$field' is empty; use 'Unknown' when it cannot be established." }
        }
        $requirements = [regex]::Match($text, '(?m)\| \|#\s+Requirements:[ \t]*\n.*\| \|#[ \t]*(.*)$')
        if (-not $requirements.Success) { & $addFinding 'Header' "Missing 'Requirements:' section." }
        elseif ([string]::IsNullOrWhiteSpace($requirements.Groups[1].Value)) { & $addFinding 'Header' 'Requirements are empty.' }

        $notes = [regex]::Match($text, '(?m)^##\s+Description/Notes:[ \t]*\n((?:##.*\n)*?)#{20,}')
        if (-not $notes.Success) { & $addFinding 'Header' "Missing 'Description/Notes:' section." }
        elseif ($notes.Groups[1].Value -notmatch '(?m)^##\s+\S') { & $addFinding 'Header' 'Description/Notes are empty.' }
        if ($text -notmatch '(?m)^##\s+Disclaimer:') { & $addFinding 'Header' 'Missing Bentley disclaimer.' }

        $bannerColumn = -1
        $misaligned = for ($index = 0; $index -lt $lines.Count; $index++) {
            $column = $lines[$index].IndexOf('| |#')
            if ($column -lt 0) { continue }
            if ($bannerColumn -lt 0) { $bannerColumn = $column }
            elseif ($column -ne $bannerColumn) { $index + 1 }
        }
        $misaligned = @($misaligned)
        if ($misaligned.Count -gt 0) {
            & $addFinding 'Header' "The '| |#' column differs from the banner on line(s) $($misaligned[0])-$($misaligned[-1]) ($($misaligned.Count) line(s)); a second or legacy banner may remain."
        }
    }

    #endregion Header

    $isModule = $file.Extension -eq '.psm1'
    $isException = @($exceptionPatterns | Where-Object { $relativePath -like $_ }).Count -gt 0

    #region Help

    if (-not $isModule) {
        $help = $ast.GetHelpContent()
        if ($null -eq $help) {
            & $addFinding 'Help' 'No script-level comment-based help; leave one blank line between the banner and <# .SYNOPSIS.'
        }
        else {
            if ([string]::IsNullOrWhiteSpace($help.Synopsis)) { & $addFinding 'Help' 'Missing .SYNOPSIS.' }
            if ([string]::IsNullOrWhiteSpace($help.Description)) { & $addFinding 'Help' 'Missing .DESCRIPTION.' }
            if ($help.Examples.Count -eq 0) { & $addFinding 'Help' 'Missing .EXAMPLE.' }
            if ([string]::IsNullOrWhiteSpace($help.Notes)) { & $addFinding 'Help' 'Missing .NOTES.' }
            # GetHelpContent stores parameter keys upper-cased.
            $documented = @($help.Parameters.Keys | ForEach-Object { $_.ToUpperInvariant() })
            $declared = @(if ($ast.ParamBlock) { $ast.ParamBlock.Parameters | Where-Object { $_ } | ForEach-Object { $_.Name.VariablePath.UserPath } })
            foreach ($parameter in $declared) {
                if ($documented -notcontains $parameter.ToUpperInvariant()) { & $addFinding 'Help' "Missing .PARAMETER $parameter." }
            }
        }
    }

    #endregion Help

    #region Lifecycle and logging

    if (-not $isModule -and -not $isException) {
        $hasCmdletBinding = $null -ne $ast.ParamBlock -and @($ast.ParamBlock.Attributes | Where-Object { $_.TypeName.Name -eq 'CmdletBinding' }).Count -gt 0
        if (-not $hasCmdletBinding) { & $addFinding 'Lifecycle' 'Missing [CmdletBinding()] param block.' }
        foreach ($blockName in 'Begin', 'Process', 'End') {
            if ($null -eq $ast."${blockName}Block") { & $addFinding 'Lifecycle' "Missing $($blockName.ToLowerInvariant()) block." }
        }

        $callsCommand = {
            param($Block, [string] $CommandName)
            $null -ne $Block -and $null -ne $Block.Find({ param($node) $node -is [System.Management.Automation.Language.CommandAst] -and $node.GetCommandName() -eq $CommandName }, $true)
        }
        if (-not (& $callsCommand $ast.BeginBlock 'Start-Transcript')) { & $addFinding 'Logging' 'begin block does not call Start-Transcript.' }
        if (-not (& $callsCommand $ast.EndBlock 'Stop-Transcript')) { & $addFinding 'Logging' 'end block does not call Stop-Transcript.' }
    }

    #endregion Lifecycle and logging

    $findings
    $findingCount += $findings.Count
}

Write-Host "Checked $checked file(s); $findingCount finding(s)."
