---
description: "Bring PowerShell scripts up to the repository script standard: banner header, comment-based help, begin/process/end, transcript logging."
argument-hint: "Script file(s) or folder to standardize"
agent: "agent"
---
Standardize the target script(s) using [the script standard instructions](../instructions/powershell-script-standard.instructions.md).

1. Run `& .\.github\scripts\Test-ScriptStandard.ps1 -Path <target>` to list the gaps. For a folder, treat the distinct files in its output as the work list.
2. For each file, read the whole script and describe its actual behavior in the header and help. Preserve existing logic, parameters, datasource names, and credential conventions.
3. Rerun the checker on each file you changed. Repeat until it reports no findings for the target.
4. Do not run scripts that log in to ProjectWise or call Bentley services unless asked; parsing and the checker are the validation.

Report the files changed, any header values set to `Unknown`, and any behavior you had to change to fit the lifecycle (for example, moving top-level code into `process`).
