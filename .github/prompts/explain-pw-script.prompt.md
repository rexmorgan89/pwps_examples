---
description: "Explain a ProjectWise PowerShell script section by section, including what it assumes about the datasource, and list risks by impact scope: ProjectWise (datasource, project, folder, document, and their metadata), Session, or External - [System]. Read-only: changes no code."
name: "ProjectWise PowerShell Script Explainer"
argument-hint: "Script file to explain (defaults to the active file)"
agent: "agent"
tools: [read, search]
---

Explain the target script: the attached file, or the active editor file if none is attached. Do not edit any file or run any command.

## Reference resolution and read-only tools

- This is a multi-root workspace. All cmdlet references and custom helper definitions listed below belong to the `pwps_bentley` workspace folder, regardless of which workspace contains the target script or active file.
- Resolve the reference links relative to this prompt file in `.github/prompts/`: `../../docs/` means `pwps_bentley/docs/`, and `../../Modules/` means `pwps_bentley/Modules/`. Do not resolve these paths relative to the target script, terminal working directory, or another workspace root.
- The frontmatter enables the `read` and `search` tool sets, including file reading. Use search tools scoped to `pwps_bentley` to locate exact cmdlet names or function definitions, then use file-reading tools to read the complete relevant entries or definitions. Links identify sources; they do not automatically supply their contents.
- Reading and searching workspace files is allowed. Do not execute terminal commands, import modules, invoke scripts, or modify files to verify behavior.
- If a reference or required tool is unavailable, report **"Unverified: reference unavailable"**, identify the inaccessible source, and briefly state the limitation. Do not claim a search was performed when it was not.
- Only report **"Unverified: not found in references"** after searching the applicable module references and custom helper definitions without finding the command. Do not infer behavior or module ownership from a command name alone.

## Verifying cmdlets

PWPS cmdlets come from any module whose name starts with `pwps`, including `pwps`, `pwps_dab`, `pwps_cloud`, `pwps_pwdm`, and `pwps_masterdoc`. Before describing what one does, verify it:

1. Find the cmdlet's entry in the reference for its module and read it. Search within the file for the cmdlet name, because the files are large:
   - `pwps`: [pwps-Reference.txt](../../docs/pwps-Reference.txt)
   - `pwps_dab`: [pwps_dab-Reference.txt](../../docs/pwps_dab-Reference.txt)
   - `pwps_cloud`: [pwps_cloud-Reference.txt](../../docs/pwps_cloud-Reference.txt)
   - `pwps_masterdoc`: [pwps_masterdoc-Reference.txt](../../docs/pwps_masterdoc-Reference.txt)
   - `pwps_pwdm`: [pwps_pwdm-Reference.txt](../../docs/pwps_pwdm-Reference.txt)

   If you don't know which module a cmdlet belongs to, check `pwps_dab` first, then the others in the order listed.

2. For custom helpers (for example `*-RSH*`), search recursively under [pwps_bentley/Modules](../../Modules/) for the function definition and read it. Read relevant delegated helper definitions when needed to establish side effects and session handling. Identify the owning module from the source rather than guessing from the function name.
3. If neither source covers the cmdlet after searching, write **"Unverified: not found in references"**. If the required sources cannot be accessed, write **"Unverified: reference unavailable"** instead. Do not infer its behavior from the cmdlet name.

## 1. Section walkthrough

Split the script into logical sections: `#region` blocks, `begin`/`process`/`end`, major loops or branches, and `### PART` banners. For each section, give:

- **Lines**: a link to the line range
- **What it does**: describe it in ProjectWise terms (work areas or rich projects, folders, documents, attributes, user lists, groups, access control, views, managed workspace, connected projects, storage areas)
- **Datasource assumptions**: datasource or host names, storage areas, templates, environments, attributes, folder paths, user or group names, UNC paths, and required login state
- **Sessions**: every login, connection, or token this section opens, uses, or closes
- **External systems**: any non-ProjectWise system the section reads from or writes to
- **Cmdlets used**: each cmdlet, with its module, marked as verified or unverified

## 2. Risks

Assess every scope a risk affects. If one risk affects more than one scope, write one row for each scope and give the rows a shared ID (for example R3a, R3b). Verbose output is fine. Do not merge rows to save space.

### ProjectWise scopes (widest to narrowest)

| Impact scope          | Meaning                                                                            | Examples                                                                                                                                         |
| --------------------- | ---------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------ |
| **Datasource**        | Affects many projects or the whole datasource                                      | Datasource views, environments, users, groups, user lists, workflows, storage areas, managed workspace config blocks, project types or templates |
| **Project**           | Affects the structure, security, or configuration of one work area or rich project | Work area creation, work area view, project access control, managed workspace association, connected project link                                |
| **Project Metadata**  | Affects only the properties of one work area or rich project                       | Project number, client, status, project manager, opportunity fields                                                                              |
| **Folder**            | Affects the structure or security of one folder or its subtree                     | Folder creation or removal, folder security, workspace association on a folder                                                                   |
| **Folder Metadata**   | Affects only folder properties or attributes                                       | Folder name, description, environment, folder attributes                                                                                         |
| **Document**          | Affects a document's file or existence                                             | Create, import, check in, replace file, delete                                                                                                   |
| **Document Metadata** | Affects only document properties or attributes                                     | Document name, file name, description, environment attribute values                                                                              |

Classify a change to an object that is shared across projects (for example a user list) as **Datasource**, even if this script changes it while working on one project.

### Session scope

**Session - [System]**: a login, connection, or token that is opened, reused, left open, expired, or exposed. Examples:

- **Session - ProjectWise**: `New-PWLogin` or a helper that logs in without a matching `Undo-PWLogin` on every exit path, or a datasource selection that comes back null
- **Session - SharePoint**: a PnP connection that is never opened, or is left open
- **Session - Bentley Infrastructure Cloud**: a connection client token that expires during a long loop, or is written to the transcript

### External scopes

**External - [System Name]**: a read or write against a system outside ProjectWise. Name the system as the script refers to it. Examples:

- **External - SharePoint** (PnP list reads or writes)
- **External - Bentley Infrastructure Cloud** (iTwin or connected project creation, membership, roles)
- **External - File Share** (UNC or GFS paths)
- **External - Entra ID** or **External - Active Directory**
- **External - Local Machine** (transcripts, shortcuts, local files, interactive prompts)

### Grading risks that cross systems

If one operation can fail or misbehave in more than one system, grade it in each system it affects. Example: the script writes a ProjectWise link to SharePoint after the work area failed to create. That produces:

- **R1a, Project**: the work area is missing or incomplete, but later steps continue against it
- **R1b, External - SharePoint**: `Set-PnPListItem` writes an invalid link, or fails because the PnP session or row ID is missing

Also add a **Session** row when the cause is a missing or invalid session.

### Risk table

Use these columns: ID, Impact scope, Category, Location (line link), Issue, and Why it matters. Order the rows as follows: ProjectWise scopes from widest to narrowest, then Session, then External, with Session and External rows grouped by system. Rows that share an ID stay in their scope's position. Check for:

- **Error handling**: missing `try`/`catch`, `Continue` or `catch` blocks that hide failures, unchecked `$null` returns, and use of variables that may not be set on every branch
- **Session cleanup**: no `Undo-PWLogin` on error paths, no `Stop-Transcript` in a `finally`, and external connections left open
- **Unsafe writes**: create, update, remove, or import calls without `-WhatIf`, `SupportsShouldProcess`, or a confirmation step
- **Hard-coded values**: datasource or host names, storage areas, template names, UNC paths, SharePoint list or field names, and client regexes
- **Scale**: per-item round trips inside loops, `-Slow` lookups, interactive prompts (`Read-Host`, `Out-GridView`) in loops or unattended runs, unbounded queries, and token expiry during long runs
- **Secrets**: tokens or credentials that could reach the transcript or console
- **Typos and dead code**: misspelled variable names, unreachable code, and code commented out by mistake

## 3. Summary

Give the 3 to 5 most important risks, ranked by their widest impact scope. For each, list every scope its rows cover (for example "R1: Project + External - SharePoint"). Do not propose rewritten code. You may describe in one line what a fix would involve.
