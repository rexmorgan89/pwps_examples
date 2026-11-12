# iTwin Membership Inventory

`AddRemoveUsersConnectedProjects.ps1` inventories and manages Bentley Infrastructure Cloud/iTwin project memberships through the Access Control API v2. It uses Bentley IMS interactive authentication through `Get-OIDCToken`; never put tokens in scripts, CSV files, or the database.

## Dependency and configuration

Run in **Windows PowerShell 5.1 x64**. Install the PowerShell 5.1-compatible SQLite provider once:

```powershell
Install-Module PSSQLite -Scope CurrentUser
```

Put allowed projects in [iTwinProjects.csv](iTwinProjects.csv). `ProjectGUID` is required and `ProjectName` is optional. The committed template contains headers only.

The database is created automatically at `ConnectedProjects\iTwinMembershipInventory.sqlite`. `Active = 1` represents a current membership and `Active = 0` preserves an absent membership as history. A project is deactivated only after all its membership pages succeed.

Member details include email, given name, surname, organization, and Bentley member ID. Each item in the API `roles` array is recorded separately, including role ID, display name, and assigned date/time, so one user can hold multiple roles in the same project.

## Commands

```powershell
# List iTwins available to the authenticated user.
.\AddRemoveUsersConnectedProjects.ps1 -Mode Discover

# List Project iTwins across the operator's Bentley organization.
.\AddRemoveUsersConnectedProjects.ps1 -Mode Discover -QueryScope OrganizationAdmin

# List only Trial or Inactive Project iTwins.
.\AddRemoveUsersConnectedProjects.ps1 -Mode Discover -ProjectStatus Trial
.\AddRemoveUsersConnectedProjects.ps1 -Mode Discover -ProjectStatus Inactive

# List Active, Trial, and Inactive Project iTwins.
.\AddRemoveUsersConnectedProjects.ps1 -Mode Discover -ProjectStatus All

# List the authenticated user's favorite Project iTwins.
.\AddRemoveUsersConnectedProjects.ps1 -Mode Discover -FavoritesOnly

# Favorite discovery supports the same status filter; name filtering happens locally.
.\AddRemoveUsersConnectedProjects.ps1 -Mode Discover -FavoritesOnly -ProjectStatus All -ProjectNameFilter "Bridge"

# Filter by name and export selected project GUIDs for polling and bulk actions.
.\AddRemoveUsersConnectedProjects.ps1 -Mode Discover -ProjectNameFilter "Bridge" -ExportProjectConfigPath .\iTwinProjects.csv

# Poll configured projects. A failed project does not stop others.
.\AddRemoveUsersConnectedProjects.ps1 -Mode Poll

# Show each active project, role, Bentley member ID, and last observation for an email.
.\AddRemoveUsersConnectedProjects.ps1 -Mode Report -Email user@example.invalid

# Open the local-inventory membership manager. Windows PowerShell 5.1 x64 with STA is required.
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File .\AddRemoveUsersConnectedProjects.ps1 -Mode Gui

# Write GUI action outcomes to a chosen CSV audit path.
.\AddRemoveUsersConnectedProjects.ps1 -Mode Gui -AuditLogPath .\iTwinMembershipManager-audit.csv

# Validate only: no IMS authentication, API call, database creation, or database update.
.\AddRemoveUsersConnectedProjects.ps1 -Mode Bulk -ActionCsvPath .\iTwinMembershipActions.csv -WhatIf

# Assign roles using their display names; RoleId remains available where names are ambiguous.
.\AddRemoveUsersConnectedProjects.ps1 -Mode Bulk -ActionCsvPath .\iTwinMembershipActions.csv
```

## Action CSV

[iTwinMembershipActions.csv](iTwinMembershipActions.csv) is a header-only safe template. Required columns are `ProjectGUID`, `User`, and `Action`; `RoleId` and `RoleDisplayName` are optional. For `AssignRole` and `UnassignRole`, provide either `RoleId` or `RoleDisplayName`. A display name is resolved case-insensitively against the roles available for that project, including enterprise roles; ambiguous names fail safely and require `RoleId`.

Use `AssignRole` to add a role, `UnassignRole` to remove one role while retaining the user and other roles, and `RemoveMember` to remove the user entirely. Legacy `Add` and `Remove` are accepted as aliases for `AssignRole` and `RemoveMember`. Successful actions refresh the project inventory. Each row reports `Success`, `Skipped`, or `Failed`; failures continue to later rows.

## Membership manager GUI

`-Mode Gui` opens a three-pane WPF interface over the local SQLite inventory. It does not poll projects automatically; run `Poll` first so the GUI has current membership observations to display.

Use **Discover projects** to retrieve iTwins available to the signed-in Bentley account. Select one or more unconfigured entries and choose **Add selected to configuration**. The GUI merges those projects into [iTwinProjects.csv](iTwinProjects.csv), retaining existing rows and preventing duplicate project GUIDs. Discovery requires Bentley IMS authentication but does not change project memberships.

Discovery is member-scoped by default. Select **Show organization projects** to query all Project iTwins in the operator's Bentley organization. This requires an Organization Administrator entitlement, such as Account Administrator, Co-Administrator, or CONNECT Services Administrator. The script cannot determine that entitlement from the IMS token alone: it requests the `OrganizationAdmin` query scope and Bentley validates it. When Bentley returns insufficient permissions, the control clears automatically and discovery falls back to projects available to the operator's account; the dialog and CLI both state that fallback explicitly.

Organization discovery uses server-side search and page navigation rather than loading every project into the GUI. Discovery defaults to **Active** projects; choose **Trial** or **Inactive** to return only that state, or **All** to include active, trial, and inactive projects. Each discovery row displays its returned Bentley project status. Choose a page size from 100 through 1000, enter a search term for the project number or name, and select **Search**. Use **Previous** and **Next** to browse the server result pages. Selections are retained when moving between pages, so **Add selected to configuration** adds all selected unconfigured projects, not only those shown on the current page.

The **My Favorites** discovery tab lists Project iTwins favorited by the authenticated user. Favorites are always user-scoped and cannot use OrganizationAdmin scope. The Bentley favorites endpoint has no server-side search, so that tab does not show the organization or search controls; status selection and paging remain available. CLI `-FavoritesOnly` follows the same rules and applies `-ProjectNameFilter` locally after retrieving favorites.

The direct iTwins list API can return at most 10,000 projects per query; its individual pages are limited to 1000 projects. Narrow broad organization results with search before browsing. For organization inventories larger than that limit, use the iTwins `/exports` API outside this tool to create and download an OrganizationAdmin export.

Use **Poll configured projects** to poll every project in the configured CSV and refresh the local inventory. The GUI shows a result dialog with per-project success or failure details, then refreshes the inventory browser. Polling reads the same project CSV as `-Mode Poll`; it does not add projects automatically.

The Browse selector drives the left pane:

- `Projects` lists every project currently configured in [iTwinProjects.csv](iTwinProjects.csv), including projects that have not yet been polled or currently have no active local membership observations. Select **Favorites only** to intersect that local list with the authenticated user's live Bentley favorites. The returned favorite GUIDs are cached while the window is open and reload when **Refresh inventory** is selected. The middle pane lists their related users and organizations, and the right pane lists matching memberships.
- `Users` lists unique email addresses. The middle pane lists related projects and the right pane lists matching memberships.
- `Organizations` lists nonblank observed organizations. The middle pane lists related projects and the right pane lists users on those projects.

All panes support extended multi-selection. The Organization filter applies to Projects and Users. Active observations are shown by default; select **Include history** to include inactive historical rows.

To remove access, select the exact membership rows in the right pane and choose **Remove selected memberships**. The confirmation dialog states that the action is not reversible and shows the unique user, project, and assignment counts. In organization mode, select the organization, related projects, and users in the right pane to remove those users from all selected projects. The GUI deliberately requires those exact right-pane selections rather than deleting from an organization name alone.

To add a user, select one or more projects in Projects mode and choose **Add user to selected projects**. The GUI obtains each project role list live. It uses an exact `Team Member` display-name match automatically; when that role is unavailable or ambiguous, it requires an explicit live role selection for that project. It then uses the Access Control API v2 and refreshes the affected local inventory after each action.

### Stop tracking projects

In Projects mode, choose one or more projects and select **Stop tracking selected projects** to remove them from local polling. The confirmation dialog lists each project, the local membership-record count, and the local role-record count that will be permanently deleted.

This operation removes the selected rows from [iTwinProjects.csv](iTwinProjects.csv) and purges their matching `MembershipObservation` and `MembershipRoleObservation` records from the local SQLite inventory. It does **not** remove project members, alter project roles, or delete projects in Bentley Infrastructure Cloud. A timestamped `.bak` copy of both the project CSV and SQLite database is created beside the original files before the local purge runs.

GUI action results are appended to `iTwinMembershipManager-audit.csv` next to the database by default, or to `-AuditLogPath` when supplied. Audit rows include timestamp, Windows user, project GUID, email, action, role ID, status, and API detail. Tokens are never written to the log.

## Role assignment GUI

[AssignRolesConnectedProjects.ps1](AssignRolesConnectedProjects.ps1) assigns live iTwin roles to users on one iTwin at a time. It works directly against Bentley and does not require a prior poll.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File .\AssignRolesConnectedProjects.ps1
```

1. **Select an iTwin.** Browse **All iTwins** or **My Favorites** with the same status, search, organization-scope, and paging rules as project discovery.
2. **Select users.** The list shows current members with their roles. **Search organization** adds users from your organization (Users API, 3 to 200 characters). **Enter email...** adds any other address; Bentley invites users from other organizations instead of adding them immediately.
3. **Select roles.** Choose one or more roles from the iTwin's live role list. Enterprise roles, defined on the organization's Account iTwin and assignable on every iTwin it owns, are listed with an **(Enterprise)** suffix. Reading them requires `administration_manage_roles` on the Account iTwin or an Organization Administrator role; otherwise a warning is logged and only iTwin roles are listed.

**Assign selected roles to selected users** previews each user's change and asks for confirmation. Existing members keep their current roles and receive only missing ones; non-members are added with all selected roles. Results go to the same audit CSV as the membership manager. When the iTwin is listed in [iTwinProjects.csv](iTwinProjects.csv) and the inventory database exists, that iTwin's local inventory is refreshed afterwards.

## Membership console

[ManageUsersConnectedProjects.ps1](ManageUsersConnectedProjects.ps1) combines the GUIs above in one tabbed window. It accepts `-ProjectConfigPath`, `-AuditLogPath`, and `-RoleCopyAuditLogPath`, and writes a `.LastRun.log` transcript beside the script. New users should start with the [user guide](Help/ManageUsersConnectedProjects-Help.html), which has screenshots of every tab and opens from the **?** button in the console or with F1.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File .\ManageUsersConnectedProjects.ps1
```

| Tab                    | Behavior                                                                                                                                                                                                                                                      |
| ---------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Manage Memberships     | The membership manager GUI. When the inventory does not exist yet, the tab offers **Discover projects** and **Poll configured projects** instead.                                                                                                             |
| Assign User Roles      | The role assignment GUI. iTwins load the first time the tab is opened.                                                                                                                                                                                        |
| Copy Role Definitions  | Copies iTwin or enterprise (account) roles and their permissions to other iTwins and accounts after a confirmed plan. Members and role assignments are not copied. Results go to `iTwinRoleCopy-audit.csv` or `-RoleCopyAuditLogPath`.                        |
| User Membership Report | `Report` mode for one email address, shown in a grid.                                                                                                                                                                                                         |
| Bulk Changes (CSV)     | `Bulk` mode for an [action CSV](#action-csv). **Download CSV Template** saves a template with placeholder rows and selects it; **Preview** validates without calling Bentley; **Apply to Bentley** asks for confirmation and writes results to the audit CSV. |

Role assignments and bulk changes reload the Manage Memberships tab after the local inventory is refreshed. Use `AddRemoveUsersConnectedProjects.ps1 -Mode Poll` for scheduled polling.

The tabs are built from `New-iTwinMembershipManagerView` in [iTwinMembershipManager.psm1](iTwinMembershipManager.psm1), `New-iTwinRoleAssignmentView` in [iTwinRoleAssignment.psm1](iTwinRoleAssignment.psm1), and `New-iTwinRoleCopyView` in [iTwinRoleCopy.psm1](iTwinRoleCopy.psm1); `-Mode Gui` and [AssignRolesConnectedProjects.ps1](AssignRolesConnectedProjects.ps1) host the first two views in their own windows.
