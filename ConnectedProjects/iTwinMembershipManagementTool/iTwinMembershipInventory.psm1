Set-StrictMode -Version Latest

$script:AccessControlBaseUri = 'https://api.bentley.com/accesscontrol/itwins'
$script:iTwinBaseUri = 'https://api.bentley.com/itwins'
$script:UsersBaseUri = 'https://api.bentley.com/users'
$script:ImsScopes = 'itwin-platform itwins:read itwins:modify rbac-user:external-client users:read email openid profile'

function Import-iTwinSqliteDependency {
    if (-not (Get-Module -ListAvailable -Name PSSQLite)) {
        throw 'PSSQLite is required. Install it with: Install-Module PSSQLite -Scope CurrentUser'
    }
    Import-Module -Name PSSQLite -ErrorAction Stop
}

function Invoke-iTwinSqliteQuery {
    param([string]$DatabasePath, [string]$Query, [hashtable]$SqlParameters = @{})
    Invoke-SqliteQuery -DataSource $DatabasePath -Query $Query -SqlParameters $SqlParameters -ErrorAction Stop
}

function Initialize-iTwinMembershipDatabase {
    param([string]$DatabasePath)
    Import-iTwinSqliteDependency
    Invoke-iTwinSqliteQuery $DatabasePath 'CREATE TABLE IF NOT EXISTS MembershipObservation (ObservationId INTEGER PRIMARY KEY AUTOINCREMENT, ProjectGuid TEXT NOT NULL, ProjectName TEXT, Email TEXT NOT NULL COLLATE NOCASE, BentleyMemberId TEXT NOT NULL, GivenName TEXT, Surname TEXT, Organization TEXT, FirstObservedUtc TEXT NOT NULL, LastObservedUtc TEXT NOT NULL, Active INTEGER NOT NULL CHECK (Active IN (0, 1)));'
    $columns = @(Invoke-iTwinSqliteQuery $DatabasePath 'PRAGMA table_info(MembershipObservation);' | ForEach-Object { $_.name })
    foreach ($column in @('GivenName', 'Surname', 'Organization')) {
        if ($columns -notcontains $column) { Invoke-iTwinSqliteQuery $DatabasePath "ALTER TABLE MembershipObservation ADD COLUMN $column TEXT;" | Out-Null }
    }
    Invoke-iTwinSqliteQuery $DatabasePath 'CREATE UNIQUE INDEX IF NOT EXISTS UX_MembershipObservation_Active ON MembershipObservation(ProjectGuid, Email, BentleyMemberId) WHERE Active = 1;'
    Invoke-iTwinSqliteQuery $DatabasePath 'CREATE INDEX IF NOT EXISTS IX_MembershipObservation_ActiveEmail ON MembershipObservation(Email, Active);'
    Invoke-iTwinSqliteQuery $DatabasePath 'CREATE INDEX IF NOT EXISTS IX_MembershipObservation_ActiveProjectOrganization ON MembershipObservation(ProjectGuid, Organization) WHERE Active = 1;'
    Invoke-iTwinSqliteQuery $DatabasePath 'CREATE INDEX IF NOT EXISTS IX_MembershipObservation_ActiveOrganizationEmail ON MembershipObservation(Organization, Email) WHERE Active = 1;'
    Invoke-iTwinSqliteQuery $DatabasePath 'CREATE TABLE IF NOT EXISTS MembershipRoleObservation (RoleObservationId INTEGER PRIMARY KEY AUTOINCREMENT, ProjectGuid TEXT NOT NULL, BentleyMemberId TEXT NOT NULL, RoleId TEXT NOT NULL, RoleName TEXT, AssignedDateTimeUtc TEXT, FirstObservedUtc TEXT NOT NULL, LastObservedUtc TEXT NOT NULL, Active INTEGER NOT NULL CHECK (Active IN (0, 1)));' | Out-Null
    Invoke-iTwinSqliteQuery $DatabasePath 'CREATE UNIQUE INDEX IF NOT EXISTS UX_MembershipRoleObservation_Active ON MembershipRoleObservation(ProjectGuid, BentleyMemberId, RoleId) WHERE Active = 1;' | Out-Null
    Invoke-iTwinSqliteQuery $DatabasePath 'CREATE INDEX IF NOT EXISTS IX_MembershipRoleObservation_ActiveMember ON MembershipRoleObservation(ProjectGuid, BentleyMemberId, Active);' | Out-Null
}

function Get-iTwinAccessToken {
    try {
        $token = Get-OIDCToken 'powershelltokengen3' -Authority 'https://ims.bentley.com' -Scopes $script:ImsScopes -UseCacheIfAvailable -UnattendedMode -ErrorAction Stop

    }
    catch {
        try { $token = Get-OIDCToken 'powershelltokengen3' -Authority 'https://ims.bentley.com' -Scopes $script:ImsScopes -UseCacheIfAvailable -ErrorAction Stop } catch { throw "Failed to acquire iTwin access token: $_" }
    }
    if ([string]::IsNullOrWhiteSpace($token.access_token) -or [string]::IsNullOrWhiteSpace($token.token_type)) { throw 'Bentley IMS authentication returned no usable access token.' }
    @{ Accept = 'application/vnd.bentley.itwin-platform.v2+json'; Authorization = "$($token.token_type) $($token.access_token)" }
}

function Test-iTwinGuid {
    param([string]$Value)
    $parsedGuid = [guid]::Empty
    [guid]::TryParse($Value, [ref]$parsedGuid)
}

function Get-iTwinProjectConfig {
    param([string]$ProjectConfigPath)
    if (-not (Test-Path -LiteralPath $ProjectConfigPath -PathType Leaf)) { throw "Project configuration CSV was not found: '$ProjectConfigPath'." }
    $projects = @(Import-Csv -LiteralPath $ProjectConfigPath)
    if ($projects.Count -eq 0) { throw 'Project configuration CSV contains no project rows.' }
    if ($projects[0].PSObject.Properties.Name -notcontains 'ProjectGUID') { throw 'Project configuration CSV requires ProjectGUID; ProjectName is optional.' }
    $seen = @{}
    foreach ($project in $projects) {
        $guid = [string]$project.ProjectGUID
        if (-not (Test-iTwinGuid $guid)) { throw "Invalid ProjectGUID '$guid' in project configuration CSV." }
        if ($seen.ContainsKey($guid)) { throw "Duplicate ProjectGUID '$guid' in project configuration CSV." }
        $seen[$guid] = $true
    }
    $projects
}

function Get-iTwinMembers {
    param([string]$ProjectGuid, [hashtable]$Headers)
    $members = @(); $nextUri = "${script:AccessControlBaseUri}/${ProjectGuid}/members/users?`$top=100&`$skip=0"; $pageCount = 0
    while ($nextUri) {
        $pageCount++
        if ($pageCount -gt 1000) { throw "Membership pagination exceeded 1000 pages for project '$ProjectGuid'." }
        $response = Invoke-RestMethod -Method Get -Uri $nextUri -Headers $Headers -ErrorAction Stop
        $pageMembers = @($response.members); $members += $pageMembers; $nextUri = $null
        foreach ($property in @('nextLink', 'next', 'nextPageLink')) {
            if (($response.PSObject.Properties.Name -contains $property) -and -not [string]::IsNullOrWhiteSpace([string]$response.$property)) { $nextUri = [string]$response.$property; break }
        }
        if (-not $nextUri) {
            $linksProperty = $response.PSObject.Properties['_links']

            if ($null -ne $linksProperty) {
                $nextProperty = $linksProperty.Value.PSObject.Properties['next']

                if ($null -ne $nextProperty -and -not [string]::IsNullOrWhiteSpace([string]$nextProperty.Value.href)) {
                    $nextUri = [string]$nextProperty.Value.href
                }
            }
        }
        if ($nextUri -and $nextUri -notmatch '^https?://') { $nextUri = "https://api.bentley.com$nextUri" }
        if (-not $nextUri -and $pageMembers.Count -eq 100) { $nextUri = "${script:AccessControlBaseUri}/${ProjectGuid}/members/users?`$top=100&`$skip=$($members.Count)" }
    }
    return , $members
}

function Get-iTwinRoles {
    param([string]$ProjectGuid, [hashtable]$Headers)
    $response = Invoke-RestMethod -Method Get -Uri "${script:AccessControlBaseUri}/${ProjectGuid}/roles" -Headers $Headers -ErrorAction Stop
    @(if ($response.PSObject.Properties.Name -contains 'roles') { $response.roles } else { $response })
}

function Get-iTwinAccountId {
    param([string]$ProjectGuid, [hashtable]$Headers)
    $accountHeaders = $Headers.Clone()
    $accountHeaders.Accept = 'application/vnd.bentley.itwin-platform.v1+json'
    $response = Invoke-RestMethod -Method Get -Uri "${script:iTwinBaseUri}/${ProjectGuid}/account" -Headers $accountHeaders -ErrorAction Stop
    [string]$response.iTwin.id
}

function Get-iTwinAssignableRoles {
    param([string]$ProjectGuid, [hashtable]$Headers)
    $roles = @(Get-iTwinRoles $ProjectGuid $Headers | Where-Object { $null -ne $_ } | ForEach-Object {
            [pscustomobject]@{ id = Get-iTwinOptionalValue $_ 'id'; displayName = Get-iTwinOptionalValue $_ 'displayName'; Source = 'iTwin' }
        })
    try {
        $accountId = Get-iTwinAccountId $ProjectGuid $Headers
        if (-not [string]::IsNullOrWhiteSpace($accountId) -and $accountId -ne $ProjectGuid) {
            # Enterprise roles are defined on the owning Account iTwin and are assignable on every iTwin under it.
            $roles += @(Get-iTwinRoles $accountId $Headers | Where-Object { $null -ne $_ } | ForEach-Object {
                    [pscustomobject]@{ id = Get-iTwinOptionalValue $_ 'id'; displayName = Get-iTwinOptionalValue $_ 'displayName'; Source = 'Account' }
                })
        }
    }
    catch { Write-Warning "Enterprise (account) roles for iTwin '$ProjectGuid' are unavailable: $($_.Exception.Message)" }
    @($roles | Where-Object { $_.id } | Sort-Object id -Unique)
}

function Get-iTwinAvailableProjects {
    param(
        [string]$DisplayName,
        [ValidateSet('MemberOfiTwin', 'OrganizationAdmin')][string]$QueryScope = 'MemberOfiTwin',
        [ValidateSet('Active', 'Trial', 'Inactive', 'All')][string]$ProjectStatus = 'Active',
        [int]$MaxReturn = 10000
    )
    if ($MaxReturn -lt 1 -or $MaxReturn -gt 10000) { throw 'MaxReturn must be between 1 and 10000.' }
    $projects = @()
    $nextUri = $null
    while ($true) {
        $page = Get-iTwinAvailableProjectPage -DisplayName $DisplayName -QueryScope $QueryScope -ProjectStatus $ProjectStatus -MaxReturn $MaxReturn -NextUri $nextUri
        $projects += @($page.Projects)
        if ([string]::IsNullOrWhiteSpace([string]$page.NextUri)) { break }
        $nextUri = [string]$page.NextUri
    }
    $projects
}

function Get-iTwinAvailableProjectPage {
    param(
        [string]$DisplayName,
        [ValidateSet('MemberOfiTwin', 'OrganizationAdmin')][string]$QueryScope = 'MemberOfiTwin',
        [ValidateSet('Active', 'Trial', 'Inactive', 'All')][string]$ProjectStatus = 'Active',
        [ValidateRange(1, 1000)][int]$PageSize = 250,
        [ValidateRange(1, 10000)][int]$MaxReturn = 10000,
        [string]$NextUri,
        [ValidateSet('Project', 'Account')][string]$SubClass = 'Project'
    )
    $headers = Get-iTwinAccessToken
    $headers.Accept = 'application/vnd.bentley.itwin-platform.v1+json'
    $headers['X-Max-Return'] = [string]$MaxReturn
    $headers['X-iTwin-Query-Scope'] = $QueryScope
    if ([string]::IsNullOrWhiteSpace($NextUri)) {
        $query = @("subClass=$SubClass", "`$top=$([Math]::Min($PageSize, $MaxReturn))")
        if ($ProjectStatus -eq 'All') { $query += 'includeInactive=true' } else { $query += "status=$ProjectStatus" }
        if (-not [string]::IsNullOrWhiteSpace($DisplayName)) {
            $query += "`$search=$([Uri]::EscapeDataString($DisplayName))"
        }
        $requestUri = "${script:iTwinBaseUri}/?$($query -join '&')"
    }
    else {
        $requestUri = $NextUri
        if ($requestUri -notmatch '^https?://') { $requestUri = "https://api.bentley.com$requestUri" }
    }
    $response = Invoke-RestMethod -Method Get -Uri $requestUri -Headers $headers -ErrorAction Stop
    $pageProjects = @(if ($response.PSObject.Properties.Name -contains 'iTwins') { $response.iTwins } else { $response })
    $projects = @($pageProjects | ForEach-Object {
            [pscustomobject]@{ ProjectGUID = [string]$_.id; ProjectName = [string]$_.displayName; Status = [string]$_.status }
        })
    $nextPageUri = $null
    $linksProperty = $response.PSObject.Properties['_links']

    if ($null -ne $linksProperty) {
        $nextProperty = $linksProperty.Value.PSObject.Properties['next']

        if ($null -ne $nextProperty) {
            $hrefProperty = $nextProperty.Value.PSObject.Properties['href']
            if ($null -ne $hrefProperty) { $nextPageUri = [string]$hrefProperty.Value }
        }
    }
    [pscustomobject]@{ Projects = $projects; NextUri = $nextPageUri; QueryScope = $QueryScope; ProjectStatus = $ProjectStatus; PageSize = $PageSize; MaxReturn = $MaxReturn }
}

function Get-iTwinPrimaryAccount {
    $headers = Get-iTwinAccessToken
    $headers.Accept = 'application/vnd.bentley.itwin-platform.v1+json'
    $response = Invoke-RestMethod -Method Get -Uri "${script:iTwinBaseUri}/myprimaryaccount" -Headers $headers -ErrorAction Stop
    $account = $response.iTwin
    [pscustomobject]@{ ProjectGUID = Get-iTwinOptionalValue $account 'id'; ProjectName = Get-iTwinOptionalValue $account 'displayName'; Status = Get-iTwinOptionalValue $account 'status' }
}

function Get-iTwinFavoriteProjects {
    param([ValidateSet('Active', 'Trial', 'Inactive', 'All')][string]$ProjectStatus = 'Active')
    $projects = @()
    $nextUri = $null
    while ($true) {
        $page = Get-iTwinFavoriteProjectPage -ProjectStatus $ProjectStatus -NextUri $nextUri
        $projects += @($page.Projects)
        if ([string]::IsNullOrWhiteSpace([string]$page.NextUri)) { break }
        $nextUri = [string]$page.NextUri
    }
    $projects
}

function Get-iTwinFavoriteProjectPage {
    param(
        [ValidateSet('Active', 'Trial', 'Inactive', 'All')][string]$ProjectStatus = 'Active',
        [ValidateRange(1, 1000)][int]$PageSize = 250,
        [string]$NextUri
    )
    $headers = Get-iTwinAccessToken
    $headers.Accept = 'application/vnd.bentley.itwin-platform.v1+json'
    if ([string]::IsNullOrWhiteSpace($NextUri)) {
        $query = @('subClass=Project', "`$top=$PageSize")
        if ($ProjectStatus -eq 'All') { $query += 'includeInactive=true' } else { $query += "status=$ProjectStatus" }
        $requestUri = "${script:iTwinBaseUri}/favorites?$($query -join '&')"
    }
    else {
        $requestUri = $NextUri
        if ($requestUri -notmatch '^https?://') { $requestUri = "https://api.bentley.com$requestUri" }
    }
    $response = Invoke-RestMethod -Method Get -Uri $requestUri -Headers $headers -ErrorAction Stop
    $pageProjects = @(if ($response.PSObject.Properties.Name -contains 'iTwins') { $response.iTwins } else { $response })
    $projects = @($pageProjects | ForEach-Object {
            [pscustomobject]@{ ProjectGUID = [string]$_.id; ProjectName = [string]$_.displayName; Status = [string]$_.status }
        })
    $nextPageUri = $null
    $linksProperty = $response.PSObject.Properties['_links']
    if ($null -ne $linksProperty) {
        $nextProperty = $linksProperty.Value.PSObject.Properties['next']
        if ($null -ne $nextProperty) {
            $hrefProperty = $nextProperty.Value.PSObject.Properties['href']
            if ($null -ne $hrefProperty) { $nextPageUri = [string]$hrefProperty.Value }
        }
    }
    [pscustomobject]@{ Projects = $projects; NextUri = $nextPageUri; ProjectStatus = $ProjectStatus; PageSize = $PageSize }
}

function Get-iTwinProjectRoles {
    param([string]$ProjectGuid)
    if (-not (Test-iTwinGuid $ProjectGuid)) { throw "Invalid ProjectGUID '$ProjectGuid'." }
    $headers = Get-iTwinAccessToken
    @(Get-iTwinAssignableRoles $ProjectGuid $headers | ForEach-Object {
            [pscustomobject]@{
                ProjectGUID     = $ProjectGuid
                RoleId          = $_.id
                RoleDisplayName = $_.displayName
                RoleSource      = $_.Source
                Label           = if ($_.Source -eq 'Account') { "$($_.displayName) (Enterprise)" } else { $_.displayName }
            }
        } | Sort-Object RoleDisplayName, RoleSource, RoleId)
}

function Get-iTwinRolePermissions {
    param($Role)
    $property = $Role.PSObject.Properties['permissions']
    if ($null -eq $property -or $null -eq $property.Value) { return , [string[]]@() }
    , [string[]]@($property.Value | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { [string]$_ } | Sort-Object -Unique)
}

function Get-iTwinErrorText {
    param($ErrorRecord)
    $text = $ErrorRecord.Exception.Message
    if ($null -ne $ErrorRecord.ErrorDetails -and -not [string]::IsNullOrWhiteSpace($ErrorRecord.ErrorDetails.Message)) { $text = "$text $($ErrorRecord.ErrorDetails.Message)" }
    $text
}

function Invoke-iTwinJsonRequest {
    param([string]$Method, [string]$Uri, [hashtable]$Headers, [hashtable]$Body)
    # Windows PowerShell 5.1 sends string bodies as ISO-8859-1, so role names are sent as UTF-8 bytes.
    $bytes = [Text.Encoding]::UTF8.GetBytes(($Body | ConvertTo-Json -Depth 4 -Compress))
    Invoke-RestMethod -Method $Method -Uri $Uri -Headers $Headers -ContentType 'application/json; charset=utf-8' -Body $bytes -ErrorAction Stop
}

function Get-iTwinRoleDefinitions {
    param([Parameter(Mandatory = $true)][string]$ProjectGuid)
    if (-not (Test-iTwinGuid $ProjectGuid)) { throw "Invalid ProjectGUID '$ProjectGuid'." }
    $headers = Get-iTwinAccessToken
    @(Get-iTwinRoles $ProjectGuid $headers | Where-Object { $null -ne $_ } | ForEach-Object {
            [pscustomobject]@{
                ProjectGUID = $ProjectGuid
                RoleId      = Get-iTwinOptionalValue $_ 'id'
                DisplayName = Get-iTwinOptionalValue $_ 'displayName'
                Description = Get-iTwinOptionalValue $_ 'description'
                Type        = Get-iTwinOptionalValue $_ 'type'
                Permissions = Get-iTwinRolePermissions $_
            }
        } | Where-Object { $_.RoleId -and $_.DisplayName } | Sort-Object DisplayName, RoleId)
}

function Copy-iTwinRoleDefinitions {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory = $true)][string]$SourceProjectGuid,
        [Parameter(Mandatory = $true)][object[]]$Role,
        [Parameter(Mandatory = $true)][string[]]$DestinationProjectGuid,
        [ValidateSet('Skip', 'Merge', 'Replace')][string]$ExistingRoleAction = 'Skip',
        # Overwrites this destination role instead of matching roles by display name.
        [string]$DestinationRoleId,
        # Copies the single source role under this display name.
        [string]$NewRoleName
    )
    if (-not (Test-iTwinGuid $SourceProjectGuid)) { throw "Invalid source ProjectGUID '$SourceProjectGuid'." }
    $destinations = @($DestinationProjectGuid | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ } | Select-Object -Unique)
    if ($destinations.Count -eq 0) { throw 'At least one destination iTwin is required.' }
    foreach ($guid in $destinations) { if (-not (Test-iTwinGuid $guid)) { throw "Invalid destination ProjectGUID '$guid'." } }
    $roles = @($Role | Where-Object { $null -ne $_ -and -not [string]::IsNullOrWhiteSpace([string]$_.DisplayName) })
    if ($roles.Count -eq 0) { throw 'At least one source role is required.' }
    $duplicateNames = @($roles | Group-Object { ([string]$_.DisplayName).ToLowerInvariant() } | Where-Object Count -gt 1 | ForEach-Object { $_.Group[0].DisplayName })
    if ($duplicateNames.Count -gt 0) { throw "Selected source roles share display names: $($duplicateNames -join ', ')." }
    # The loop below reuses $destinationRoleId, which is the same variable as this parameter.
    $targetRoleId = ([string]$DestinationRoleId).Trim()
    $overwrite = -not [string]::IsNullOrWhiteSpace($targetRoleId)
    if ($overwrite) {
        if (-not (Test-iTwinGuid $targetRoleId)) { throw "Invalid destination role ID '$targetRoleId'." }
        if ($roles.Count -ne 1 -or $destinations.Count -ne 1) { throw 'Overwriting a destination role requires exactly one source role and one destination.' }
        if ($ExistingRoleAction -eq 'Skip') { throw 'Overwriting a destination role requires ExistingRoleAction Merge or Replace.' }
    }
    $newName = ([string]$NewRoleName).Trim()
    $createNew = -not [string]::IsNullOrWhiteSpace($newName)
    $sourceRoleName = ''
    if ($createNew) {
        if ($overwrite) { throw 'NewRoleName and DestinationRoleId cannot be used together.' }
        if ($roles.Count -ne 1) { throw 'Creating a new role requires exactly one source role.' }
        $sourceRoleName = [string]$roles[0].DisplayName
        $renamedRole = $roles[0] | Select-Object -Property *
        $renamedRole.DisplayName = $newName
        $roles = @($renamedRole)
    }

    $headers = Get-iTwinAccessToken
    # Reads $destinationGuid from the calling loop scope.
    $newResult = {
        param($SourceRole, [string]$Action, [string]$Status, [string]$Detail, [string]$DestinationRoleId)
        [pscustomobject]@{ SourceProjectGUID = $SourceProjectGuid; DestinationProjectGUID = $destinationGuid; RoleName = [string]$SourceRole.DisplayName; SourceRoleId = [string]$SourceRole.RoleId; DestinationRoleId = $DestinationRoleId; Action = $Action; Status = $Status; Detail = $Detail }
    }
    foreach ($destinationGuid in $destinations) {
        if ($destinationGuid -eq $SourceProjectGuid -and -not $overwrite -and -not $createNew) {
            foreach ($sourceRole in $roles) { & $newResult $sourceRole 'None' 'Skipped' 'Destination is the source iTwin.' '' }
            continue
        }
        try { $existingRoles = @(Get-iTwinRoles $destinationGuid $headers | Where-Object { $null -ne $_ }) }
        catch {
            $detail = "Destination roles could not be read: $(Get-iTwinErrorText $_)"
            foreach ($sourceRole in $roles) { & $newResult $sourceRole 'None' 'Failed' $detail '' }
            continue
        }

        foreach ($sourceRole in $roles) {
            $sourcePermissions = @($sourceRole.Permissions | ForEach-Object { [string]$_ })
            $action = 'None'
            $destinationRoleId = ''
            try {
                if ($overwrite) {
                    $roleMatches = @($existingRoles | Where-Object { (Get-iTwinOptionalValue $_ 'id') -eq $targetRoleId })
                    if ($roleMatches.Count -eq 0) { throw "Destination role '$targetRoleId' was not found." }
                }
                else {
                    $roleMatches = @($existingRoles | Where-Object { (Get-iTwinOptionalValue $_ 'displayName') -ieq $sourceRole.DisplayName })
                    if ($roleMatches.Count -gt 1) { throw "Destination has $($roleMatches.Count) roles named '$($sourceRole.DisplayName)'." }
                }
                if ($roleMatches.Count -gt 0 -and $destinationGuid -eq $SourceProjectGuid -and (Get-iTwinOptionalValue $roleMatches[0] 'id') -eq [string]$sourceRole.RoleId) {
                    & $newResult $sourceRole 'None' 'Skipped' 'The destination role is the source role.' ([string]$sourceRole.RoleId)
                    continue
                }
                $fromText = if ($createNew) { " from '$sourceRoleName'" } else { '' }
                if ($roleMatches.Count -eq 0) {
                    $action = 'Create'
                    if (-not $PSCmdlet.ShouldProcess("Role '$($sourceRole.DisplayName)' on iTwin $destinationGuid", 'Create role')) {
                        & $newResult $sourceRole $action 'Planned' "Create role$fromText with $($sourcePermissions.Count) permission(s)." ''
                        continue
                    }
                    $description = if (-not [string]::IsNullOrWhiteSpace([string]$sourceRole.Description)) { [string]$sourceRole.Description }
                    elseif ($createNew) { "Created from role '$sourceRoleName' on iTwin $SourceProjectGuid." }
                    else { "Copied from iTwin $SourceProjectGuid." }
                    $response = Invoke-iTwinJsonRequest Post "${script:AccessControlBaseUri}/${destinationGuid}/roles" $headers @{ displayName = [string]$sourceRole.DisplayName; description = $description }
                    $destinationRoleId = [string]$response.role.id
                    if ($sourcePermissions.Count -gt 0) {
                        try { Invoke-iTwinJsonRequest Patch "${script:AccessControlBaseUri}/${destinationGuid}/roles/${destinationRoleId}" $headers @{ permissions = $sourcePermissions } | Out-Null }
                        catch { throw "Role was created without permissions: $(Get-iTwinErrorText $_)" }
                    }
                    & $newResult $sourceRole $action 'Success' "Created role$fromText with $($sourcePermissions.Count) permission(s)." $destinationRoleId
                    continue
                }

                $destinationRole = $roleMatches[0]
                $destinationRoleId = Get-iTwinOptionalValue $destinationRole 'id'
                $destinationRoleName = Get-iTwinOptionalValue $destinationRole 'displayName'
                $roleText = if ($destinationRoleName -ieq $sourceRole.DisplayName) { 'existing role' } else { "existing role '$destinationRoleName'" }
                $currentPermissions = Get-iTwinRolePermissions $destinationRole
                $missing = @($sourcePermissions | Where-Object { $currentPermissions -notcontains $_ })
                $extra = @($currentPermissions | Where-Object { $sourcePermissions -notcontains $_ })
                if ($ExistingRoleAction -eq 'Skip') {
                    & $newResult $sourceRole 'None' 'Skipped' "Role already exists with $($currentPermissions.Count) permission(s); $($missing.Count) missing source permission(s) were not added." $destinationRoleId
                    continue
                }
                if ($ExistingRoleAction -eq 'Merge') {
                    $action = 'AddPermissions'
                    $desired = @($currentPermissions) + $missing
                    $change = "add $($missing.Count) missing permission(s)"
                    $unchanged = $missing.Count -eq 0
                }
                else {
                    $action = 'ReplacePermissions'
                    $desired = $sourcePermissions
                    $change = "add $($missing.Count) and remove $($extra.Count) permission(s)"
                    $unchanged = $missing.Count -eq 0 -and $extra.Count -eq 0
                }
                if ($unchanged) {
                    & $newResult $sourceRole $action 'Skipped' "The $roleText already has the source permissions." $destinationRoleId
                    continue
                }
                if (-not $PSCmdlet.ShouldProcess("Role '$destinationRoleName' on iTwin $destinationGuid", "Update ${roleText}: $change")) {
                    & $newResult $sourceRole $action 'Planned' "Update ${roleText}: $change." $destinationRoleId
                    continue
                }
                Invoke-iTwinJsonRequest Patch "${script:AccessControlBaseUri}/${destinationGuid}/roles/${destinationRoleId}" $headers @{ permissions = $desired } | Out-Null
                & $newResult $sourceRole $action 'Success' "Updated ${roleText}: $change." $destinationRoleId
            }
            catch { & $newResult $sourceRole $action 'Failed' (Get-iTwinErrorText $_) $destinationRoleId }
        }
    }
}

function Get-iTwinOptionalValue {
    param($InputObject, [string]$Name)
    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) { return '' }
    [string]$property.Value
}

function Get-iTwinProjectMembers {
    param([string]$ProjectGuid)
    if (-not (Test-iTwinGuid $ProjectGuid)) { throw "Invalid ProjectGUID '$ProjectGuid'." }
    $headers = Get-iTwinAccessToken
    $members = Get-iTwinMembers $ProjectGuid $headers
    @($members | ForEach-Object {
            [pscustomobject]@{
                ProjectGUID  = $ProjectGuid
                MemberId     = Get-iTwinOptionalValue $_ 'id'
                Email        = Get-iTwinOptionalValue $_ 'email'
                GivenName    = Get-iTwinOptionalValue $_ 'givenName'
                Surname      = Get-iTwinOptionalValue $_ 'surname'
                Organization = Get-iTwinOptionalValue $_ 'organization'
                Roles        = @(@($_.roles) | Where-Object { $null -ne $_ } | ForEach-Object { [pscustomobject]@{ RoleId = Get-iTwinOptionalValue $_ 'id'; RoleDisplayName = Get-iTwinOptionalValue $_ 'displayName' } })
            }
        } | Sort-Object Email)
}

function Find-iTwinOrganizationUsers {
    param([string]$SearchText)
    $search = ([string]$SearchText).Trim()
    if ($search.Length -lt 3 -or $search.Length -gt 200) { throw 'Organization user search requires 3 to 200 characters.' }
    $headers = Get-iTwinAccessToken
    $headers.Accept = 'application/vnd.bentley.itwin-platform.v1+json'
    $response = Invoke-RestMethod -Method Get -Uri "${script:UsersBaseUri}/?`$search=$([Uri]::EscapeDataString($search))" -Headers $headers -ErrorAction Stop
    $users = if ($response.PSObject.Properties.Name -contains 'users') { @($response.users) } else { @() }
    @($users | Where-Object { $null -ne $_ } | ForEach-Object {
            [pscustomobject]@{
                UserId       = Get-iTwinOptionalValue $_ 'id'
                Email        = Get-iTwinOptionalValue $_ 'email'
                GivenName    = Get-iTwinOptionalValue $_ 'givenName'
                Surname      = Get-iTwinOptionalValue $_ 'surname'
                Organization = Get-iTwinOptionalValue $_ 'organizationName'
            }
        } | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Email) } | Sort-Object Email)
}

function Add-iTwinMemberRoles {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory = $true)][string]$ProjectGuid,
        [Parameter(Mandatory = $true)][string[]]$Email,
        [Parameter(Mandatory = $true)][string[]]$RoleId
    )
    if (-not (Test-iTwinGuid $ProjectGuid)) { throw "Invalid ProjectGUID '$ProjectGuid'." }
    $roleIds = @($RoleId | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ } | Select-Object -Unique)
    if ($roleIds.Count -eq 0) { throw 'At least one role is required.' }
    # Access Control v2 accepts at most 50 role assignments per add request.
    if ($roleIds.Count -gt 50) { throw 'At most 50 roles can be assigned to a user in one request.' }
    foreach ($id in $roleIds) { if (-not (Test-iTwinGuid $id)) { throw "RoleId '$id' must be a GUID." } }

    $headers = Get-iTwinAccessToken
    $members = Get-iTwinMembers $ProjectGuid $headers
    foreach ($address in @($Email | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ } | Select-Object -Unique)) {
        $appliedRoleIds = @()
        try {
            try { $mailAddress = New-Object System.Net.Mail.MailAddress($address) } catch { throw "Invalid email '$address'." }
            if ($mailAddress.Address -ne $address) { throw "Email '$address' must be a plain address." }
            $existing = @($members | Where-Object { $_.email -ieq $address })
            if ($existing.Count -eq 0) {
                $appliedRoleIds = $roleIds
                if (-not $PSCmdlet.ShouldProcess("$address on iTwin $ProjectGuid", "Add member with $($roleIds.Count) role(s)")) {
                    $result = [pscustomobject]@{ Status = 'Skipped'; Detail = 'WhatIf: would add the user with the selected role(s).' }
                }
                else {
                    $body = @{ members = @(@{ email = $address; roleIds = @($roleIds) }) } | ConvertTo-Json -Depth 4 -Compress
                    $response = Invoke-RestMethod -Method Post -Uri "${script:AccessControlBaseUri}/${ProjectGuid}/members/users" -Headers $headers -ContentType 'application/json' -Body $body -ErrorAction Stop
                    $invitations = @(if ($response.PSObject.Properties.Name -contains 'invitations') { $response.invitations | Where-Object { $null -ne $_ } })
                    $detail = if ($invitations.Count -gt 0) { "Invitation sent with $($roleIds.Count) role(s); membership starts when the user accepts." } else { "User added with $($roleIds.Count) role(s)." }
                    $result = [pscustomobject]@{ Status = 'Success'; Detail = $detail }
                }
            }
            else {
                $member = $existing[0]
                $currentRoleIds = @(@($member.roles) | Where-Object { $null -ne $_ } | ForEach-Object { [string]$_.id })
                $appliedRoleIds = @($roleIds | Where-Object { $currentRoleIds -notcontains $_ })
                if ($appliedRoleIds.Count -eq 0) {
                    $result = [pscustomobject]@{ Status = 'Skipped'; Detail = 'User already has the selected role(s).' }
                }
                elseif (-not $PSCmdlet.ShouldProcess("$address on iTwin $ProjectGuid", "Assign $($appliedRoleIds.Count) role(s)")) {
                    $result = [pscustomobject]@{ Status = 'Skipped'; Detail = 'WhatIf: would assign the missing role(s).' }
                }
                else {
                    # PATCH replaces the member's role list, so existing roles are resent.
                    $body = @{ roleIds = @($currentRoleIds + $appliedRoleIds) } | ConvertTo-Json -Compress
                    Invoke-RestMethod -Method Patch -Uri "${script:AccessControlBaseUri}/${ProjectGuid}/members/users/$($member.id)" -Headers $headers -ContentType 'application/json' -Body $body -ErrorAction Stop | Out-Null
                    $result = [pscustomobject]@{ Status = 'Success'; Detail = "Assigned $($appliedRoleIds.Count) role(s); $($roleIds.Count - $appliedRoleIds.Count) already assigned." }
                }
            }
        }
        catch { $result = [pscustomobject]@{ Status = 'Failed'; Detail = $_.Exception.Message } }
        [pscustomobject]@{ ProjectGUID = $ProjectGuid; User = $address; Action = 'AssignRole'; RoleId = ($appliedRoleIds -join ';'); Status = $result.Status; Detail = $result.Detail }
    }
}

function Update-iTwinProjectInventory {
    param($Project, [array]$Members, [string]$DatabasePath)
    $now = [DateTime]::UtcNow.ToString('o'); $projectGuid = [string]$Project.ProjectGUID
    $projectName = if ($Project.PSObject.Properties.Name -contains 'ProjectName') { [string]$Project.ProjectName } else { $null }
    foreach ($member in $Members) {
        $email = [string]$member.email
        $memberId = if ($member.PSObject.Properties.Name -contains 'id') { [string]$member.id } else { [string]$member.memberId }
        if ([string]::IsNullOrWhiteSpace($email) -or [string]::IsNullOrWhiteSpace($memberId)) { throw "Project '$projectGuid' returned a membership without email or member ID. Inventory was not changed." }
        $parameters = @{ projectGuid = $projectGuid; projectName = $projectName; email = $email; memberId = $memberId; givenName = [string]$member.givenName; surname = [string]$member.surname; organization = [string]$member.organization; now = $now }
        Invoke-iTwinSqliteQuery $DatabasePath 'UPDATE MembershipObservation SET Active = 0 WHERE ProjectGuid = @projectGuid AND Email = @email AND BentleyMemberId <> @memberId AND Active = 1;' $parameters | Out-Null
        Invoke-iTwinSqliteQuery $DatabasePath 'UPDATE MembershipObservation SET ProjectName = @projectName, GivenName = @givenName, Surname = @surname, Organization = @organization, LastObservedUtc = @now, Active = 1 WHERE ProjectGuid = @projectGuid AND Email = @email AND BentleyMemberId = @memberId;' $parameters | Out-Null
        Invoke-iTwinSqliteQuery $DatabasePath 'INSERT INTO MembershipObservation (ProjectGuid, ProjectName, Email, BentleyMemberId, GivenName, Surname, Organization, FirstObservedUtc, LastObservedUtc, Active) SELECT @projectGuid, @projectName, @email, @memberId, @givenName, @surname, @organization, @now, @now, 1 WHERE NOT EXISTS (SELECT 1 FROM MembershipObservation WHERE ProjectGuid = @projectGuid AND Email = @email AND BentleyMemberId = @memberId);' $parameters | Out-Null
        foreach ($role in @($member.roles)) {
            $roleId = [string]$role.id
            if ([string]::IsNullOrWhiteSpace($roleId)) { throw "Project '$projectGuid' returned a role without an ID. Inventory was not changed." }
            $roleParameters = @{ projectGuid = $projectGuid; memberId = $memberId; roleId = $roleId; roleName = [string]$role.displayName; assignedDateTimeUtc = [string]$role.assignedDateTime; now = $now }
            Invoke-iTwinSqliteQuery $DatabasePath 'UPDATE MembershipRoleObservation SET RoleName = @roleName, AssignedDateTimeUtc = @assignedDateTimeUtc, LastObservedUtc = @now, Active = 1 WHERE ProjectGuid = @projectGuid AND BentleyMemberId = @memberId AND RoleId = @roleId;' $roleParameters | Out-Null
            Invoke-iTwinSqliteQuery $DatabasePath 'INSERT INTO MembershipRoleObservation (ProjectGuid, BentleyMemberId, RoleId, RoleName, AssignedDateTimeUtc, FirstObservedUtc, LastObservedUtc, Active) SELECT @projectGuid, @memberId, @roleId, @roleName, @assignedDateTimeUtc, @now, @now, 1 WHERE NOT EXISTS (SELECT 1 FROM MembershipRoleObservation WHERE ProjectGuid = @projectGuid AND BentleyMemberId = @memberId AND RoleId = @roleId);' $roleParameters | Out-Null
        }
        Invoke-iTwinSqliteQuery $DatabasePath 'UPDATE MembershipRoleObservation SET Active = 0 WHERE ProjectGuid = @projectGuid AND BentleyMemberId = @memberId AND Active = 1 AND LastObservedUtc <> @now;' @{ projectGuid = $projectGuid; memberId = $memberId; now = $now } | Out-Null
    }
    Invoke-iTwinSqliteQuery $DatabasePath 'UPDATE MembershipObservation SET Active = 0 WHERE ProjectGuid = @projectGuid AND Active = 1 AND LastObservedUtc <> @now;' @{ projectGuid = $projectGuid; now = $now } | Out-Null
    Invoke-iTwinSqliteQuery $DatabasePath 'UPDATE MembershipRoleObservation SET Active = 0 WHERE ProjectGuid = @projectGuid AND Active = 1 AND LastObservedUtc <> @now;' @{ projectGuid = $projectGuid; now = $now } | Out-Null
}

function Invoke-iTwinMembershipPoll {
    param([string]$ProjectConfigPath, [string]$DatabasePath)
    Invoke-iTwinMembershipPollProjects -Projects (Get-iTwinProjectConfig $ProjectConfigPath) -DatabasePath $DatabasePath
}

function Invoke-iTwinMembershipPollProjects {
    param([array]$Projects, [string]$DatabasePath)
    if ($Projects.Count -eq 0) { throw 'At least one project is required for polling.' }
    foreach ($project in $Projects) {
        if (-not (Test-iTwinGuid ([string]$project.ProjectGUID))) { throw "Invalid ProjectGUID '$($project.ProjectGUID)' for polling." }
    }
    Initialize-iTwinMembershipDatabase $DatabasePath
    $headers = Get-iTwinAccessToken
    foreach ($project in $projects) {
        try {
            $members = Get-iTwinMembers $project.ProjectGUID $headers
            Update-iTwinProjectInventory $project $members $DatabasePath
            [pscustomobject]@{ ProjectGUID = $project.ProjectGUID; Status = 'Success'; Members = $members.Count; Error = $null }
        }
        catch { [pscustomobject]@{ ProjectGUID = $project.ProjectGUID; Status = 'Failed'; Members = $null; Error = $_.Exception.Message } }
    }
}

function Get-iTwinActiveMembershipReport {
    param([string]$DatabasePath, [string]$Email)
    if (-not (Test-Path -LiteralPath $DatabasePath -PathType Leaf)) { throw "Inventory database was not found: '$DatabasePath'. Run Poll mode first." }
    Import-iTwinSqliteDependency
    Invoke-iTwinSqliteQuery $DatabasePath 'SELECT m.ProjectGuid, m.ProjectName, m.Email, m.BentleyMemberId, m.GivenName, m.Surname, m.Organization, r.RoleId, r.RoleName, r.AssignedDateTimeUtc, m.LastObservedUtc FROM MembershipObservation m LEFT JOIN MembershipRoleObservation r ON r.ProjectGuid = m.ProjectGuid AND r.BentleyMemberId = m.BentleyMemberId AND r.Active = 1 WHERE m.Email = @email AND m.Active = 1 ORDER BY m.ProjectName, m.ProjectGuid, r.RoleName;' @{ email = $Email }
}

function Test-iTwinMembershipInventoryDatabase {
    param([string]$DatabasePath)
    if (-not (Test-Path -LiteralPath $DatabasePath -PathType Leaf)) {
        throw "Inventory database was not found: '$DatabasePath'. Run Poll mode first."
    }
    Import-iTwinSqliteDependency
    $tables = @(Invoke-iTwinSqliteQuery $DatabasePath "SELECT name FROM sqlite_master WHERE type = 'table' AND name IN ('MembershipObservation', 'MembershipRoleObservation');" | ForEach-Object { [string]$_.name })
    if ($tables -notcontains 'MembershipObservation' -or $tables -notcontains 'MembershipRoleObservation') {
        throw "Inventory database '$DatabasePath' does not contain the required membership tables. Run Poll mode first."
    }
}

function Get-iTwinInventoryProjects {
    param([string]$DatabasePath, [string]$Organization, [switch]$IncludeHistory)
    Test-iTwinMembershipInventoryDatabase $DatabasePath
    $parameters = @{ organization = ([string]$Organization).Trim() }
    $activeClause = if ($IncludeHistory) { '' } else { 'AND m.Active = 1' }
    Invoke-iTwinSqliteQuery $DatabasePath "SELECT m.ProjectGuid, COALESCE(NULLIF(m.ProjectName, ''), m.ProjectGuid) AS ProjectName, COUNT(DISTINCT m.Email) AS MemberCount, MAX(m.LastObservedUtc) AS LastObservedUtc FROM MembershipObservation m WHERE (@organization = '' OR m.Organization = @organization) $activeClause GROUP BY m.ProjectGuid, m.ProjectName ORDER BY ProjectName COLLATE NOCASE, m.ProjectGuid;" $parameters
}

function Get-iTwinConfiguredProjectInventory {
    param([string]$ProjectConfigPath, [string]$DatabasePath, [string]$Organization, [switch]$IncludeHistory)
    if (-not (Test-Path -LiteralPath $ProjectConfigPath -PathType Leaf)) { throw "Project configuration CSV was not found: '$ProjectConfigPath'." }
    $projects = @(Import-Csv -LiteralPath $ProjectConfigPath)
    if ($projects.Count -eq 0) { return }
    if ($projects[0].PSObject.Properties.Name -notcontains 'ProjectGUID') { throw 'Project configuration CSV requires ProjectGUID; ProjectName is optional.' }
    foreach ($project in $projects) {
        if (-not (Test-iTwinGuid ([string]$project.ProjectGUID))) { throw "Invalid ProjectGUID '$($project.ProjectGUID)' in project configuration CSV." }
    }
    Test-iTwinMembershipInventoryDatabase $DatabasePath
    $organization = ([string]$Organization).Trim()
    foreach ($project in $projects) {
        $parameters = @{ projectGuid = [string]$project.ProjectGUID; organization = $organization }
        $observations = @(Invoke-iTwinSqliteQuery $DatabasePath 'SELECT COUNT(DISTINCT CASE WHEN Active = 1 THEN Email END) AS ActiveMemberCount, COUNT(*) AS ObservationCount, MAX(LastObservedUtc) AS LastObservedUtc, COUNT(CASE WHEN @organization = '''' OR Organization = @organization THEN 1 END) AS OrganizationMatchCount FROM MembershipObservation WHERE ProjectGuid = @projectGuid;' $parameters)
        $observation = if ($observations.Count -gt 0) { $observations[0] } else { $null }
        if (-not [string]::IsNullOrWhiteSpace($organization) -and ($null -eq $observation -or [int]$observation.OrganizationMatchCount -eq 0)) { continue }
        $activeMemberCount = if ($null -eq $observation) { 0 } else { [int]$observation.ActiveMemberCount }
        $observationCount = if ($null -eq $observation) { 0 } else { [int]$observation.ObservationCount }
        $inventoryState = if ($observationCount -eq 0) { 'NeverPolled' } elseif ($activeMemberCount -gt 0) { 'Current' } else { 'HistoryOnly' }
        [pscustomobject]@{
            ProjectGUID       = [string]$project.ProjectGUID
            ProjectName       = if ([string]::IsNullOrWhiteSpace([string]$project.ProjectName)) { [string]$project.ProjectGUID } else { [string]$project.ProjectName }
            ActiveMemberCount = $activeMemberCount
            ObservationCount  = $observationCount
            LastObservedUtc   = if ($null -eq $observation) { $null } else { [string]$observation.LastObservedUtc }
            InventoryState    = $inventoryState
        }
    }
}

function Remove-iTwinTrackedProjects {
    param([string[]]$ProjectGuid, [string]$ProjectConfigPath, [string]$DatabasePath, [switch]$WhatIf)
    if ($ProjectGuid.Count -eq 0) { throw 'At least one ProjectGUID is required.' }
    $projects = @(Get-iTwinProjectConfig $ProjectConfigPath)
    Test-iTwinMembershipInventoryDatabase $DatabasePath
    $targets = @($ProjectGuid | ForEach-Object { ([string]$_).Trim() } | Select-Object -Unique)
    foreach ($guid in $targets) {
        if (-not (Test-iTwinGuid $guid)) { throw "Invalid ProjectGUID '$guid'." }
        if (@($projects | Where-Object { $_.ProjectGUID -eq $guid }).Count -ne 1) { throw "Project '$guid' is not present in '$ProjectConfigPath'." }
    }
    $plans = @()
    foreach ($guid in $targets) {
        $counts = @(Invoke-iTwinSqliteQuery $DatabasePath 'SELECT (SELECT COUNT(*) FROM MembershipObservation WHERE ProjectGuid = @projectGuid) AS MembershipRecords, (SELECT COUNT(*) FROM MembershipRoleObservation WHERE ProjectGuid = @projectGuid) AS RoleRecords;' @{ projectGuid = $guid })[0]
        $project = @($projects | Where-Object { $_.ProjectGUID -eq $guid })[0]
        $plans += [pscustomobject]@{ ProjectGUID = $guid; ProjectName = [string]$project.ProjectName; MembershipRecords = [int]$counts.MembershipRecords; RoleRecords = [int]$counts.RoleRecords; Status = if ($WhatIf) { 'Planned' } else { 'Pending' }; Detail = $null }
    }
    if ($WhatIf) { return $plans }

    $timestamp = Get-Date -Format 'yyyyMMddHHmmss'
    $csvBackupPath = "${ProjectConfigPath}.${timestamp}.bak"
    $databaseBackupPath = "${DatabasePath}.${timestamp}.bak"
    $temporaryCsvPath = "${ProjectConfigPath}.${timestamp}.tmp"
    Copy-Item -LiteralPath $ProjectConfigPath -Destination $csvBackupPath -ErrorAction Stop
    Copy-Item -LiteralPath $DatabasePath -Destination $databaseBackupPath -ErrorAction Stop
    $remainingProjects = @($projects | Where-Object { $targets -notcontains [string]$_.ProjectGUID })
    try {
        $remainingProjects | Select-Object ProjectGUID, ProjectName | Export-Csv -LiteralPath $temporaryCsvPath -NoTypeInformation
        foreach ($plan in $plans) {
            Invoke-iTwinSqliteQuery $DatabasePath 'DELETE FROM MembershipRoleObservation WHERE ProjectGuid = @projectGuid;' @{ projectGuid = $plan.ProjectGUID } | Out-Null
            Invoke-iTwinSqliteQuery $DatabasePath 'DELETE FROM MembershipObservation WHERE ProjectGuid = @projectGuid;' @{ projectGuid = $plan.ProjectGUID } | Out-Null
        }
        Move-Item -LiteralPath $temporaryCsvPath -Destination $ProjectConfigPath -Force -ErrorAction Stop
        foreach ($plan in $plans) { $plan.Status = 'Success'; $plan.Detail = "Removed from configuration and purged $($plan.MembershipRecords) membership and $($plan.RoleRecords) role records." }
    }
    catch {
        if (Test-Path -LiteralPath $temporaryCsvPath -PathType Leaf) { Remove-Item -LiteralPath $temporaryCsvPath -Force -ErrorAction SilentlyContinue }
        if (Test-Path -LiteralPath $databaseBackupPath -PathType Leaf) { Copy-Item -LiteralPath $databaseBackupPath -Destination $DatabasePath -Force -ErrorAction SilentlyContinue }
        throw "Tracked project removal failed. The original CSV is backed up at '$csvBackupPath' and the local database was restored from '$databaseBackupPath'. $($_.Exception.Message)"
    }
    $plans
}

function Get-iTwinInventoryUsers {
    param([string]$DatabasePath, [string]$Organization, [switch]$IncludeHistory)
    Test-iTwinMembershipInventoryDatabase $DatabasePath
    $parameters = @{ organization = ([string]$Organization).Trim() }
    $activeClause = if ($IncludeHistory) { '' } else { 'AND m.Active = 1' }
    Invoke-iTwinSqliteQuery $DatabasePath "SELECT m.Email, MAX(m.GivenName) AS GivenName, MAX(m.Surname) AS Surname, MAX(m.Organization) AS Organization, COUNT(DISTINCT m.ProjectGuid) AS ProjectCount, MAX(m.LastObservedUtc) AS LastObservedUtc FROM MembershipObservation m WHERE (@organization = '' OR m.Organization = @organization) $activeClause GROUP BY m.Email ORDER BY m.Email COLLATE NOCASE;" $parameters
}

function Get-iTwinInventoryOrganizations {
    param([string]$DatabasePath, [switch]$IncludeHistory)
    Test-iTwinMembershipInventoryDatabase $DatabasePath
    $activeClause = if ($IncludeHistory) { '' } else { 'AND m.Active = 1' }
    Invoke-iTwinSqliteQuery $DatabasePath "SELECT m.Organization, COUNT(DISTINCT m.ProjectGuid) AS ProjectCount, COUNT(DISTINCT m.Email) AS UserCount, MAX(m.LastObservedUtc) AS LastObservedUtc FROM MembershipObservation m WHERE 1 = 1 $activeClause AND NULLIF(TRIM(m.Organization), '') IS NOT NULL GROUP BY m.Organization ORDER BY m.Organization COLLATE NOCASE;"
}

function Get-iTwinInventoryRelationships {
    param([string]$DatabasePath, [string]$ProjectGuid, [string]$Email, [string]$Organization, [switch]$IncludeHistory)
    Test-iTwinMembershipInventoryDatabase $DatabasePath
    $parameters = @{ projectGuid = ([string]$ProjectGuid).Trim(); email = ([string]$Email).Trim(); organization = ([string]$Organization).Trim(); includeHistory = if ($IncludeHistory) { 1 } else { 0 } }
    $activeClause = if ($IncludeHistory) { '' } else { 'AND m.Active = 1' }
    Invoke-iTwinSqliteQuery $DatabasePath "SELECT m.ProjectGuid, COALESCE(NULLIF(m.ProjectName, ''), m.ProjectGuid) AS ProjectName, m.Email, m.BentleyMemberId, m.GivenName, m.Surname, m.Organization, m.Active, m.FirstObservedUtc, m.LastObservedUtc, GROUP_CONCAT(CASE WHEN r.Active = 1 OR @includeHistory = 1 THEN r.RoleName END, '; ') AS Roles FROM MembershipObservation m LEFT JOIN MembershipRoleObservation r ON r.ProjectGuid = m.ProjectGuid AND r.BentleyMemberId = m.BentleyMemberId WHERE (@projectGuid = '' OR m.ProjectGuid = @projectGuid) AND (@email = '' OR m.Email = @email) AND (@organization = '' OR m.Organization = @organization) $activeClause GROUP BY m.ObservationId ORDER BY ProjectName COLLATE NOCASE, m.Email COLLATE NOCASE;" $parameters
}

function Get-iTwinBulkActions {
    param([string]$ActionCsvPath, [string]$DefaultRoleId)
    if (-not (Test-Path -LiteralPath $ActionCsvPath -PathType Leaf)) { throw "Action CSV was not found: '$ActionCsvPath'." }
    $actions = @(Import-Csv -LiteralPath $ActionCsvPath)
    if ($actions.Count -eq 0) { throw 'Action CSV contains no action rows.' }
    $headers = @($actions[0].PSObject.Properties.Name)
    foreach ($requiredHeader in @('ProjectGUID', 'User', 'Action')) { if ($headers -notcontains $requiredHeader) { throw "Action CSV is missing required header '$requiredHeader'." } }
    $seen = @{}; $validated = @()
    foreach ($row in $actions) {
        $guid = [string]$row.ProjectGUID; $email = ([string]$row.User).Trim(); $action = ([string]$row.Action).Trim()
        $roleId = if ($headers -contains 'RoleId') { ([string]$row.RoleId).Trim() } else { '' }
        $roleDisplayName = if ($headers -contains 'RoleDisplayName') { ([string]$row.RoleDisplayName).Trim() } else { '' }
        if (-not (Test-iTwinGuid $guid)) { throw "Invalid ProjectGUID '$guid' in action CSV." }
        try { $emailAddress = New-Object System.Net.Mail.MailAddress($email) } catch { throw "Invalid email '$email' in action CSV." }
        if ($emailAddress.Address -ne $email) { throw "Email '$email' in action CSV must be a plain address." }
        switch -Regex ($action) {
            '^(?i:Add|AssignRole)$' { $action = 'AssignRole'; break }
            '^(?i:UnassignRole)$' { $action = 'UnassignRole'; break }
            '^(?i:Remove|RemoveMember)$' { $action = 'RemoveMember'; break }
            default { throw "Action '$action' must be AssignRole, UnassignRole, or RemoveMember." }
        }
        if ($action -eq 'AssignRole' -and [string]::IsNullOrWhiteSpace($roleId) -and [string]::IsNullOrWhiteSpace($roleDisplayName)) { $roleId = $DefaultRoleId }
        if ($action -in @('AssignRole', 'UnassignRole') -and [string]::IsNullOrWhiteSpace($roleId) -and [string]::IsNullOrWhiteSpace($roleDisplayName)) { throw "$action for '$email' requires RoleId, RoleDisplayName, or -DefaultRoleId." }
        if (-not [string]::IsNullOrWhiteSpace($roleId) -and -not (Test-iTwinGuid $roleId)) { throw "RoleId '$roleId' for '$email' must be a GUID." }
        $key = "${guid}|$($email.ToLowerInvariant())|$action|$($roleId.ToLowerInvariant())|$($roleDisplayName.ToLowerInvariant())"
        if ($seen.ContainsKey($key)) { throw "Duplicate or conflicting action for project '$guid' and user '$email'." }
        $seen[$key] = $true; $validated += [pscustomobject]@{ ProjectGUID = $guid; User = $email; Action = $action; RoleId = $roleId; RoleDisplayName = $roleDisplayName }
    }
    $validated
}

function Resolve-iTwinRoleId {
    param([string]$ProjectGuid, [string]$RoleId, [string]$RoleDisplayName, [hashtable]$Headers)
    if (-not [string]::IsNullOrWhiteSpace($RoleId)) { return $RoleId }
    $roleMatches = @(Get-iTwinAssignableRoles $ProjectGuid $Headers | Where-Object { $_.displayName -ieq $RoleDisplayName })
    if ($roleMatches.Count -eq 0) { throw "Project '$ProjectGuid' does not have an iTwin or enterprise role named '$RoleDisplayName'." }
    if ($roleMatches.Count -gt 1) { throw "Project '$ProjectGuid' has multiple iTwin or enterprise roles named '$RoleDisplayName'. Use RoleId instead." }
    [string]$roleMatches[0].id
}

function Invoke-iTwinMembershipBulkAction {
    param([string]$ActionCsvPath, [string]$ProjectConfigPath, [string]$DatabasePath, [string]$DefaultRoleId, [switch]$WhatIf)
    $actions = Get-iTwinBulkActions $ActionCsvPath $DefaultRoleId
    $configuredProjects = @((Get-iTwinProjectConfig $ProjectConfigPath) | ForEach-Object { [string]$_.ProjectGUID })
    foreach ($action in $actions) { if ($configuredProjects -notcontains $action.ProjectGUID) { throw "Action project '$($action.ProjectGUID)' is not present in the project configuration CSV." } }
    if ($WhatIf) { foreach ($action in $actions) { [pscustomobject]@{ ProjectGUID = $action.ProjectGUID; User = $action.User; Action = $action.Action; Status = 'Skipped'; Detail = "WhatIf: would call the API and refresh '$DatabasePath'. RoleDisplayName resolution requires a live run." } }; return }
    Initialize-iTwinMembershipDatabase $DatabasePath; $headers = Get-iTwinAccessToken
    foreach ($action in $actions) {
        try {
            $roleId = if ($action.Action -eq 'RemoveMember') { $null } else { Resolve-iTwinRoleId $action.ProjectGUID $action.RoleId $action.RoleDisplayName $headers }
            $members = Get-iTwinMembers $action.ProjectGUID $headers; $existing = @($members | Where-Object { $_.email -ieq $action.User })
            if ($action.Action -eq 'AssignRole') {
                if ($existing.Count -eq 0) {
                    $body = @{ members = @(@{ email = $action.User; roleIds = @($roleId) }) } | ConvertTo-Json -Depth 4 -Compress
                    Invoke-RestMethod -Method Post -Uri "${script:AccessControlBaseUri}/$($action.ProjectGUID)/members/users" -Headers $headers -ContentType 'application/json' -Body $body -ErrorAction Stop | Out-Null
                    $result = [pscustomobject]@{ Status = 'Success'; Detail = "User added with role '$roleId'." }
                }
                else {
                    $member = $existing[0]; $currentRoleIds = @($member.roles | ForEach-Object { [string]$_.id })
                    if ($currentRoleIds -contains $roleId) { $result = [pscustomobject]@{ Status = 'Skipped'; Detail = "User already has role '$roleId'." } }
                    else { $body = @{ roleIds = @($currentRoleIds + $roleId) } | ConvertTo-Json -Compress; Invoke-RestMethod -Method Patch -Uri "${script:AccessControlBaseUri}/$($action.ProjectGUID)/members/users/$($member.id)" -Headers $headers -ContentType 'application/json' -Body $body -ErrorAction Stop | Out-Null; $result = [pscustomobject]@{ Status = 'Success'; Detail = "Role '$roleId' assigned." } }
                }
            }
            elseif ($action.Action -eq 'UnassignRole') {
                if ($existing.Count -eq 0) { $result = [pscustomobject]@{ Status = 'Skipped'; Detail = 'User is absent.' } }
                else {
                    $member = $existing[0]; $currentRoleIds = @($member.roles | ForEach-Object { [string]$_.id })
                    if ($currentRoleIds -notcontains $roleId) { $result = [pscustomobject]@{ Status = 'Skipped'; Detail = "User does not have role '$roleId'." } }
                    else { $body = @{ roleIds = @($currentRoleIds | Where-Object { $_ -ne $roleId }) } | ConvertTo-Json -Compress; Invoke-RestMethod -Method Patch -Uri "${script:AccessControlBaseUri}/$($action.ProjectGUID)/members/users/$($member.id)" -Headers $headers -ContentType 'application/json' -Body $body -ErrorAction Stop | Out-Null; $result = [pscustomobject]@{ Status = 'Success'; Detail = "Role '$roleId' unassigned." } }
                }
            }
            elseif ($existing.Count -eq 0) { $result = [pscustomobject]@{ Status = 'Skipped'; Detail = 'User is absent.' } }
            else { Invoke-RestMethod -Method Delete -Uri "${script:AccessControlBaseUri}/$($action.ProjectGUID)/members/users/$($existing[0].id)" -Headers $headers -ErrorAction Stop | Out-Null; $result = [pscustomobject]@{ Status = 'Success'; Detail = 'User removed.' } }
            $project = [pscustomobject]@{ ProjectGUID = $action.ProjectGUID; ProjectName = $null }
            Update-iTwinProjectInventory $project (Get-iTwinMembers $action.ProjectGUID $headers) $DatabasePath
            [pscustomobject]@{ ProjectGUID = $action.ProjectGUID; User = $action.User; Action = $action.Action; RoleId = $roleId; Status = $result.Status; Detail = $result.Detail }
        }
        catch { [pscustomobject]@{ ProjectGUID = $action.ProjectGUID; User = $action.User; Action = $action.Action; RoleId = $null; Status = 'Failed'; Detail = $_.Exception.Message } }
    }
}

function Invoke-iTwinMembershipActions {
    param([array]$Actions, [string]$DatabasePath, [switch]$WhatIf)
    if ($Actions.Count -eq 0) { throw 'At least one membership action is required.' }
    if ($WhatIf) {
        foreach ($action in $Actions) {
            [pscustomobject]@{ ProjectGUID = [string]$action.ProjectGUID; User = [string]$action.User; Action = [string]$action.Action; RoleId = [string]$action.RoleId; Status = 'Skipped'; Detail = 'WhatIf: would call the API and refresh the inventory.' }
        }
        return
    }
    Initialize-iTwinMembershipDatabase $DatabasePath
    $headers = Get-iTwinAccessToken
    foreach ($action in $Actions) {
        $projectGuid = ([string]$action.ProjectGUID).Trim()
        $email = ([string]$action.User).Trim()
        $actionName = ([string]$action.Action).Trim()
        $roleId = ([string]$action.RoleId).Trim()
        $roleDisplayName = ([string]$action.RoleDisplayName).Trim()
        try {
            if (-not (Test-iTwinGuid $projectGuid)) { throw "Invalid ProjectGUID '$projectGuid'." }
            try { $emailAddress = New-Object System.Net.Mail.MailAddress($email) } catch { throw "Invalid email '$email'." }
            if ($emailAddress.Address -ne $email) { throw "Email '$email' must be a plain address." }
            switch -Regex ($actionName) {
                '^(?i:Add|AssignRole)$' { $actionName = 'AssignRole'; break }
                '^(?i:Remove|RemoveMember)$' { $actionName = 'RemoveMember'; break }
                default { throw "Action '$actionName' must be AssignRole or RemoveMember." }
            }
            if ($actionName -eq 'AssignRole') { $roleId = Resolve-iTwinRoleId $projectGuid $roleId $roleDisplayName $headers }
            $members = Get-iTwinMembers $projectGuid $headers
            $existing = @($members | Where-Object { $_.email -ieq $email })
            if ($actionName -eq 'AssignRole') {
                if ($existing.Count -eq 0) {
                    $body = @{ members = @(@{ email = $email; roleIds = @($roleId) }) } | ConvertTo-Json -Depth 4 -Compress
                    Invoke-RestMethod -Method Post -Uri "${script:AccessControlBaseUri}/${projectGuid}/members/users" -Headers $headers -ContentType 'application/json' -Body $body -ErrorAction Stop | Out-Null
                    $result = [pscustomobject]@{ Status = 'Success'; Detail = "User added with role '$roleId'." }
                }
                else {
                    $member = $existing[0]
                    $currentRoleIds = @($member.roles | ForEach-Object { [string]$_.id })
                    if ($currentRoleIds -contains $roleId) { $result = [pscustomobject]@{ Status = 'Skipped'; Detail = "User already has role '$roleId'." } }
                    else {
                        $body = @{ roleIds = @($currentRoleIds + $roleId) } | ConvertTo-Json -Compress
                        Invoke-RestMethod -Method Patch -Uri "${script:AccessControlBaseUri}/${projectGuid}/members/users/$($member.id)" -Headers $headers -ContentType 'application/json' -Body $body -ErrorAction Stop | Out-Null
                        $result = [pscustomobject]@{ Status = 'Success'; Detail = "Role '$roleId' assigned." }
                    }
                }
            }
            elseif ($existing.Count -eq 0) { $result = [pscustomobject]@{ Status = 'Skipped'; Detail = 'User is absent.' } }
            else {
                Invoke-RestMethod -Method Delete -Uri "${script:AccessControlBaseUri}/${projectGuid}/members/users/$($existing[0].id)" -Headers $headers -ErrorAction Stop | Out-Null
                $result = [pscustomobject]@{ Status = 'Success'; Detail = 'User removed.' }
            }
            $project = [pscustomobject]@{ ProjectGUID = $projectGuid; ProjectName = $null }
            Update-iTwinProjectInventory $project (Get-iTwinMembers $projectGuid $headers) $DatabasePath
            [pscustomobject]@{ ProjectGUID = $projectGuid; User = $email; Action = $actionName; RoleId = $roleId; Status = $result.Status; Detail = $result.Detail }
        }
        catch { [pscustomobject]@{ ProjectGUID = $projectGuid; User = $email; Action = $actionName; RoleId = $roleId; Status = 'Failed'; Detail = $_.Exception.Message } }
    }
}

Export-ModuleMember -Function Invoke-iTwinMembershipPoll, Invoke-iTwinMembershipPollProjects, Get-iTwinActiveMembershipReport, Invoke-iTwinMembershipBulkAction, Invoke-iTwinMembershipActions, Get-iTwinAvailableProjects, Get-iTwinAvailableProjectPage, Get-iTwinPrimaryAccount, Get-iTwinFavoriteProjects, Get-iTwinFavoriteProjectPage, Get-iTwinProjectRoles, Get-iTwinRoleDefinitions, Copy-iTwinRoleDefinitions, Get-iTwinProjectMembers, Find-iTwinOrganizationUsers, Add-iTwinMemberRoles, Test-iTwinMembershipInventoryDatabase, Get-iTwinInventoryProjects, Get-iTwinConfiguredProjectInventory, Remove-iTwinTrackedProjects, Get-iTwinInventoryUsers, Get-iTwinInventoryOrganizations, Get-iTwinInventoryRelationships
