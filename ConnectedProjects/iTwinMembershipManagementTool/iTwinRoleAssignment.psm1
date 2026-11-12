Set-StrictMode -Version Latest

# WPF event handlers run in this module's scope, so the single view's controls and state live here.
$script:root = $null
$script:controls = $null
$script:state = $null
$script:databasePath = $null
$script:projectConfigPath = $null
$script:auditLogPath = $null
$script:onAssigned = $null

#region Helpers

function Test-OrganizationAccessDenied {
    param($ErrorRecord)
    $exception = $ErrorRecord.Exception
    $responseProperty = $exception.PSObject.Properties['Response']
    if ($null -ne $responseProperty -and $null -ne $responseProperty.Value) {
        try { return [int]$responseProperty.Value.StatusCode -eq 403 } catch { }
    }
    [string]$exception.Message -match 'InsufficientPermissions|\b403\b'
}

function Write-RoleAssignmentAudit {
    param([array]$Results)
    $rows = @($Results | ForEach-Object {
            [pscustomobject]@{
                TimestampUtc = [DateTime]::UtcNow.ToString('o')
                Operator     = [Environment]::UserName
                ProjectGUID  = [string]$_.ProjectGUID
                User         = [string]$_.User
                Action       = [string]$_.Action
                RoleId       = [string]$_.RoleId
                Status       = [string]$_.Status
                Detail       = [string]$_.Detail
            }
        })
    if ($rows.Count -gt 0) { $rows | Export-Csv -LiteralPath $auditLogPath -NoTypeInformation -Append }
}

function Get-UserDisplayName {
    param($User)
    $name = ("{0} {1}" -f $User.GivenName, $User.Surname).Trim()
    if ($name) { "$name <$($User.Email)>" } else { [string]$User.Email }
}

function Show-Message {
    param([string]$Text, [string]$Title, [string]$Icon = 'Information')
    [System.Windows.MessageBox]::Show([System.Windows.Window]::GetWindow($root), $Text, $Title, 'OK', $Icon) | Out-Null
}

function Invoke-WithWaitCursor {
    param([scriptblock]$Action)
    [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
    # Let pending status text render before the API call blocks the UI thread.
    $root.Dispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::Background)
    try { & $Action } finally { [System.Windows.Input.Mouse]::OverrideCursor = $null }
}

#endregion Helpers

#region Load data

function Update-iTwinPage {
    $isFavorites = $controls.SourceTabs.SelectedIndex -eq 1
    $queryScope = if ([bool]$controls.OrganizationScopeBox.IsChecked) { 'OrganizationAdmin' } else { 'MemberOfiTwin' }
    $projectStatus = [string]$controls.ProjectStatusBox.SelectedItem.Content
    $searchText = ([string]$controls.SearchBox.Text).Trim()
    $nextUri = $state.NextUris[$state.PageIndex]
    $fallbackNote = ''
    foreach ($button in @($controls.SearchButton, $controls.PreviousButton, $controls.NextButton)) { $button.IsEnabled = $false }
    $controls.iTwinStatusText.Text = 'Retrieving iTwins...'
    try {
        $page = Invoke-WithWaitCursor {
            if ($isFavorites) { return Get-iTwinFavoriteProjectPage -ProjectStatus $projectStatus -PageSize 250 -NextUri $nextUri }
            try { Get-iTwinAvailableProjectPage -DisplayName $searchText -QueryScope $queryScope -ProjectStatus $projectStatus -PageSize 250 -NextUri $nextUri }
            catch {
                if ($queryScope -ne 'OrganizationAdmin' -or -not (Test-OrganizationAccessDenied $_)) { throw }
                $null
            }
        }
        if ($null -eq $page) {
            $controls.OrganizationScopeBox.IsChecked = $false
            $queryScope = 'MemberOfiTwin'
            $state.PageIndex = 0
            $state.NextUris.Clear(); $state.NextUris.Add('')
            $fallbackNote = 'Organization access is unavailable; showing your iTwins. '
            Write-Warning 'Organization access is unavailable. Showing iTwins available to your account.'
            $page = Invoke-WithWaitCursor { Get-iTwinAvailableProjectPage -DisplayName $searchText -QueryScope $queryScope -ProjectStatus $projectStatus -PageSize 250 }
        }

        $items = @($page.Projects | Sort-Object ProjectName | ForEach-Object {
                [pscustomobject]@{ ProjectGUID = [string]$_.ProjectGUID; ProjectName = [string]$_.ProjectName; Status = [string]$_.Status; Label = "$($_.ProjectName) [$($_.Status)]" }
            })
        $controls.iTwinList.ItemsSource = $items
        if ($state.NextUris.Count -eq $state.PageIndex + 1) { $state.NextUris.Add([string]$page.NextUri) } else { $state.NextUris[$state.PageIndex + 1] = [string]$page.NextUri }
        $controls.PreviousButton.IsEnabled = $state.PageIndex -gt 0
        $controls.NextButton.IsEnabled = -not [string]::IsNullOrWhiteSpace([string]$page.NextUri)
        $source = if ($isFavorites) { 'favorite' } else { $queryScope }
        $controls.iTwinStatusText.Text = "${fallbackNote}Page $($state.PageIndex + 1): $($items.Count) $projectStatus $source iTwin(s)."
    }
    catch {
        $controls.iTwinList.ItemsSource = @()
        $controls.iTwinStatusText.Text = "iTwin retrieval failed: $($_.Exception.Message)"
        Write-Warning "iTwin retrieval failed: $($_.Exception.Message)"
    }
    finally { $controls.SearchButton.IsEnabled = $true }
}

function Reset-iTwinPage {
    $state.PageIndex = 0
    $state.NextUris.Clear()
    $state.NextUris.Add('')
    Update-iTwinPage
}

function Update-UserList {
    $selectedEmails = @($controls.UserList.SelectedItems | ForEach-Object { [string]$_.Email })
    $memberEmails = @{}
    $items = [System.Collections.Generic.List[object]]::new()
    foreach ($member in $state.Members) {
        $memberEmails[$member.Email.ToLowerInvariant()] = $true
        $roleNames = @($member.Roles | ForEach-Object { $_.RoleDisplayName }) -join ', '
        $label = "{0} | {1} | Roles: {2}" -f (Get-UserDisplayName $member), $member.Organization, $(if ($roleNames) { $roleNames } else { 'none' })
        $items.Add([pscustomobject]@{ Email = $member.Email; Member = $member; Label = $label })
    }
    foreach ($user in $state.ExtraUsers.Values) {
        if ($memberEmails.ContainsKey($user.Email.ToLowerInvariant())) { continue }
        $organization = if ($user.Organization) { " | $($user.Organization)" } else { '' }
        $items.Add([pscustomobject]@{ Email = $user.Email; Member = $null; Label = "[Not a member] $(Get-UserDisplayName $user)$organization" })
    }
    $controls.UserList.ItemsSource = $items
    foreach ($item in $items) {
        if ($selectedEmails -contains $item.Email) { $controls.UserList.SelectedItems.Add($item) | Out-Null }
    }
    $controls.UserStatusText.Text = "$($state.Members.Count) member(s); $($items.Count - $state.Members.Count) non-member(s) added to the list."
}

function Update-SelectediTwin {
    $itwin = $controls.iTwinList.SelectedItem
    $state.SelectediTwin = $itwin
    $state.Members = @()
    $state.Roles = @()
    $state.ExtraUsers.Clear()
    $controls.UserList.ItemsSource = @()
    $controls.RoleList.ItemsSource = @()
    $controls.UserStatusText.Text = ''
    $controls.RoleStatusText.Text = ''
    foreach ($control in @($controls.ReloadButton, $controls.UserSearchBox, $controls.UserSearchButton, $controls.AddEmailButton, $controls.AssignButton)) { $control.IsEnabled = $false }
    if ($null -eq $itwin) {
        $controls.StatusText.Text = 'Select an iTwin to load its members and roles.'
        return
    }

    $controls.StatusText.Text = "Loading members and roles for '$($itwin.ProjectName)'..."
    try {
        Invoke-WithWaitCursor {
            $state.Members = @(Get-iTwinProjectMembers -ProjectGuid $itwin.ProjectGUID)
            $state.Roles = @(Get-iTwinProjectRoles -ProjectGuid $itwin.ProjectGUID)
        }
        Update-UserList
        $controls.RoleList.ItemsSource = $state.Roles
        $controls.RoleStatusText.Text = "$($state.Roles.Count) role(s) available."
        foreach ($control in @($controls.ReloadButton, $controls.UserSearchBox, $controls.UserSearchButton, $controls.AddEmailButton, $controls.AssignButton)) { $control.IsEnabled = $true }
        $controls.StatusText.Text = "'$($itwin.ProjectName)' ($($itwin.ProjectGUID)). Select users and roles, then assign."
        Write-Host "Loaded iTwin '$($itwin.ProjectName)' ($($itwin.ProjectGUID)): $($state.Members.Count) member(s), $($state.Roles.Count) role(s)."
    }
    catch {
        $controls.ReloadButton.IsEnabled = $true
        $controls.StatusText.Text = "Could not load '$($itwin.ProjectName)': $($_.Exception.Message)"
        Write-Warning "Could not load iTwin '$($itwin.ProjectName)' ($($itwin.ProjectGUID)): $($_.Exception.Message)"
    }
}

function Test-iTwinTracked {
    param([string]$ProjectGuid)
    if (-not (Test-Path -LiteralPath $projectConfigPath -PathType Leaf)) { return $false }
    @(Import-Csv -LiteralPath $projectConfigPath | Where-Object { [string]$_.ProjectGUID -eq $ProjectGuid }).Count -gt 0
}

#endregion Load data

#region Assign roles

function Invoke-RoleAssignment {
    $itwin = $state.SelectediTwin
    $users = @($controls.UserList.SelectedItems | ForEach-Object { $_ })
    $roles = @($controls.RoleList.SelectedItems | ForEach-Object { $_ })
    if ($null -eq $itwin -or $users.Count -eq 0 -or $roles.Count -eq 0) {
        Show-Message 'Select an iTwin, one or more users, and one or more roles.' 'Assign roles'
        return
    }

    $roleIds = @($roles | ForEach-Object { [string]$_.RoleId })
    $roleNames = @($roles | ForEach-Object { $_.Label }) -join ', '
    $preview = foreach ($user in $users) {
        if ($null -eq $user.Member) {
            "  $($user.Email): not a member; will be added, or invited if outside your organization"
            continue
        }
        $currentRoleIds = @($user.Member.Roles | ForEach-Object { $_.RoleId })
        $missing = @($roles | Where-Object { $currentRoleIds -notcontains $_.RoleId } | ForEach-Object { $_.Label })
        if ($missing.Count -eq 0) { "  $($user.Email): already has all selected roles (skipped)" }
        else { "  $($user.Email): adds $($missing -join ', ')" }
    }
    $preview = @($preview)
    $previewText = (@($preview | Select-Object -First 25) -join "`n")
    if ($preview.Count -gt 25) { $previewText += "`n  ...and $($preview.Count - 25) more user(s)" }
    $message = "iTwin: $($itwin.ProjectName)`nRoles: ${roleNames}`n`n$previewText`n`nExisting roles are kept. Continue?"
    if ([System.Windows.MessageBox]::Show([System.Windows.Window]::GetWindow($root), $message, 'Confirm role assignment', 'YesNo', 'Question') -ne 'Yes') {
        Write-Host "Role assignment cancelled for iTwin '$($itwin.ProjectName)'."
        return
    }

    Write-Host "Assigning role(s) '${roleNames}' to $($users.Count) user(s) on iTwin '$($itwin.ProjectName)' ($($itwin.ProjectGUID))."
    $controls.AssignButton.IsEnabled = $false
    $controls.StatusText.Text = "Assigning roles to $($users.Count) user(s)..."
    try {
        $results = @(Invoke-WithWaitCursor { Add-iTwinMemberRoles -ProjectGuid $itwin.ProjectGUID -Email @($users | ForEach-Object { $_.Email }) -RoleId $roleIds -Confirm:$false })
        Write-RoleAssignmentAudit -Results $results
        foreach ($result in $results) { Write-Host "  $($result.Status) | $($result.User) | $($result.Detail)" }

        $refreshNote = ''
        if (@($results | Where-Object { $_.Status -eq 'Success' }).Count -gt 0 -and (Test-iTwinTracked $itwin.ProjectGUID) -and (Test-Path -LiteralPath $databasePath -PathType Leaf)) {
            try {
                $poll = @(Invoke-iTwinMembershipPollProjects -Projects @([pscustomobject]@{ ProjectGUID = $itwin.ProjectGUID; ProjectName = $itwin.ProjectName }) -DatabasePath $databasePath)
                $refreshNote = if ($poll[0].Status -eq 'Success') { "`n`nLocal inventory refreshed." } else { "`n`nLocal inventory was not refreshed: $($poll[0].Error)" }
                if ($poll[0].Status -eq 'Success' -and $null -ne $onAssigned) { & $onAssigned }
            }
            catch { $refreshNote = "`n`nLocal inventory was not refreshed: $($_.Exception.Message)" }
            Write-Host $refreshNote.Trim()
        }

        $summary = @($results | ForEach-Object { "$($_.Status) | $($_.User) | $($_.Detail)" }) -join "`n"
        $icon = if (@($results | Where-Object { $_.Status -eq 'Failed' }).Count -gt 0) { 'Warning' } else { 'Information' }
        Show-Message "$summary$refreshNote`n`nResults were written to '$auditLogPath'." 'Role assignment results' $icon
        $controls.StatusText.Text = "Assignment completed: $(@($results | Where-Object Status -eq 'Success').Count) succeeded, $(@($results | Where-Object Status -eq 'Skipped').Count) skipped, $(@($results | Where-Object Status -eq 'Failed').Count) failed."
        Write-Host $controls.StatusText.Text
    }
    catch {
        Write-Warning "Role assignment failed before any user result was returned: $($_.Exception.Message)"
        Show-Message $_.Exception.Message 'Assign roles' 'Error'
        $controls.StatusText.Text = 'Role assignment failed before any user result was returned.'
        return
    }
    finally { $controls.AssignButton.IsEnabled = $true }

    try {
        $state.Members = @(Invoke-WithWaitCursor { Get-iTwinProjectMembers -ProjectGuid $itwin.ProjectGUID })
        Update-UserList
    }
    catch { $controls.StatusText.Text += " Member list was not reloaded: $($_.Exception.Message)" }
}

#endregion Assign roles

#region View

function New-iTwinRoleAssignmentView {
    param(
        [Parameter(Mandatory = $true)][string]$DatabasePath,
        [Parameter(Mandatory = $true)][string]$ProjectConfigPath,
        [Parameter(Mandatory = $true)][string]$AuditLogPath,
        [scriptblock]$OnAssigned
    )
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, Microsoft.VisualBasic
    $script:databasePath = $DatabasePath
    $script:projectConfigPath = $ProjectConfigPath
    $script:auditLogPath = $AuditLogPath
    $script:onAssigned = $OnAssigned

    [xml]$xaml = @'
<Grid xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Margin="14">
  <Grid.RowDefinitions><RowDefinition Height="Auto" /><RowDefinition Height="Auto" /><RowDefinition Height="*" /><RowDefinition Height="Auto" /></Grid.RowDefinitions>
  <DockPanel Grid.Row="0" LastChildFill="False" Margin="0,0,0,10">
    <TextBlock Text="iTwin Role Assignment" FontSize="22" FontWeight="SemiBold" Foreground="#17324D" VerticalAlignment="Center" DockPanel.Dock="Left" />
    <Button x:Name="ReloadButton" Content="Reload members and roles" Padding="12,5" DockPanel.Dock="Right" IsEnabled="False" />
  </DockPanel>
  <Grid Grid.Row="1" Margin="0,0,0,10">
    <Grid.ColumnDefinitions><ColumnDefinition Width="210" /><ColumnDefinition Width="Auto" /><ColumnDefinition Width="Auto" /><ColumnDefinition Width="*" /><ColumnDefinition Width="Auto" /><ColumnDefinition Width="Auto" /><ColumnDefinition Width="Auto" /></Grid.ColumnDefinitions>
    <TabControl x:Name="SourceTabs" Height="30" SelectedIndex="0"><TabItem Header="All iTwins" /><TabItem Header="My Favorites" /></TabControl>
    <CheckBox x:Name="OrganizationScopeBox" Grid.Column="1" Content="Show organization iTwins" Margin="14,0,0,0" VerticalAlignment="Center" />
    <ComboBox x:Name="ProjectStatusBox" Grid.Column="2" Width="88" Margin="14,0,8,0" SelectedIndex="0"><ComboBoxItem Content="Active" /><ComboBoxItem Content="Trial" /><ComboBoxItem Content="Inactive" /><ComboBoxItem Content="All" /></ComboBox>
    <TextBox x:Name="SearchBox" Grid.Column="3" Margin="0,0,8,0" MinWidth="220" VerticalContentAlignment="Center" ToolTip="Search iTwin number or name" />
    <Button x:Name="SearchButton" Grid.Column="4" Content="Search" Padding="12,5" IsDefault="True" />
    <Button x:Name="PreviousButton" Grid.Column="5" Content="Previous" Padding="12,5" Margin="8,0,0,0" IsEnabled="False" />
    <Button x:Name="NextButton" Grid.Column="6" Content="Next" Padding="12,5" Margin="8,0,0,0" IsEnabled="False" />
  </Grid>
  <Grid Grid.Row="2">
    <Grid.ColumnDefinitions><ColumnDefinition Width="*" MinWidth="240" /><ColumnDefinition Width="6" /><ColumnDefinition Width="1.5*" MinWidth="320" /><ColumnDefinition Width="6" /><ColumnDefinition Width="*" MinWidth="220" /></Grid.ColumnDefinitions>
    <GroupBox Header="1. Select an iTwin" Grid.Column="0" Padding="7" BorderBrush="#AEB8C2">
      <DockPanel>
        <TextBlock x:Name="iTwinStatusText" DockPanel.Dock="Bottom" Margin="0,6,0,0" TextWrapping="Wrap" Foreground="#3E4C59" />
        <ListBox x:Name="iTwinList" DisplayMemberPath="Label" SelectionMode="Single" VirtualizingStackPanel.IsVirtualizing="True" VirtualizingStackPanel.VirtualizationMode="Recycling" />
      </DockPanel>
    </GroupBox>
    <GridSplitter Grid.Column="1" Width="6" HorizontalAlignment="Stretch" Background="#B6C0CA" />
    <GroupBox Header="2. Select users" Grid.Column="2" Padding="7" BorderBrush="#AEB8C2">
      <DockPanel>
        <Grid DockPanel.Dock="Top" Margin="0,0,0,6">
          <Grid.ColumnDefinitions><ColumnDefinition Width="*" /><ColumnDefinition Width="Auto" /><ColumnDefinition Width="Auto" /></Grid.ColumnDefinitions>
          <TextBox x:Name="UserSearchBox" VerticalContentAlignment="Center" Margin="0,0,8,0" ToolTip="Name or email of a user in your organization (3 to 200 characters)" IsEnabled="False" />
          <Button x:Name="UserSearchButton" Grid.Column="1" Content="Search organization" Padding="10,4" IsEnabled="False" />
          <Button x:Name="AddEmailButton" Grid.Column="2" Content="Enter email..." Padding="10,4" Margin="8,0,0,0" IsEnabled="False" />
        </Grid>
        <TextBlock x:Name="UserStatusText" DockPanel.Dock="Bottom" Margin="0,6,0,0" TextWrapping="Wrap" Foreground="#3E4C59" />
        <ListBox x:Name="UserList" DisplayMemberPath="Label" SelectionMode="Extended" />
      </DockPanel>
    </GroupBox>
    <GridSplitter Grid.Column="3" Width="6" HorizontalAlignment="Stretch" Background="#B6C0CA" />
    <GroupBox Header="3. Select roles" Grid.Column="4" Padding="7" BorderBrush="#AEB8C2">
      <DockPanel>
        <TextBlock x:Name="RoleStatusText" DockPanel.Dock="Bottom" Margin="0,6,0,0" TextWrapping="Wrap" Foreground="#3E4C59" />
        <ListBox x:Name="RoleList" DisplayMemberPath="Label" SelectionMode="Extended" />
      </DockPanel>
    </GroupBox>
  </Grid>
  <DockPanel Grid.Row="3" Margin="0,10,0,0">
    <Button x:Name="AssignButton" Content="Assign selected roles to selected users" Padding="12,5" Margin="8,0,0,0" DockPanel.Dock="Right" Background="#176B4D" Foreground="White" IsEnabled="False" />
    <TextBlock x:Name="StatusText" Text="Select an iTwin to load its members and roles." VerticalAlignment="Center" Foreground="#3E4C59" TextWrapping="Wrap" />
  </DockPanel>
</Grid>
'@

    $script:root = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
    $script:controls = @{}
    foreach ($name in @('ReloadButton', 'SourceTabs', 'OrganizationScopeBox', 'ProjectStatusBox', 'SearchBox', 'SearchButton', 'PreviousButton', 'NextButton', 'iTwinStatusText', 'iTwinList', 'UserSearchBox', 'UserSearchButton', 'AddEmailButton', 'UserStatusText', 'UserList', 'RoleStatusText', 'RoleList', 'AssignButton', 'StatusText')) {
        $controls[$name] = $root.FindName($name)
    }
    $script:state = @{
        Loaded        = $false
        PageIndex     = 0
        NextUris      = [System.Collections.Generic.List[string]]::new()
        SelectediTwin = $null
        Members       = @()
        ExtraUsers    = [ordered]@{}
        Roles         = @()
    }
    $state.NextUris.Add('')

    $controls.SearchButton.Add_Click({ Reset-iTwinPage })
    $controls.OrganizationScopeBox.Add_Click({ Reset-iTwinPage })
    $controls.ProjectStatusBox.Add_SelectionChanged({ Reset-iTwinPage })
    $controls.SourceTabs.Add_SelectionChanged({
            param($source, $e)
            if ($e.OriginalSource -ne $controls.SourceTabs) { return }
            $isFavorites = $controls.SourceTabs.SelectedIndex -eq 1
            foreach ($control in @($controls.OrganizationScopeBox, $controls.SearchBox, $controls.SearchButton)) {
                $control.Visibility = if ($isFavorites) { 'Collapsed' } else { 'Visible' }
            }
            Reset-iTwinPage
        })
    $controls.PreviousButton.Add_Click({ $state.PageIndex--; Update-iTwinPage })
    $controls.NextButton.Add_Click({ $state.PageIndex++; Update-iTwinPage })
    $controls.iTwinList.Add_SelectionChanged({ Update-SelectediTwin })
    $controls.ReloadButton.Add_Click({ Update-SelectediTwin })

    $controls.UserSearchButton.Add_Click({
            try {
                $found = @(Invoke-WithWaitCursor { Find-iTwinOrganizationUsers -SearchText $controls.UserSearchBox.Text })
                foreach ($user in $found) { $state.ExtraUsers[$user.Email.ToLowerInvariant()] = $user }
                Update-UserList
                $controls.StatusText.Text = "Organization search returned $($found.Count) user(s); non-members were added to the user list."
                Write-Host "Organization user search '$($controls.UserSearchBox.Text)' returned $($found.Count) user(s)."
            }
            catch {
                Write-Warning "Organization user search failed: $($_.Exception.Message)"
                Show-Message $_.Exception.Message 'Search organization' 'Warning'
            }
        })

    $controls.AddEmailButton.Add_Click({
            $text = [Microsoft.VisualBasic.Interaction]::InputBox('Enter one or more email addresses separated by commas or semicolons. Users outside your organization are invited by email.', 'Enter email')
            if ([string]::IsNullOrWhiteSpace($text)) { return }
            $invalid = @()
            foreach ($address in @($text -split '[,;\s]+' | Where-Object { $_ })) {
                try { $mailAddress = New-Object System.Net.Mail.MailAddress($address) } catch { $mailAddress = $null }
                if ($null -eq $mailAddress -or $mailAddress.Address -ne $address) { $invalid += $address; continue }
                if (-not $state.ExtraUsers.Contains($address.ToLowerInvariant())) {
                    $state.ExtraUsers[$address.ToLowerInvariant()] = [pscustomobject]@{ Email = $address; GivenName = ''; Surname = ''; Organization = '' }
                }
            }
            Update-UserList
            if ($invalid.Count -gt 0) { Show-Message "These entries are not valid email addresses:`n$($invalid -join "`n")" 'Enter email' 'Warning' }
        })

    $controls.AssignButton.Add_Click({ Invoke-RoleAssignment })

    # Load iTwins the first time the view is visible; Loaded also fires for unselected tab content.
    $root.Add_IsVisibleChanged({
            if ($state.Loaded -or -not $root.IsVisible) { return }
            $state.Loaded = $true
            $null = $root.Dispatcher.BeginInvoke([Action] { Reset-iTwinPage }, [System.Windows.Threading.DispatcherPriority]::ContextIdle)
        })

    [pscustomobject]@{ Root = $root }
}

function Show-iTwinRoleAssignment {
    param(
        [Parameter(Mandatory = $true)][string]$DatabasePath,
        [Parameter(Mandatory = $true)][string]$ProjectConfigPath,
        [Parameter(Mandatory = $true)][string]$AuditLogPath
    )
    if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
        throw 'The role assignment window requires an STA PowerShell host. Start Windows PowerShell with -STA.'
    }
    $view = New-iTwinRoleAssignmentView -DatabasePath $DatabasePath -ProjectConfigPath $ProjectConfigPath -AuditLogPath $AuditLogPath
    $window = New-Object System.Windows.Window -Property @{
        Title = 'iTwin Role Assignment'; Height = 760; Width = 1320; MinHeight = 520; MinWidth = 980
        WindowStartupLocation = 'CenterScreen'; Background = '#F4F6F8'; Content = $view.Root
    }
    $window.ShowDialog() | Out-Null
}

#endregion View

Export-ModuleMember -Function New-iTwinRoleAssignmentView, Show-iTwinRoleAssignment
