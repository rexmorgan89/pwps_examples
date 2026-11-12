Set-StrictMode -Version Latest

# WPF event handlers run in this module's scope, so the single view's controls and state live here.
$script:root = $null
$script:controls = $null
$script:state = $null
$script:pickers = $null
$script:auditLogPath = $null
$script:onCopied = $null

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

function Write-RoleCopyAudit {
    param([array]$Results)
    $timestamp = [DateTime]::UtcNow.ToString('o')
    $rows = @($Results | ForEach-Object {
            [pscustomobject]@{
                TimestampUtc           = $timestamp
                Operator               = [Environment]::UserName
                SourceProjectGUID      = [string]$_.SourceProjectGUID
                DestinationProjectGUID = [string]$_.DestinationProjectGUID
                RoleName               = [string]$_.RoleName
                SourceRoleId           = [string]$_.SourceRoleId
                DestinationRoleId      = [string]$_.DestinationRoleId
                Action                 = [string]$_.Action
                Status                 = [string]$_.Status
                Detail                 = [string]$_.Detail
            }
        })
    if ($rows.Count -gt 0) { $rows | Export-Csv -LiteralPath $auditLogPath -NoTypeInformation -Append -Encoding UTF8 }
}

function Get-TargetText {
    param($Item, [switch]$NoGuid)
    $kind = if ($Item.Kind -eq 'Account') { 'Account' } else { 'iTwin' }
    if ($NoGuid) { return "[$kind] $($Item.ProjectName)" }
    "[$kind] $($Item.ProjectName) ($($Item.ProjectGUID))"
}

#endregion Helpers

#region iTwin pickers

function Set-PickerItems {
    param([string]$Name, [array]$Items)
    $list = $controls["${Name}List"]
    $pickers[$Name].Refreshing = $true
    try {
        $list.ItemsSource = $Items
        foreach ($item in $Items) {
            if ($Name -eq 'Source') {
                if ($null -ne $state.SourceiTwin -and $item.ProjectGUID -eq $state.SourceiTwin.ProjectGUID) { $list.SelectedItem = $item }
            }
            elseif ($state.Destinations.Contains($item.ProjectGUID)) { $list.SelectedItems.Add($item) | Out-Null }
        }
    }
    finally { $pickers[$Name].Refreshing = $false }
}

function Update-PickerPage {
    param([string]$Name)
    $picker = $pickers[$Name]
    $statusText = $controls["${Name}StatusText"]
    $tabIndex = $controls["${Name}Tabs"].SelectedIndex
    $isFavorites = $tabIndex -eq 1
    $subClass = if ($tabIndex -eq 2) { 'Account' } else { 'Project' }
    $noun = if ($subClass -eq 'Account') { 'account(s)' } else { 'iTwin(s)' }
    $queryScope = if ([bool]$controls["${Name}OrganizationScopeBox"].IsChecked) { 'OrganizationAdmin' } else { 'MemberOfiTwin' }
    $projectStatus = [string]$controls["${Name}StatusBox"].SelectedItem.Content
    $searchText = ([string]$controls["${Name}SearchBox"].Text).Trim()
    $nextUri = $picker.NextUris[$picker.PageIndex]
    $fallbackNote = ''
    foreach ($part in @('SearchButton', 'PreviousButton', 'NextButton')) { $controls["$Name$part"].IsEnabled = $false }
    $statusText.Text = "Retrieving $noun..."
    try {
        $page = Invoke-WithWaitCursor {
            if ($isFavorites) { return Get-iTwinFavoriteProjectPage -ProjectStatus $projectStatus -PageSize 250 -NextUri $nextUri }
            try { Get-iTwinAvailableProjectPage -DisplayName $searchText -QueryScope $queryScope -ProjectStatus $projectStatus -PageSize 250 -NextUri $nextUri -SubClass $subClass }
            catch {
                if ($queryScope -ne 'OrganizationAdmin' -or -not (Test-OrganizationAccessDenied $_)) { throw }
                $null
            }
        }
        if ($null -eq $page) {
            $controls["${Name}OrganizationScopeBox"].IsChecked = $false
            $queryScope = 'MemberOfiTwin'
            $picker.PageIndex = 0
            $picker.NextUris.Clear(); $picker.NextUris.Add('')
            $fallbackNote = "Organization access is unavailable; showing your $noun. "
            Write-Warning "Organization access is unavailable. Showing $noun available to your account."
            $page = Invoke-WithWaitCursor { Get-iTwinAvailableProjectPage -DisplayName $searchText -QueryScope $queryScope -ProjectStatus $projectStatus -PageSize 250 -SubClass $subClass }
        }

        $pageItems = @($page.Projects)
        if ($subClass -eq 'Account' -and $picker.PageIndex -eq 0 -and -not $searchText -and $projectStatus -in @('Active', 'All')) {
            # Project members are not always members of the Account iTwin itself, so add the primary account explicitly.
            try {
                $primary = Invoke-WithWaitCursor { Get-iTwinPrimaryAccount }
                if ($primary.ProjectGUID -and @($pageItems | Where-Object { $_.ProjectGUID -eq $primary.ProjectGUID }).Count -eq 0) { $pageItems += $primary }
            }
            catch { Write-Warning "Primary account could not be retrieved: $($_.Exception.Message)" }
        }
        $items = @($pageItems | Sort-Object ProjectName | ForEach-Object {
                $label = if ($subClass -eq 'Account') { "$($_.ProjectName) [Account]" } else { "$($_.ProjectName) [$($_.Status)]" }
                [pscustomobject]@{ ProjectGUID = [string]$_.ProjectGUID; ProjectName = [string]$_.ProjectName; Status = [string]$_.Status; Kind = $subClass; Label = $label }
            })
        Set-PickerItems $Name $items
        if ($picker.NextUris.Count -eq $picker.PageIndex + 1) { $picker.NextUris.Add([string]$page.NextUri) } else { $picker.NextUris[$picker.PageIndex + 1] = [string]$page.NextUri }
        $controls["${Name}PreviousButton"].IsEnabled = $picker.PageIndex -gt 0
        $controls["${Name}NextButton"].IsEnabled = -not [string]::IsNullOrWhiteSpace([string]$page.NextUri)
        $scopeText = if ($isFavorites) { 'favorite' } else { $queryScope }
        $statusText.Text = "${fallbackNote}Page $($picker.PageIndex + 1): $($items.Count) $projectStatus $scopeText $noun."
    }
    catch {
        Set-PickerItems $Name @()
        $statusText.Text = "Retrieval of $noun failed: $($_.Exception.Message)"
        Write-Warning "Retrieval of $noun failed: $($_.Exception.Message)"
    }
    finally { $controls["${Name}SearchButton"].IsEnabled = $true }
}

function Reset-PickerPage {
    param([string]$Name)
    $picker = $pickers[$Name]
    $picker.PageIndex = 0
    $picker.NextUris.Clear()
    $picker.NextUris.Add('')
    Update-PickerPage $Name
}

#endregion iTwin pickers

#region Selection

function Get-DestinationTargets {
    # Overwriting or creating a new role may target the source itself.
    if ($state.Overwrite -or $state.CreateNew) { return @($state.Destinations.Values) }
    @($state.Destinations.Values | Where-Object { $null -eq $state.SourceiTwin -or $_.ProjectGUID -ne $state.SourceiTwin.ProjectGUID })
}

function Get-SingleSourceRole {
    if (@($controls.RoleList.SelectedItems).Count -eq 1) { return $controls.RoleList.SelectedItem }
    $null
}

function Update-CopyState {
    $roleCount = @($controls.RoleList.SelectedItems).Count
    $destinationCount = @(Get-DestinationTargets).Count
    if ($state.Overwrite) {
        $sourceRole = Get-SingleSourceRole
        $targetRole = $controls.OverwriteRoleList.SelectedItem
        $isSameRole = $null -ne $sourceRole -and $null -ne $targetRole -and $sourceRole.RoleId -eq $targetRole.RoleId
        $controls.CopyButton.IsEnabled = $null -ne $state.SourceiTwin -and $null -ne $sourceRole -and $destinationCount -eq 1 -and $null -ne $targetRole -and -not $isSameRole
        $controls.StatusText.Text = if ($controls.CopyButton.IsEnabled) { "Ready to overwrite '$($targetRole.DisplayName)' with the permissions of '$($sourceRole.DisplayName)'." }
        elseif ($isSameRole) { 'The role to overwrite is the source role. Select a different role.' }
        else { 'Select a source iTwin or account, one role, one destination, and the existing role to overwrite.' }
        return
    }
    if ($state.CreateNew) {
        $sourceRole = Get-SingleSourceRole
        $newName = ([string]$controls.NewRoleNameBox.Text).Trim()
        $controls.CopyButton.IsEnabled = $null -ne $state.SourceiTwin -and $null -ne $sourceRole -and $destinationCount -gt 0 -and $newName
        $controls.StatusText.Text = if ($controls.CopyButton.IsEnabled) { "Ready to create '$newName' from '$($sourceRole.DisplayName)' on $destinationCount destination(s)." }
        elseif ($roleCount -gt 1) { 'Select only one role to create a new role from.' }
        elseif ($null -ne $sourceRole -and -not $newName) { 'Enter a name for the new role.' }
        else { 'Select a source iTwin or account, one role, one or more destinations, and enter the new role name.' }
        return
    }
    $controls.CopyButton.IsEnabled = $null -ne $state.SourceiTwin -and $roleCount -gt 0 -and $destinationCount -gt 0
    $sourceIsDestination = $null -ne $state.SourceiTwin -and $state.Destinations.Contains($state.SourceiTwin.ProjectGUID)
    $controls.StatusText.Text = if ($controls.CopyButton.IsEnabled) {
        "Ready to copy $roleCount role(s) to $destinationCount destination(s)."
    }
    else { 'Select a source iTwin or account, one or more roles, and one or more destinations.' }
    if ($sourceIsDestination) { $controls.StatusText.Text += ' The source is also selected as a destination and will be ignored.' }
}

function Update-RolePreview {
    $roles = @($controls.RoleList.SelectedItems | ForEach-Object { $_ })
    $textInfo = [Globalization.CultureInfo]::InvariantCulture.TextInfo
    $items = @($roles | Sort-Object DisplayName | ForEach-Object {
            # Bentley permission names start with their service area, for example imodels_read.
            $groups = @($_.Permissions | Group-Object { ($_ -split '_', 2)[0] } | Sort-Object Name | ForEach-Object {
                    [pscustomobject]@{ Name = $textInfo.ToTitleCase($_.Name); Permissions = [string[]]@($_.Group) }
                })
            [pscustomobject]@{ DisplayName = $_.DisplayName; Summary = $_.Summary; Description = $_.Description; Groups = $groups }
        })
    $controls.PreviewList.ItemsSource = $items
    $controls.PreviewEmptyText.Visibility = if ($items.Count -eq 0) { 'Visible' } else { 'Collapsed' }
    $uniqueCount = @($roles | ForEach-Object { $_.Permissions } | Select-Object -Unique).Count
    $controls.PreviewSummaryText.Text = if ($items.Count -eq 0) { '' } else { "$($items.Count) role(s) selected | $uniqueCount unique permission(s)" }
    if (-not $state.CreateNew) { $controls.NewRoleNameBox.Text = @($roles | ForEach-Object { $_.DisplayName }) -join ', ' }
    Update-OverwritePreview
}

function Update-SourceRoles {
    $itwin = $state.SourceiTwin
    $state.Roles = @()
    $controls.RoleList.ItemsSource = $null
    foreach ($control in @($controls.ReloadRolesButton, $controls.SelectAllRolesButton, $controls.ClearRolesButton)) { $control.IsEnabled = $false }
    Update-RolePreview
    if ($null -eq $itwin) { return }

    $roleKind = if ($itwin.Kind -eq 'Account') { 'enterprise' } else { 'iTwin' }
    $controls.SourceSummaryText.Text = "Source: $(Get-TargetText $itwin)"
    $controls.RoleStatusText.Text = 'Loading roles...'
    try {
        $roles = @(Invoke-WithWaitCursor { Get-iTwinRoleDefinitions -ProjectGuid $itwin.ProjectGUID })
        $state.Roles = @($roles | ForEach-Object {
                $typeText = if ($_.Type) { "$($_.Type) role" } else { 'Role' }
                $_ | Add-Member -NotePropertyName Summary -NotePropertyValue "$typeText | $($_.Permissions.Count) permission(s)" -PassThru
            })
        $controls.RoleList.ItemsSource = $state.Roles
        $controls.RoleStatusText.Text = "$($state.Roles.Count) $roleKind role(s). Click roles to select them."
        Write-Host "Loaded $($state.Roles.Count) $roleKind role(s) from source $(Get-TargetText $itwin)."
    }
    catch {
        $controls.RoleStatusText.Text = "Roles could not be loaded: $($_.Exception.Message)"
        Write-Warning "Roles could not be loaded from $(Get-TargetText $itwin): $($_.Exception.Message)"
    }
    finally { $controls.ReloadRolesButton.IsEnabled = $true }
    $controls.SelectAllRolesButton.IsEnabled = $state.Roles.Count -gt 0
    $controls.ClearRolesButton.IsEnabled = $state.Roles.Count -gt 0
}

function Update-SourceSelection {
    if ($pickers.Source.Refreshing) { return }
    $itwin = $controls.SourceList.SelectedItem
    if ($null -eq $itwin) { return }
    if ($null -ne $state.SourceiTwin -and $state.SourceiTwin.ProjectGUID -eq $itwin.ProjectGUID) { return }
    $state.SourceiTwin = $itwin
    Update-SourceRoles
}

function Update-DestinationSummary {
    $names = @($state.Destinations.Values | ForEach-Object { Get-TargetText $_ -NoGuid })
    $text = 'No destinations selected.'
    if ($names.Count -gt 0) {
        $text = "$($names.Count) selected: $(@($names | Select-Object -First 3) -join ', ')"
        if ($names.Count -gt 3) { $text += ", +$($names.Count - 3) more" }
    }
    $controls.DestinationSelectionText.Text = $text
    $controls.DestinationSelectionText.ToolTip = if ($names.Count -gt 0) { $names -join "`n" } else { $null }
    $controls.DestinationClearButton.IsEnabled = $names.Count -gt 0
    Update-OverwriteRoles
}

function Update-DestinationSelection {
    param($SelectionArgs)
    if ($pickers.Destination.Refreshing) { return }
    foreach ($item in @($SelectionArgs.RemovedItems)) { $state.Destinations.Remove([string]$item.ProjectGUID) }
    foreach ($item in @($SelectionArgs.AddedItems)) { $state.Destinations[[string]$item.ProjectGUID] = $item }
    Update-DestinationSummary
}

#endregion Selection

#region Overwrite existing role

function Update-OverwritePreview {
    $sourceRole = Get-SingleSourceRole
    $targetRole = $controls.OverwriteRoleList.SelectedItem
    $text = if ($null -eq $targetRole) { 'Select the destination role whose permissions will be overwritten.' }
    elseif ($null -eq $sourceRole) { 'Select exactly one role to copy permissions from.' }
    else {
        $sourcePermissions = @($sourceRole.Permissions)
        $currentPermissions = @($targetRole.Permissions)
        $added = @($sourcePermissions | Where-Object { $currentPermissions -notcontains $_ })
        $extra = @($currentPermissions | Where-Object { $sourcePermissions -notcontains $_ })
        $lines = @("'$($targetRole.DisplayName)' gets the permissions of '$($sourceRole.DisplayName)'.")
        $lines += if ($added.Count -gt 0) { "Add ($($added.Count)): $($added -join ', ')" } else { 'Add: none' }
        if ([string]$controls.ExistingRoleBox.SelectedItem.Tag -eq 'Replace') {
            $lines += if ($extra.Count -gt 0) { "Remove ($($extra.Count)): $($extra -join ', ')" } else { 'Remove: none' }
        }
        elseif ($extra.Count -gt 0) { $lines += "Keep ($($extra.Count)) permission(s) not in the source role." }
        $lines -join [Environment]::NewLine
    }
    $controls.OverwriteDiffText.Text = $text
    Update-CopyState
}

function Update-OverwriteRoles {
    param([switch]$Force)
    if (-not $state.Overwrite) {
        Update-CopyState
        return
    }
    $destinations = @($state.Destinations.Values)
    $target = if ($destinations.Count -eq 1) { $destinations[0] } else { $null }
    $targetGuid = if ($null -ne $target) { [string]$target.ProjectGUID } else { '' }
    if (-not $Force -and $targetGuid -eq $state.OverwriteProjectGuid) {
        Update-OverwritePreview
        return
    }
    $state.OverwriteProjectGuid = $targetGuid
    $selectedRoleId = if ($null -ne $controls.OverwriteRoleList.SelectedItem) { [string]$controls.OverwriteRoleList.SelectedItem.RoleId } else { '' }
    $controls.OverwriteRoleList.ItemsSource = $null
    $controls.ReloadOverwriteRolesButton.IsEnabled = $null -ne $target
    $controls.OverwriteStatusText.Text = ''
    if ($null -eq $target) {
        $controls.OverwriteSummaryText.Text = if ($destinations.Count -eq 0) { 'Select one destination to load its roles.' } else { "Select only one destination; $($destinations.Count) are selected." }
        Update-OverwritePreview
        return
    }

    $roleKind = if ($target.Kind -eq 'Account') { 'enterprise' } else { 'iTwin' }
    $controls.OverwriteSummaryText.Text = "Destination: $(Get-TargetText $target)"
    $controls.OverwriteStatusText.Text = 'Loading roles...'
    try {
        $roles = @(Invoke-WithWaitCursor { Get-iTwinRoleDefinitions -ProjectGuid $targetGuid })
        $roles = @($roles | ForEach-Object {
                $typeText = if ($_.Type) { "$($_.Type) role" } else { 'Role' }
                $_ | Add-Member -NotePropertyName Summary -NotePropertyValue "$typeText | $($_.Permissions.Count) permission(s)" -PassThru
            })
        $controls.OverwriteRoleList.ItemsSource = $roles
        $controls.OverwriteRoleList.SelectedItem = $roles | Where-Object { $selectedRoleId -and $_.RoleId -eq $selectedRoleId } | Select-Object -First 1
        $controls.OverwriteStatusText.Text = "$($roles.Count) $roleKind role(s). Select the role to overwrite."
        Write-Host "Loaded $($roles.Count) $roleKind role(s) from destination $(Get-TargetText $target) for overwrite."
    }
    catch {
        $controls.OverwriteStatusText.Text = "Roles could not be loaded: $($_.Exception.Message)"
        Write-Warning "Roles could not be loaded from $(Get-TargetText $target): $($_.Exception.Message)"
    }
    Update-OverwritePreview
}

function Set-OverwriteMode {
    $state.Overwrite = [bool]$controls.OverwriteCheckBox.IsChecked
    $on = $state.Overwrite
    if ($on -and $state.CreateNew) {
        $controls.CreateNewCheckBox.IsChecked = $false
        Set-CreateNewMode
    }
    $controls.OverwriteSplitterColumn.Width = [System.Windows.GridLength]::new($(if ($on) { 8 } else { 0 }))
    $controls.OverwriteColumn.Width = if ($on) { [System.Windows.GridLength]::new(1, 'Star') } else { [System.Windows.GridLength]::new(0) }
    $controls.OverwriteColumn.MinWidth = if ($on) { 220 } else { 0 }
    foreach ($control in @($controls.OverwriteSplitter, $controls.OverwritePane)) { $control.Visibility = if ($on) { 'Visible' } else { 'Collapsed' } }
    $controls.PreviewTitleText.Text = if ($on) { '5. Permission preview' } else { '4. Permission preview' }
    $controls.CopyButton.Content = if ($on) { 'Overwrite role...' } else { 'Copy selected roles...' }
    $controls.ExistingRoleLabel.Text = if ($on) { 'Overwrite mode:' } else { 'Existing roles:' }
    # Skip does nothing when overwriting, so default to Replace and restore the previous choice afterwards.
    $controls.ExistingRoleBox.Items[0].IsEnabled = -not $on
    if ($on) {
        $state.ExistingRoleIndex = $controls.ExistingRoleBox.SelectedIndex
        if ($state.ExistingRoleIndex -eq 0) { $controls.ExistingRoleBox.SelectedIndex = 2 }
        Update-OverwriteRoles -Force
    }
    else {
        $controls.ExistingRoleBox.SelectedIndex = $state.ExistingRoleIndex
        Update-CopyState
    }
}

#endregion Overwrite existing role

#region Create new role

function Set-CreateNewMode {
    $state.CreateNew = [bool]$controls.CreateNewCheckBox.IsChecked
    $on = $state.CreateNew
    if ($on -and $state.Overwrite) {
        $controls.OverwriteCheckBox.IsChecked = $false
        Set-OverwriteMode
    }
    $controls.NewRoleNameBox.IsReadOnly = -not $on
    $controls.NewRoleNameBox.Background = if ($on) { 'White' } else { '#F4F6F8' }
    $controls.NewRoleNameBox.Foreground = if ($on) { '#17324D' } else { '#52606D' }
    $controls.NewRoleNameLabel.Text = if ($on) { 'New role name:' } else { 'Role to copy:' }
    $controls.CopyButton.Content = if ($on) { 'Create role...' } else { 'Copy selected roles...' }
    if ($on) {
        $controls.NewRoleNameBox.Focus() | Out-Null
        $controls.NewRoleNameBox.SelectAll()
        Update-CopyState
    }
    else { Update-RolePreview }
}

#endregion Create new role

#region Copy roles

function Format-RoleCopyText {
    param($Source, [array]$Roles, [array]$Destinations, [string]$ExistingRoleText, [string]$Heading, [array]$Rows, $TargetRole, [string]$NewRoleName)
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add('SOURCE')
    $lines.Add("  $(Get-TargetText $Source)")
    $lines.Add('')
    $lines.Add("DESTINATIONS: $($Destinations.Count)")
    foreach ($destination in $Destinations) { $lines.Add("  $(Get-TargetText $destination)") }
    $lines.Add('')
    $lines.Add("ROLES: $($Roles.Count)")
    foreach ($role in $Roles) { $lines.Add("  $($role.DisplayName) ($($role.Permissions.Count) permission(s))") }
    $lines.Add('')
    if ($null -ne $TargetRole) {
        $lines.Add('ROLE TO OVERWRITE')
        $lines.Add("  $($TargetRole.DisplayName) ($($TargetRole.Permissions.Count) permission(s), $($TargetRole.RoleId))")
        $lines.Add('')
    }
    if ($NewRoleName) {
        $lines.Add('NEW ROLE NAME')
        $lines.Add("  $NewRoleName")
        $lines.Add('')
    }
    $lines.Add("EXISTING ROLES: $ExistingRoleText")
    $lines.Add('')
    $lines.Add($Heading)
    foreach ($destination in $Destinations) {
        $lines.Add("  $(Get-TargetText $destination -NoGuid)")
        foreach ($row in @($Rows | Where-Object { $_.DestinationProjectGUID -eq $destination.ProjectGUID })) {
            $lines.Add(('    {0,-8} {1,-32} {2}' -f $row.Status, $row.RoleName, $row.Detail))
        }
    }
    $lines -join [Environment]::NewLine
}

function Show-RoleCopyDialog {
    param([string]$Title, [string]$Heading, [string]$Body, [string]$ConfirmText)
    [xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Width="820" Height="620" MinWidth="560" MinHeight="380" WindowStartupLocation="CenterOwner" ShowInTaskbar="False" ResizeMode="CanResizeWithGrip" Background="#F4F6F8">
  <Grid Margin="16">
    <Grid.RowDefinitions><RowDefinition Height="Auto" /><RowDefinition Height="*" /><RowDefinition Height="Auto" /></Grid.RowDefinitions>
    <TextBlock x:Name="HeadingText" FontSize="15" FontWeight="SemiBold" Foreground="#17324D" TextWrapping="Wrap" Margin="0,0,0,10" />
    <Border Grid.Row="1" Background="White" BorderBrush="#D6DCE2" BorderThickness="1" CornerRadius="6" Padding="4">
      <TextBox x:Name="BodyBox" IsReadOnly="True" BorderThickness="0" FontFamily="Consolas" FontSize="12" TextWrapping="NoWrap" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" />
    </Border>
    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0">
      <Button x:Name="CancelButton" Content="Cancel" Padding="14,6" MinWidth="90" IsCancel="True" />
      <Button x:Name="ConfirmButton" Padding="14,6" MinWidth="90" Margin="8,0,0,0" Background="#176B4D" BorderBrush="#176B4D" Foreground="White" FontWeight="SemiBold" />
    </StackPanel>
  </Grid>
</Window>
'@
    $dialog = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
    $dialog.Title = $Title
    $dialog.Owner = [System.Windows.Window]::GetWindow($root)
    $dialog.FindName('HeadingText').Text = $Heading
    $dialog.FindName('BodyBox').Text = $Body
    $confirmButton = $dialog.FindName('ConfirmButton')
    if ([string]::IsNullOrWhiteSpace($ConfirmText)) {
        $dialog.FindName('CancelButton').Content = 'Close'
        $confirmButton.Visibility = 'Collapsed'
    }
    else {
        $confirmButton.Content = $ConfirmText
        $confirmButton.Add_Click({ param($control) [System.Windows.Window]::GetWindow($control).DialogResult = $true })
    }
    [bool]$dialog.ShowDialog()
}

function Invoke-RoleCopy {
    $source = $state.SourceiTwin
    $roles = @($controls.RoleList.SelectedItems | ForEach-Object { $_ } | Sort-Object DisplayName)
    $destinations = @(Get-DestinationTargets | Sort-Object ProjectName)
    if ($null -eq $source -or $roles.Count -eq 0 -or $destinations.Count -eq 0) {
        Show-Message 'Select a source iTwin or account, one or more roles, and one or more destinations.' 'Copy roles'
        return
    }
    $existingRoleItem = $controls.ExistingRoleBox.SelectedItem
    $existingRoleText = [string]$existingRoleItem.Content
    $copyParameters = @{
        SourceProjectGuid      = $source.ProjectGUID
        Role                   = $roles
        DestinationProjectGuid = @($destinations | ForEach-Object { $_.ProjectGUID })
        ExistingRoleAction     = [string]$existingRoleItem.Tag
    }
    $targetRole = $null
    if ($state.Overwrite) {
        $targetRole = $controls.OverwriteRoleList.SelectedItem
        if ($roles.Count -ne 1 -or $destinations.Count -ne 1 -or $null -eq $targetRole) {
            Show-Message 'Select one role, one destination, and the existing role to overwrite.' 'Overwrite role'
            return
        }
        $copyParameters.DestinationRoleId = [string]$targetRole.RoleId
    }
    $newRoleName = ''
    if ($state.CreateNew) {
        $newRoleName = ([string]$controls.NewRoleNameBox.Text).Trim()
        if ($roles.Count -ne 1 -or -not $newRoleName) {
            Show-Message 'Select one role and enter a name for the new role.' 'Create role'
            return
        }
        $copyParameters.NewRoleName = $newRoleName
    }
    $title = if ($state.Overwrite) { 'Overwrite role' } elseif ($state.CreateNew) { 'Create role' } else { 'Copy roles' }

    $summary = $null
    $controls.CopyButton.IsEnabled = $false
    try {
        $controls.StatusText.Text = "Checking existing roles on $($destinations.Count) destination(s)..."
        try { $plan = @(Invoke-WithWaitCursor { Copy-iTwinRoleDefinitions @copyParameters -WhatIf }) }
        catch {
            Write-Warning "Role copy plan failed: $($_.Exception.Message)"
            Show-Message $_.Exception.Message $title 'Error'
            return
        }
        $plannedCount = @($plan | Where-Object Status -eq 'Planned').Count
        $planText = Format-RoleCopyText $source $roles $destinations $existingRoleText 'PLANNED CHANGES' $plan -TargetRole $targetRole -NewRoleName $newRoleName
        if ($plannedCount -eq 0) {
            Show-RoleCopyDialog -Title $title -Heading 'No changes are needed. Review the plan below for skipped or failed roles.' -Body $planText | Out-Null
            return
        }
        if ($state.Overwrite) {
            $heading = "Overwrite the permissions of '$($targetRole.DisplayName)' on '$($destinations[0].ProjectName)' with those of '$($roles[0].DisplayName)'? $plannedCount change(s) will be made in Bentley."
            if ($destinations[0].Kind -eq 'Account') { $heading += ' Enterprise role changes apply on every iTwin in that account.' }
            $confirmText = 'Overwrite role'
        }
        elseif ($state.CreateNew) {
            $heading = "Create role '$newRoleName' from '$($roles[0].DisplayName)' in '$($source.ProjectName)' on $($destinations.Count) destination(s)? $plannedCount change(s) will be made in Bentley."
            if (@($destinations | Where-Object Kind -eq 'Account').Count -gt 0) {
                $heading += ' Roles created on an account become enterprise roles available on every iTwin in that account.'
            }
            $confirmText = 'Create role'
        }
        else {
            $heading = "Copy $($roles.Count) role(s) from '$($source.ProjectName)' to $($destinations.Count) destination(s)? $plannedCount change(s) will be made in Bentley."
            if (@($destinations | Where-Object Kind -eq 'Account').Count -gt 0) {
                $heading += ' Roles copied to an account become enterprise roles available on every iTwin in that account.'
            }
            if ($source.Kind -eq 'Account' -and @($destinations | Where-Object Kind -ne 'Account').Count -gt 0) {
                $heading += " iTwins in '$($source.ProjectName)' already inherit its enterprise roles; copying creates separate iTwin roles with the same names."
            }
            $confirmText = 'Copy roles'
        }
        if (-not (Show-RoleCopyDialog -Title "Confirm $($title.ToLowerInvariant())" -Heading $heading -Body $planText -ConfirmText $confirmText)) {
            Write-Host "Role copy cancelled for source $(Get-TargetText $source)."
            return
        }

        if ($state.Overwrite) {
            Write-Host "Overwriting role '$($targetRole.DisplayName)' ($($targetRole.RoleId)) on $(Get-TargetText $destinations[0]) with role '$($roles[0].DisplayName)' from $(Get-TargetText $source); mode: $existingRoleText."
        }
        elseif ($state.CreateNew) {
            Write-Host "Creating role '$newRoleName' from role '$($roles[0].DisplayName)' in $(Get-TargetText $source) on $($destinations.Count) destination(s); existing roles: $existingRoleText."
        }
        else {
            Write-Host "Copying $($roles.Count) role(s) from $(Get-TargetText $source) to $($destinations.Count) destination(s); existing roles: $existingRoleText."
        }
        $controls.StatusText.Text = "Copying roles to $($destinations.Count) destination(s)..."
        try { $results = @(Invoke-WithWaitCursor { Copy-iTwinRoleDefinitions @copyParameters -Confirm:$false }) }
        catch {
            Write-Warning "Role copy failed before any result was returned: $($_.Exception.Message)"
            Show-Message $_.Exception.Message $title 'Error'
            return
        }
        foreach ($result in $results) { Write-Host "  $($result.Status) | $($result.DestinationProjectGUID) | $($result.RoleName) | $($result.Detail)" }
        $auditNote = "Results were written to '$auditLogPath'."
        try { Write-RoleCopyAudit -Results $results }
        catch {
            $auditNote = "Results were not written to the audit CSV: $($_.Exception.Message)"
            Write-Warning $auditNote
        }
        $summary = "Role copy completed: $(@($results | Where-Object Status -eq 'Success').Count) succeeded, $(@($results | Where-Object Status -eq 'Skipped').Count) skipped, $(@($results | Where-Object Status -eq 'Failed').Count) failed."
        Write-Host $summary
        Show-RoleCopyDialog -Title "$title results" -Heading "$summary $auditNote" -Body (Format-RoleCopyText $source $roles $destinations $existingRoleText 'RESULTS' $results -TargetRole $targetRole -NewRoleName $newRoleName) | Out-Null
        if ($state.Overwrite) {
            # Show the updated permissions; the source list is stale too when the role was overwritten on the source itself.
            Update-OverwriteRoles -Force
            if ($destinations[0].ProjectGUID -eq $source.ProjectGUID) { Update-SourceRoles }
        }
        elseif ($state.CreateNew -and @($destinations | Where-Object ProjectGUID -eq $source.ProjectGUID).Count -gt 0) { Update-SourceRoles }
        if ($null -ne $onCopied) { & $onCopied }
    }
    finally {
        Update-CopyState
        if ($null -ne $summary) { $controls.StatusText.Text = $summary }
    }
}

#endregion Copy roles

#region View

function New-iTwinRoleCopyView {
    param(
        [Parameter(Mandatory = $true)][string]$AuditLogPath,
        [scriptblock]$OnCopied
    )
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    $script:auditLogPath = $AuditLogPath
    $script:onCopied = $OnCopied

    $pickerTemplate = @'
    <Border Style="{StaticResource Card}" Grid.Column="{COLUMN}">
      <DockPanel>
        <TextBlock DockPanel.Dock="Top" Style="{StaticResource CardTitle}" Text="{TITLE}" />
        <TabControl x:Name="{P}Tabs" DockPanel.Dock="Top" Height="28" SelectedIndex="0" Margin="0,0,0,6"><TabItem Header="All iTwins" /><TabItem Header="My Favorites" /><TabItem Header="Accounts" ToolTip="Account iTwins that hold enterprise roles" /></TabControl>
        <DockPanel DockPanel.Dock="Top" Margin="0,0,0,6">
          <ComboBox x:Name="{P}StatusBox" DockPanel.Dock="Right" Width="88" SelectedIndex="0"><ComboBoxItem Content="Active" /><ComboBoxItem Content="Trial" /><ComboBoxItem Content="Inactive" /><ComboBoxItem Content="All" /></ComboBox>
          <CheckBox x:Name="{P}OrganizationScopeBox" Content="Show organization iTwins" VerticalAlignment="Center" />
        </DockPanel>
        <DockPanel DockPanel.Dock="Top" Margin="0,0,0,6">
          <Button x:Name="{P}SearchButton" DockPanel.Dock="Right" Content="Search" Margin="6,0,0,0" />
          <TextBox x:Name="{P}SearchBox" VerticalContentAlignment="Center" Padding="4,3" ToolTip="Search iTwin or account number or name, then press Enter" />
        </DockPanel>
        <DockPanel DockPanel.Dock="Bottom" Margin="0,6,0,0">
          <Button x:Name="{P}NextButton" DockPanel.Dock="Right" Content="Next" Margin="6,0,0,0" IsEnabled="False" />
          <Button x:Name="{P}PreviousButton" DockPanel.Dock="Right" Content="Previous" IsEnabled="False" />
          <TextBlock x:Name="{P}StatusText" Style="{StaticResource Hint}" VerticalAlignment="Center" />
        </DockPanel>
        {EXTRA}
        <ListBox x:Name="{P}List" DisplayMemberPath="Label" SelectionMode="{MODE}" BorderBrush="#D6DCE2" VirtualizingStackPanel.IsVirtualizing="True" VirtualizingStackPanel.VirtualizationMode="Recycling" />
      </DockPanel>
    </Border>
'@
    $destinationExtra = @'
<DockPanel DockPanel.Dock="Bottom" Margin="0,6,0,0">
          <Button x:Name="DestinationClearButton" DockPanel.Dock="Right" Content="Clear selection" Margin="6,0,0,0" IsEnabled="False" />
          <TextBlock x:Name="DestinationSelectionText" Style="{StaticResource Hint}" FontWeight="SemiBold" VerticalAlignment="Center" Text="No destinations selected." />
        </DockPanel>
'@
    $sourcePicker = $pickerTemplate.Replace('{P}', 'Source').Replace('{COLUMN}', '0').Replace('{TITLE}', '1. Source iTwin or account').Replace('{MODE}', 'Single').Replace('{EXTRA}', '')
    $destinationPicker = $pickerTemplate.Replace('{P}', 'Destination').Replace('{COLUMN}', '4').Replace('{TITLE}', '3. Destinations (click to select several)').Replace('{MODE}', 'Multiple').Replace('{EXTRA}', $destinationExtra)

    $layout = @'
<Grid xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Margin="14">
  <Grid.Resources>
    <Style TargetType="Button">
      <Setter Property="Padding" Value="12,5" />
      <Setter Property="Background" Value="White" />
      <Setter Property="BorderBrush" Value="#AEB8C2" />
      <Setter Property="Foreground" Value="#17324D" />
      <Setter Property="Cursor" Value="Hand" />
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Chrome" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="4" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" />
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Chrome" Property="BorderBrush" Value="#2F6FAF" /></Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter TargetName="Chrome" Property="Opacity" Value="0.45" /></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="PrimaryButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
      <Setter Property="Background" Value="#176B4D" />
      <Setter Property="BorderBrush" Value="#176B4D" />
      <Setter Property="Foreground" Value="White" />
      <Setter Property="FontWeight" Value="SemiBold" />
    </Style>
    <Style x:Key="Card" TargetType="Border">
      <Setter Property="Background" Value="White" />
      <Setter Property="BorderBrush" Value="#D6DCE2" />
      <Setter Property="BorderThickness" Value="1" />
      <Setter Property="CornerRadius" Value="6" />
      <Setter Property="Padding" Value="10" />
    </Style>
    <Style x:Key="CardTitle" TargetType="TextBlock">
      <Setter Property="FontSize" Value="15" />
      <Setter Property="FontWeight" Value="SemiBold" />
      <Setter Property="Foreground" Value="#17324D" />
      <Setter Property="Margin" Value="0,0,0,8" />
    </Style>
    <Style x:Key="Hint" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#52606D" />
      <Setter Property="TextWrapping" Value="Wrap" />
    </Style>
  </Grid.Resources>
  <Grid.RowDefinitions><RowDefinition Height="Auto" /><RowDefinition Height="*" MinHeight="260" /><RowDefinition Height="8" /><RowDefinition Height="240" MinHeight="110" /><RowDefinition Height="Auto" /></Grid.RowDefinitions>
  <StackPanel Grid.Row="0" Margin="0,0,0,10">
    <TextBlock Text="iTwin Role Copy" FontSize="22" FontWeight="SemiBold" Foreground="#17324D" />
    <TextBlock Style="{StaticResource Hint}" Text="Copy role definitions and their permissions between iTwins and accounts. Use the Accounts tab to copy enterprise (account-level) roles, or to send roles to an account. Members and role assignments are not copied." />
  </StackPanel>
  <Grid Grid.Row="1">
    <Grid.ColumnDefinitions><ColumnDefinition Width="*" MinWidth="260" /><ColumnDefinition Width="8" /><ColumnDefinition Width="*" MinWidth="240" /><ColumnDefinition Width="8" /><ColumnDefinition Width="*" MinWidth="260" /><ColumnDefinition x:Name="OverwriteSplitterColumn" Width="0" /><ColumnDefinition x:Name="OverwriteColumn" Width="0" /></Grid.ColumnDefinitions>
{SOURCE_PICKER}
    <GridSplitter Grid.Column="1" Width="8" HorizontalAlignment="Stretch" Background="Transparent" />
    <Border Style="{StaticResource Card}" Grid.Column="2">
      <DockPanel>
        <DockPanel DockPanel.Dock="Top" Margin="0,0,0,8">
          <Button x:Name="ReloadRolesButton" DockPanel.Dock="Right" Content="Reload" IsEnabled="False" />
          <TextBlock Style="{StaticResource CardTitle}" Margin="0" VerticalAlignment="Center" Text="2. Roles to copy" />
        </DockPanel>
        <TextBlock x:Name="SourceSummaryText" DockPanel.Dock="Top" Style="{StaticResource Hint}" Margin="0,0,0,6" Text="Select a source iTwin or account to load its roles." />
        <DockPanel DockPanel.Dock="Bottom" Margin="0,6,0,0">
          <Button x:Name="ClearRolesButton" DockPanel.Dock="Right" Content="Clear" Margin="6,0,0,0" IsEnabled="False" />
          <Button x:Name="SelectAllRolesButton" DockPanel.Dock="Right" Content="Select all" IsEnabled="False" />
          <TextBlock x:Name="RoleStatusText" Style="{StaticResource Hint}" VerticalAlignment="Center" />
        </DockPanel>
        <ListBox x:Name="RoleList" SelectionMode="Multiple" HorizontalContentAlignment="Stretch" BorderBrush="#D6DCE2">
          <ListBox.ItemTemplate>
            <DataTemplate>
              <StackPanel Margin="2,3">
                <TextBlock Text="{Binding DisplayName}" FontWeight="SemiBold" Foreground="#17324D" />
                <TextBlock Text="{Binding Summary}" FontSize="11" Foreground="#52606D" />
              </StackPanel>
            </DataTemplate>
          </ListBox.ItemTemplate>
        </ListBox>
      </DockPanel>
    </Border>
    <GridSplitter Grid.Column="3" Width="8" HorizontalAlignment="Stretch" Background="Transparent" />
{DESTINATION_PICKER}
    <GridSplitter x:Name="OverwriteSplitter" Grid.Column="5" Width="8" HorizontalAlignment="Stretch" Background="Transparent" Visibility="Collapsed" />
    <Border x:Name="OverwritePane" Style="{StaticResource Card}" Grid.Column="6" Visibility="Collapsed">
      <DockPanel>
        <DockPanel DockPanel.Dock="Top" Margin="0,0,0,8">
          <Button x:Name="ReloadOverwriteRolesButton" DockPanel.Dock="Right" Content="Reload" IsEnabled="False" />
          <TextBlock Style="{StaticResource CardTitle}" Margin="0" VerticalAlignment="Center" Text="4. Existing role to overwrite" />
        </DockPanel>
        <TextBlock x:Name="OverwriteSummaryText" DockPanel.Dock="Top" Style="{StaticResource Hint}" Margin="0,0,0,6" Text="Select one destination to load its roles." />
        <Border DockPanel.Dock="Bottom" Background="#FFF8E6" BorderBrush="#F0D58C" BorderThickness="1" CornerRadius="4" Padding="8,6" Margin="0,6,0,0" MaxHeight="150">
          <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled"><TextBlock x:Name="OverwriteDiffText" TextWrapping="Wrap" Foreground="#3E4C59" /></ScrollViewer>
        </Border>
        <TextBlock x:Name="OverwriteStatusText" DockPanel.Dock="Bottom" Style="{StaticResource Hint}" Margin="0,6,0,0" />
        <ListBox x:Name="OverwriteRoleList" SelectionMode="Single" HorizontalContentAlignment="Stretch" BorderBrush="#D6DCE2" ToolTip="Its permissions are overwritten; its name, description, and members are kept">
          <ListBox.ItemTemplate>
            <DataTemplate>
              <StackPanel Margin="2,3">
                <TextBlock Text="{Binding DisplayName}" FontWeight="SemiBold" Foreground="#17324D" />
                <TextBlock Text="{Binding Summary}" FontSize="11" Foreground="#52606D" />
              </StackPanel>
            </DataTemplate>
          </ListBox.ItemTemplate>
        </ListBox>
      </DockPanel>
    </Border>
  </Grid>
  <GridSplitter Grid.Row="2" Height="8" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Background="Transparent" ResizeDirection="Rows" />
  <Border Grid.Row="3" Style="{StaticResource Card}">
    <DockPanel>
      <DockPanel DockPanel.Dock="Top" Margin="0,0,0,6">
        <TextBlock x:Name="PreviewSummaryText" DockPanel.Dock="Right" Style="{StaticResource Hint}" VerticalAlignment="Center" />
        <TextBlock x:Name="PreviewTitleText" Style="{StaticResource CardTitle}" Margin="0" Text="4. Permission preview" />
      </DockPanel>
      <Grid>
        <TextBlock x:Name="PreviewEmptyText" Style="{StaticResource Hint}" Text="Select one or more roles to preview their permissions." />
        <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
          <ItemsControl x:Name="PreviewList">
            <ItemsControl.ItemTemplate>
              <DataTemplate>
                <Border BorderBrush="#E4E8EC" BorderThickness="0,0,0,1" Padding="0,6,0,8">
                  <StackPanel>
                    <TextBlock><Run Text="{Binding DisplayName, Mode=OneWay}" FontWeight="SemiBold" FontSize="13" Foreground="#17324D" /><Run Text="   " /><Run Text="{Binding Summary, Mode=OneWay}" Foreground="#52606D" /></TextBlock>
                    <TextBlock Text="{Binding Description}" Style="{StaticResource Hint}" Margin="0,2,0,4" />
                    <ItemsControl ItemsSource="{Binding Groups}">
                      <ItemsControl.ItemTemplate>
                        <DataTemplate>
                          <DockPanel Margin="0,2">
                            <TextBlock Text="{Binding Name}" Width="150" FontWeight="SemiBold" Foreground="#3E4C59" VerticalAlignment="Top" Margin="0,3,0,0" />
                            <ItemsControl ItemsSource="{Binding Permissions}">
                              <ItemsControl.ItemsPanel><ItemsPanelTemplate><WrapPanel /></ItemsPanelTemplate></ItemsControl.ItemsPanel>
                              <ItemsControl.ItemTemplate>
                                <DataTemplate>
                                  <Border Background="#E8F1FB" BorderBrush="#C5DCF3" BorderThickness="1" CornerRadius="10" Padding="8,1" Margin="0,2,6,2">
                                    <TextBlock Text="{Binding}" FontSize="11" Foreground="#17324D" />
                                  </Border>
                                </DataTemplate>
                              </ItemsControl.ItemTemplate>
                            </ItemsControl>
                          </DockPanel>
                        </DataTemplate>
                      </ItemsControl.ItemTemplate>
                    </ItemsControl>
                  </StackPanel>
                </Border>
              </DataTemplate>
            </ItemsControl.ItemTemplate>
          </ItemsControl>
        </ScrollViewer>
      </Grid>
    </DockPanel>
  </Border>
  <Border Grid.Row="4" Style="{StaticResource Card}" Margin="0,10,0,0" Padding="10,8">
    <DockPanel>
      <DockPanel DockPanel.Dock="Top" Margin="0,0,0,8">
        <CheckBox x:Name="CreateNewCheckBox" DockPanel.Dock="Left" Content="Create New Role" Width="180" VerticalAlignment="Center" ToolTip="Copy one source role to the destinations under a new name" />
        <TextBlock x:Name="NewRoleNameLabel" DockPanel.Dock="Left" Text="Role to copy:" Width="98" VerticalAlignment="Center" Foreground="#3E4C59" />
        <TextBox x:Name="NewRoleNameBox" IsReadOnly="True" Width="290" HorizontalAlignment="Left" Padding="4,3" VerticalContentAlignment="Center" Background="#F4F6F8" Foreground="#52606D" BorderBrush="#AEB8C2" ToolTip="Check Create New Role to enter a new name; otherwise this shows the roles that would be copied" />
      </DockPanel>
      <Button x:Name="CopyButton" DockPanel.Dock="Right" Style="{StaticResource PrimaryButton}" Content="Copy selected roles..." Padding="16,6" Margin="8,0,0,0" IsEnabled="False" />
      <CheckBox x:Name="OverwriteCheckBox" DockPanel.Dock="Left" Content="Overwrite Existing Role" Width="180" VerticalAlignment="Center" ToolTip="Overwrite the permissions of one chosen destination role, whatever its name, with one source role's permissions" />
      <TextBlock x:Name="ExistingRoleLabel" DockPanel.Dock="Left" Text="Existing roles:" Width="98" VerticalAlignment="Center" Foreground="#3E4C59" />
      <ComboBox x:Name="ExistingRoleBox" DockPanel.Dock="Left" Width="290" SelectedIndex="0" VerticalContentAlignment="Center" ToolTip="Destination roles are matched to source roles by display name, ignoring case">
        <ComboBoxItem Content="Skip roles that already exist" Tag="Skip" />
        <ComboBoxItem Content="Add missing permissions to existing roles" Tag="Merge" />
        <ComboBoxItem Content="Replace permissions of existing roles" Tag="Replace" />
      </ComboBox>
      <TextBlock x:Name="StatusText" Style="{StaticResource Hint}" Margin="14,0,0,0" VerticalAlignment="Center" />
    </DockPanel>
  </Border>
</Grid>
'@
    [xml]$xaml = $layout.Replace('{SOURCE_PICKER}', $sourcePicker).Replace('{DESTINATION_PICKER}', $destinationPicker)

    $script:root = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
    $script:controls = @{}
    $pickerParts = @('Tabs', 'StatusBox', 'OrganizationScopeBox', 'SearchButton', 'SearchBox', 'NextButton', 'PreviousButton', 'StatusText', 'List')
    $names = @('ReloadRolesButton', 'SourceSummaryText', 'ClearRolesButton', 'SelectAllRolesButton', 'RoleStatusText', 'RoleList', 'PreviewSummaryText', 'PreviewEmptyText', 'PreviewList', 'CopyButton', 'ExistingRoleBox', 'StatusText', 'DestinationClearButton', 'DestinationSelectionText', 'PreviewTitleText', 'OverwriteCheckBox', 'ExistingRoleLabel', 'OverwriteSplitterColumn', 'OverwriteColumn', 'OverwriteSplitter', 'OverwritePane', 'ReloadOverwriteRolesButton', 'OverwriteSummaryText', 'OverwriteDiffText', 'OverwriteStatusText', 'OverwriteRoleList', 'CreateNewCheckBox', 'NewRoleNameLabel', 'NewRoleNameBox')
    foreach ($pickerName in @('Source', 'Destination')) { $names += @($pickerParts | ForEach-Object { "$pickerName$_" }) }
    foreach ($name in $names) {
        $controls[$name] = $root.FindName($name)
        if ($null -eq $controls[$name]) { throw "Role copy view control '$name' was not found." }
    }
    $script:state = @{
        Loaded               = $false
        SourceiTwin          = $null
        Roles                = @()
        Destinations         = [ordered]@{}
        Overwrite            = $false
        CreateNew            = $false
        OverwriteProjectGuid = $null
        ExistingRoleIndex    = 0
    }
    $script:pickers = @{}
    foreach ($pickerName in @('Source', 'Destination')) {
        $pickers[$pickerName] = @{ PageIndex = 0; NextUris = [System.Collections.Generic.List[string]]::new(); Refreshing = $false }
        $pickers[$pickerName].NextUris.Add('')
        # Shared handlers find their picker through the control Tag.
        foreach ($part in $pickerParts) { $controls["$pickerName$part"].Tag = $pickerName }
    }
    Update-CopyState

    #region Events

    foreach ($pickerName in @('Source', 'Destination')) {
        $controls["${pickerName}SearchButton"].Add_Click({ param($control) Reset-PickerPage ([string]$control.Tag) })
        $controls["${pickerName}OrganizationScopeBox"].Add_Click({ param($control) Reset-PickerPage ([string]$control.Tag) })
        $controls["${pickerName}StatusBox"].Add_SelectionChanged({ param($control) Reset-PickerPage ([string]$control.Tag) })
        $controls["${pickerName}SearchBox"].Add_KeyDown({
                param($control, $e)
                if ($e.Key -eq [System.Windows.Input.Key]::Enter) { Reset-PickerPage ([string]$control.Tag) }
            })
        $controls["${pickerName}Tabs"].Add_SelectionChanged({
                param($control, $e)
                if ($e.OriginalSource -ne $control) { return }
                $name = [string]$control.Tag
                $visibility = if ($control.SelectedIndex -eq 1) { 'Collapsed' } else { 'Visible' }
                foreach ($part in @('OrganizationScopeBox', 'SearchBox', 'SearchButton')) { $controls["$name$part"].Visibility = $visibility }
                Reset-PickerPage $name
            })
        $controls["${pickerName}PreviousButton"].Add_Click({
                param($control)
                $picker = $pickers[[string]$control.Tag]
                $picker.PageIndex = $picker.PageIndex - 1
                Update-PickerPage ([string]$control.Tag)
            })
        $controls["${pickerName}NextButton"].Add_Click({
                param($control)
                $picker = $pickers[[string]$control.Tag]
                $picker.PageIndex = $picker.PageIndex + 1
                Update-PickerPage ([string]$control.Tag)
            })
    }
    $controls.SourceList.Add_SelectionChanged({ Update-SourceSelection })
    $controls.DestinationList.Add_SelectionChanged({ param($control, $e) Update-DestinationSelection $e })
    $controls.DestinationClearButton.Add_Click({
            $state.Destinations.Clear()
            $controls.DestinationList.UnselectAll()
            Update-DestinationSummary
        })
    $controls.ReloadRolesButton.Add_Click({ Update-SourceRoles })
    $controls.RoleList.Add_SelectionChanged({ Update-RolePreview })
    $controls.SelectAllRolesButton.Add_Click({ $controls.RoleList.SelectAll() })
    $controls.ClearRolesButton.Add_Click({ $controls.RoleList.UnselectAll() })
    $controls.CopyButton.Add_Click({ Invoke-RoleCopy })
    $controls.OverwriteCheckBox.Add_Click({ Set-OverwriteMode })
    $controls.ReloadOverwriteRolesButton.Add_Click({ Update-OverwriteRoles -Force })
    $controls.OverwriteRoleList.Add_SelectionChanged({ Update-OverwritePreview })
    $controls.ExistingRoleBox.Add_SelectionChanged({ Update-OverwritePreview })
    $controls.CreateNewCheckBox.Add_Click({ Set-CreateNewMode })
    $controls.NewRoleNameBox.Add_TextChanged({ if ($state.CreateNew) { Update-CopyState } })

    # Load iTwins the first time the view is visible; Loaded also fires for unselected tab content.
    $root.Add_IsVisibleChanged({
            if ($state.Loaded -or -not $root.IsVisible) { return }
            $state.Loaded = $true
            $null = $root.Dispatcher.BeginInvoke([Action] { Reset-PickerPage 'Source'; Reset-PickerPage 'Destination' }, [System.Windows.Threading.DispatcherPriority]::ContextIdle)
        })

    #endregion Events

    [pscustomobject]@{ Root = $root }
}

function Show-iTwinRoleCopy {
    param([Parameter(Mandatory = $true)][string]$AuditLogPath)
    if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
        throw 'The role copy window requires an STA PowerShell host. Start Windows PowerShell with -STA.'
    }
    $view = New-iTwinRoleCopyView -AuditLogPath $AuditLogPath
    $window = New-Object System.Windows.Window -Property @{
        Title = 'iTwin Role Copy'; Height = 880; Width = 1440; MinHeight = 620; MinWidth = 1080
        WindowStartupLocation = 'CenterScreen'; Background = '#F4F6F8'; Content = $view.Root
    }
    $window.ShowDialog() | Out-Null
}

#endregion View

Export-ModuleMember -Function New-iTwinRoleCopyView, Show-iTwinRoleCopy
