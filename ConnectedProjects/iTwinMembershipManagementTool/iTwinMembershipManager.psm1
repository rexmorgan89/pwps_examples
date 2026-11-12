Set-StrictMode -Version Latest

function New-iTwinMembershipManagerItem {
    param([string]$Kind, $Value, [string]$Label)
    [pscustomobject]@{ Kind = $Kind; Value = $Value; Label = $Label }
}

function Get-iTwinMembershipManagerSelectedItems {
    param($Control)
    @($Control.SelectedItems | ForEach-Object { $_ })
}

function Write-iTwinMembershipManagerAudit {
    param([string]$AuditLogPath, [array]$Results)
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
    if ($rows.Count -eq 0) { return }
    $rows | Export-Csv -LiteralPath $AuditLogPath -NoTypeInformation -Append
}

function Select-iTwinMembershipManagerRole {
    param([string]$ProjectName, [array]$Roles)
    [xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Select project role" Height="350" Width="440" MinHeight="240" MinWidth="380" WindowStartupLocation="CenterOwner" ResizeMode="CanResizeWithGrip">
    <Grid Margin="14">
        <Grid.RowDefinitions><RowDefinition Height="Auto" /><RowDefinition Height="*" /><RowDefinition Height="Auto" /></Grid.RowDefinitions>
        <TextBlock x:Name="MessageText" TextWrapping="Wrap" Margin="0,0,0,10" />
        <ListBox x:Name="RoleList" Grid.Row="1" DisplayMemberPath="Label" />
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0"><Button x:Name="CancelButton" Content="Cancel" Padding="12,5" Margin="0,0,8,0" /><Button x:Name="SelectButton" Content="Use selected role" Padding="12,5" IsDefault="True" /></StackPanel>
    </Grid>
</Window>
'@
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $dialog = [Windows.Markup.XamlReader]::Load($reader)
    $dialog.FindName('MessageText').Text = "Team Member is unavailable or ambiguous for '$ProjectName'. Select a current project role."
    $roleList = $dialog.FindName('RoleList')
    $roleList.ItemsSource = $Roles
    if ($Roles.Count -gt 0) { $roleList.SelectedIndex = 0 }
    $selectedRole = $null
    $dialog.FindName('CancelButton').Add_Click({ $dialog.DialogResult = $false })
    $dialog.FindName('SelectButton').Add_Click({ if ($null -ne $roleList.SelectedItem) { $selectedRole = $roleList.SelectedItem; $dialog.DialogResult = $true } })
    $dialog.ShowDialog() | Out-Null
    $selectedRole
}

function Show-iTwinMembershipManagerResults {
    param([string]$Title, [array]$Results)
    $summary = @($Results | ForEach-Object { "{0} | {1} | {2}" -f $_.ProjectGUID, $_.Status, $_.Error } ) -join [Environment]::NewLine
    [System.Windows.MessageBox]::Show($summary, $Title, 'OK', 'Information') | Out-Null
}

function Show-iTwinProjectDiscoveryDialog {
    param([string]$ProjectConfigPath)
    [xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Discover and configure projects" Height="660" Width="1040" MinHeight="460" MinWidth="760" WindowStartupLocation="CenterOwner">
  <Grid Margin="14">
        <Grid.RowDefinitions><RowDefinition Height="Auto" /><RowDefinition Height="Auto" /><RowDefinition Height="*" /><RowDefinition Height="Auto" /></Grid.RowDefinitions>
        <Grid Grid.Row="0" Margin="0,0,0,10">
            <Grid.ColumnDefinitions><ColumnDefinition Width="210" /><ColumnDefinition Width="Auto" /><ColumnDefinition Width="Auto" /><ColumnDefinition Width="*" /><ColumnDefinition Width="Auto" /><ColumnDefinition Width="Auto" /></Grid.ColumnDefinitions>
            <TabControl x:Name="SourceTabs" Height="30" SelectedIndex="0"><TabItem Header="All Projects" /><TabItem Header="My Favorites" /></TabControl>
            <CheckBox x:Name="OrganizationScopeBox" Grid.Column="1" Content="Show organization projects" Margin="14,0,0,0" VerticalAlignment="Center" />
            <ComboBox x:Name="ProjectStatusBox" Grid.Column="2" Width="88" Margin="14,0,8,0" SelectedIndex="0"><ComboBoxItem Content="Active" /><ComboBoxItem Content="Trial" /><ComboBoxItem Content="Inactive" /><ComboBoxItem Content="All" /></ComboBox>
            <TextBox x:Name="SearchBox" Grid.Column="3" Margin="0,0,8,0" MinWidth="220" VerticalContentAlignment="Center" />
            <ComboBox x:Name="PageSizeBox" Grid.Column="4" Width="72" SelectedIndex="1"><ComboBoxItem Content="100" /><ComboBoxItem Content="250" /><ComboBoxItem Content="500" /><ComboBoxItem Content="1000" /></ComboBox>
            <Button x:Name="SearchButton" Grid.Column="5" Content="Search" Margin="8,0,0,0" Padding="12,5" />
        </Grid>
        <TextBlock x:Name="StatusText" Grid.Row="1" Text="Retrieving available iTwins..." Margin="0,0,0,10" />
        <ListBox x:Name="ProjectList" Grid.Row="2" DisplayMemberPath="Label" SelectionMode="Extended" ScrollViewer.CanContentScroll="True" VirtualizingStackPanel.IsVirtualizing="True" VirtualizingStackPanel.VirtualizationMode="Recycling" />
        <DockPanel Grid.Row="3" LastChildFill="False" Margin="0,12,0,0"><Button x:Name="CancelButton" Content="Close" Padding="12,5" Margin="8,0,0,0" DockPanel.Dock="Right" /><Button x:Name="AddButton" Content="Add selected to configuration" Padding="12,5" Margin="8,0,0,0" DockPanel.Dock="Right" IsDefault="True" /><Button x:Name="NextButton" Content="Next" Padding="12,5" Margin="8,0,0,0" DockPanel.Dock="Right" /><Button x:Name="PreviousButton" Content="Previous" Padding="12,5" DockPanel.Dock="Right" /></DockPanel>
  </Grid>
</Window>
'@
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $dialog = [Windows.Markup.XamlReader]::Load($reader)
    $statusText = $dialog.FindName('StatusText')
    $projectList = $dialog.FindName('ProjectList')
    $addButton = $dialog.FindName('AddButton')
    $sourceTabs = $dialog.FindName('SourceTabs')
    $organizationScopeBox = $dialog.FindName('OrganizationScopeBox')
    $projectStatusBox = $dialog.FindName('ProjectStatusBox')
    $searchBox = $dialog.FindName('SearchBox')
    $pageSizeBox = $dialog.FindName('PageSizeBox')
    $searchButton = $dialog.FindName('SearchButton')
    $previousButton = $dialog.FindName('PreviousButton')
    $nextButton = $dialog.FindName('NextButton')
    $configured = @{}
    if (Test-Path -LiteralPath $ProjectConfigPath -PathType Leaf) {
        foreach ($project in @(Import-Csv -LiteralPath $ProjectConfigPath)) { $configured[[string]$project.ProjectGUID] = $true }
    }
    $state = @{ Sources = @{ All = @{ PageIndex = 0; NextUris = @($null); CurrentItems = @() }; Favorites = @{ PageIndex = 0; NextUris = @($null); CurrentItems = @() } }; CurrentSourceName = 'All'; Selected = @{}; QueryScope = 'MemberOfiTwin'; ProjectStatus = 'Active' }
    $getSourceName = { if ($sourceTabs.SelectedIndex -eq 1) { 'Favorites' } else { 'All' } }.GetNewClosure()
    $isOrganizationAccessDenied = {
        param($ErrorRecord)
        $exception = $ErrorRecord.Exception
        $responseProperty = $exception.PSObject.Properties['Response']
        if ($null -ne $responseProperty -and $null -ne $responseProperty.Value) {
            try { return [int]$responseProperty.Value.StatusCode -eq 403 } catch { }
        }
        [string]$exception.Message -match 'InsufficientPermissions|\b403\b'
    }.GetNewClosure()
    $saveCurrentSelection = {
        $sourceState = $state.Sources[$state.CurrentSourceName]
        foreach ($item in $sourceState.CurrentItems) { $state.Selected.Remove([string]$item.ProjectGUID) | Out-Null }
        foreach ($item in @($projectList.SelectedItems | Where-Object { -not $_.Configured })) { $state.Selected[[string]$item.ProjectGUID] = $item }
    }.GetNewClosure()
    $loadPage = {
        $sourceName = & $getSourceName
        $state.CurrentSourceName = $sourceName
        $sourceState = $state.Sources[$sourceName]
        $pageSize = [int](($pageSizeBox.SelectedItem).Content)
        $queryScope = if ([bool]$organizationScopeBox.IsChecked) { 'OrganizationAdmin' } else { 'MemberOfiTwin' }
        $projectStatus = [string](($projectStatusBox.SelectedItem).Content)
        $state.QueryScope = $queryScope; $state.ProjectStatus = $projectStatus
        $searchButton.IsEnabled = $false; $previousButton.IsEnabled = $false; $nextButton.IsEnabled = $false; $addButton.IsEnabled = $false
        try {
            $page = $null
            $fallbackOccurred = $false
            while ($true) {
                try {
                    if ($sourceName -eq 'Favorites') {
                        $page = Get-iTwinFavoriteProjectPage -ProjectStatus $projectStatus -PageSize $pageSize -NextUri ([string]$sourceState.NextUris[$sourceState.PageIndex])
                    }
                    else {
                        $page = Get-iTwinAvailableProjectPage -DisplayName ([string]$searchBox.Text).Trim() -QueryScope $queryScope -ProjectStatus $projectStatus -PageSize $pageSize -NextUri ([string]$sourceState.NextUris[$sourceState.PageIndex])
                    }
                    break
                }
                catch {
                    if ($queryScope -eq 'OrganizationAdmin' -and (& $isOrganizationAccessDenied $_)) {
                        $organizationScopeBox.IsChecked = $false
                        $sourceState.PageIndex = 0; $sourceState.NextUris = @($null)
                        $queryScope = 'MemberOfiTwin'; $state.QueryScope = $queryScope; $fallbackOccurred = $true
                        continue
                    }
                    throw
                }
            }
            $sourceState.CurrentItems = @($page.Projects | Sort-Object ProjectName | ForEach-Object {
                    $projectGuid = [string]$_.ProjectGUID
                    [pscustomobject]@{ ProjectGUID = $projectGuid; ProjectName = [string]$_.ProjectName; Status = [string]$_.Status; Configured = $configured.ContainsKey($projectGuid); Label = "$(if ($configured.ContainsKey($projectGuid)) { '[Configured] ' })$($_.ProjectName) [$($_.Status)] ($projectGuid)" }
                })
            $projectList.ItemsSource = $sourceState.CurrentItems
            foreach ($item in $sourceState.CurrentItems) { if ($state.Selected.ContainsKey([string]$item.ProjectGUID)) { $projectList.SelectedItems.Add($item) | Out-Null } }
            if ($sourceState.NextUris.Count -eq ($sourceState.PageIndex + 1)) { $sourceState.NextUris += [string]$page.NextUri } else { $sourceState.NextUris[$sourceState.PageIndex + 1] = [string]$page.NextUri }
            $previousButton.IsEnabled = $sourceState.PageIndex -gt 0
            $nextButton.IsEnabled = -not [string]::IsNullOrWhiteSpace([string]$page.NextUri)
            $addButton.IsEnabled = $true
            $statusText.Text = "$(if ($fallbackOccurred) { 'Organization access is unavailable. ' })Page $($sourceState.PageIndex + 1): $($sourceState.CurrentItems.Count) $($state.ProjectStatus) $(if ($sourceName -eq 'Favorites') { 'favorite' } else { $state.QueryScope }) project(s). $($state.Selected.Count) selected across pages."
        }
        catch {
            $projectList.ItemsSource = @()
            $sourceState.CurrentItems = @()
            $statusText.Text = "Discovery failed: $($_.Exception.Message)"
        }
        finally { $searchButton.IsEnabled = $true }
    }.GetNewClosure()
    $dialog.FindName('CancelButton').Add_Click({ $dialog.Close() })
    $resetCurrentSource = { & $saveCurrentSelection; $sourceState = $state.Sources[(& $getSourceName)]; $sourceState.PageIndex = 0; $sourceState.NextUris = @($null); & $loadPage }.GetNewClosure()
    $searchButton.Add_Click({ & $resetCurrentSource })
    $organizationScopeBox.Add_Click({ & $resetCurrentSource })
    $projectStatusBox.Add_SelectionChanged({ & $resetCurrentSource })
    $pageSizeBox.Add_SelectionChanged({ & $resetCurrentSource })
    $sourceTabs.Add_SelectionChanged({
            & $saveCurrentSelection
            $isFavorites = (& $getSourceName) -eq 'Favorites'
            $organizationScopeBox.Visibility = if ($isFavorites) { 'Collapsed' } else { 'Visible' }
            $searchBox.Visibility = if ($isFavorites) { 'Collapsed' } else { 'Visible' }
            $searchButton.Visibility = if ($isFavorites) { 'Collapsed' } else { 'Visible' }
            & $loadPage
        }.GetNewClosure())
    $previousButton.Add_Click({ & $saveCurrentSelection; $sourceState = $state.Sources[(& $getSourceName)]; $sourceState.PageIndex--; & $loadPage }.GetNewClosure())
    $nextButton.Add_Click({ & $saveCurrentSelection; $sourceState = $state.Sources[(& $getSourceName)]; $sourceState.PageIndex++; & $loadPage }.GetNewClosure())
    $addButton.Add_Click({
            & $saveCurrentSelection
            $selected = @($state.Selected.Values)
            if ($selected.Count -eq 0) {
                [System.Windows.MessageBox]::Show('Select one or more projects that are not already configured.', 'Configure projects', 'OK', 'Information') | Out-Null
                return
            }
            $existing = if (Test-Path -LiteralPath $ProjectConfigPath -PathType Leaf) { @(Import-Csv -LiteralPath $ProjectConfigPath) } else { @() }
            $merged = @($existing + @($selected | ForEach-Object { [pscustomobject]@{ ProjectGUID = $_.ProjectGUID; ProjectName = $_.ProjectName } }) | Sort-Object ProjectGUID -Unique)
            $merged | Select-Object ProjectGUID, ProjectName | Export-Csv -LiteralPath $ProjectConfigPath -NoTypeInformation
            $statusText.Text = "Added $($selected.Count) project(s). The configuration now contains $($merged.Count) project(s)."
            foreach ($project in $selected) { $project.Configured = $true; $project.Label = "[Configured] $($project.ProjectName) ($($project.ProjectGUID))"; $configured[[string]$project.ProjectGUID] = $true; $state.Selected.Remove([string]$project.ProjectGUID) | Out-Null }
            $projectList.Items.Refresh()
        }.GetNewClosure())
    & $loadPage
    $dialog.ShowDialog() | Out-Null
}

function New-iTwinMembershipManagerView {
    param([Parameter(Mandatory = $true)][string]$DatabasePath, [Parameter(Mandatory = $true)][string]$ProjectConfigPath, [string]$AuditLogPath)

    Test-iTwinMembershipInventoryDatabase -DatabasePath $DatabasePath
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, Microsoft.VisualBasic
    if ([string]::IsNullOrWhiteSpace($AuditLogPath)) {
        $AuditLogPath = Join-Path -Path (Split-Path -Parent $DatabasePath) -ChildPath 'iTwinMembershipManager-audit.csv'
    }

    [xml]$xaml = @'
  <Grid xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Margin="14">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto" />
      <RowDefinition Height="Auto" />
      <RowDefinition Height="*" />
      <RowDefinition Height="Auto" />
    </Grid.RowDefinitions>
    <Grid.ColumnDefinitions><ColumnDefinition Width="*" /></Grid.ColumnDefinitions>
    <DockPanel Grid.Row="0" LastChildFill="False" Margin="0,0,0,10">
      <TextBlock Text="iTwin Membership Manager" FontSize="22" FontWeight="SemiBold" Foreground="#17324D" VerticalAlignment="Center" DockPanel.Dock="Left" />
    <Button x:Name="RefreshButton" Content="Refresh inventory" Margin="8,0,0,0" Padding="12,5" DockPanel.Dock="Right" />
    <Button x:Name="PollButton" Content="Poll configured projects" Margin="8,0,0,0" Padding="12,5" DockPanel.Dock="Right" />
    <Button x:Name="DiscoverButton" Content="Discover projects" Margin="18,0,0,0" Padding="12,5" DockPanel.Dock="Right" />
    </DockPanel>
    <Grid Grid.Row="1" Margin="0,0,0,10">
    <Grid.ColumnDefinitions><ColumnDefinition Width="Auto" /><ColumnDefinition Width="180" /><ColumnDefinition Width="18" /><ColumnDefinition Width="Auto" /><ColumnDefinition Width="220" /><ColumnDefinition Width="18" /><ColumnDefinition Width="Auto" /><ColumnDefinition Width="18" /><ColumnDefinition Width="Auto" /></Grid.ColumnDefinitions>
      <TextBlock Text="Browse" VerticalAlignment="Center" />
      <ComboBox x:Name="ModeBox" Grid.Column="1" SelectedIndex="0" Margin="7,0,0,0"><ComboBoxItem Content="Projects" /><ComboBoxItem Content="Users" /><ComboBoxItem Content="Organizations" /></ComboBox>
      <TextBlock Text="Organization filter" Grid.Column="3" VerticalAlignment="Center" />
      <ComboBox x:Name="OrganizationBox" Grid.Column="4" Margin="7,0,0,0" IsEditable="True" />
      <CheckBox x:Name="HistoryBox" Grid.Column="6" Content="Include history" VerticalAlignment="Center" />
    <CheckBox x:Name="FavoritesOnlyBox" Grid.Column="8" Content="Favorites only" VerticalAlignment="Center" />
    </Grid>
    <Grid Grid.Row="2">
      <Grid.ColumnDefinitions><ColumnDefinition Width="*" MinWidth="220" /><ColumnDefinition Width="6" /><ColumnDefinition Width="*" MinWidth="220" /><ColumnDefinition Width="6" /><ColumnDefinition Width="2*" MinWidth="330" /></Grid.ColumnDefinitions>
      <GroupBox Header="Browse" Grid.Column="0" Padding="7" BorderBrush="#AEB8C2"><ListBox x:Name="LeftList" DisplayMemberPath="Label" SelectionMode="Extended" /></GroupBox>
      <GridSplitter Grid.Column="1" Width="6" HorizontalAlignment="Stretch" Background="#B6C0CA" />
      <GroupBox Header="Related selection" Grid.Column="2" Padding="7" BorderBrush="#AEB8C2"><ListBox x:Name="MiddleList" DisplayMemberPath="Label" SelectionMode="Extended" /></GroupBox>
      <GridSplitter Grid.Column="3" Width="6" HorizontalAlignment="Stretch" Background="#B6C0CA" />
      <GroupBox Header="Matching memberships" Grid.Column="4" Padding="7" BorderBrush="#AEB8C2"><ListBox x:Name="RightList" DisplayMemberPath="Label" SelectionMode="Extended" /></GroupBox>
    </Grid>
    <DockPanel Grid.Row="3" Margin="0,10,0,0">
      <TextBlock x:Name="StatusText" Text="Select records to view membership relationships." VerticalAlignment="Center" Foreground="#3E4C59" DockPanel.Dock="Left" />
      <Button x:Name="RemoveButton" Content="Remove selected memberships" Padding="12,5" Margin="8,0,0,0" DockPanel.Dock="Right" Background="#B42318" Foreground="White" />
            <Button x:Name="RemoveProjectButton" Content="Stop tracking selected projects" Padding="12,5" Margin="8,0,0,0" DockPanel.Dock="Right" Background="#8A1C1C" Foreground="White" />
      <Button x:Name="AddButton" Content="Add user to selected projects" Padding="12,5" Margin="8,0,0,0" DockPanel.Dock="Right" Background="#176B4D" Foreground="White" />
    </DockPanel>
  </Grid>
'@

    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $root = [Windows.Markup.XamlReader]::Load($reader)
    $modeBox = $root.FindName('ModeBox')
    $organizationBox = $root.FindName('OrganizationBox')
    $historyBox = $root.FindName('HistoryBox')
    $favoritesOnlyBox = $root.FindName('FavoritesOnlyBox')
    $leftList = $root.FindName('LeftList')
    $middleList = $root.FindName('MiddleList')
    $rightList = $root.FindName('RightList')
    $statusText = $root.FindName('StatusText')
    $refreshButton = $root.FindName('RefreshButton')
    $discoverButton = $root.FindName('DiscoverButton')
    $pollButton = $root.FindName('PollButton')
    $addButton = $root.FindName('AddButton')
    $removeButton = $root.FindName('RemoveButton')
    $removeProjectButton = $root.FindName('RemoveProjectButton')

    $newItem = ${function:New-iTwinMembershipManagerItem}.GetNewClosure()
    $getSelectedItems = ${function:Get-iTwinMembershipManagerSelectedItems}.GetNewClosure()
    $writeAudit = ${function:Write-iTwinMembershipManagerAudit}.GetNewClosure()
    $selectRole = ${function:Select-iTwinMembershipManagerRole}.GetNewClosure()
    $showDiscovery = ${function:Show-iTwinProjectDiscoveryDialog}.GetNewClosure()
    $showResults = ${function:Show-iTwinMembershipManagerResults}.GetNewClosure()
    $state = @{ Busy = $false; Relationships = @(); FavoriteProjectGuids = $null; OrganizationTextRefreshTimer = New-Object Windows.Threading.DispatcherTimer }
    $state.OrganizationTextRefreshTimer.Interval = [TimeSpan]::FromMilliseconds(350)
    $getMode = { [string](($modeBox.SelectedItem).Content) }.GetNewClosure()
    $getOrganization = { ([string]$organizationBox.Text).Trim() }.GetNewClosure()
    $getHistory = { [bool]$historyBox.IsChecked }.GetNewClosure()

    $refreshRight = {
        $mode = & $getMode
        $leftItems = @(& $getSelectedItems $leftList)
        $middleItems = @(& $getSelectedItems $middleList)
        $relationships = @()
        foreach ($leftItem in $leftItems) {
            if ($mode -eq 'Projects') { $relationships += @(Get-iTwinInventoryRelationships -DatabasePath $DatabasePath -ProjectGuid $leftItem.Value -IncludeHistory:(& $getHistory)) }
            elseif ($mode -eq 'Users') { $relationships += @(Get-iTwinInventoryRelationships -DatabasePath $DatabasePath -Email $leftItem.Value -IncludeHistory:(& $getHistory)) }
            else {
                foreach ($project in $middleItems) { $relationships += @(Get-iTwinInventoryRelationships -DatabasePath $DatabasePath -ProjectGuid $project.Value -Organization $leftItem.Value -IncludeHistory:(& $getHistory)) }
            }
        }
        if ($mode -eq 'Projects' -and $middleItems.Count -gt 0) {
            $emails = @($middleItems | Where-Object { $_.Kind -eq 'User' } | ForEach-Object { [string]$_.Value })
            $organizations = @($middleItems | Where-Object { $_.Kind -eq 'Organization' } | ForEach-Object { [string]$_.Value })
            $relationships = @($relationships | Where-Object { $emails -contains $_.Email -or $organizations -contains $_.Organization })
        }
        elseif ($mode -eq 'Users' -and $middleItems.Count -gt 0) {
            $projects = @($middleItems | ForEach-Object { [string]$_.Value })
            $relationships = @($relationships | Where-Object { $projects -contains $_.ProjectGuid })
        }
        $state.Relationships = @($relationships | Sort-Object ProjectName, Email -Unique)
        $rightList.ItemsSource = @($state.Relationships | ForEach-Object { & $newItem 'Relationship' $_ ("{0} | {1} | {2} | {3}" -f $_.ProjectName, $_.Email, $_.Organization, $_.Roles) })
        $statusText.Text = "Showing $($state.Relationships.Count) membership relationship(s)."
    }.GetNewClosure()

    $refreshMiddle = {
        $middleList.ItemsSource = @()
        $mode = & $getMode
        $leftItems = @(& $getSelectedItems $leftList)
        $items = @()
        foreach ($leftItem in $leftItems) {
            if ($mode -eq 'Projects') {
                $relationships = @(Get-iTwinInventoryRelationships -DatabasePath $DatabasePath -ProjectGuid $leftItem.Value -IncludeHistory:(& $getHistory))
                $items += @($relationships | ForEach-Object { & $newItem 'User' $_.Email $_.Email })
                $items += @($relationships | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.Organization) } | ForEach-Object { & $newItem 'Organization' $_.Organization $_.Organization })
            }
            elseif ($mode -eq 'Users') {
                $relationships = @(Get-iTwinInventoryRelationships -DatabasePath $DatabasePath -Email $leftItem.Value -IncludeHistory:(& $getHistory))
                $items += @($relationships | ForEach-Object { & $newItem 'Project' $_.ProjectGuid ("{0} ({1})" -f $_.ProjectName, $_.ProjectGuid) })
            }
            else {
                $projects = @(Get-iTwinInventoryProjects -DatabasePath $DatabasePath -Organization $leftItem.Value -IncludeHistory:(& $getHistory))
                $items += @($projects | ForEach-Object { & $newItem 'Project' $_.ProjectGuid ("{0} ({1})" -f $_.ProjectName, $_.ProjectGuid) })
            }
        }
        $middleList.ItemsSource = @($items | Sort-Object Kind, Label -Unique)
        & $refreshRight
    }.GetNewClosure()

    $refreshLeft = {
        $mode = & $getMode
        $organization = & $getOrganization
        $history = & $getHistory
        $favoriteProjectGuids = $null
        if ($mode -eq 'Projects' -and [bool]$favoritesOnlyBox.IsChecked) {
            try {
                if ($null -eq $state.FavoriteProjectGuids) {
                    $state.FavoriteProjectGuids = @{}
                    foreach ($project in @(Get-iTwinFavoriteProjects -ProjectStatus All)) { $state.FavoriteProjectGuids[[string]$project.ProjectGUID] = $true }
                }
                $favoriteProjectGuids = $state.FavoriteProjectGuids
            }
            catch {
                $statusText.Text = "Could not retrieve favorites: $($_.Exception.Message)"
                return
            }
        }
        $leftList.ItemsSource = @()
        $middleList.ItemsSource = @()
        $rightList.ItemsSource = @()
        if ($mode -eq 'Projects') {
            $projects = @(Get-iTwinConfiguredProjectInventory -ProjectConfigPath $ProjectConfigPath -DatabasePath $DatabasePath -Organization $organization -IncludeHistory:$history)
            if ($null -ne $favoriteProjectGuids) { $projects = @($projects | Where-Object { $favoriteProjectGuids.ContainsKey([string]$_.ProjectGUID) }) }
            $leftList.ItemsSource = @($projects | ForEach-Object { & $newItem 'Project' $_.ProjectGUID ("{0} ({1} active members, {2})" -f $_.ProjectName, $_.ActiveMemberCount, $_.InventoryState) })
        }
        elseif ($mode -eq 'Users') {
            $leftList.ItemsSource = @(Get-iTwinInventoryUsers -DatabasePath $DatabasePath -Organization $organization -IncludeHistory:$history | ForEach-Object { & $newItem 'User' $_.Email ("{0} | {1} | {2} projects" -f $_.Email, $_.Organization, $_.ProjectCount) })
        }
        else {
            $leftList.ItemsSource = @(Get-iTwinInventoryOrganizations -DatabasePath $DatabasePath -IncludeHistory:$history | ForEach-Object { & $newItem 'Organization' $_.Organization ("{0} ({1} users, {2} projects)" -f $_.Organization, $_.UserCount, $_.ProjectCount) })
        }
        $statusText.Text = 'Select records to view membership relationships.'
    }.GetNewClosure()

    $refreshOrganizations = {
        $organizationBox.ItemsSource = @('') + @(Get-iTwinInventoryOrganizations -DatabasePath $DatabasePath -IncludeHistory:(& $getHistory) | ForEach-Object { [string]$_.Organization })
        & $refreshLeft
    }.GetNewClosure()

    $state.OrganizationTextRefreshTimer.Add_Tick({
            $state.OrganizationTextRefreshTimer.Stop()
            & $refreshLeft
        }.GetNewClosure())
    $null = $organizationBox.ApplyTemplate()
    $organizationTextBox = $organizationBox.Template.FindName('PART_EditableTextBox', $organizationBox)
    if ($null -ne $organizationTextBox) {
        $organizationTextBox.Add_TextChanged({
                $state.OrganizationTextRefreshTimer.Stop()
                $state.OrganizationTextRefreshTimer.Start()
            }.GetNewClosure())
    }

    $leftList.Add_SelectionChanged({ & $refreshMiddle }.GetNewClosure())
    $middleList.Add_SelectionChanged({ & $refreshRight }.GetNewClosure())
    $modeBox.Add_SelectionChanged({ $favoritesOnlyBox.Visibility = if ((& $getMode) -eq 'Projects') { 'Visible' } else { 'Collapsed' }; & $refreshLeft }.GetNewClosure())
    $organizationBox.Add_SelectionChanged({ $state.OrganizationTextRefreshTimer.Stop(); & $refreshLeft }.GetNewClosure())
    $historyBox.Add_Click({ & $refreshOrganizations }.GetNewClosure())
    $favoritesOnlyBox.Add_Click({ & $refreshLeft }.GetNewClosure())
    $refreshButton.Add_Click({ $state.FavoriteProjectGuids = $null; & $refreshOrganizations }.GetNewClosure())
    $discoverButton.Add_Click({
            & $showDiscovery -ProjectConfigPath $ProjectConfigPath
            $statusText.Text = "Project configuration: '$ProjectConfigPath'. Poll configured projects to inventory new selections."
        }.GetNewClosure())
    $pollButton.Add_Click({
            if (-not (Test-Path -LiteralPath $ProjectConfigPath -PathType Leaf)) {
                [System.Windows.MessageBox]::Show("Project configuration was not found: '$ProjectConfigPath'. Use Discover projects first.", 'Poll projects', 'OK', 'Information') | Out-Null
                return
            }
            try {
                $projects = @(Import-Csv -LiteralPath $ProjectConfigPath)
                if ($projects.Count -eq 0) { throw 'The project configuration contains no projects.' }
                $pollButton.IsEnabled = $false
                $statusText.Text = "Polling $($projects.Count) configured project(s)..."
                $results = @(Invoke-iTwinMembershipPollProjects -Projects $projects -DatabasePath $DatabasePath)
                & $showResults -Title 'Project poll results' -Results $results
                $statusText.Text = "Polling completed: $(@($results | Where-Object { $_.Status -eq 'Success' }).Count) succeeded, $(@($results | Where-Object { $_.Status -eq 'Failed' }).Count) failed."
                & $refreshOrganizations
            }
            catch {
                [System.Windows.MessageBox]::Show($_.Exception.Message, 'Poll projects', 'OK', 'Error') | Out-Null
                $statusText.Text = 'Polling failed before any project result was returned.'
            }
            finally { $pollButton.IsEnabled = $true }
        }.GetNewClosure())
    $removeProjectButton.Add_Click({
            if ((& $getMode) -ne 'Projects') {
                [System.Windows.MessageBox]::Show('Select Projects mode before removing tracked projects.', 'Stop tracking projects', 'OK', 'Information') | Out-Null
                return
            }
            $projects = @(& $getSelectedItems $leftList | Where-Object { $_.Kind -eq 'Project' })
            if ($projects.Count -eq 0) {
                [System.Windows.MessageBox]::Show('Select one or more tracked projects in the left pane.', 'Stop tracking projects', 'OK', 'Information') | Out-Null
                return
            }
            try {
                $preview = @(Remove-iTwinTrackedProjects -ProjectGuid @($projects | ForEach-Object { $_.Value }) -ProjectConfigPath $ProjectConfigPath -DatabasePath $DatabasePath -WhatIf)
                $membershipCount = @($preview | Measure-Object -Property MembershipRecords -Sum).Sum
                $roleCount = @($preview | Measure-Object -Property RoleRecords -Sum).Sum
                $projectSummary = @($preview | ForEach-Object { "  $($_.ProjectName) ($($_.ProjectGUID))" }) -join [Environment]::NewLine
                $message = "This action is not reversible.`n`nIt stops local polling and permanently deletes local SQLite membership history. It does not change Bentley memberships or projects.`n`nProjects: $($preview.Count)`nMembership records to delete: $membershipCount`nRole records to delete: $roleCount`n`n$projectSummary"
                if ([System.Windows.MessageBox]::Show($message, 'Confirm local project removal', 'YesNo', 'Warning') -ne 'Yes') { return }
                $removeProjectButton.IsEnabled = $false
                $results = @(Remove-iTwinTrackedProjects -ProjectGuid @($projects | ForEach-Object { $_.Value }) -ProjectConfigPath $ProjectConfigPath -DatabasePath $DatabasePath)
                & $writeAudit -AuditLogPath $AuditLogPath -Results @($results | ForEach-Object { [pscustomobject]@{ ProjectGUID = $_.ProjectGUID; User = ''; Action = 'RemoveProject'; RoleId = ''; Status = $_.Status; Detail = $_.Detail } })
                $statusText.Text = "Stopped tracking $(@($results | Where-Object { $_.Status -eq 'Success' }).Count) project(s)."
                & $refreshOrganizations
            }
            catch {
                [System.Windows.MessageBox]::Show($_.Exception.Message, 'Stop tracking projects', 'OK', 'Error') | Out-Null
            }
            finally { $removeProjectButton.IsEnabled = $true }
        }.GetNewClosure())

    $addButton.Add_Click({
            $projects = @(& $getSelectedItems $leftList | Where-Object { $_.Kind -eq 'Project' })
            if ((& $getMode) -ne 'Projects' -or $projects.Count -eq 0) {
                [System.Windows.MessageBox]::Show('Select one or more projects in Projects mode before adding a user.', 'Add user', 'OK', 'Information') | Out-Null
                return
            }
            $email = [Microsoft.VisualBasic.Interaction]::InputBox('Enter the email address to add. Team Member is selected when it exists; otherwise you will choose a live project role.', 'Add user')
            if ([string]::IsNullOrWhiteSpace($email)) { return }
            $actions = @()
            foreach ($project in $projects) {
                try {
                    $roles = @(Get-iTwinProjectRoles -ProjectGuid $project.Value)
                    $teamMemberRoles = @($roles | Where-Object { $_.RoleDisplayName -ieq 'Team Member' })
                    $selectedRole = if ($teamMemberRoles.Count -eq 1) { $teamMemberRoles[0] } else { & $selectRole -ProjectName $project.Label -Roles $roles }
                    if ($null -eq $selectedRole) { return }
                    $actions += [pscustomobject]@{ ProjectGUID = $project.Value; User = $email.Trim(); Action = 'AssignRole'; RoleId = [string]$selectedRole.RoleId; RoleDisplayName = [string]$selectedRole.RoleDisplayName }
                }
                catch {
                    [System.Windows.MessageBox]::Show("Could not retrieve roles for '$($project.Label)'.`n$($_.Exception.Message)", 'Add user', 'OK', 'Error') | Out-Null
                    return
                }
            }
            $result = @(Invoke-iTwinMembershipActions -Actions $actions -DatabasePath $DatabasePath)
            & $writeAudit -AuditLogPath $AuditLogPath -Results $result
            $statusText.Text = "Add operation completed: $($result.Count) project action(s)."
            & $refreshLeft
        }.GetNewClosure())

    $removeButton.Add_Click({
            $targets = @(& $getSelectedItems $rightList | ForEach-Object { $_.Value })
            if ($targets.Count -eq 0) {
                [System.Windows.MessageBox]::Show('Select one or more exact membership relationships in the right pane before removing.', 'Remove memberships', 'OK', 'Information') | Out-Null
                return
            }
            $projectCount = @($targets | ForEach-Object { $_.ProjectGuid } | Select-Object -Unique).Count
            $userCount = @($targets | ForEach-Object { $_.Email } | Select-Object -Unique).Count
            $summary = "This action is not reversible.`n`nRemove $($targets.Count) membership assignment(s) for $userCount user(s) across $projectCount project(s)?"
            if ([System.Windows.MessageBox]::Show($summary, 'Confirm irreversible removal', 'YesNo', 'Warning') -ne 'Yes') { return }
            $actions = @($targets | ForEach-Object { [pscustomobject]@{ ProjectGUID = $_.ProjectGuid; User = $_.Email; Action = 'RemoveMember'; RoleId = ''; RoleDisplayName = '' } })
            $result = @(Invoke-iTwinMembershipActions -Actions $actions -DatabasePath $DatabasePath)
            & $writeAudit -AuditLogPath $AuditLogPath -Results $result
            $statusText.Text = "Removal completed: $($result.Count) project action(s)."
            & $refreshLeft
        }.GetNewClosure())

    $null = $refreshOrganizations.Invoke()
    [pscustomobject]@{
        Root    = $root
        Refresh = $refreshOrganizations
        Close   = { $state.OrganizationTextRefreshTimer.Stop() }.GetNewClosure()
    }
}

function Show-iTwinMembershipManager {
    param([Parameter(Mandatory = $true)][string]$DatabasePath, [Parameter(Mandatory = $true)][string]$ProjectConfigPath, [string]$AuditLogPath)

    if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
        throw 'GUI mode requires an STA PowerShell host. Start Windows PowerShell with -STA.'
    }
    $view = New-iTwinMembershipManagerView -DatabasePath $DatabasePath -ProjectConfigPath $ProjectConfigPath -AuditLogPath $AuditLogPath
    $window = New-Object System.Windows.Window -Property @{
        Title = 'iTwin Membership Manager'; Height = 760; Width = 1320; MinHeight = 520; MinWidth = 900
        WindowStartupLocation = 'CenterScreen'; Background = '#F4F6F8'; Content = $view.Root
    }
    $window.Add_Closed({ & $view.Close }.GetNewClosure())
    $window.ShowDialog() | Out-Null
}

Export-ModuleMember -Function New-iTwinMembershipManagerView, Show-iTwinMembershipManager, Show-iTwinProjectDiscoveryDialog, Write-iTwinMembershipManagerAudit
