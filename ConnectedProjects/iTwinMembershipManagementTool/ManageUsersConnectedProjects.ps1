#####################################################################################################################################################################################################################
#########################################################################################################################| |#
####                 #############################################  #####    ############################################| |#  Name                     : ManageUsersConnectedProjects
####                   #########################################    #####    ############################################| |#  Type                     : Script
####    ##########     #########################################    #####    ############################################| |#  Purpose                  : Manage iTwin memberships, roles, role copies, reports, and bulk actions in one tabbed WPF window.
####    ###########    #######         ######    #         ####         #    ######          ######    ########    ######| |#  Author                   : GitHub Copilot
####    ##########-    ####               ###                ##         #    ####              ####    ########    ######| |#  Creation Date            : 10-22-2026
#                     #####     ######    ###      #####     ###    #####    ###     ######     ###    ########    ######| |#  Modified By              : GitHub Copilot
#                    #####    #########    ##    ########     ##    #####    ##     ########    ###    ########    ######| |#  Modified Date            : 10-30-2026
####    ##########      ##                 ##    #########    ##    #####    ##                  ##    ########    ######| |#  Script Version           : 1.2
####    ############    ##                 ##    #########    ##    #####    ##                  ##    ########    ######| |#  PowerShell Version       : Windows PowerShell 5.1 x64 (STA)
####    ############    ##    ###############    #########    ##    #####    ##     ###############    #######     ######| |#  ProjectWise Version      : Not required
####    ###########     ###     ######    ###    #########    ##    #####    ###     #######    ###     #####      ######| |#  PWPS Module Version      : Not required
####                   ####               ###    #########    ##    #####    ####              #####               ######| |#  PWPS_DAB Module Version  : Repository reference: 24.0.2
####                  #######           #####    #########    ##    #####    ######           ########        #    ######| |#
##############################################################################################################    #######| |#  Requirements:
###################################################################################################     ###       #######| |#  pwps_dab (Get-OIDCToken), iTwinMembershipInventory.psm1, iTwinMembershipManager.psm1, iTwinRoleAssignment.psm1, iTwinRoleCopy.psm1.
##################################################################################################              #########| |#  PSSQLite for the Manage Memberships, User Membership Report, and Bulk Changes (CSV) tabs.
#######################################################################################################      ############| |#  Copy Role Definitions needs administration_manage_roles or Org Admin on each destination.
#########################################################################################################################| |#
#####################################################################################################################################################################################################################
##
##  Description/Notes:
##  Combines AddRemoveUsersConnectedProjects.ps1 (Gui, Report, and Bulk modes), AssignRolesConnectedProjects.ps1, and the
##  former CopyRolesConnectedProjects.ps1 in one window.
##  Module functions used: New-iTwinMembershipManagerView, Show-iTwinProjectDiscoveryDialog, Write-iTwinMembershipManagerAudit,
##  New-iTwinRoleAssignmentView, New-iTwinRoleCopyView, Invoke-iTwinMembershipPoll, Get-iTwinActiveMembershipReport,
##  Invoke-iTwinMembershipBulkAction (Access Control v2, iTwins, and Users APIs).
##  Changes Bentley iTwin memberships from the Manage Memberships, Assign User Roles, and Bulk Changes (CSV) tabs after confirmation.
##  The Copy Role Definitions tab creates iTwin or enterprise (account) roles on destinations after confirmation and, when chosen,
##  adds or replaces permissions of same-named destination roles; replacing can remove permissions. Role assignments are not copied.
##  With Overwrite Existing Role checked, it adds or replaces the permissions of one chosen destination role from one source role.
##  With Create New Role checked, it copies one source role to the destinations under a new name.
##  Writes membership changes and role copy results to separate audit CSVs, and a timestamped .LastRun.log transcript beside the script.
##  User guide with screenshots: Help\ManageUsersConnectedProjects-Help.html, opened from the ? button or F1.
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
Manages iTwin memberships, role assignments, role copies, membership reports, and bulk actions in one tabbed window.

.DESCRIPTION
Opens a WPF window with five tabs:
  Manage Memberships      Browses the local SQLite inventory by project, user, or organization; discovers
                          and polls projects; adds users; removes memberships; and stops tracking projects.
                          When the inventory does not exist yet, the tab offers Discover and Poll instead.
  Assign User Roles       Selects one live iTwin, its users, and its iTwin and enterprise roles, then adds
                          only the missing roles. Existing roles are never removed.
  Copy Role Definitions   Selects a source iTwin or account, its roles, and one or more destination iTwins
                          and accounts, previews the permissions, and copies the roles after confirmation.
                          Destination roles are matched by display name, ignoring case; existing roles are
                          skipped by default or can have permissions added or replaced. Enterprise roles are
                          the roles of an Account iTwin. Members and role assignments are not copied.
                          Overwrite Existing Role shows a fourth pane to pick one existing role on a single
                          destination, whatever its name, and add or replace its permissions with those of
                          one source role. The role's name, description, and members are kept.
                          Create New Role makes the role name field editable and copies one source role
                          to the destinations, including the source itself, under the entered name. When
                          it is cleared, the field is read-only and shows the selected role names.
  User Membership Report  Lists the active iTwin memberships and roles of one user from the local inventory.
  Bulk Changes (CSV)      Previews or applies AssignRole, UnassignRole, and RemoveMember rows from an action CSV,
                          and saves an action CSV template with placeholder rows.

Changes refresh the local inventory of the affected tracked iTwins, and the Manage Memberships tab
reloads afterwards. Every change result is appended to the audit CSV and written to the transcript.
Role copy results are appended to a separate audit CSV because their columns differ.

.PARAMETER ProjectConfigPath
Project configuration CSV with a ProjectGUID column and an optional ProjectName column.
Defaults to iTwinProjects.csv beside this script.

.PARAMETER AuditLogPath
CSV file that receives one row per change result. Defaults to iTwinMembershipManager-audit.csv
beside this script.

.PARAMETER RoleCopyAuditLogPath
CSV file that receives one row per role and destination result from the Copy Role Definitions tab.
Defaults to iTwinRoleCopy-audit.csv beside this script.

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File .\ManageUsersConnectedProjects.ps1
Open the tabbed membership window.

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File .\ManageUsersConnectedProjects.ps1 -AuditLogPath .\Membership-audit.csv
Open the window and write change results to a separate audit CSV.

.INPUTS
None.

.OUTPUTS
Membership and role copy audit CSV rows and a timestamped .LastRun.log transcript beside the script.

.NOTES
Requires Windows PowerShell 5.1 x64 in an STA host and Bentley IMS sign-in through Get-OIDCToken,
as used by iTwinMembershipInventory.psm1. Membership changes require the iTwin administration_invite_member
permission or an Organization Administrator role. Copying roles requires the administration_manage_roles
permission or an Organization Administrator role on each destination iTwin or account. Bentley API calls run on the UI thread, so the
window is unresponsive while a call is in progress. Tokens are never written to the audit CSV or transcript.
Scheduled polling remains available through AddRemoveUsersConnectedProjects.ps1 -Mode Poll.

User guide with screenshots: Help\ManageUsersConnectedProjects-Help.html beside this script. Select the ? button
at the top right of the window, or press F1, to open it in the default browser.
References: iTwinMembershipInventory.md, Bentley Access Control API v2, iTwins API, and Users API.
#>

[CmdletBinding()]
param(
    [string]$ProjectConfigPath,
    [string]$AuditLogPath,
    [string]$RoleCopyAuditLogPath
)

begin {
    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
        throw 'This script requires an STA PowerShell host. Start Windows PowerShell with -STA.'
    }

    $InformationPreference = 'Continue'
    # $VerbosePreference = 'Continue' # Uncomment this line if you want to see verbose messages.
    # Log all script output to a file for easy reference later if needed.
    [string] $lastRunLogFilePath = "$PSCommandPath.$(Get-Date -Format 'yyyyMMdd_HHmmss').LastRun.log"
    Start-Transcript -Path $lastRunLogFilePath -ErrorAction Stop

    # Display the time that this script started running.
    [DateTime] $startTime = Get-Date
    Write-Host "Starting script at '$($startTime.ToString('u'))'."

    try {
        if ([string]::IsNullOrWhiteSpace($ProjectConfigPath)) {
            $ProjectConfigPath = Join-Path -Path $PSScriptRoot -ChildPath 'iTwinProjects.csv'
        }
        if ([string]::IsNullOrWhiteSpace($AuditLogPath)) {
            $AuditLogPath = Join-Path -Path $PSScriptRoot -ChildPath 'iTwinMembershipManager-audit.csv'
        }
        if ([string]::IsNullOrWhiteSpace($RoleCopyAuditLogPath)) {
            $RoleCopyAuditLogPath = Join-Path -Path $PSScriptRoot -ChildPath 'iTwinRoleCopy-audit.csv'
        }
        $databasePath = Join-Path -Path $PSScriptRoot -ChildPath 'iTwinMembershipInventory.sqlite'

        foreach ($moduleName in @('iTwinMembershipInventory.psm1', 'iTwinMembershipManager.psm1', 'iTwinRoleAssignment.psm1', 'iTwinRoleCopy.psm1')) {
            Import-Module -Name (Join-Path -Path $PSScriptRoot -ChildPath $moduleName) -Force
        }
        Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
        Write-Host "Audit CSV: '${AuditLogPath}'."
        Write-Host "Role copy audit CSV: '${RoleCopyAuditLogPath}'."
    }
    catch {
        $null = Stop-Transcript
        throw
    }

    #region Helpers

    function Show-Message {
        param([string]$Text, [string]$Title, [string]$Icon = 'Information')
        [System.Windows.MessageBox]::Show($window, $Text, $Title, 'OK', $Icon) | Out-Null
    }

    function Invoke-WithWaitCursor {
        param([scriptblock]$Action)
        [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
        # Let pending status text render before the API call blocks the UI thread.
        $window.Dispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::Background)
        try { & $Action } finally { [System.Windows.Input.Mouse]::OverrideCursor = $null }
    }

    function Open-UserGuide {
        $helpPath = Join-Path -Path $PSScriptRoot -ChildPath 'Help\ManageUsersConnectedProjects-Help.html'
        if (-not (Test-Path -LiteralPath $helpPath -PathType Leaf)) {
            Show-Message "The user guide was not found: '$helpPath'." 'User guide' 'Warning'
            return
        }
        try { Start-Process -FilePath $helpPath }
        catch {
            Write-Warning "The user guide could not be opened: $($_.Exception.Message)"
            Show-Message $_.Exception.Message 'User guide' 'Error'
        }
    }

    #endregion Helpers

    #region Membership Manager tab

    function Initialize-ManagerTab {
        try {
            $state.ManagerView = New-iTwinMembershipManagerView -DatabasePath $databasePath -ProjectConfigPath $ProjectConfigPath -AuditLogPath $AuditLogPath
            $controls.ManagerTab.Content = $state.ManagerView.Root
        }
        catch {
            Write-Warning "The Manage Memberships tab is unavailable: $($_.Exception.Message)"
            $controls.ManagerUnavailableText.Text = "$($_.Exception.Message)`n`nDiscover projects to build the project configuration, then poll them to create the local inventory."
            $controls.ManagerTab.Content = $controls.ManagerPlaceholder
        }
    }

    function Update-InventoryViews {
        if ($null -eq $state.ManagerView) {
            Initialize-ManagerTab
            return
        }
        try { & $state.ManagerView.Refresh }
        catch { Write-Warning "The Manage Memberships tab was not refreshed: $($_.Exception.Message)" }
    }

    function Invoke-PlaceholderPoll {
        if (-not (Test-Path -LiteralPath $ProjectConfigPath -PathType Leaf)) {
            Show-Message "Project configuration was not found: '$ProjectConfigPath'. Use Discover projects first." 'Poll projects'
            return
        }
        try {
            $results = @(Invoke-WithWaitCursor { Invoke-iTwinMembershipPoll -ProjectConfigPath $ProjectConfigPath -DatabasePath $databasePath })
            foreach ($result in $results) { Write-Host "  $($result.Status) | $($result.ProjectGUID) | $($result.Members) member(s) $($result.Error)" }
            $failed = @($results | Where-Object Status -eq 'Failed').Count
            Show-Message "Polled $($results.Count) project(s): $($results.Count - $failed) succeeded, $failed failed." 'Poll projects'
            Initialize-ManagerTab
        }
        catch {
            Write-Warning "Polling failed: $($_.Exception.Message)"
            Show-Message $_.Exception.Message 'Poll projects' 'Error'
        }
    }

    #endregion Membership Manager tab

    #region Report and bulk tabs

    function Invoke-MembershipReport {
        $email = ([string]$controls.ReportEmailBox.Text).Trim()
        if ([string]::IsNullOrWhiteSpace($email)) {
            Show-Message 'Enter the email address to report on.' 'Membership report'
            return
        }
        try {
            $rows = @(Invoke-WithWaitCursor { Get-iTwinActiveMembershipReport -DatabasePath $databasePath -Email $email })
            $controls.ReportGrid.ItemsSource = @($rows | Select-Object ProjectName, ProjectGuid, Email, GivenName, Surname, Organization, RoleName, AssignedDateTimeUtc, LastObservedUtc)
            $controls.ReportStatusText.Text = "$($rows.Count) active membership role row(s) for '$email' in the local inventory."
            Write-Host "Membership report for '$email' returned $($rows.Count) row(s)."
        }
        catch {
            $controls.ReportGrid.ItemsSource = @()
            $controls.ReportStatusText.Text = "Report failed: $($_.Exception.Message)"
            Write-Warning "Membership report for '$email' failed: $($_.Exception.Message)"
        }
    }

    function Invoke-BulkActionTab {
        param([switch]$Preview)
        $actionCsvPath = ([string]$controls.BulkCsvBox.Text).Trim()
        $defaultRoleId = ([string]$controls.BulkDefaultRoleBox.Text).Trim()
        if ([string]::IsNullOrWhiteSpace($actionCsvPath)) {
            Show-Message 'Select an action CSV.' 'Bulk changes'
            return
        }
        if (-not $Preview) {
            $message = "Apply every row in '$actionCsvPath' to Bentley?`n`nRemoveMember and UnassignRole rows are not reversible. Use Preview first to validate the file."
            if ([System.Windows.MessageBox]::Show($window, $message, 'Confirm bulk changes', 'YesNo', 'Warning') -ne 'Yes') { return }
        }
        $verb = if ($Preview) { 'Preview' } else { 'Run' }
        Write-Host "$verb bulk actions from '$actionCsvPath'."
        foreach ($button in @($controls.BulkPreviewButton, $controls.BulkRunButton)) { $button.IsEnabled = $false }
        $controls.BulkStatusText.Text = "${verb}: processing '$actionCsvPath'..."
        try {
            $results = @(Invoke-WithWaitCursor { Invoke-iTwinMembershipBulkAction -ActionCsvPath $actionCsvPath -ProjectConfigPath $ProjectConfigPath -DatabasePath $databasePath -DefaultRoleId $defaultRoleId -WhatIf:$Preview })
            $controls.BulkGrid.ItemsSource = @($results | Select-Object ProjectGUID, User, Action, RoleId, Status, Detail)
            foreach ($result in $results) { Write-Host "  $($result.Status) | $($result.Action) | $($result.User) | $($result.Detail)" }
            if ($Preview) {
                $controls.BulkStatusText.Text = "Preview: $($results.Count) valid action row(s). No changes were made."
            }
            else {
                Write-iTwinMembershipManagerAudit -AuditLogPath $AuditLogPath -Results $results
                $controls.BulkStatusText.Text = "Bulk actions completed: $(@($results | Where-Object Status -eq 'Success').Count) succeeded, $(@($results | Where-Object Status -eq 'Skipped').Count) skipped, $(@($results | Where-Object Status -eq 'Failed').Count) failed. Results were written to '$AuditLogPath'."
                Update-InventoryViews
            }
            Write-Host $controls.BulkStatusText.Text
        }
        catch {
            $controls.BulkGrid.ItemsSource = @()
            $controls.BulkStatusText.Text = "${verb} failed: $($_.Exception.Message)"
            Write-Warning "Bulk actions failed before any row result was returned: $($_.Exception.Message)"
        }
        finally { foreach ($button in @($controls.BulkPreviewButton, $controls.BulkRunButton)) { $button.IsEnabled = $true } }
    }

    function Save-BulkActionTemplate {
        $dialog = New-Object Microsoft.Win32.SaveFileDialog -Property @{ Filter = 'CSV files (*.csv)|*.csv'; DefaultExt = '.csv'; FileName = 'iTwinMembershipActions.csv'; InitialDirectory = $PSScriptRoot; OverwritePrompt = $true }
        if (-not $dialog.ShowDialog($window)) { return }
        # Placeholder GUIDs are deliberately invalid so Preview and Apply reject the rows until they are replaced.
        $projectGuid = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'
        $roleGuid = 'yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy'
        $rows = @(
            [pscustomobject]@{ ProjectGUID = $projectGuid; User = 'first.user@example.invalid'; Action = 'AssignRole'; RoleId = ''; RoleDisplayName = 'Team Member' }
            [pscustomobject]@{ ProjectGUID = $projectGuid; User = 'second.user@example.invalid'; Action = 'AssignRole'; RoleId = $roleGuid; RoleDisplayName = '' }
            [pscustomobject]@{ ProjectGUID = $projectGuid; User = 'third.user@example.invalid'; Action = 'UnassignRole'; RoleId = ''; RoleDisplayName = 'Reviewer' }
            [pscustomobject]@{ ProjectGUID = $projectGuid; User = 'fourth.user@example.invalid'; Action = 'RemoveMember'; RoleId = ''; RoleDisplayName = '' }
        )
        try {
            $rows | Select-Object ProjectGUID, User, Action, RoleId, RoleDisplayName | Export-Csv -LiteralPath $dialog.FileName -NoTypeInformation -Encoding UTF8
            $controls.BulkCsvBox.Text = $dialog.FileName
            $controls.BulkStatusText.Text = "Template saved to '$($dialog.FileName)'. Replace the placeholder project GUIDs, emails, and roles before running Preview."
            Write-Host $controls.BulkStatusText.Text
        }
        catch {
            Write-Warning "Bulk action template was not saved: $($_.Exception.Message)"
            Show-Message $_.Exception.Message 'Download CSV Template' 'Error'
        }
    }

    #endregion Report and bulk tabs
}

process {
    try {
        [xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="iTwin Membership Console" Height="880" Width="1440" MinHeight="620" MinWidth="1080" WindowStartupLocation="CenterScreen" Background="#F4F6F8">
  <Grid>
  <TabControl x:Name="MainTabs" Margin="8" Background="#F4F6F8">
    <TabItem x:Name="ManagerTab" Header="Manage Memberships" ToolTip="Browse the local inventory, add users, remove memberships, and choose tracked projects">
      <StackPanel x:Name="ManagerPlaceholder" Margin="24" MaxWidth="760" HorizontalAlignment="Left">
        <TextBlock Text="The local membership inventory is not available" FontSize="18" FontWeight="SemiBold" Foreground="#17324D" />
        <TextBlock x:Name="ManagerUnavailableText" Margin="0,10,0,14" TextWrapping="Wrap" Foreground="#3E4C59" />
        <StackPanel Orientation="Horizontal">
          <Button x:Name="PlaceholderDiscoverButton" Content="Discover projects" Padding="12,5" />
          <Button x:Name="PlaceholderPollButton" Content="Poll configured projects" Padding="12,5" Margin="8,0,0,0" />
        </StackPanel>
      </StackPanel>
    </TabItem>
    <TabItem x:Name="AssignTab" Header="Assign User Roles" ToolTip="Give users on one live iTwin additional iTwin or enterprise roles" />
    <TabItem x:Name="CopyRolesTab" Header="Copy Role Definitions" ToolTip="Copy roles and their permissions between iTwins and accounts; members are not copied" />
    <TabItem Header="User Membership Report" ToolTip="List one user's active memberships and roles from the local inventory">
      <Grid Margin="14">
        <Grid.RowDefinitions><RowDefinition Height="Auto" /><RowDefinition Height="*" /><RowDefinition Height="Auto" /></Grid.RowDefinitions>
        <DockPanel Grid.Row="0" Margin="0,0,0,10">
          <TextBlock Text="User email" VerticalAlignment="Center" Margin="0,0,8,0" />
          <Button x:Name="ReportButton" Content="Run report" Padding="12,5" Margin="8,0,0,0" DockPanel.Dock="Right" />
          <TextBox x:Name="ReportEmailBox" VerticalContentAlignment="Center" />
        </DockPanel>
        <DataGrid x:Name="ReportGrid" Grid.Row="1" IsReadOnly="True" AutoGenerateColumns="True" CanUserAddRows="False" />
        <TextBlock x:Name="ReportStatusText" Grid.Row="2" Margin="0,8,0,0" TextWrapping="Wrap" Foreground="#3E4C59" Text="Reports read the local inventory. Poll projects first for current results." />
      </Grid>
    </TabItem>
    <TabItem Header="Bulk Changes (CSV)" ToolTip="Preview or apply membership changes listed in an action CSV">
      <Grid Margin="14">
        <Grid.RowDefinitions><RowDefinition Height="Auto" /><RowDefinition Height="Auto" /><RowDefinition Height="*" /><RowDefinition Height="Auto" /></Grid.RowDefinitions>
        <DockPanel Grid.Row="0" Margin="0,0,0,8">
          <TextBlock Text="Action CSV" Width="100" VerticalAlignment="Center" />
          <Button x:Name="BulkTemplateButton" Content="Download CSV Template" Padding="12,5" Margin="8,0,0,0" DockPanel.Dock="Right" ToolTip="Save an action CSV with placeholder rows that show each action" />
          <Button x:Name="BulkBrowseButton" Content="Browse..." Padding="12,5" Margin="8,0,0,0" DockPanel.Dock="Right" />
          <TextBox x:Name="BulkCsvBox" VerticalContentAlignment="Center" />
        </DockPanel>
        <DockPanel Grid.Row="1" Margin="0,0,0,10">
          <TextBlock Text="Default role ID" Width="100" VerticalAlignment="Center" />
          <Button x:Name="BulkRunButton" Content="Apply to Bentley" Padding="12,5" Margin="8,0,0,0" DockPanel.Dock="Right" Background="#B42318" Foreground="White" />
          <Button x:Name="BulkPreviewButton" Content="Preview" Padding="12,5" Margin="8,0,0,0" DockPanel.Dock="Right" />
          <TextBox x:Name="BulkDefaultRoleBox" VerticalContentAlignment="Center" ToolTip="Optional role GUID for AssignRole rows without RoleId or RoleDisplayName" />
        </DockPanel>
        <DataGrid x:Name="BulkGrid" Grid.Row="2" IsReadOnly="True" AutoGenerateColumns="True" CanUserAddRows="False" />
        <TextBlock x:Name="BulkStatusText" Grid.Row="3" Margin="0,8,0,0" TextWrapping="Wrap" Foreground="#3E4C59" Text="Columns: ProjectGUID, User, Action (AssignRole, UnassignRole, RemoveMember), and optional RoleId or RoleDisplayName. Every ProjectGUID must be in the project configuration." />
      </Grid>
    </TabItem>
  </TabControl>
  <Button x:Name="HelpButton" HorizontalAlignment="Right" VerticalAlignment="Top" Margin="0,9,14,0" Cursor="Hand" ToolTip="Open the user guide (F1)" AutomationProperties.Name="Help">
    <Button.Template>
      <ControlTemplate TargetType="Button">
        <Grid Width="22" Height="22">
          <Ellipse x:Name="Circle" Fill="White" Stroke="#2F6FAF" StrokeThickness="1.5" />
          <TextBlock Text="?" FontSize="14" FontWeight="Bold" Foreground="#2F6FAF" HorizontalAlignment="Center" VerticalAlignment="Center" />
        </Grid>
        <ControlTemplate.Triggers>
          <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Circle" Property="Fill" Value="#E8F1FB" /></Trigger>
          <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="Circle" Property="StrokeThickness" Value="2.5" /></Trigger>
        </ControlTemplate.Triggers>
      </ControlTemplate>
    </Button.Template>
  </Button>
  </Grid>
</Window>
'@

        $window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
        $controls = @{}
        foreach ($name in @('HelpButton', 'ManagerTab', 'ManagerPlaceholder', 'ManagerUnavailableText', 'PlaceholderDiscoverButton', 'PlaceholderPollButton', 'AssignTab', 'CopyRolesTab', 'ReportButton', 'ReportEmailBox', 'ReportGrid', 'ReportStatusText', 'BulkBrowseButton', 'BulkTemplateButton', 'BulkCsvBox', 'BulkRunButton', 'BulkPreviewButton', 'BulkDefaultRoleBox', 'BulkGrid', 'BulkStatusText')) {
            $controls[$name] = $window.FindName($name)
        }

        # Event handlers run in this script's scope, so shared mutable state lives in one hashtable.
        $state = @{ ManagerView = $null; RoleView = $null; RoleCopyView = $null }

        Initialize-ManagerTab
        $state.RoleView = New-iTwinRoleAssignmentView -DatabasePath $databasePath -ProjectConfigPath $ProjectConfigPath -AuditLogPath $AuditLogPath -OnAssigned { Update-InventoryViews }
        $controls.AssignTab.Content = $state.RoleView.Root
        $state.RoleCopyView = New-iTwinRoleCopyView -AuditLogPath $RoleCopyAuditLogPath
        $controls.CopyRolesTab.Content = $state.RoleCopyView.Root

        #region Events

        $controls.HelpButton.Add_Click({ Open-UserGuide })
        $window.Add_PreviewKeyDown({
                param($source, $e)
                if ($e.Key -eq [System.Windows.Input.Key]::F1) { $e.Handled = $true; Open-UserGuide }
            })
        $controls.PlaceholderDiscoverButton.Add_Click({ Show-iTwinProjectDiscoveryDialog -ProjectConfigPath $ProjectConfigPath })
        $controls.PlaceholderPollButton.Add_Click({ Invoke-PlaceholderPoll })
        $controls.ReportButton.Add_Click({ Invoke-MembershipReport })
        $controls.ReportEmailBox.Add_KeyDown({
                param($source, $e)
                if ($e.Key -eq [System.Windows.Input.Key]::Enter) { Invoke-MembershipReport }
            })
        $controls.BulkBrowseButton.Add_Click({
                $dialog = New-Object Microsoft.Win32.OpenFileDialog -Property @{ Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'; InitialDirectory = $PSScriptRoot }
                if ($dialog.ShowDialog($window)) { $controls.BulkCsvBox.Text = $dialog.FileName }
            })
        $controls.BulkTemplateButton.Add_Click({ Save-BulkActionTemplate })
        $controls.BulkPreviewButton.Add_Click({ Invoke-BulkActionTab -Preview })
        $controls.BulkRunButton.Add_Click({ Invoke-BulkActionTab })
        $window.Add_Closed({ if ($null -ne $state.ManagerView) { & $state.ManagerView.Close } })

        #endregion Events

        $window.ShowDialog() | Out-Null
    }
    catch {
        $null = Stop-Transcript
        throw
    }
}

end {
    try {
        # Display the time that this script finished running, and how long it took to run.
        [DateTime] $finishTime = Get-Date
        [TimeSpan] $elapsedTime = $finishTime - $startTime
        Write-Host "Finished script at '$($finishTime.ToString('u'))'. Took '$elapsedTime' to run."
    }
    finally {
        Stop-Transcript
    }
}
