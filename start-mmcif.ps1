<#
.SYNOPSIS
    MahApps.Metro WPF shell for MMC-If - plugin tools that replace mmc.exe snap-ins.

.DESCRIPTION
    MMC-If is a modular WPF application that provides alternatives to Windows
    administration tools hosted in mmc.exe. The underlying .NET/WMI APIs work
    without elevation even when endpoint policy blocks mmc.exe itself.

    The shell is read-only except where user permissions allow (e.g., HKCU
    registry writes, services the user owns). No UAC elevation is ever
    attempted - the tool targets non-admin endpoint users.

.EXAMPLE
    .\start-mmcif.ps1

.NOTES
    Requirements:
      - PowerShell 5.1
      - .NET Framework 4.8+
      - MahApps.Metro DLLs in .\Lib\

    ScriptName : start-mmcif.ps1
    Purpose    : WPF shell for MMC alternative tools
    Version    : 1.0.0
    Updated    : 2026-04-17
#>

param()

$ErrorActionPreference = 'Stop'

# =============================================================================
# Assembly loading (must happen before XAML parse)
# =============================================================================
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

$libDir = Join-Path $PSScriptRoot 'Lib'
[System.Reflection.Assembly]::LoadFrom((Join-Path $libDir 'Microsoft.Xaml.Behaviors.dll')) | Out-Null
[System.Reflection.Assembly]::LoadFrom((Join-Path $libDir 'ControlzEx.dll')) | Out-Null
[System.Reflection.Assembly]::LoadFrom((Join-Path $libDir 'MahApps.Metro.dll')) | Out-Null

# =============================================================================
# Shared helpers
# =============================================================================
$moduleRoot = Join-Path $PSScriptRoot 'Module'
Import-Module (Join-Path $moduleRoot 'MmcIfCommon.psd1') -Force -DisableNameChecking

$toolLogFolder = Join-Path $PSScriptRoot 'Logs'
if (-not (Test-Path -LiteralPath $toolLogFolder)) {
    New-Item -ItemType Directory -Path $toolLogFolder -Force | Out-Null
}
$toolLogPath = Join-Path $toolLogFolder ("MmcIf-{0}.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
Initialize-Logging -LogPath $toolLogPath

# =============================================================================
# Preferences
# =============================================================================
function Get-MmcIfPreference {
    $path = Join-Path $PSScriptRoot 'MmcIf.prefs.json'
    $defaults = [pscustomobject]@{
        DarkMode          = $true
        DisabledModules   = @()
        RegistryFavorites = @()
    }
    if (-not (Test-Path -LiteralPath $path)) { return $defaults }
    try {
        $raw = Get-Content -LiteralPath $path -Raw
        if ([string]::IsNullOrWhiteSpace($raw)) { return $defaults }
        $data = $raw | ConvertFrom-Json
        if ($null -ne $data.DarkMode)          { $defaults.DarkMode          = [bool]$data.DarkMode }
        if ($null -ne $data.DisabledModules)   { $defaults.DisabledModules   = @($data.DisabledModules) }
        if ($null -ne $data.RegistryFavorites) { $defaults.RegistryFavorites = @($data.RegistryFavorites) }
    }
    catch { }
    return $defaults
}

function Save-MmcIfPreference {
    param([pscustomobject]$Prefs)
    $path = Join-Path $PSScriptRoot 'MmcIf.prefs.json'
    $Prefs | ConvertTo-Json | Set-Content -LiteralPath $path -Encoding UTF8
}

$script:Prefs = Get-MmcIfPreference

# =============================================================================
# Window state (scriptblocks so WPF event handlers can invoke them across scopes)
# =============================================================================
$script:GetWindowStatePath = { Join-Path $PSScriptRoot 'MmcIf.windowstate.json' }

$script:SaveWindowState = {
    param([System.Windows.Window]$Window)
    $state = @{
        X         = [int]$Window.Left
        Y         = [int]$Window.Top
        Width     = [int]$Window.Width
        Height    = [int]$Window.Height
        Maximized = ($Window.WindowState -eq [System.Windows.WindowState]::Maximized)
    }
    try { $state | ConvertTo-Json | Set-Content -LiteralPath (& $script:GetWindowStatePath) -Encoding UTF8 }
    catch { }
}

$script:RestoreWindowState = {
    param([System.Windows.Window]$Window)
    $path = & $script:GetWindowStatePath
    if (-not (Test-Path -LiteralPath $path)) { return }
    try {
        $state = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
        if ($null -ne $state.Width -and $state.Width -gt 600)   { $Window.Width  = [double]$state.Width }
        if ($null -ne $state.Height -and $state.Height -gt 400) { $Window.Height = [double]$state.Height }
        if ($null -ne $state.X -and $null -ne $state.Y) {
            $Window.WindowStartupLocation = [System.Windows.WindowStartupLocation]::Manual
            $Window.Left = [double]$state.X
            $Window.Top  = [double]$state.Y
        }
        if ($state.Maximized) { $Window.WindowState = [System.Windows.WindowState]::Maximized }
    }
    catch { }
}

# =============================================================================
# XAML parse
# =============================================================================
$xamlPath = Join-Path $PSScriptRoot 'MainWindow.xaml'
$xamlRaw  = Get-Content -LiteralPath $xamlPath -Raw

$reader = New-Object System.Xml.XmlNodeReader ([xml]$xamlRaw)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Helper: find by name
function Get-Control { param([string]$Name) $window.FindName($Name) }

# Write-Log reference (MmcIfCommon module function) - captured as a scriptblock
# so handlers can invoke it after function-lookup falls out of scope.
$script:WriteLogSb = ${function:Write-Log}

$txtAppVersion    = Get-Control 'txtAppVersion'
$txtModuleTitle   = Get-Control 'txtModuleTitle'
$txtModuleSubtitle = Get-Control 'txtModuleSubtitle'
$contentHost      = Get-Control 'contentHost'
$txtLog           = Get-Control 'txtLog'
$txtStatus        = Get-Control 'txtStatus'
$toggleTheme      = Get-Control 'toggleTheme'

# =============================================================================
# Version stamp
# =============================================================================
$manifestPath = Join-Path $PSScriptRoot 'project-manifest.json'
if (Test-Path -LiteralPath $manifestPath) {
    try {
        $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
        if ($manifest.Version) { $txtAppVersion.Text = "v$($manifest.Version)" }
    }
    catch { }
}

# =============================================================================
# Log helpers (scriptblocks)
# =============================================================================
$script:AddLogLine = {
    param([string]$Message, [string]$Level = 'INFO')
    $ts = Get-Date -Format 'HH:mm:ss'
    $line = "{0}  [{1,-5}] {2}" -f $ts, $Level, $Message
    if ([string]::IsNullOrWhiteSpace($txtLog.Text)) { $txtLog.Text = $line }
    else { $txtLog.AppendText([Environment]::NewLine + $line) }
    $txtLog.ScrollToEnd()
    if ($script:WriteLogSb) { & $script:WriteLogSb -Message $Message -Level $Level -Quiet }
}

$script:SetStatus = {
    param([string]$Message)
    $txtStatus.Text = $Message
}

# =============================================================================
# Module registry - placeholders until WPF ports land (sessions 3-8)
# =============================================================================
$script:ModuleInfo = @{
    RegistryBrowser  = @{ Title = 'Registry';          Subtitle = 'Browse HKLM/HKCR/HKU/HKCC read-only; HKCU read-write.'; PortedIn = 'Session 3' }
    EventLogViewer   = @{ Title = 'Event Logs';        Subtitle = 'Application / System / Security / Setup via Get-WinEvent.'; PortedIn = 'Session 3' }
    DeviceInfo       = @{ Title = 'Device Info';       Subtitle = 'WMI-based system, BIOS, CPU, memory, storage, network, devices.'; PortedIn = 'Session 4' }
    FileExplorer     = @{ Title = 'Files';             Subtitle = 'File browser with admin-friendly defaults and 7-Zip archive traversal.'; PortedIn = 'Session 4' }
    Services         = @{ Title = 'Services';          Subtitle = 'View and control services where user permissions allow.'; PortedIn = 'Session 5' }
    CertificateStore = @{ Title = 'Certificates';      Subtitle = 'CurrentUser / LocalMachine certificate stores, read-only.'; PortedIn = 'Session 5' }
    UsersGroups      = @{ Title = 'Users & Groups';    Subtitle = 'Local users and groups, read-only.'; PortedIn = 'Session 6' }
    TaskScheduler    = @{ Title = 'Task Scheduler';    Subtitle = 'Browse scheduled tasks, triggers, actions, and run history. Read-only.'; PortedIn = 'Session 7' }
    Networking       = @{ Title = 'Networking';        Subtitle = 'Adapters, IP config, DNS. Read-only with ping/tracert actions.'; PortedIn = 'Session 7' }
    Disks            = @{ Title = 'Disks';             Subtitle = 'Physical disks, partitions, volumes. Read-only.'; PortedIn = 'Session 8' }
    GroupPolicy      = @{ Title = 'Group Policy';      Subtitle = 'Local GPO viewer via gpresult /h. Read-only.'; PortedIn = 'Session 8' }
}

$script:LoadedModules = @{}

$script:GetModuleManifest = {
    param([string]$ModuleKey)
    $manifestPath = Join-Path $PSScriptRoot "Modules\$ModuleKey\module.json"
    if (-not (Test-Path -LiteralPath $manifestPath)) { return $null }
    try { Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json }
    catch { return $null }
}

$script:ShowModule = {
    param([string]$Key)
    $info = $script:ModuleInfo[$Key]
    if (-not $info) { return }

    $txtModuleTitle.Text    = $info.Title
    $txtModuleSubtitle.Text = $info.Subtitle

    $manifest = & $script:GetModuleManifest -ModuleKey $Key

    if ($manifest -and $manifest.UI -eq 'WPF' -and $manifest.EntryScript -and $manifest.InitFunction) {
        try {
            if (-not $script:LoadedModules[$Key]) {
                $entry = Join-Path $PSScriptRoot "Modules\$Key\$($manifest.EntryScript)"
                . $entry
                # Capture the module's init function as a scriptblock so we can
                # invoke it across dispatcher scopes without a name-based lookup.
                $fn = Get-Item -LiteralPath "function:\$($manifest.InitFunction)" -ErrorAction Stop
                $script:LoadedModules[$Key] = $fn.ScriptBlock
            }
            $initSb = $script:LoadedModules[$Key]
            $ctx = @{
                Prefs     = $script:Prefs
                SavePrefs = $script:SaveMmcIfPreferenceSb
                SetStatus = $script:SetStatus
                Log       = $script:AddLogLine
                Window    = $window
            }
            $userControl = & $initSb -Context $ctx
            if ($userControl) {
                $contentHost.Content = $userControl
                & $script:AddLogLine -Message "Loaded module: $($info.Title)"
                & $script:SetStatus -Message "Ready - $($info.Title)"
                return
            }
        }
        catch {
            & $script:AddLogLine -Message "Failed to load $($info.Title): $($_.Exception.Message)" -Level 'ERROR'
            & $script:SetStatus -Message "Error loading $($info.Title)"
        }
    }

    $placeholder = New-Object System.Windows.Controls.TextBlock
    $placeholder.Text = "Module under construction - scheduled for $($info.PortedIn) of the v1.0 waterfall."
    $placeholder.FontSize = 13
    $placeholder.Margin = '16,24,16,16'
    $placeholder.TextWrapping = [System.Windows.TextWrapping]::Wrap
    $placeholder.Foreground = [System.Windows.Media.Brushes]::Gray

    $contentHost.Content = $placeholder

    & $script:AddLogLine -Message "Selected module: $($info.Title)"
    & $script:SetStatus -Message "Ready - $($info.Title)"
}

# =============================================================================
# Sidebar button wiring
# =============================================================================
$moduleButtons = @(
    'btnRegistry', 'btnEventLogs', 'btnDeviceInfo', 'btnFiles', 'btnServices',
    'btnCertificates', 'btnUsersGroups', 'btnTaskScheduler', 'btnNetworking',
    'btnDisks', 'btnGroupPolicy'
)
foreach ($btnName in $moduleButtons) {
    $btn = Get-Control $btnName
    if ($btn) {
        $btn.Add_Click({
            $sender = $args[0]
            & $script:ShowModule -Key $sender.Tag
        }.GetNewClosure())
    }
}

# Capture Save-MmcIfPreference as a scriptblock for the theme-toggle handler.
$script:SaveMmcIfPreferenceSb = ${function:Save-MmcIfPreference}

# =============================================================================
# Theme (live hot-swap via ControlzEx ThemeManager + button color swap)
# =============================================================================
# The MahApps ChangeTheme call flips the theme-aware brushes, but mmc-if's
# sidebar buttons hard-code #1E1E1E / #555555 in their Style (needed for the
# dark-mode aesthetic). In light mode those black buttons on a white window
# look broken, so we also swap Background / BorderBrush on every sidebar
# button when the theme flips. Matches app-packager's Set-ButtonTheme pattern
# (which is the canonical reference - see reference_srl_wpf_brand.md).
$txtThemeLabel = Get-Control 'txtThemeLabel'

$script:DarkButtonBg         = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#1E1E1E')
$script:DarkButtonBorder     = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#555555')
$script:LightButtonBg        = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#0078D4')  # Windows blue
$script:LightButtonBorder    = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#006CBE')
$script:TitleBarBlue         = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#0078D4')
$script:TitleBarBlueInactive = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#4BA3E0')

$script:SidebarButtons = @(
    'btnRegistry','btnEventLogs','btnDeviceInfo','btnFiles','btnServices',
    'btnCertificates','btnUsersGroups','btnTaskScheduler','btnNetworking',
    'btnDisks','btnGroupPolicy','btnPreferences'
) | ForEach-Object { Get-Control $_ } | Where-Object { $_ }

$script:ApplyTheme = {
    param([bool]$IsDark)
    $themeName = if ($IsDark) { 'Dark.Steel' } else { 'Light.Blue' }
    [ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($window, $themeName) | Out-Null
    if ($txtThemeLabel) { $txtThemeLabel.Text = if ($IsDark) { 'Dark Theme' } else { 'Light Theme' } }

    $bg     = if ($IsDark) { $script:DarkButtonBg }     else { $script:LightButtonBg }
    $border = if ($IsDark) { $script:DarkButtonBorder } else { $script:LightButtonBorder }
    foreach ($b in $script:SidebarButtons) {
        $b.Background  = $bg
        $b.BorderBrush = $border
    }

    # Title bar: dark mode keeps the Steel theme default; light mode paints
    # Windows blue active / desaturated blue inactive (matches app-packager).
    if ($IsDark) {
        $window.ClearValue([MahApps.Metro.Controls.MetroWindow]::WindowTitleBrushProperty)
        $window.ClearValue([MahApps.Metro.Controls.MetroWindow]::NonActiveWindowTitleBrushProperty)
    } else {
        $window.WindowTitleBrush          = $script:TitleBarBlue
        $window.NonActiveWindowTitleBrush = $script:TitleBarBlueInactive
    }

    # Any currently-loaded module's UserControl also needs the theme applied so
    # its DynamicResource bindings pick up the new brushes.
    if ($contentHost -and $contentHost.Content) {
        try { [ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($contentHost.Content, $themeName) | Out-Null } catch { }
    }
}

# Apply persisted theme BEFORE the window paints for the first time.
& $script:ApplyTheme -IsDark ([bool]$script:Prefs.DarkMode)

$toggleTheme.IsOn = [bool]$script:Prefs.DarkMode

$toggleTheme.Add_Toggled({
    $newState = [bool]$toggleTheme.IsOn
    if ($script:Prefs.DarkMode -eq $newState) { return }
    $script:Prefs.DarkMode = $newState
    & $script:SaveMmcIfPreferenceSb -Prefs $script:Prefs
    & $script:ApplyTheme -IsDark $newState
    & $script:SetStatus -Message ("Theme: " + $(if ($newState) { 'Dark' } else { 'Light' }))
    & $script:AddLogLine -Message ("Theme changed -> " + $(if ($newState) { 'Dark' } else { 'Light' }))
}.GetNewClosure())

# =============================================================================
# Preferences dialog (in-app, MetroWindow)
# =============================================================================
function Show-MmcIfPreferencesDialog {
    param([Parameter(Mandatory)]$Owner)

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="Preferences"
    Width="620" Height="560"
    MinWidth="560" MinHeight="480"
    WindowStartupLocation="CenterOwner"
    TitleCharacterCasing="Normal"
    ShowIconOnTitleBar="False"
    ResizeMode="CanResize"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1">

    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml"/>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml"/>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml"/>
            </ResourceDictionary.MergedDictionaries>
        </ResourceDictionary>
    </Window.Resources>

    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Appearance -->
        <TextBlock Grid.Row="0" Text="Appearance" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,8"/>
        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="4,0,0,16">
            <Controls:ToggleSwitch x:Name="togDark" Header="Dark theme" OnContent="On" OffContent="Off" VerticalAlignment="Center"/>
            <TextBlock Text="(applied on OK)" Margin="10,0,0,0" VerticalAlignment="Center"
                       Foreground="{DynamicResource MahApps.Brushes.Gray5}"/>
        </StackPanel>

        <!-- Favorites -->
        <TextBlock Grid.Row="2" Text="Registry Favorites" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,8"/>
        <DataGrid Grid.Row="3" x:Name="gridFavs"
                  AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False"
                  IsReadOnly="True" SelectionMode="Extended"
                  HeadersVisibility="Column" GridLinesVisibility="Horizontal"
                  RowHeaderWidth="0" Margin="0,0,0,8" MinHeight="140">
            <DataGrid.Columns>
                <DataGridTextColumn Header="Name" Width="160" Binding="{Binding Name}"/>
                <DataGridTextColumn Header="Path" Width="*"   Binding="{Binding Path}"/>
            </DataGrid.Columns>
        </DataGrid>
        <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Left" Margin="0,0,0,16">
            <Button Controls:ControlsHelper.ContentCharacterCasing="Normal" x:Name="btnFavRemove"
                    Content="Remove Selected" MinWidth="140" Height="28"
                    Style="{DynamicResource MahApps.Styles.Button.Square}"/>
        </StackPanel>

        <!-- About -->
        <StackPanel Grid.Row="5" Margin="0,0,0,16">
            <TextBlock Text="About" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,8"/>
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="110"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <TextBlock Grid.Row="0" Grid.Column="0" Text="Version:"      Foreground="{DynamicResource MahApps.Brushes.Gray5}" Margin="4,0,0,4"/>
                <TextBlock Grid.Row="0" Grid.Column="1" x:Name="txtVersion"  Margin="0,0,0,4" FontFamily="Cascadia Code, Consolas"/>
                <TextBlock Grid.Row="1" Grid.Column="0" Text="Prefs file:"   Foreground="{DynamicResource MahApps.Brushes.Gray5}" Margin="4,0,0,4"/>
                <TextBlock Grid.Row="1" Grid.Column="1" x:Name="txtPrefsPath" Margin="0,0,0,4" FontFamily="Cascadia Code, Consolas" TextWrapping="NoWrap" TextTrimming="CharacterEllipsis"/>
                <TextBlock Grid.Row="2" Grid.Column="0" Text="Log folder:"   Foreground="{DynamicResource MahApps.Brushes.Gray5}" Margin="4,0,0,4"/>
                <TextBlock Grid.Row="2" Grid.Column="1" x:Name="txtLogFolder" Margin="0,0,0,4" FontFamily="Cascadia Code, Consolas" TextWrapping="NoWrap" TextTrimming="CharacterEllipsis"/>
                <StackPanel Grid.Row="3" Grid.Column="1" Orientation="Horizontal" Margin="0,8,0,0">
                    <Button Controls:ControlsHelper.ContentCharacterCasing="Normal" x:Name="btnOpenLog"
                            Content="Open Log Folder" MinWidth="130" Height="28" Margin="0,0,8,0"
                            Style="{DynamicResource MahApps.Styles.Button.Square}"/>
                    <Button Controls:ControlsHelper.ContentCharacterCasing="Normal" x:Name="btnEditRaw"
                            Content="Edit Prefs File (advanced)" MinWidth="180" Height="28"
                            Style="{DynamicResource MahApps.Styles.Button.Square}"/>
                </StackPanel>
            </Grid>
        </StackPanel>

        <!-- OK / Cancel -->
        <StackPanel Grid.Row="6" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Controls:ControlsHelper.ContentCharacterCasing="Normal" x:Name="btnOK"
                    Content="OK" MinWidth="90" Height="32" Margin="0,0,8,0" IsDefault="True"
                    Style="{DynamicResource MahApps.Styles.Button.Square.Accent}"/>
            <Button Controls:ControlsHelper.ContentCharacterCasing="Normal" x:Name="btnCancel"
                    Content="Cancel" MinWidth="90" Height="32" IsCancel="True"
                    Style="{DynamicResource MahApps.Styles.Button.Square}"/>
        </StackPanel>
    </Grid>
</Controls:MetroWindow>
'@

    [xml]$dlgXml = $dlgXaml
    $dlgReader = New-Object System.Xml.XmlNodeReader $dlgXml
    $dlg = [System.Windows.Markup.XamlReader]::Load($dlgReader)

    $currentTheme = [ControlzEx.Theming.ThemeManager]::Current.DetectTheme($Owner)
    if ($currentTheme) { [ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($dlg, $currentTheme) }
    $dlg.Owner = $Owner

    $togDark       = $dlg.FindName('togDark')
    $gridFavs      = $dlg.FindName('gridFavs')
    $btnFavRemove  = $dlg.FindName('btnFavRemove')
    $txtVersion    = $dlg.FindName('txtVersion')
    $txtPrefsPath  = $dlg.FindName('txtPrefsPath')
    $txtLogFolder  = $dlg.FindName('txtLogFolder')
    $btnOpenLog    = $dlg.FindName('btnOpenLog')
    $btnEditRaw    = $dlg.FindName('btnEditRaw')
    $btnOK         = $dlg.FindName('btnOK')
    $btnCancel     = $dlg.FindName('btnCancel')

    # Seed current values
    $togDark.IsOn        = [bool]$script:Prefs.DarkMode
    $txtVersion.Text     = $txtAppVersion.Text
    $txtPrefsPath.Text   = Join-Path $PSScriptRoot 'MmcIf.prefs.json'
    $txtLogFolder.Text   = Join-Path $PSScriptRoot 'Logs'

    $favItems = New-Object System.Collections.ObjectModel.ObservableCollection[object]
    foreach ($f in @($script:Prefs.RegistryFavorites)) {
        if ($f.Name -and $f.Path) {
            $favItems.Add([pscustomobject]@{ Name = [string]$f.Name; Path = [string]$f.Path })
        }
    }
    $gridFavs.ItemsSource = $favItems

    $btnFavRemove.Add_Click({
        $selected = @($gridFavs.SelectedItems)
        foreach ($sel in $selected) { [void]$favItems.Remove($sel) }
    }.GetNewClosure())

    $btnOpenLog.Add_Click({
        $logFolder = $txtLogFolder.Text
        if (-not (Test-Path -LiteralPath $logFolder)) { New-Item -ItemType Directory -Path $logFolder -Force | Out-Null }
        Start-Process -FilePath 'explorer.exe' -ArgumentList "`"$logFolder`""
    }.GetNewClosure())

    $btnEditRaw.Add_Click({
        $p = $txtPrefsPath.Text
        if (-not (Test-Path -LiteralPath $p)) { Save-MmcIfPreference -Prefs $script:Prefs }
        Start-Process -FilePath 'notepad.exe' -ArgumentList "`"$p`""
    }.GetNewClosure())

    $script:__mmcifPrefsSaved = $false
    $btnOK.Add_Click({
        $newDark = [bool]$togDark.IsOn
        $darkChanged = ($script:Prefs.DarkMode -ne $newDark)
        # Use Add-Member -Force so the write succeeds whether or not the
        # property currently exists on the pscustomobject. Direct property
        # assignment via dot-notation fails with "property cannot be found"
        # on certain PSCustomObject instances after a round-trip through
        # ConvertFrom-Json + partial property updates in PS 5.1.
        $newFavs = @($favItems | ForEach-Object { @{ Name = $_.Name; Path = $_.Path } })
        $script:Prefs | Add-Member -NotePropertyName 'DarkMode'          -NotePropertyValue $newDark -Force
        $script:Prefs | Add-Member -NotePropertyName 'RegistryFavorites' -NotePropertyValue $newFavs -Force
        & $script:SaveMmcIfPreferenceSb -Prefs $script:Prefs
        $script:__mmcifPrefsSaved = $true
        if ($darkChanged) {
            & $script:ApplyTheme -IsDark $newDark
            if ($toggleTheme) { $toggleTheme.IsOn = $newDark }
        }
        & $script:SetStatus -Message 'Preferences saved.'
        $dlg.Close()
    }.GetNewClosure())

    [void]$dlg.ShowDialog()
    return $script:__mmcifPrefsSaved
}

$script:ShowMmcIfPreferencesDialogSb = ${function:Show-MmcIfPreferencesDialog}

# =============================================================================
# Window lifecycle
# =============================================================================
& $script:RestoreWindowState -Window $window

$window.Add_Closing({
    & $script:SaveWindowState -Window $window
    & $script:AddLogLine -Message 'MMC-If shell closing.'
}.GetNewClosure())

# Initial state: no module selected - welcome panel only
& $script:AddLogLine -Message 'MMC-If shell initialized.'
& $script:AddLogLine -Message ("Log: {0}" -f $toolLogPath)
& $script:SetStatus -Message 'Ready. Select a tool from the sidebar.'

# Global safety net: log any unhandled dispatcher exception so we see what's
# killing the host instead of the terminal vanishing silently.
$dispatcher = $window.Dispatcher
$dispatcher.add_UnhandledException({
    param($s, $e)
    try {
        $msg = "UNHANDLED: {0}`r`n{1}" -f $e.Exception.Message, $e.Exception.StackTrace
        if ($script:WriteLogSb) { & $script:WriteLogSb -Message $msg -Level 'ERROR' }
        & $script:AddLogLine -Message $msg -Level 'ERROR'
    } catch { }
    $e.Handled = $true
}.GetNewClosure())

try {
    [void]$window.ShowDialog()
}
catch {
    $msg = "ShowDialog crash: {0}`r`n{1}" -f $_.Exception.Message, $_.ScriptStackTrace
    if ($script:WriteLogSb) { try { & $script:WriteLogSb -Message $msg -Level 'ERROR' } catch { } }
    try { Set-Content -LiteralPath (Join-Path $toolLogFolder 'crash.log') -Value $msg -Encoding UTF8 } catch { }
}
