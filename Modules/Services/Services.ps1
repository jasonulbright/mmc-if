<#
.SYNOPSIS
    Services Manager module for MMC-If (WPF).

.DESCRIPTION
    View Windows services via CIM. Start/stop/restart where user permissions
    allow (no UAC elevation). Services that require admin will simply fail
    gracefully with the native error message.
#>

function New-ServicesView {
    param([hashtable]$Context)

    $setStatus = $Context.SetStatus
    $log       = $Context.Log

    $xamlPath = Join-Path $PSScriptRoot 'Services.xaml'
    $xamlRaw  = Get-Content -LiteralPath $xamlPath -Raw
    $reader   = New-Object System.Xml.XmlNodeReader ([xml]$xamlRaw)
    $view     = [Windows.Markup.XamlReader]::Load($reader)

    $txtFilter      = $view.FindName('txtFilter')
    $btnRefresh     = $view.FindName('btnRefresh')
    $btnStart       = $view.FindName('btnStart')
    $btnStop        = $view.FindName('btnStop')
    $btnRestart     = $view.FindName('btnRestart')
    $gridSvc        = $view.FindName('gridSvc')
    $txtDetail      = $view.FindName('txtDetail')
    $mnuCopyName    = $view.FindName('mnuCopyName')
    $mnuCopyDisplay = $view.FindName('mnuCopyDisplay')
    $mnuCopyPath    = $view.FindName('mnuCopyPath')

    $state = @{
        AllServices = New-Object System.Collections.ObjectModel.ObservableCollection[object]
        Filtered    = $null
    }

    $applyFilter = {
        $ft = $txtFilter.Text
        if ([string]::IsNullOrWhiteSpace($ft)) {
            $gridSvc.ItemsSource = $state.AllServices
            return
        }
        $needle = $ft.ToLowerInvariant()
        $filtered = New-Object System.Collections.ObjectModel.ObservableCollection[object]
        foreach ($svc in $state.AllServices) {
            $hay = "$($svc.Name) $($svc.DisplayName) $($svc.FullDescription)".ToLowerInvariant()
            if ($hay.Contains($needle)) { [void]$filtered.Add($svc) }
        }
        $gridSvc.ItemsSource = $filtered
    }.GetNewClosure()

    $loadServices = {
        & $setStatus 'Loading services...'
        $view.Cursor = [System.Windows.Input.Cursors]::Wait
        $state.AllServices.Clear()
        $txtDetail.Clear()

        try {
            $services = Get-CimInstance Win32_Service -ErrorAction Stop | Sort-Object DisplayName
            foreach ($svc in $services) {
                $desc = if ($svc.Description) { $svc.Description } else { '' }
                $short = if ($desc.Length -gt 100) { $desc.Substring(0, 97) + '...' } else { $desc }
                $state.AllServices.Add([pscustomobject]@{
                    Name             = $svc.Name
                    DisplayName      = if ($svc.DisplayName) { $svc.DisplayName } else { $svc.Name }
                    Status           = $svc.State
                    StartupType      = $svc.StartMode
                    Account          = if ($svc.StartName) { $svc.StartName } else { '' }
                    Description      = $short
                    PathName         = if ($svc.PathName) { $svc.PathName } else { '' }
                    ProcessId        = if ($svc.ProcessId) { $svc.ProcessId.ToString() } else { '' }
                    FullDescription  = $desc
                })
            }
            & $setStatus "$($services.Count) services loaded"
            & $log "Loaded $($services.Count) services"
        }
        catch {
            & $setStatus "Error loading services: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show("Failed to load services: $($_.Exception.Message)", 'Error',
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
        }

        & $applyFilter
        $view.Cursor = [System.Windows.Input.Cursors]::Arrow
    }.GetNewClosure()

    $btnRefresh.Add_Click({ & $loadServices }.GetNewClosure())
    $txtFilter.Add_TextChanged({ & $applyFilter }.GetNewClosure())

    $gridSvc.Add_SelectionChanged({
        $sel = $gridSvc.SelectedItem
        if (-not $sel) {
            $txtDetail.Clear()
            $btnStart.IsEnabled = $false
            $btnStop.IsEnabled = $false
            $btnRestart.IsEnabled = $false
            return
        }

        $lines = @(
            "Service: $($sel.Name)",
            "Display Name: $($sel.DisplayName)",
            "Status: $($sel.Status)  |  Startup: $($sel.StartupType)  |  Account: $($sel.Account)",
            "Path: $($sel.PathName)",
            "PID: $($sel.ProcessId)",
            '',
            'Description:',
            $sel.FullDescription,
            ''
        )
        try {
            $svcObj = Get-Service -Name $sel.Name -ErrorAction Stop
            $deps = @($svcObj.ServicesDependedOn | ForEach-Object { "  $($_.Name) ($($_.DisplayName))" })
            $dependents = @($svcObj.DependentServices | ForEach-Object { "  $($_.Name) ($($_.DisplayName))" })
            $lines += 'Depends On:'
            if ($deps.Count -gt 0) { $lines += $deps } else { $lines += '  (none)' }
            $lines += ''
            $lines += 'Dependent Services:'
            if ($dependents.Count -gt 0) { $lines += $dependents } else { $lines += '  (none)' }
        }
        catch { $lines += 'Dependencies: (could not query)' }

        $txtDetail.Text = $lines -join "`r`n"

        # Hide buttons that don't apply to the selected service's state so
        # we never render a greyed-out disabled button (MahApps disabled-state
        # contrast fails WCAG AA and the visual is non-brand).
        $btnStart.Visibility   = if ($sel.Status -eq 'Stopped') { 'Visible' } else { 'Collapsed' }
        $btnStop.Visibility    = if ($sel.Status -eq 'Running') { 'Visible' } else { 'Collapsed' }
        $btnRestart.Visibility = if ($sel.Status -eq 'Running') { 'Visible' } else { 'Collapsed' }
    }.GetNewClosure())

    $doServiceAction = {
        param([string]$Action, [string]$Name, [string]$DisplayName)
        if ($Action -ne 'Start') {
            $res = [System.Windows.MessageBox]::Show("$Action service '$DisplayName' ($Name)?", 'Confirm',
                [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
            if ($res -ne [System.Windows.MessageBoxResult]::Yes) { return }
        }
        $view.Cursor = [System.Windows.Input.Cursors]::Wait
        try {
            switch ($Action) {
                'Start'   { Start-Service   -Name $Name -ErrorAction Stop }
                'Stop'    { Stop-Service    -Name $Name -Force -ErrorAction Stop }
                'Restart' { Restart-Service -Name $Name -Force -ErrorAction Stop }
            }
            & $setStatus "$Action ed service: $Name"
            & $log "$Action service: $Name"
            & $loadServices
        }
        catch {
            [System.Windows.MessageBox]::Show("Failed to $($Action.ToLower()) service '$Name': $($_.Exception.Message)",
                'Error', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
            & $setStatus "Failed to $($Action.ToLower()) $Name"
        }
        $view.Cursor = [System.Windows.Input.Cursors]::Arrow
    }.GetNewClosure()

    $btnStart.Add_Click({
        $sel = $gridSvc.SelectedItem
        if ($sel) { & $doServiceAction 'Start' $sel.Name $sel.DisplayName }
    }.GetNewClosure())
    $btnStop.Add_Click({
        $sel = $gridSvc.SelectedItem
        if ($sel) { & $doServiceAction 'Stop' $sel.Name $sel.DisplayName }
    }.GetNewClosure())
    $btnRestart.Add_Click({
        $sel = $gridSvc.SelectedItem
        if ($sel) { & $doServiceAction 'Restart' $sel.Name $sel.DisplayName }
    }.GetNewClosure())

    $mnuCopyName.Add_Click({
        $sel = $gridSvc.SelectedItem
        if ($sel) { [System.Windows.Clipboard]::SetText([string]$sel.Name) }
    }.GetNewClosure())
    $mnuCopyDisplay.Add_Click({
        $sel = $gridSvc.SelectedItem
        if ($sel) { [System.Windows.Clipboard]::SetText([string]$sel.DisplayName) }
    }.GetNewClosure())
    $mnuCopyPath.Add_Click({
        $sel = $gridSvc.SelectedItem
        if ($sel -and $sel.PathName) { [System.Windows.Clipboard]::SetText([string]$sel.PathName) }
    }.GetNewClosure())

    $view.Add_Loaded({ if ($state.AllServices.Count -eq 0) { & $loadServices } }.GetNewClosure())

    return $view
}
