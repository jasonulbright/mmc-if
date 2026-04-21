<#
.SYNOPSIS
    Event Log Viewer module for MMC-If (WPF).

.DESCRIPTION
    Browse and filter Windows event logs using Get-WinEvent.
    Does not require Event Viewer (mmc.exe) or elevation for Application/System logs.
    Read-only: does not write to the event log.
#>

function New-EventLogViewerView {
    param([hashtable]$Context)

    $setStatus = $Context.SetStatus
    $log       = $Context.Log

    $xamlPath = Join-Path $PSScriptRoot 'EventLogViewer.xaml'
    $xamlRaw  = Get-Content -LiteralPath $xamlPath -Raw
    $reader   = New-Object System.Xml.XmlNodeReader ([xml]$xamlRaw)
    $view     = [Windows.Markup.XamlReader]::Load($reader)

    $cboLog         = $view.FindName('cboLog')
    $cboLevel       = $view.FindName('cboLevel')
    $cboTime        = $view.FindName('cboTime')
    $txtMax         = $view.FindName('txtMax')
    $btnLoad        = $view.FindName('btnLoad')
    $txtFilter      = $view.FindName('txtFilter')
    $gridEvents     = $view.FindName('gridEvents')
    $txtDetail      = $view.FindName('txtDetail')
    $mnuCopyMessage = $view.FindName('mnuCopyMessage')
    $mnuCopyRow     = $view.FindName('mnuCopyRow')

    $state = @{
        AllEvents = New-Object System.Collections.ObjectModel.ObservableCollection[object]
        Filtered  = $null
    }

    $applyFilter = {
        $filterText = $txtFilter.Text
        if ([string]::IsNullOrWhiteSpace($filterText)) {
            $gridEvents.ItemsSource = $state.AllEvents
            return
        }
        $needle = $filterText.ToLowerInvariant()
        $filtered = New-Object System.Collections.ObjectModel.ObservableCollection[object]
        foreach ($evt in $state.AllEvents) {
            $hay = "$($evt.Source) $($evt.Message) $($evt.EventId)".ToLowerInvariant()
            if ($hay.Contains($needle)) { $filtered.Add($evt) }
        }
        $gridEvents.ItemsSource = $filtered
    }.GetNewClosure()

    $loadEvents = {
        $logName   = $cboLog.SelectedItem.Content
        $levelSel  = $cboLevel.SelectedItem.Content
        $timeSel   = $cboTime.SelectedItem.Content
        $maxParsed = 0
        $maxEvents = if ([int]::TryParse($txtMax.Text, [ref]$maxParsed)) {
            [Math]::Max(50, [Math]::Min(10000, $maxParsed))
        } else { 500 }
        $txtMax.Text = $maxEvents.ToString()

        & $setStatus "Loading $logName events..."
        $view.Cursor = [System.Windows.Input.Cursors]::Wait
        $state.AllEvents.Clear()
        $txtDetail.Clear()

        $filter = @{ LogName = $logName }
        switch ($levelSel) {
            'Critical'    { $filter['Level'] = 1 }
            'Error'       { $filter['Level'] = 2 }
            'Warning'     { $filter['Level'] = 3 }
            'Information' { $filter['Level'] = 4 }
        }
        switch ($timeSel) {
            'Last 1 Hour'   { $filter['StartTime'] = (Get-Date).AddHours(-1) }
            'Last 24 Hours' { $filter['StartTime'] = (Get-Date).AddDays(-1) }
            'Last 7 Days'   { $filter['StartTime'] = (Get-Date).AddDays(-7) }
            'Last 30 Days'  { $filter['StartTime'] = (Get-Date).AddDays(-30) }
        }

        try {
            $events = Get-WinEvent -FilterHashtable $filter -MaxEvents $maxEvents -ErrorAction Stop
            foreach ($evt in $events) {
                $levelStr = switch ($evt.Level) {
                    1 { 'Critical' } 2 { 'Error' } 3 { 'Warning' } 4 { 'Information' }
                    0 { 'Information' } 5 { 'Verbose' }
                    default { "Level $($evt.Level)" }
                }
                $msg = if ($evt.Message) { $evt.Message } else { '(No message)' }
                $msgOneLine = ($msg -split "`n")[0].Trim()

                $state.AllEvents.Add([pscustomobject]@{
                    Level       = $levelStr
                    Time        = $evt.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')
                    Source      = $evt.ProviderName
                    EventId     = $evt.Id.ToString()
                    Message     = $msgOneLine
                    FullMessage = $msg
                })
            }
            & $setStatus "$($logName): $($events.Count) event(s) loaded"
            & $log "Loaded $($events.Count) events from $logName log"
        }
        catch {
            $errMsg = $_.Exception.Message
            if ($errMsg -like '*No events were found*') {
                & $setStatus "$($logName): No events found for '$timeSel' - try expanding the time range"
            }
            elseif ($errMsg -like '*access*' -or $errMsg -like '*privilege*') {
                & $setStatus "$($logName): Access denied - Security log typically requires Event Log Readers group membership"
                $state.AllEvents.Add([pscustomobject]@{
                    Level       = 'Error'
                    Time        = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                    Source      = 'MMC-If'
                    EventId     = '0'
                    Message     = "Access denied reading $logName log."
                    FullMessage = $errMsg
                })
            }
            else {
                & $setStatus "Error loading events: $errMsg"
                [System.Windows.MessageBox]::Show("Failed to load events: $errMsg", 'Error',
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Error) | Out-Null
            }
        }

        & $applyFilter
        $view.Cursor = [System.Windows.Input.Cursors]::Arrow
    }.GetNewClosure()

    $btnLoad.Add_Click({ & $loadEvents }.GetNewClosure())

    $txtFilter.Add_TextChanged({ & $applyFilter }.GetNewClosure())

    $gridEvents.Add_SelectionChanged({
        $sel = $gridEvents.SelectedItem
        if ($sel) { $txtDetail.Text = [string]$sel.FullMessage }
        else { $txtDetail.Clear() }
    }.GetNewClosure())

    $mnuCopyMessage.Add_Click({
        if ($txtDetail.Text) {
            [System.Windows.Clipboard]::SetText($txtDetail.Text)
        }
    }.GetNewClosure())

    $mnuCopyRow.Add_Click({
        $sel = $gridEvents.SelectedItem
        if ($sel) {
            $summary = "[{0}] {1} - Source: {2}, ID: {3}`r`n{4}" -f `
                $sel.Level, $sel.Time, $sel.Source, $sel.EventId, $sel.FullMessage
            [System.Windows.Clipboard]::SetText($summary)
        }
    }.GetNewClosure())

    $gridEvents.ItemsSource = $state.AllEvents

    return $view
}
