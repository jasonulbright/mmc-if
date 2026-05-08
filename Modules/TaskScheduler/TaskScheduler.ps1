<#
.SYNOPSIS
    Task Scheduler module for MMC-If (WPF).

.DESCRIPTION
    Browse scheduled tasks via Get-ScheduledTask. Read-only view of triggers,
    actions, conditions, and run history. No enable/disable/run-on-demand
    (those require admin for most tasks; this tool targets non-admins).
#>

function New-TaskSchedulerView {
    param([hashtable]$Context)

    $setStatus = $Context.SetStatus
    $log       = $Context.Log

    $xamlPath = Join-Path $PSScriptRoot 'TaskScheduler.xaml'
    $xamlRaw  = Get-Content -LiteralPath $xamlPath -Raw
    $reader   = New-Object System.Xml.XmlNodeReader ([xml]$xamlRaw)
    $view     = [Windows.Markup.XamlReader]::Load($reader)

    $btnRefresh      = $view.FindName('btnRefresh')
    $txtFilter       = $view.FindName('txtFilter')
    $chkHideMicrosoft = $view.FindName('chkHideMicrosoft')
    $gridTasks       = $view.FindName('gridTasks')
    $txtDetail       = $view.FindName('txtDetail')
    $mnuCopyName     = $view.FindName('mnuCopyName')
    $mnuCopyPath     = $view.FindName('mnuCopyPath')
    $mnuCopyDetail   = $view.FindName('mnuCopyDetail')

    $state = @{
        All      = New-Object System.Collections.ObjectModel.ObservableCollection[object]
    }

    $formatResult = {
        param([int]$Code)
        switch ($Code) {
            0           { 'Success (0)' }
            0x00041301  { 'Task ready (0x00041301)' }
            0x00041303  { 'Task not yet run (0x00041303)' }
            0x00041304  { 'No triggers defined (0x00041304)' }
            0x00041306  { 'Task terminated by user (0x00041306)' }
            267008      { 'Task ready' }
            267009      { 'Task running' }
            267011      { 'Task not yet run' }
            267014      { 'Task terminated by user' }
            default     {
                if ($Code -eq 0) { 'Success (0)' }
                else { '0x{0:X8} ({0})' -f $Code }
            }
        }
    }

    $applyFilter = {
        $ft = $txtFilter.Text
        $hideMs = [bool]$chkHideMicrosoft.IsChecked
        $filtered = New-Object System.Collections.ObjectModel.ObservableCollection[object]
        foreach ($t in $state.All) {
            if ($hideMs -and $t.TaskPath -like '\Microsoft\*') { continue }
            if ($hideMs -and $t.TaskPath -eq '\Microsoft\') { continue }
            if (-not [string]::IsNullOrWhiteSpace($ft)) {
                $needle = $ft.ToLowerInvariant()
                $hay = "$($t.TaskName) $($t.TaskPath) $($t.Author)".ToLowerInvariant()
                if (-not $hay.Contains($needle)) { continue }
            }
            [void]$filtered.Add($t)
        }
        $gridTasks.ItemsSource = $filtered
    }.GetNewClosure()

    $loadTasks = {
        & $setStatus 'Loading scheduled tasks...'
        $view.Cursor = [System.Windows.Input.Cursors]::Wait
        $state.All.Clear()
        $txtDetail.Clear()

        try {
            $tasks = Get-ScheduledTask -ErrorAction Stop | Sort-Object TaskPath, TaskName
            foreach ($t in $tasks) {
                $info = $null
                $lastRun = ''
                $nextRun = ''
                $lastCode = 0
                try {
                    $info = Get-ScheduledTaskInfo -TaskPath $t.TaskPath -TaskName $t.TaskName -ErrorAction Stop
                    if ($info.LastRunTime) { $lastRun = $info.LastRunTime.ToString('yyyy-MM-dd HH:mm') }
                    if ($info.NextRunTime) { $nextRun = $info.NextRunTime.ToString('yyyy-MM-dd HH:mm') }
                    $lastCode = [int]$info.LastTaskResult
                } catch { }

                $state.All.Add([pscustomobject]@{
                    TaskName       = $t.TaskName
                    TaskPath       = $t.TaskPath
                    State          = $t.State.ToString()
                    Author         = if ($t.Author) { [string]$t.Author } else { '' }
                    Description    = if ($t.Description) { [string]$t.Description } else { '' }
                    LastRunTime    = $lastRun
                    NextRunTime    = $nextRun
                    LastResultText = & $formatResult $lastCode
                    LastResultCode = $lastCode
                    Task           = $t
                })
            }
            & $setStatus "$($tasks.Count) scheduled tasks"
            & $log "Loaded $($tasks.Count) scheduled tasks"
        }
        catch {
            & $setStatus "Error loading tasks: $($_.Exception.Message)"
        }

        & $applyFilter
        $view.Cursor = [System.Windows.Input.Cursors]::Arrow
    }.GetNewClosure()

    $btnRefresh.Add_Click({ & $loadTasks }.GetNewClosure())
    $txtFilter.Add_TextChanged({ & $applyFilter }.GetNewClosure())
    $chkHideMicrosoft.Add_Checked({ & $applyFilter }.GetNewClosure())
    $chkHideMicrosoft.Add_Unchecked({ & $applyFilter }.GetNewClosure())

    $gridTasks.Add_SelectionChanged({
        $sel = $gridTasks.SelectedItem
        if (-not $sel) { $txtDetail.Clear(); return }
        $t = $sel.Task
        $lines = @(
            "Task: $($sel.TaskPath)$($sel.TaskName)",
            "State: $($sel.State)  |  Author: $($sel.Author)",
            "Last Run: $($sel.LastRunTime)  |  Next Run: $($sel.NextRunTime)",
            "Last Result: $($sel.LastResultText)",
            '',
            'Description:',
            $sel.Description,
            ''
        )

        if ($t.Triggers -and $t.Triggers.Count -gt 0) {
            $lines += "Triggers ($($t.Triggers.Count)):"
            foreach ($trg in $t.Triggers) {
                $kind = $trg.CimClass.CimClassName -replace '^MSFT_Task', ''
                $enabled = if ($trg.Enabled) { 'enabled' } else { 'disabled' }
                $lines += "  [$kind] $enabled"
                if ($trg.StartBoundary) { $lines += "    Start: $($trg.StartBoundary)" }
                if ($trg.EndBoundary)   { $lines += "    End:   $($trg.EndBoundary)" }
                if ($trg.ExecutionTimeLimit) { $lines += "    Limit: $($trg.ExecutionTimeLimit)" }
                if ($trg.Repetition -and $trg.Repetition.Interval) {
                    $lines += "    Repeat: every $($trg.Repetition.Interval) for $($trg.Repetition.Duration)"
                }
                if ($trg.DaysOfWeek) { $lines += "    Days: $($trg.DaysOfWeek)" }
            }
            $lines += ''
        }

        if ($t.Actions -and $t.Actions.Count -gt 0) {
            $lines += "Actions ($($t.Actions.Count)):"
            foreach ($a in $t.Actions) {
                if ($a.Execute) {
                    $lines += "  Exec: $($a.Execute) $($a.Arguments)"
                    if ($a.WorkingDirectory) { $lines += "    Working Dir: $($a.WorkingDirectory)" }
                }
                elseif ($a.ClassId) {
                    $lines += "  COM: $($a.ClassId)"
                }
            }
            $lines += ''
        }

        if ($t.Principal) {
            $lines += 'Principal:'
            $lines += "  RunAs: $($t.Principal.UserId)  |  LogonType: $($t.Principal.LogonType)"
            $lines += "  RunLevel: $($t.Principal.RunLevel)"
            $lines += ''
        }

        if ($t.Settings) {
            $lines += 'Settings:'
            $lines += "  Enabled: $($t.Settings.Enabled)  |  Hidden: $($t.Settings.Hidden)"
            $lines += "  StartIfOnBattery: $($t.Settings.DisallowStartIfOnBatteries -eq $false)"
            $lines += "  WakeToRun: $($t.Settings.WakeToRun)"
            if ($t.Settings.ExecutionTimeLimit) { $lines += "  ExecutionTimeLimit: $($t.Settings.ExecutionTimeLimit)" }
        }

        $txtDetail.Text = $lines -join "`r`n"
    }.GetNewClosure())

    $mnuCopyName.Add_Click({
        $s = $gridTasks.SelectedItem
        if ($s) { [System.Windows.Clipboard]::SetText([string]$s.TaskName) }
    }.GetNewClosure())
    $mnuCopyPath.Add_Click({
        $s = $gridTasks.SelectedItem
        if ($s) { [System.Windows.Clipboard]::SetText("$($s.TaskPath)$($s.TaskName)") }
    }.GetNewClosure())
    $mnuCopyDetail.Add_Click({
        if ($txtDetail.Text) { [System.Windows.Clipboard]::SetText($txtDetail.Text) }
    }.GetNewClosure())

    $view.Add_Loaded({ if ($state.All.Count -eq 0) { & $loadTasks } }.GetNewClosure())

    return $view
}
