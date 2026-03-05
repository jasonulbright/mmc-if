<#
.SYNOPSIS
    Event Log Viewer module for MMC-If.

.DESCRIPTION
    Browse and filter Windows event logs using Get-WinEvent.
    Does not require Event Viewer (mmc.exe) or elevation for Application/System logs.
#>

function Initialize-EventLogViewer {
    param([hashtable]$Context)

    $tabPage      = $Context.TabPage
    $theme        = $Context.Theme
    $prefs        = $Context.Prefs
    $newGrid      = $Context.NewGridFunction
    $btnStyle     = $Context.ButtonStyleFunction
    $dblBuffer    = $Context.DoubleBufferFunction
    $statusFunc   = $Context.StatusFunction

    # Theme color shortcuts
    $clrFormBg   = $theme.FormBg
    $clrPanelBg  = $theme.PanelBg
    $clrAccent   = $theme.Accent
    $clrText     = $theme.Text
    $clrGridText = $theme.GridText
    $clrDetailBg = $theme.DetailBg
    $clrSepLine  = $theme.SepLine
    $clrHint     = $theme.Hint
    $clrErrText  = $theme.ErrText
    $clrWarnText = $theme.WarnText
    $clrOkText   = $theme.OkText
    $clrInfoText = $theme.InfoText
    $clrGridAlt  = $theme.GridAlt
    $clrGridLine = $theme.GridLine
    $isDark      = $theme.DarkMode

    # -----------------------------------------------------------------------
    # State
    # -----------------------------------------------------------------------
    $evtEvents = New-Object System.Data.DataTable
    [void]$evtEvents.Columns.Add("Level", [string])
    [void]$evtEvents.Columns.Add("Date/Time", [string])
    [void]$evtEvents.Columns.Add("Source", [string])
    [void]$evtEvents.Columns.Add("Event ID", [string])
    [void]$evtEvents.Columns.Add("Message", [string])
    [void]$evtEvents.Columns.Add("_RawLevel", [int])
    [void]$evtEvents.Columns.Add("_FullMessage", [string])

    # -----------------------------------------------------------------------
    # Filter bar (Dock:Top)
    # -----------------------------------------------------------------------

    $pnlFilter = New-Object System.Windows.Forms.Panel
    $pnlFilter.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlFilter.Height = 40
    $pnlFilter.BackColor = $clrPanelBg
    $pnlFilter.Padding = New-Object System.Windows.Forms.Padding(6, 4, 6, 4)

    # Log selector
    $lblLog = New-Object System.Windows.Forms.Label
    $lblLog.Text = "Log:"; $lblLog.AutoSize = $true
    $lblLog.Location = New-Object System.Drawing.Point(8, 10); $lblLog.ForeColor = $clrText
    $lblLog.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $pnlFilter.Controls.Add($lblLog)

    $cboLog = New-Object System.Windows.Forms.ComboBox
    $cboLog.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cboLog.Location = New-Object System.Drawing.Point(38, 7); $cboLog.Width = 130
    $cboLog.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $cboLog.BackColor = $clrDetailBg; $cboLog.ForeColor = $clrText
    @('Application', 'System', 'Security', 'Setup') | ForEach-Object { $cboLog.Items.Add($_) | Out-Null }
    $cboLog.SelectedIndex = 0
    $pnlFilter.Controls.Add($cboLog)

    # Level filter
    $lblLevel = New-Object System.Windows.Forms.Label
    $lblLevel.Text = "Level:"; $lblLevel.AutoSize = $true
    $lblLevel.Location = New-Object System.Drawing.Point(178, 10); $lblLevel.ForeColor = $clrText
    $lblLevel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $pnlFilter.Controls.Add($lblLevel)

    $cboLevel = New-Object System.Windows.Forms.ComboBox
    $cboLevel.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cboLevel.Location = New-Object System.Drawing.Point(218, 7); $cboLevel.Width = 100
    $cboLevel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $cboLevel.BackColor = $clrDetailBg; $cboLevel.ForeColor = $clrText
    @('All Levels', 'Critical', 'Error', 'Warning', 'Information') | ForEach-Object { $cboLevel.Items.Add($_) | Out-Null }
    $cboLevel.SelectedIndex = 0
    $pnlFilter.Controls.Add($cboLevel)

    # Time range
    $lblTime = New-Object System.Windows.Forms.Label
    $lblTime.Text = "Time:"; $lblTime.AutoSize = $true
    $lblTime.Location = New-Object System.Drawing.Point(328, 10); $lblTime.ForeColor = $clrText
    $lblTime.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $pnlFilter.Controls.Add($lblTime)

    $cboTime = New-Object System.Windows.Forms.ComboBox
    $cboTime.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cboTime.Location = New-Object System.Drawing.Point(366, 7); $cboTime.Width = 110
    $cboTime.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $cboTime.BackColor = $clrDetailBg; $cboTime.ForeColor = $clrText
    @('Last 1 Hour', 'Last 24 Hours', 'Last 7 Days', 'Last 30 Days', 'All Time') | ForEach-Object { $cboTime.Items.Add($_) | Out-Null }
    $cboTime.SelectedIndex = 1
    $pnlFilter.Controls.Add($cboTime)

    # Max events
    $lblMax = New-Object System.Windows.Forms.Label
    $lblMax.Text = "Max:"; $lblMax.AutoSize = $true
    $lblMax.Location = New-Object System.Drawing.Point(486, 10); $lblMax.ForeColor = $clrText
    $lblMax.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $pnlFilter.Controls.Add($lblMax)

    $numMax = New-Object System.Windows.Forms.NumericUpDown
    $numMax.Location = New-Object System.Drawing.Point(520, 7); $numMax.Width = 70
    $numMax.Minimum = 50; $numMax.Maximum = 10000; $numMax.Value = 500; $numMax.Increment = 100
    $numMax.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $numMax.BackColor = $clrDetailBg; $numMax.ForeColor = $clrText
    $pnlFilter.Controls.Add($numMax)

    # Load button
    $btnLoad = New-Object System.Windows.Forms.Button
    $btnLoad.Text = "Load Events"
    $btnLoad.Location = New-Object System.Drawing.Point(600, 5)
    $btnLoad.Size = New-Object System.Drawing.Size(100, 28)
    $btnLoad.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    & $btnStyle $btnLoad $clrAccent
    $pnlFilter.Controls.Add($btnLoad)

    # Text filter
    $txtFilter = New-Object System.Windows.Forms.TextBox
    $txtFilter.Location = New-Object System.Drawing.Point(710, 7); $txtFilter.Width = 180
    $txtFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $txtFilter.BackColor = $clrDetailBg; $txtFilter.ForeColor = $clrText
    $txtFilter.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $txtFilter.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right

    # Watermark via GotFocus/LostFocus
    $txtFilter.Text = "Filter..."
    $txtFilter.ForeColor = $clrHint
    $txtFilter.Tag = "watermark"
    $txtFilter.Add_GotFocus(({
        if ($txtFilter.Tag -eq "watermark") {
            $txtFilter.Text = ""; $txtFilter.ForeColor = $clrText; $txtFilter.Tag = ""
        }
    }).GetNewClosure())
    $txtFilter.Add_LostFocus(({
        if ([string]::IsNullOrWhiteSpace($txtFilter.Text)) {
            $txtFilter.Text = "Filter..."; $txtFilter.ForeColor = $clrHint; $txtFilter.Tag = "watermark"
        }
    }).GetNewClosure())
    $pnlFilter.Controls.Add($txtFilter)

    # Separator
    $pnlFilterSep = New-Object System.Windows.Forms.Panel
    $pnlFilterSep.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlFilterSep.Height = 1; $pnlFilterSep.BackColor = $clrSepLine

    # -----------------------------------------------------------------------
    # SplitContainer (Events grid top, Detail panel bottom)
    # -----------------------------------------------------------------------

    $splitEvt = New-Object System.Windows.Forms.SplitContainer
    $splitEvt.Dock = [System.Windows.Forms.DockStyle]::Fill
    $splitEvt.Orientation = [System.Windows.Forms.Orientation]::Horizontal
    $splitEvt.SplitterWidth = 6
    $splitEvt.BackColor = $clrSepLine
    $splitEvt.Panel1.BackColor = $clrPanelBg
    $splitEvt.Panel2.BackColor = $clrPanelBg
    $splitEvt.Panel1MinSize = 100
    $splitEvt.Panel2MinSize = 80

    # -----------------------------------------------------------------------
    # Events grid (top panel)
    # -----------------------------------------------------------------------

    $gridEvents = & $newGrid
    $gridEvents.AllowUserToOrderColumns = $true

    $colLevel = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colLevel.HeaderText = "Level"; $colLevel.DataPropertyName = "Level"; $colLevel.Width = 80
    $colTime = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colTime.HeaderText = "Date/Time"; $colTime.DataPropertyName = "Date/Time"; $colTime.Width = 150
    $colSource = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colSource.HeaderText = "Source"; $colSource.DataPropertyName = "Source"; $colSource.Width = 180
    $colEvtId = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colEvtId.HeaderText = "Event ID"; $colEvtId.DataPropertyName = "Event ID"; $colEvtId.Width = 70
    $colMsg = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colMsg.HeaderText = "Message"; $colMsg.DataPropertyName = "Message"; $colMsg.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill

    $gridEvents.Columns.Add($colLevel) | Out-Null
    $gridEvents.Columns.Add($colTime) | Out-Null
    $gridEvents.Columns.Add($colSource) | Out-Null
    $gridEvents.Columns.Add($colEvtId) | Out-Null
    $gridEvents.Columns.Add($colMsg) | Out-Null

    $gridEvents.DataSource = $evtEvents
    $splitEvt.Panel1.Controls.Add($gridEvents)

    # Color-code rows by level
    $gridEvents.Add_CellFormatting(({
        param($s, $e)
        if ($e.RowIndex -lt 0 -or $e.ColumnIndex -ne 0) { return }
        $level = $s.Rows[$e.RowIndex].Cells[0].Value
        switch ($level) {
            'Critical' { $e.CellStyle.ForeColor = $clrErrText }
            'Error'    { $e.CellStyle.ForeColor = $clrErrText }
            'Warning'  { $e.CellStyle.ForeColor = $clrWarnText }
        }
    }).GetNewClosure())

    # -----------------------------------------------------------------------
    # Detail panel (bottom)
    # -----------------------------------------------------------------------

    $txtDetail = New-Object System.Windows.Forms.TextBox
    $txtDetail.Dock = [System.Windows.Forms.DockStyle]::Fill
    $txtDetail.Multiline = $true; $txtDetail.ReadOnly = $true
    $txtDetail.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $txtDetail.Font = New-Object System.Drawing.Font("Consolas", 9.5)
    $txtDetail.BackColor = $clrDetailBg; $txtDetail.ForeColor = $clrText
    $txtDetail.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $txtDetail.WordWrap = $true
    $splitEvt.Panel2.Controls.Add($txtDetail)

    # Show full message on row select
    $gridEvents.Add_SelectionChanged(({
        if ($gridEvents.CurrentRow -and $gridEvents.CurrentRow.Index -ge 0) {
            $idx = $gridEvents.CurrentRow.Index
            $view = $evtEvents.DefaultView
            if ($idx -lt $view.Count) {
                $rowView = $view[$idx]
                $txtDetail.Text = [string]$rowView["_FullMessage"]
            }
        }
    }).GetNewClosure())

    # -----------------------------------------------------------------------
    # Load events handler
    # -----------------------------------------------------------------------

    $loadEvents = {
        $logName = $cboLog.SelectedItem.ToString()
        $maxEvents = [int]$numMax.Value

        & $statusFunc "Loading $logName events..."
        $tabPage.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        [System.Windows.Forms.Application]::DoEvents()

        $evtEvents.Clear()

        $filter = @{ LogName = $logName }

        # Level filter
        $levelSel = $cboLevel.SelectedItem.ToString()
        switch ($levelSel) {
            'Critical'    { $filter['Level'] = 1 }
            'Error'       { $filter['Level'] = 2 }
            'Warning'     { $filter['Level'] = 3 }
            'Information' { $filter['Level'] = 4 }
        }

        # Time filter
        $timeSel = $cboTime.SelectedItem.ToString()
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
                    1 { 'Critical' }
                    2 { 'Error' }
                    3 { 'Warning' }
                    4 { 'Information' }
                    0 { 'Information' }
                    5 { 'Verbose' }
                    default { "Level $($evt.Level)" }
                }

                $msgOneLine = if ($evt.Message) { ($evt.Message -split "`n")[0].Trim() } else { '' }
                $fullMsg = if ($evt.Message) { $evt.Message } else { '(No message)' }

                $row = $evtEvents.NewRow()
                $row["Level"] = $levelStr
                $row["Date/Time"] = $evt.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
                $row["Source"] = $evt.ProviderName
                $row["Event ID"] = $evt.Id.ToString()
                $row["Message"] = $msgOneLine
                $row["_RawLevel"] = [int]$evt.Level
                $row["_FullMessage"] = $fullMsg
                $evtEvents.Rows.Add($row)
            }

            & $statusFunc "$($logName): $($events.Count) event(s) loaded"
            Write-Log "Loaded $($events.Count) events from $logName log"

        } catch [Exception] {
            $errMsg = $_.Exception.Message
            if ($errMsg -like '*No events were found*') {
                $timeInfo = $cboTime.SelectedItem.ToString()
                & $statusFunc "$($logName): No events found for '$timeInfo' - try expanding the time range"
            } elseif ($errMsg -like '*access*' -or $errMsg -like '*privilege*') {
                & $statusFunc "$($logName): Access denied - may require elevated permissions"
                $row = $evtEvents.NewRow()
                $row["Level"] = "Error"
                $row["Date/Time"] = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                $row["Source"] = "MMC-If"
                $row["Event ID"] = "0"
                $row["Message"] = "Access denied reading $logName log. Security log typically requires admin or Event Log Readers group membership."
                $row["_RawLevel"] = 2
                $row["_FullMessage"] = $errMsg
                $evtEvents.Rows.Add($row)
            } else {
                & $statusFunc "Error loading events: $errMsg"
                [System.Windows.Forms.MessageBox]::Show("Failed to load events: $errMsg", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
            }
        }

        $tabPage.Cursor = [System.Windows.Forms.Cursors]::Default
    }.GetNewClosure()

    $btnLoad.Add_Click(({ & $loadEvents }).GetNewClosure())

    # -----------------------------------------------------------------------
    # Text filter (live filtering via DataView RowFilter)
    # -----------------------------------------------------------------------

    $txtFilter.Add_TextChanged(({
        if ($txtFilter.Tag -eq "watermark") { return }
        $filterText = $txtFilter.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($filterText)) {
            $evtEvents.DefaultView.RowFilter = ''
        } else {
            $escaped = $filterText.Replace("'", "''").Replace("[", "[[]").Replace("%", "[%]").Replace("*", "[*]")
            $evtEvents.DefaultView.RowFilter = "Source LIKE '%$escaped%' OR Message LIKE '%$escaped%' OR [Event ID] LIKE '%$escaped%'"
        }
    }).GetNewClosure())

    # -----------------------------------------------------------------------
    # Context menu: Copy
    # -----------------------------------------------------------------------

    $ctxEvents = New-Object System.Windows.Forms.ContextMenuStrip
    if ($isDark -and $script:DarkRenderer) { $ctxEvents.Renderer = $script:DarkRenderer }
    $ctxEvents.BackColor = $clrPanelBg; $ctxEvents.ForeColor = $clrText

    $ctxCopyMsg = New-Object System.Windows.Forms.ToolStripMenuItem("Copy Full Message")
    $ctxCopyMsg.ForeColor = $clrText
    $ctxCopyMsg.Add_Click(({
        if ($txtDetail.Text) { [System.Windows.Forms.Clipboard]::SetText($txtDetail.Text) }
    }).GetNewClosure())
    $ctxEvents.Items.Add($ctxCopyMsg) | Out-Null

    $ctxCopyRow = New-Object System.Windows.Forms.ToolStripMenuItem("Copy Row Summary")
    $ctxCopyRow.ForeColor = $clrText
    $ctxCopyRow.Add_Click(({
        if ($gridEvents.CurrentRow) {
            $r = $gridEvents.CurrentRow
            $summary = "[{0}] {1} - Source: {2}, ID: {3}`r`n{4}" -f $r.Cells[0].Value, $r.Cells[1].Value, $r.Cells[2].Value, $r.Cells[3].Value, $txtDetail.Text
            [System.Windows.Forms.Clipboard]::SetText($summary)
        }
    }).GetNewClosure())
    $ctxEvents.Items.Add($ctxCopyRow) | Out-Null

    $gridEvents.ContextMenuStrip = $ctxEvents

    # -----------------------------------------------------------------------
    # Assemble layout
    # -----------------------------------------------------------------------

    $tabPage.Controls.Add($splitEvt)
    $splitEvt.BringToFront()

    $tabPage.Controls.Add($pnlFilterSep)
    $pnlFilterSep.SendToBack()

    $tabPage.Controls.Add($pnlFilter)
    $pnlFilter.SendToBack()

    # Defer SplitterDistance until the control has a valid size
    $evtSplitFlag = @{ Done = $false }
    $splitEvt.Add_SizeChanged(({
        if (-not $evtSplitFlag.Done -and $splitEvt.Height -gt ($splitEvt.Panel1MinSize + $splitEvt.Panel2MinSize + $splitEvt.SplitterWidth)) {
            $evtSplitFlag.Done = $true
            $splitEvt.SplitterDistance = [Math]::Max($splitEvt.Panel1MinSize, [int]($splitEvt.Height * 0.7))
        }
    }).GetNewClosure())
}
