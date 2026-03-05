<#
.SYNOPSIS
    Services Manager module for MMC-If.

.DESCRIPTION
    View and control Windows services using Get-CimInstance Win32_Service.
    Start, stop, restart services without services.msc (mmc.exe).
#>

function Initialize-Services {
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
    $clrGridAlt  = $theme.GridAlt
    $clrGridLine = $theme.GridLine
    $isDark      = $theme.DarkMode

    # -------------------------------------------------------------------
    # State
    # -------------------------------------------------------------------
    $svcTable = New-Object System.Data.DataTable
    [void]$svcTable.Columns.Add("Name", [string])
    [void]$svcTable.Columns.Add("DisplayName", [string])
    [void]$svcTable.Columns.Add("Status", [string])
    [void]$svcTable.Columns.Add("StartupType", [string])
    [void]$svcTable.Columns.Add("Account", [string])
    [void]$svcTable.Columns.Add("Description", [string])
    [void]$svcTable.Columns.Add("_PathName", [string])
    [void]$svcTable.Columns.Add("_ProcessId", [string])
    [void]$svcTable.Columns.Add("_FullDesc", [string])

    # -------------------------------------------------------------------
    # Toolbar (Dock:Top)
    # -------------------------------------------------------------------

    $pnlToolbar = New-Object System.Windows.Forms.Panel
    $pnlToolbar.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlToolbar.Height = 40
    $pnlToolbar.BackColor = $clrPanelBg
    $pnlToolbar.Padding = New-Object System.Windows.Forms.Padding(6, 4, 6, 4)

    # Filter textbox with watermark
    $txtFilter = New-Object System.Windows.Forms.TextBox
    $txtFilter.Location = New-Object System.Drawing.Point(8, 7)
    $txtFilter.Width = 220
    $txtFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $txtFilter.BackColor = $clrDetailBg; $txtFilter.ForeColor = $clrHint
    $txtFilter.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $txtFilter.Text = "Filter services..."
    $txtFilter.Tag = "watermark"
    $txtFilter.Add_GotFocus(({
        if ($txtFilter.Tag -eq "watermark") {
            $txtFilter.Text = ""; $txtFilter.ForeColor = $clrText; $txtFilter.Tag = ""
        }
    }).GetNewClosure())
    $txtFilter.Add_LostFocus(({
        if ([string]::IsNullOrWhiteSpace($txtFilter.Text)) {
            $txtFilter.Text = "Filter services..."; $txtFilter.ForeColor = $clrHint; $txtFilter.Tag = "watermark"
        }
    }).GetNewClosure())
    $pnlToolbar.Controls.Add($txtFilter)

    # Refresh button
    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = "Refresh"
    $btnRefresh.Location = New-Object System.Drawing.Point(240, 5)
    $btnRefresh.Size = New-Object System.Drawing.Size(80, 28)
    $btnRefresh.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    & $btnStyle $btnRefresh $clrAccent
    $pnlToolbar.Controls.Add($btnRefresh)

    # Start button
    $btnStart = New-Object System.Windows.Forms.Button
    $btnStart.Text = "Start"
    $btnStart.Location = New-Object System.Drawing.Point(340, 5)
    $btnStart.Size = New-Object System.Drawing.Size(70, 28)
    $btnStart.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $btnStart.Enabled = $false
    & $btnStyle $btnStart $clrAccent
    $pnlToolbar.Controls.Add($btnStart)

    # Stop button
    $btnStop = New-Object System.Windows.Forms.Button
    $btnStop.Text = "Stop"
    $btnStop.Location = New-Object System.Drawing.Point(416, 5)
    $btnStop.Size = New-Object System.Drawing.Size(70, 28)
    $btnStop.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $btnStop.Enabled = $false
    & $btnStyle $btnStop $clrAccent
    $pnlToolbar.Controls.Add($btnStop)

    # Restart button
    $btnRestart = New-Object System.Windows.Forms.Button
    $btnRestart.Text = "Restart"
    $btnRestart.Location = New-Object System.Drawing.Point(492, 5)
    $btnRestart.Size = New-Object System.Drawing.Size(80, 28)
    $btnRestart.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $btnRestart.Enabled = $false
    & $btnStyle $btnRestart $clrAccent
    $pnlToolbar.Controls.Add($btnRestart)

    # Separator
    $pnlSep = New-Object System.Windows.Forms.Panel
    $pnlSep.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlSep.Height = 1; $pnlSep.BackColor = $clrSepLine

    # -------------------------------------------------------------------
    # SplitContainer (grid top, detail bottom)
    # -------------------------------------------------------------------

    $splitSvc = New-Object System.Windows.Forms.SplitContainer
    $splitSvc.Dock = [System.Windows.Forms.DockStyle]::Fill
    $splitSvc.Orientation = [System.Windows.Forms.Orientation]::Horizontal
    $splitSvc.SplitterWidth = 6
    $splitSvc.BackColor = $clrSepLine
    $splitSvc.Panel1.BackColor = $clrPanelBg
    $splitSvc.Panel2.BackColor = $clrPanelBg
    $splitSvc.Panel1MinSize = 100
    $splitSvc.Panel2MinSize = 80

    # -------------------------------------------------------------------
    # Services grid (top panel)
    # -------------------------------------------------------------------

    $gridSvc = & $newGrid
    $gridSvc.AllowUserToOrderColumns = $true

    $colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colName.HeaderText = "Name"; $colName.DataPropertyName = "Name"; $colName.Width = 150
    $colDisplay = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colDisplay.HeaderText = "Display Name"; $colDisplay.DataPropertyName = "DisplayName"; $colDisplay.Width = 220
    $colStatus = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colStatus.HeaderText = "Status"; $colStatus.DataPropertyName = "Status"; $colStatus.Width = 80
    $colStartup = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colStartup.HeaderText = "Startup Type"; $colStartup.DataPropertyName = "StartupType"; $colStartup.Width = 100
    $colAccount = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colAccount.HeaderText = "Account"; $colAccount.DataPropertyName = "Account"; $colAccount.Width = 150
    $colDesc = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colDesc.HeaderText = "Description"; $colDesc.DataPropertyName = "Description"
    $colDesc.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill

    $gridSvc.Columns.Add($colName) | Out-Null
    $gridSvc.Columns.Add($colDisplay) | Out-Null
    $gridSvc.Columns.Add($colStatus) | Out-Null
    $gridSvc.Columns.Add($colStartup) | Out-Null
    $gridSvc.Columns.Add($colAccount) | Out-Null
    $gridSvc.Columns.Add($colDesc) | Out-Null

    $gridSvc.DataSource = $svcTable
    $splitSvc.Panel1.Controls.Add($gridSvc)

    # Color-code Status column
    $gridSvc.Add_CellFormatting(({
        param($s, $e)
        if ($e.RowIndex -lt 0 -or $e.ColumnIndex -ne 2) { return }
        $status = $s.Rows[$e.RowIndex].Cells[2].Value
        switch ($status) {
            'Running' { $e.CellStyle.ForeColor = $clrOkText }
            'Stopped' { $e.CellStyle.ForeColor = $clrErrText }
            'Paused'  { $e.CellStyle.ForeColor = $clrWarnText }
        }
    }).GetNewClosure())

    # -------------------------------------------------------------------
    # Detail panel (bottom)
    # -------------------------------------------------------------------

    $txtDetail = New-Object System.Windows.Forms.TextBox
    $txtDetail.Dock = [System.Windows.Forms.DockStyle]::Fill
    $txtDetail.Multiline = $true; $txtDetail.ReadOnly = $true
    $txtDetail.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $txtDetail.Font = New-Object System.Drawing.Font("Consolas", 9.5)
    $txtDetail.BackColor = $clrDetailBg; $txtDetail.ForeColor = $clrText
    $txtDetail.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $txtDetail.WordWrap = $true
    $splitSvc.Panel2.Controls.Add($txtDetail)

    # -------------------------------------------------------------------
    # Load services
    # -------------------------------------------------------------------

    $loadServices = {
        & $statusFunc "Loading services..."
        $tabPage.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        [System.Windows.Forms.Application]::DoEvents()

        $svcTable.Clear()

        try {
            $services = Get-CimInstance Win32_Service -ErrorAction Stop | Sort-Object DisplayName

            foreach ($svc in $services) {
                $row = $svcTable.NewRow()
                $row["Name"] = $svc.Name
                $row["DisplayName"] = if ($svc.DisplayName) { $svc.DisplayName } else { $svc.Name }
                $row["Status"] = $svc.State
                $row["StartupType"] = $svc.StartMode
                $row["Account"] = if ($svc.StartName) { $svc.StartName } else { '' }
                $desc = if ($svc.Description) { $svc.Description } else { '' }
                $row["Description"] = if ($desc.Length -gt 100) { $desc.Substring(0, 97) + '...' } else { $desc }
                $row["_PathName"] = if ($svc.PathName) { $svc.PathName } else { '' }
                $row["_ProcessId"] = if ($svc.ProcessId) { $svc.ProcessId.ToString() } else { '' }
                $row["_FullDesc"] = $desc
                $svcTable.Rows.Add($row)
            }

            & $statusFunc "$($services.Count) services loaded"
            Write-Log "Loaded $($services.Count) services"
        } catch {
            & $statusFunc "Error loading services: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to load services: $($_.Exception.Message)",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        }

        $tabPage.Cursor = [System.Windows.Forms.Cursors]::Default
    }.GetNewClosure()

    $btnRefresh.Add_Click(({ & $loadServices }).GetNewClosure())

    # -------------------------------------------------------------------
    # Detail panel on row select
    # -------------------------------------------------------------------

    $gridSvc.Add_SelectionChanged(({
        if ($gridSvc.CurrentRow -and $gridSvc.CurrentRow.Index -ge 0) {
            $view = $svcTable.DefaultView
            $idx = $gridSvc.CurrentRow.Index
            if ($idx -lt $view.Count) {
                $rowView = $view[$idx]
                $name = [string]$rowView["Name"]
                $lines = @(
                    "Service: $name",
                    "Display Name: $([string]$rowView['DisplayName'])",
                    "Status: $([string]$rowView['Status'])  |  Startup: $([string]$rowView['StartupType'])  |  Account: $([string]$rowView['Account'])",
                    "Path: $([string]$rowView['_PathName'])",
                    "PID: $([string]$rowView['_ProcessId'])",
                    "",
                    "Description:",
                    [string]$rowView['_FullDesc'],
                    ""
                )

                # Dependencies
                try {
                    $svcObj = Get-Service -Name $name -ErrorAction Stop
                    $deps = @($svcObj.ServicesDependedOn | ForEach-Object { "  $($_.Name) ($($_.DisplayName))" })
                    $dependents = @($svcObj.DependentServices | ForEach-Object { "  $($_.Name) ($($_.DisplayName))" })
                    $lines += "Depends On:"
                    if ($deps.Count -gt 0) { $lines += $deps } else { $lines += "  (none)" }
                    $lines += ""
                    $lines += "Dependent Services:"
                    if ($dependents.Count -gt 0) { $lines += $dependents } else { $lines += "  (none)" }
                } catch {
                    $lines += "Dependencies: (could not query)"
                }

                $txtDetail.Text = $lines -join "`r`n"

                # Update button states
                $status = [string]$rowView['Status']
                $btnStart.Enabled = ($status -eq 'Stopped')
                $btnStop.Enabled = ($status -eq 'Running')
                $btnRestart.Enabled = ($status -eq 'Running')
            }
        } else {
            $btnStart.Enabled = $false
            $btnStop.Enabled = $false
            $btnRestart.Enabled = $false
        }
    }).GetNewClosure())

    # -------------------------------------------------------------------
    # Service control actions
    # -------------------------------------------------------------------

    $btnStart.Add_Click(({
        if (-not $gridSvc.CurrentRow -or $gridSvc.CurrentRow.Index -lt 0) { return }
        $view = $svcTable.DefaultView
        $idx = $gridSvc.CurrentRow.Index
        if ($idx -ge $view.Count) { return }
        $name = [string]$view[$idx]["Name"]

        $tabPage.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            Start-Service -Name $name -ErrorAction Stop
            & $statusFunc "Started service: $name"
            Write-Log "Started service: $name"
            & $loadServices
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to start service '$name': $($_.Exception.Message)",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
            & $statusFunc "Failed to start $name"
        }
        $tabPage.Cursor = [System.Windows.Forms.Cursors]::Default
    }).GetNewClosure())

    $btnStop.Add_Click(({
        if (-not $gridSvc.CurrentRow -or $gridSvc.CurrentRow.Index -lt 0) { return }
        $view = $svcTable.DefaultView
        $idx = $gridSvc.CurrentRow.Index
        if ($idx -ge $view.Count) { return }
        $name = [string]$view[$idx]["Name"]
        $displayName = [string]$view[$idx]["DisplayName"]

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Stop service '$displayName' ($name)?",
            "Confirm",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $tabPage.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            Stop-Service -Name $name -Force -ErrorAction Stop
            & $statusFunc "Stopped service: $name"
            Write-Log "Stopped service: $name"
            & $loadServices
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to stop service '$name': $($_.Exception.Message)",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
            & $statusFunc "Failed to stop $name"
        }
        $tabPage.Cursor = [System.Windows.Forms.Cursors]::Default
    }).GetNewClosure())

    $btnRestart.Add_Click(({
        if (-not $gridSvc.CurrentRow -or $gridSvc.CurrentRow.Index -lt 0) { return }
        $view = $svcTable.DefaultView
        $idx = $gridSvc.CurrentRow.Index
        if ($idx -ge $view.Count) { return }
        $name = [string]$view[$idx]["Name"]
        $displayName = [string]$view[$idx]["DisplayName"]

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Restart service '$displayName' ($name)?",
            "Confirm",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $tabPage.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            Restart-Service -Name $name -Force -ErrorAction Stop
            & $statusFunc "Restarted service: $name"
            Write-Log "Restarted service: $name"
            & $loadServices
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to restart service '$name': $($_.Exception.Message)",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
            & $statusFunc "Failed to restart $name"
        }
        $tabPage.Cursor = [System.Windows.Forms.Cursors]::Default
    }).GetNewClosure())

    # -------------------------------------------------------------------
    # Text filter (DataView RowFilter)
    # -------------------------------------------------------------------

    $txtFilter.Add_TextChanged(({
        if ($txtFilter.Tag -eq "watermark") { return }
        $filterText = $txtFilter.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($filterText)) {
            $svcTable.DefaultView.RowFilter = ''
        } else {
            $escaped = $filterText.Replace("'", "''").Replace("[", "[[]").Replace("%", "[%]").Replace("*", "[*]")
            $svcTable.DefaultView.RowFilter = "Name LIKE '%$escaped%' OR DisplayName LIKE '%$escaped%' OR Description LIKE '%$escaped%'"
        }
    }).GetNewClosure())

    # -------------------------------------------------------------------
    # Context menu
    # -------------------------------------------------------------------

    $ctxSvc = New-Object System.Windows.Forms.ContextMenuStrip
    if ($isDark -and $script:DarkRenderer) { $ctxSvc.Renderer = $script:DarkRenderer }
    $ctxSvc.BackColor = $clrPanelBg; $ctxSvc.ForeColor = $clrText

    $ctxCopyName = New-Object System.Windows.Forms.ToolStripMenuItem("Copy Service Name")
    $ctxCopyName.ForeColor = $clrText
    $ctxCopyName.Add_Click(({
        if ($gridSvc.CurrentRow -and $gridSvc.CurrentRow.Index -ge 0) {
            $view = $svcTable.DefaultView
            $idx = $gridSvc.CurrentRow.Index
            if ($idx -lt $view.Count) {
                [System.Windows.Forms.Clipboard]::SetText([string]$view[$idx]["Name"])
            }
        }
    }).GetNewClosure())
    $ctxSvc.Items.Add($ctxCopyName) | Out-Null

    $ctxCopyDisplay = New-Object System.Windows.Forms.ToolStripMenuItem("Copy Display Name")
    $ctxCopyDisplay.ForeColor = $clrText
    $ctxCopyDisplay.Add_Click(({
        if ($gridSvc.CurrentRow -and $gridSvc.CurrentRow.Index -ge 0) {
            $view = $svcTable.DefaultView
            $idx = $gridSvc.CurrentRow.Index
            if ($idx -lt $view.Count) {
                [System.Windows.Forms.Clipboard]::SetText([string]$view[$idx]["DisplayName"])
            }
        }
    }).GetNewClosure())
    $ctxSvc.Items.Add($ctxCopyDisplay) | Out-Null

    $ctxCopyPath = New-Object System.Windows.Forms.ToolStripMenuItem("Copy Path")
    $ctxCopyPath.ForeColor = $clrText
    $ctxCopyPath.Add_Click(({
        if ($gridSvc.CurrentRow -and $gridSvc.CurrentRow.Index -ge 0) {
            $view = $svcTable.DefaultView
            $idx = $gridSvc.CurrentRow.Index
            if ($idx -lt $view.Count) {
                $path = [string]$view[$idx]["_PathName"]
                if ($path) { [System.Windows.Forms.Clipboard]::SetText($path) }
            }
        }
    }).GetNewClosure())
    $ctxSvc.Items.Add($ctxCopyPath) | Out-Null

    $ctxSvc.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null

    $ctxExportCsv = New-Object System.Windows.Forms.ToolStripMenuItem("Export to CSV...")
    $ctxExportCsv.ForeColor = $clrText
    $ctxExportCsv.Add_Click(({
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV Files (*.csv)|*.csv"
        $sfd.DefaultExt = "csv"
        $sfd.FileName = "Services-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Export-MmcIfCsv -DataTable $svcTable -OutputPath $sfd.FileName
            & $statusFunc "Exported to $($sfd.FileName)"
        }
        $sfd.Dispose()
    }).GetNewClosure())
    $ctxSvc.Items.Add($ctxExportCsv) | Out-Null

    $ctxExportHtml = New-Object System.Windows.Forms.ToolStripMenuItem("Export to HTML...")
    $ctxExportHtml.ForeColor = $clrText
    $ctxExportHtml.Add_Click(({
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "HTML Files (*.html)|*.html"
        $sfd.DefaultExt = "html"
        $sfd.FileName = "Services-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Export-MmcIfHtml -DataTable $svcTable -OutputPath $sfd.FileName -ReportTitle "Services Report"
            & $statusFunc "Exported to $($sfd.FileName)"
        }
        $sfd.Dispose()
    }).GetNewClosure())
    $ctxSvc.Items.Add($ctxExportHtml) | Out-Null

    $gridSvc.ContextMenuStrip = $ctxSvc

    # -------------------------------------------------------------------
    # Assemble layout
    # -------------------------------------------------------------------

    $tabPage.Controls.Add($splitSvc)
    $splitSvc.BringToFront()

    $tabPage.Controls.Add($pnlSep)
    $pnlSep.SendToBack()

    $tabPage.Controls.Add($pnlToolbar)
    $pnlToolbar.SendToBack()

    # Defer SplitterDistance
    $svcSplitFlag = @{ Done = $false }
    $splitSvc.Add_SizeChanged(({
        if (-not $svcSplitFlag.Done -and $splitSvc.Height -gt ($splitSvc.Panel1MinSize + $splitSvc.Panel2MinSize + $splitSvc.SplitterWidth)) {
            $svcSplitFlag.Done = $true
            $splitSvc.SplitterDistance = [Math]::Max($splitSvc.Panel1MinSize, [int]($splitSvc.Height * 0.65))
        }
    }).GetNewClosure())

    # Auto-load on init
    & $loadServices
}
