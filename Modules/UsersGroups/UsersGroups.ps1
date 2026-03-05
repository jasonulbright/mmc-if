<#
.SYNOPSIS
    Local Users & Groups module for MMC-If.

.DESCRIPTION
    View local user accounts and group memberships using Get-LocalUser/Get-LocalGroup.
    Check locked/disabled status without lusrmgr.msc (mmc.exe).
#>

function Initialize-UsersGroups {
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
    $ugState = @{ ViewMode = 'Users' }

    $userTable = New-Object System.Data.DataTable
    [void]$userTable.Columns.Add("Name", [string])
    [void]$userTable.Columns.Add("FullName", [string])
    [void]$userTable.Columns.Add("Enabled", [string])
    [void]$userTable.Columns.Add("PasswordLastSet", [string])
    [void]$userTable.Columns.Add("LastLogon", [string])
    [void]$userTable.Columns.Add("Description", [string])
    [void]$userTable.Columns.Add("_SID", [string])
    [void]$userTable.Columns.Add("_Source", [string])

    $groupTable = New-Object System.Data.DataTable
    [void]$groupTable.Columns.Add("Name", [string])
    [void]$groupTable.Columns.Add("Description", [string])
    [void]$groupTable.Columns.Add("Members", [string])
    [void]$groupTable.Columns.Add("_SID", [string])

    # -------------------------------------------------------------------
    # Toolbar (Dock:Top)
    # -------------------------------------------------------------------

    $pnlToolbar = New-Object System.Windows.Forms.Panel
    $pnlToolbar.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlToolbar.Height = 40
    $pnlToolbar.BackColor = $clrPanelBg
    $pnlToolbar.Padding = New-Object System.Windows.Forms.Padding(6, 4, 6, 4)

    # View toggle
    $cboView = New-Object System.Windows.Forms.ComboBox
    $cboView.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cboView.Location = New-Object System.Drawing.Point(8, 7)
    $cboView.Width = 100
    $cboView.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $cboView.BackColor = $clrDetailBg; $cboView.ForeColor = $clrText
    @('Users', 'Groups') | ForEach-Object { $cboView.Items.Add($_) | Out-Null }
    $cboView.SelectedIndex = 0
    $pnlToolbar.Controls.Add($cboView)

    # Filter textbox with watermark
    $txtFilter = New-Object System.Windows.Forms.TextBox
    $txtFilter.Location = New-Object System.Drawing.Point(118, 7)
    $txtFilter.Width = 200
    $txtFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $txtFilter.BackColor = $clrDetailBg; $txtFilter.ForeColor = $clrHint
    $txtFilter.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $txtFilter.Text = "Filter..."
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
    $pnlToolbar.Controls.Add($txtFilter)

    # Refresh button
    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = "Refresh"
    $btnRefresh.Location = New-Object System.Drawing.Point(328, 5)
    $btnRefresh.Size = New-Object System.Drawing.Size(80, 28)
    $btnRefresh.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    & $btnStyle $btnRefresh $clrAccent
    $pnlToolbar.Controls.Add($btnRefresh)

    # Separator
    $pnlSep = New-Object System.Windows.Forms.Panel
    $pnlSep.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlSep.Height = 1; $pnlSep.BackColor = $clrSepLine

    # -------------------------------------------------------------------
    # SplitContainer (grid top, detail bottom)
    # -------------------------------------------------------------------

    $splitUG = New-Object System.Windows.Forms.SplitContainer
    $splitUG.Dock = [System.Windows.Forms.DockStyle]::Fill
    $splitUG.Orientation = [System.Windows.Forms.Orientation]::Horizontal
    $splitUG.SplitterWidth = 6
    $splitUG.BackColor = $clrSepLine
    $splitUG.Panel1.BackColor = $clrPanelBg
    $splitUG.Panel2.BackColor = $clrPanelBg
    $splitUG.Panel1MinSize = 100
    $splitUG.Panel2MinSize = 80

    # -------------------------------------------------------------------
    # Grid (top panel)
    # -------------------------------------------------------------------

    $gridUG = & $newGrid
    $gridUG.AllowUserToOrderColumns = $true
    $splitUG.Panel1.Controls.Add($gridUG)

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
    $splitUG.Panel2.Controls.Add($txtDetail)

    # -------------------------------------------------------------------
    # Column definitions
    # -------------------------------------------------------------------

    $setupUserColumns = {
        $gridUG.DataSource = $null
        $gridUG.Columns.Clear()

        $c1 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $c1.HeaderText = "Name"; $c1.DataPropertyName = "Name"; $c1.Width = 130
        $c2 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $c2.HeaderText = "Full Name"; $c2.DataPropertyName = "FullName"; $c2.Width = 160
        $c3 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $c3.HeaderText = "Enabled"; $c3.DataPropertyName = "Enabled"; $c3.Width = 65
        $c4 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $c4.HeaderText = "Password Last Set"; $c4.DataPropertyName = "PasswordLastSet"; $c4.Width = 140
        $c5 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $c5.HeaderText = "Last Logon"; $c5.DataPropertyName = "LastLogon"; $c5.Width = 140
        $c6 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $c6.HeaderText = "Description"; $c6.DataPropertyName = "Description"
        $c6.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill

        $gridUG.Columns.Add($c1) | Out-Null
        $gridUG.Columns.Add($c2) | Out-Null
        $gridUG.Columns.Add($c3) | Out-Null
        $gridUG.Columns.Add($c4) | Out-Null
        $gridUG.Columns.Add($c5) | Out-Null
        $gridUG.Columns.Add($c6) | Out-Null

        $gridUG.DataSource = $userTable
    }.GetNewClosure()

    $setupGroupColumns = {
        $gridUG.DataSource = $null
        $gridUG.Columns.Clear()

        $c1 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $c1.HeaderText = "Name"; $c1.DataPropertyName = "Name"; $c1.Width = 200
        $c2 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $c2.HeaderText = "Description"; $c2.DataPropertyName = "Description"
        $c2.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
        $c3 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $c3.HeaderText = "Members"; $c3.DataPropertyName = "Members"; $c3.Width = 80

        $gridUG.Columns.Add($c1) | Out-Null
        $gridUG.Columns.Add($c2) | Out-Null
        $gridUG.Columns.Add($c3) | Out-Null

        $gridUG.DataSource = $groupTable
    }.GetNewClosure()

    # -------------------------------------------------------------------
    # Load functions
    # -------------------------------------------------------------------

    $loadUsers = {
        & $statusFunc "Loading local users..."
        $tabPage.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        [System.Windows.Forms.Application]::DoEvents()

        $userTable.Clear()

        try {
            $users = Get-LocalUser -ErrorAction Stop | Sort-Object Name

            foreach ($u in $users) {
                $row = $userTable.NewRow()
                $row["Name"] = $u.Name
                $row["FullName"] = if ($u.FullName) { $u.FullName } else { '' }
                $row["Enabled"] = if ($u.Enabled) { "Yes" } else { "No" }
                $row["PasswordLastSet"] = if ($u.PasswordLastSet) { $u.PasswordLastSet.ToString("yyyy-MM-dd HH:mm") } else { "(never)" }
                $row["LastLogon"] = if ($u.LastLogon) { $u.LastLogon.ToString("yyyy-MM-dd HH:mm") } else { "(never)" }
                $row["Description"] = if ($u.Description) { $u.Description } else { '' }
                $row["_SID"] = $u.SID.ToString()
                $row["_Source"] = if ($u.PrincipalSource) { $u.PrincipalSource.ToString() } else { 'Local' }
                $userTable.Rows.Add($row)
            }

            & $statusFunc "$($users.Count) local users"
            Write-Log "Loaded $($users.Count) local users"
        } catch {
            & $statusFunc "Error loading users: $($_.Exception.Message)"
        }

        $tabPage.Cursor = [System.Windows.Forms.Cursors]::Default
    }.GetNewClosure()

    $loadGroups = {
        & $statusFunc "Loading local groups..."
        $tabPage.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        [System.Windows.Forms.Application]::DoEvents()

        $groupTable.Clear()

        try {
            $groups = Get-LocalGroup -ErrorAction Stop | Sort-Object Name

            foreach ($g in $groups) {
                $row = $groupTable.NewRow()
                $row["Name"] = $g.Name
                $row["Description"] = if ($g.Description) { $g.Description } else { '' }

                try {
                    $members = @(Get-LocalGroupMember -Group $g.Name -ErrorAction Stop)
                    $row["Members"] = $members.Count.ToString()
                } catch {
                    $row["Members"] = "(error)"
                }

                $row["_SID"] = $g.SID.ToString()
                $groupTable.Rows.Add($row)
            }

            & $statusFunc "$($groups.Count) local groups"
            Write-Log "Loaded $($groups.Count) local groups"
        } catch {
            & $statusFunc "Error loading groups: $($_.Exception.Message)"
        }

        $tabPage.Cursor = [System.Windows.Forms.Cursors]::Default
    }.GetNewClosure()

    # -------------------------------------------------------------------
    # Color-code rows
    # -------------------------------------------------------------------

    $gridUG.Add_CellFormatting(({
        param($s, $e)
        if ($e.RowIndex -lt 0) { return }
        if ($ugState.ViewMode -eq 'Users') {
            # Dim disabled users
            $enabledIdx = 2
            if ($enabledIdx -lt $s.Columns.Count) {
                $enabledVal = $s.Rows[$e.RowIndex].Cells[$enabledIdx].Value
                if ($enabledVal -eq 'No') {
                    $e.CellStyle.ForeColor = $clrHint
                }
            }
        }
    }).GetNewClosure())

    # -------------------------------------------------------------------
    # Detail on row select
    # -------------------------------------------------------------------

    $gridUG.Add_SelectionChanged(({
        if (-not $gridUG.CurrentRow -or $gridUG.CurrentRow.Index -lt 0) { return }

        if ($ugState.ViewMode -eq 'Users') {
            $view = $userTable.DefaultView
            $idx = $gridUG.CurrentRow.Index
            if ($idx -ge $view.Count) { return }
            $r = $view[$idx]
            $name = [string]$r["Name"]

            $lines = @(
                "User: $name",
                "Full Name: $([string]$r['FullName'])",
                "SID: $([string]$r['_SID'])",
                "Enabled: $([string]$r['Enabled'])  |  Source: $([string]$r['_Source'])",
                "Password Last Set: $([string]$r['PasswordLastSet'])",
                "Last Logon: $([string]$r['LastLogon'])",
                "",
                "Description:",
                [string]$r['Description'],
                ""
            )

            # Get group memberships via net user
            try {
                $output = & net user $name 2>&1
                $inGroups = $false
                $groupLines = @()
                foreach ($line in $output) {
                    $lineStr = [string]$line
                    if ($lineStr -match 'Local Group Memberships') {
                        $inGroups = $true
                        $groupPart = ($lineStr -split '\*' | Select-Object -Skip 1) | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                        $groupLines += $groupPart
                    } elseif ($lineStr -match 'Global Group memberships') {
                        $inGroups = $false
                    } elseif ($inGroups -and $lineStr.Trim().StartsWith('*')) {
                        $groupPart = ($lineStr -split '\*' | Select-Object -Skip 1) | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                        $groupLines += $groupPart
                    }
                }
                $lines += "Group Memberships:"
                if ($groupLines.Count -gt 0) {
                    foreach ($grp in $groupLines) { $lines += "  $grp" }
                } else {
                    $lines += "  (none)"
                }
            } catch {
                $lines += "Group Memberships: (could not query)"
            }

            $txtDetail.Text = $lines -join "`r`n"
        } else {
            $view = $groupTable.DefaultView
            $idx = $gridUG.CurrentRow.Index
            if ($idx -ge $view.Count) { return }
            $r = $view[$idx]
            $name = [string]$r["Name"]

            $lines = @(
                "Group: $name",
                "SID: $([string]$r['_SID'])",
                "Description: $([string]$r['Description'])",
                ""
            )

            # Get members
            try {
                $members = @(Get-LocalGroupMember -Group $name -ErrorAction Stop)
                $lines += "Members ($($members.Count)):"
                if ($members.Count -gt 0) {
                    foreach ($m in $members) {
                        $mName = $m.Name
                        $mClass = $m.ObjectClass
                        $mSource = if ($m.PrincipalSource) { $m.PrincipalSource.ToString() } else { '' }
                        $lines += "  $mName  ($mClass, $mSource)"
                    }
                } else {
                    $lines += "  (empty)"
                }
            } catch {
                $lines += "Members: Error enumerating - $($_.Exception.Message)"
                $lines += "  (This can happen with orphaned SIDs in the group)"
            }

            $txtDetail.Text = $lines -join "`r`n"
        }
    }).GetNewClosure())

    # -------------------------------------------------------------------
    # View switching
    # -------------------------------------------------------------------

    $switchView = {
        param([string]$Mode)
        $ugState.ViewMode = $Mode
        $txtDetail.Text = ''

        # Reset filter watermark
        $txtFilter.Text = "Filter..."
        $txtFilter.ForeColor = $clrHint
        $txtFilter.Tag = "watermark"

        if ($Mode -eq 'Users') {
            & $setupUserColumns
            & $loadUsers
        } else {
            & $setupGroupColumns
            & $loadGroups
        }
    }.GetNewClosure()

    $cboView.Add_SelectedIndexChanged(({
        $mode = $cboView.SelectedItem.ToString()
        & $switchView $mode
    }).GetNewClosure())

    $btnRefresh.Add_Click(({
        $mode = $cboView.SelectedItem.ToString()
        & $switchView $mode
    }).GetNewClosure())

    # -------------------------------------------------------------------
    # Text filter (DataView RowFilter)
    # -------------------------------------------------------------------

    $txtFilter.Add_TextChanged(({
        if ($txtFilter.Tag -eq "watermark") { return }
        $filterText = $txtFilter.Text.Trim()

        if ($ugState.ViewMode -eq 'Users') {
            if ([string]::IsNullOrWhiteSpace($filterText)) {
                $userTable.DefaultView.RowFilter = ''
            } else {
                $escaped = $filterText.Replace("'", "''").Replace("[", "[[]").Replace("%", "[%]").Replace("*", "[*]")
                $userTable.DefaultView.RowFilter = "Name LIKE '%$escaped%' OR FullName LIKE '%$escaped%' OR Description LIKE '%$escaped%'"
            }
        } else {
            if ([string]::IsNullOrWhiteSpace($filterText)) {
                $groupTable.DefaultView.RowFilter = ''
            } else {
                $escaped = $filterText.Replace("'", "''").Replace("[", "[[]").Replace("%", "[%]").Replace("*", "[*]")
                $groupTable.DefaultView.RowFilter = "Name LIKE '%$escaped%' OR Description LIKE '%$escaped%'"
            }
        }
    }).GetNewClosure())

    # -------------------------------------------------------------------
    # Context menu
    # -------------------------------------------------------------------

    $ctxUG = New-Object System.Windows.Forms.ContextMenuStrip
    if ($isDark -and $script:DarkRenderer) { $ctxUG.Renderer = $script:DarkRenderer }
    $ctxUG.BackColor = $clrPanelBg; $ctxUG.ForeColor = $clrText

    $ctxCopyName = New-Object System.Windows.Forms.ToolStripMenuItem("Copy Name")
    $ctxCopyName.ForeColor = $clrText
    $ctxCopyName.Add_Click(({
        if ($gridUG.CurrentRow -and $gridUG.CurrentRow.Index -ge 0) {
            $table = if ($ugState.ViewMode -eq 'Users') { $userTable } else { $groupTable }
            $view = $table.DefaultView
            $idx = $gridUG.CurrentRow.Index
            if ($idx -lt $view.Count) {
                [System.Windows.Forms.Clipboard]::SetText([string]$view[$idx]["Name"])
            }
        }
    }).GetNewClosure())
    $ctxUG.Items.Add($ctxCopyName) | Out-Null

    $ctxCopySID = New-Object System.Windows.Forms.ToolStripMenuItem("Copy SID")
    $ctxCopySID.ForeColor = $clrText
    $ctxCopySID.Add_Click(({
        if ($gridUG.CurrentRow -and $gridUG.CurrentRow.Index -ge 0) {
            $table = if ($ugState.ViewMode -eq 'Users') { $userTable } else { $groupTable }
            $view = $table.DefaultView
            $idx = $gridUG.CurrentRow.Index
            if ($idx -lt $view.Count) {
                [System.Windows.Forms.Clipboard]::SetText([string]$view[$idx]["_SID"])
            }
        }
    }).GetNewClosure())
    $ctxUG.Items.Add($ctxCopySID) | Out-Null

    $ctxCopyDetail = New-Object System.Windows.Forms.ToolStripMenuItem("Copy Details")
    $ctxCopyDetail.ForeColor = $clrText
    $ctxCopyDetail.Add_Click(({
        if ($txtDetail.Text) { [System.Windows.Forms.Clipboard]::SetText($txtDetail.Text) }
    }).GetNewClosure())
    $ctxUG.Items.Add($ctxCopyDetail) | Out-Null

    $ctxUG.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null

    $ctxExportCsv = New-Object System.Windows.Forms.ToolStripMenuItem("Export to CSV...")
    $ctxExportCsv.ForeColor = $clrText
    $ctxExportCsv.Add_Click(({
        $table = if ($ugState.ViewMode -eq 'Users') { $userTable } else { $groupTable }
        $label = $ugState.ViewMode
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV Files (*.csv)|*.csv"
        $sfd.DefaultExt = "csv"
        $sfd.FileName = "Local$label-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Export-MmcIfCsv -DataTable $table -OutputPath $sfd.FileName
            & $statusFunc "Exported to $($sfd.FileName)"
        }
        $sfd.Dispose()
    }).GetNewClosure())
    $ctxUG.Items.Add($ctxExportCsv) | Out-Null

    $ctxExportHtml = New-Object System.Windows.Forms.ToolStripMenuItem("Export to HTML...")
    $ctxExportHtml.ForeColor = $clrText
    $ctxExportHtml.Add_Click(({
        $table = if ($ugState.ViewMode -eq 'Users') { $userTable } else { $groupTable }
        $label = $ugState.ViewMode
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "HTML Files (*.html)|*.html"
        $sfd.DefaultExt = "html"
        $sfd.FileName = "Local$label-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Export-MmcIfHtml -DataTable $table -OutputPath $sfd.FileName -ReportTitle "Local $label Report"
            & $statusFunc "Exported to $($sfd.FileName)"
        }
        $sfd.Dispose()
    }).GetNewClosure())
    $ctxUG.Items.Add($ctxExportHtml) | Out-Null

    $gridUG.ContextMenuStrip = $ctxUG

    # -------------------------------------------------------------------
    # Assemble layout
    # -------------------------------------------------------------------

    $tabPage.Controls.Add($splitUG)
    $splitUG.BringToFront()

    $tabPage.Controls.Add($pnlSep)
    $pnlSep.SendToBack()

    $tabPage.Controls.Add($pnlToolbar)
    $pnlToolbar.SendToBack()

    # Defer SplitterDistance
    $ugSplitFlag = @{ Done = $false }
    $splitUG.Add_SizeChanged(({
        if (-not $ugSplitFlag.Done -and $splitUG.Height -gt ($splitUG.Panel1MinSize + $splitUG.Panel2MinSize + $splitUG.SplitterWidth)) {
            $ugSplitFlag.Done = $true
            $splitUG.SplitterDistance = [Math]::Max($splitUG.Panel1MinSize, [int]($splitUG.Height * 0.6))
        }
    }).GetNewClosure())

    # Initialize with Users view
    & $setupUserColumns
    & $loadUsers
}
