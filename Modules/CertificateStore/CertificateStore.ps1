<#
.SYNOPSIS
    Certificate Store Browser module for MMC-If.

.DESCRIPTION
    Browse user and machine certificate stores using .NET X509Store.
    View expiry dates, thumbprints, and certificate details without certmgr.msc (mmc.exe).
#>

function Initialize-CertificateStore {
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
    $clrTreeBg   = $theme.TreeBg
    $isDark      = $theme.DarkMode

    # -------------------------------------------------------------------
    # State
    # -------------------------------------------------------------------
    $certTable = New-Object System.Data.DataTable
    [void]$certTable.Columns.Add("Subject", [string])
    [void]$certTable.Columns.Add("Issuer", [string])
    [void]$certTable.Columns.Add("Expires", [string])
    [void]$certTable.Columns.Add("Thumbprint", [string])
    [void]$certTable.Columns.Add("FriendlyName", [string])
    [void]$certTable.Columns.Add("_ExpiryDate", [string])
    [void]$certTable.Columns.Add("_FullSubject", [string])
    [void]$certTable.Columns.Add("_FullIssuer", [string])
    [void]$certTable.Columns.Add("_SerialNumber", [string])
    [void]$certTable.Columns.Add("_SigAlgo", [string])
    [void]$certTable.Columns.Add("_ValidFrom", [string])
    [void]$certTable.Columns.Add("_KeyUsage", [string])
    [void]$certTable.Columns.Add("_SAN", [string])
    [void]$certTable.Columns.Add("_Template", [string])

    # -------------------------------------------------------------------
    # Helpers
    # -------------------------------------------------------------------

    $extractCN = {
        param([string]$DistinguishedName)
        if ([string]::IsNullOrEmpty($DistinguishedName)) { return '' }
        if ($DistinguishedName -match '^CN=([^,]+)') { return $Matches[1] }
        return $DistinguishedName
    }

    $getSAN = {
        param($cert)
        $sanExt = $cert.Extensions | Where-Object { $_.Oid.Value -eq '2.5.29.17' }
        if ($sanExt) { return $sanExt.Format($false) }
        return ''
    }

    $getTemplate = {
        param($cert)
        # V2 template OID
        $templateExt = $cert.Extensions | Where-Object { $_.Oid.Value -eq '1.3.6.1.4.1.311.21.7' }
        if (-not $templateExt) {
            # V1 template OID
            $templateExt = $cert.Extensions | Where-Object { $_.Oid.Value -eq '1.3.6.1.4.1.311.20.2' }
        }
        if ($templateExt) { return $templateExt.Format($false) }
        return ''
    }

    $getKeyUsage = {
        param($cert)
        $parts = @()
        $kuExt = $cert.Extensions | Where-Object { $_ -is [System.Security.Cryptography.X509Certificates.X509KeyUsageExtension] }
        if ($kuExt) { $parts += $kuExt.KeyUsages.ToString() }
        $ekuExt = $cert.Extensions | Where-Object { $_ -is [System.Security.Cryptography.X509Certificates.X509EnhancedKeyUsageExtension] }
        if ($ekuExt) { $parts += ($ekuExt.EnhancedKeyUsages | ForEach-Object { $_.FriendlyName }) -join ', ' }
        return $parts -join '; '
    }

    # -------------------------------------------------------------------
    # Toolbar (Dock:Top)
    # -------------------------------------------------------------------

    $pnlToolbar = New-Object System.Windows.Forms.Panel
    $pnlToolbar.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlToolbar.Height = 40
    $pnlToolbar.BackColor = $clrPanelBg
    $pnlToolbar.Padding = New-Object System.Windows.Forms.Padding(6, 4, 6, 4)

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = "Refresh"
    $btnRefresh.Location = New-Object System.Drawing.Point(8, 5)
    $btnRefresh.Size = New-Object System.Drawing.Size(80, 28)
    $btnRefresh.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    & $btnStyle $btnRefresh $clrAccent
    $pnlToolbar.Controls.Add($btnRefresh)

    $lblHint = New-Object System.Windows.Forms.Label
    $lblHint.Text = "Certificate stores via .NET X509Store (read-only)"
    $lblHint.AutoSize = $true
    $lblHint.Location = New-Object System.Drawing.Point(100, 10)
    $lblHint.ForeColor = $clrHint
    $lblHint.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $pnlToolbar.Controls.Add($lblHint)

    # Separator
    $pnlSep = New-Object System.Windows.Forms.Panel
    $pnlSep.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlSep.Height = 1; $pnlSep.BackColor = $clrSepLine

    # -------------------------------------------------------------------
    # Main SplitContainer (tree left, content right)
    # -------------------------------------------------------------------

    $splitMain = New-Object System.Windows.Forms.SplitContainer
    $splitMain.Dock = [System.Windows.Forms.DockStyle]::Fill
    $splitMain.Orientation = [System.Windows.Forms.Orientation]::Vertical
    $splitMain.SplitterWidth = 6
    $splitMain.BackColor = $clrSepLine
    $splitMain.Panel1.BackColor = $clrPanelBg
    $splitMain.Panel2.BackColor = $clrPanelBg

    # -------------------------------------------------------------------
    # TreeView (left panel)
    # -------------------------------------------------------------------

    $treeCerts = New-Object System.Windows.Forms.TreeView
    $treeCerts.Dock = [System.Windows.Forms.DockStyle]::Fill
    $treeCerts.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $treeCerts.BackColor = $clrTreeBg
    $treeCerts.ForeColor = $clrText
    $treeCerts.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $treeCerts.FullRowSelect = $true
    $treeCerts.HideSelection = $false
    $treeCerts.ItemHeight = 22
    & $dblBuffer $treeCerts

    # Build tree
    $storeDefinitions = [ordered]@{
        'CurrentUser' = @(
            @{ Name = 'My';                   Label = 'Personal' }
            @{ Name = 'Root';                  Label = 'Trusted Root CA' }
            @{ Name = 'CertificateAuthority';  Label = 'Intermediate CA' }
            @{ Name = 'TrustedPeople';         Label = 'Trusted People' }
            @{ Name = 'TrustedPublisher';      Label = 'Trusted Publishers' }
            @{ Name = 'Disallowed';            Label = 'Disallowed' }
        )
        'LocalMachine' = @(
            @{ Name = 'My';                   Label = 'Personal' }
            @{ Name = 'Root';                  Label = 'Trusted Root CA' }
            @{ Name = 'CertificateAuthority';  Label = 'Intermediate CA' }
            @{ Name = 'TrustedPeople';         Label = 'Trusted People' }
            @{ Name = 'TrustedPublisher';      Label = 'Trusted Publishers' }
            @{ Name = 'Disallowed';            Label = 'Disallowed' }
            @{ Name = 'Remote Desktop';        Label = 'Remote Desktop' }
            @{ Name = 'WebHosting';            Label = 'Web Hosting' }
        )
    }

    foreach ($locationName in $storeDefinitions.Keys) {
        $locationNode = New-Object System.Windows.Forms.TreeNode($locationName)
        $locationNode.Tag = $null
        $storeLocation = [System.Security.Cryptography.X509Certificates.StoreLocation]::$locationName

        foreach ($storeDef in $storeDefinitions[$locationName]) {
            $storeNode = New-Object System.Windows.Forms.TreeNode($storeDef.Label)
            $storeNode.Tag = @{ Location = $storeLocation; StoreName = $storeDef.Name }
            $locationNode.Nodes.Add($storeNode) | Out-Null
        }

        $treeCerts.Nodes.Add($locationNode) | Out-Null
        $locationNode.Expand()
    }

    $splitMain.Panel1.Controls.Add($treeCerts)

    # -------------------------------------------------------------------
    # Right panel: nested split (grid top, detail bottom)
    # -------------------------------------------------------------------

    $splitRight = New-Object System.Windows.Forms.SplitContainer
    $splitRight.Dock = [System.Windows.Forms.DockStyle]::Fill
    $splitRight.Orientation = [System.Windows.Forms.Orientation]::Horizontal
    $splitRight.SplitterWidth = 6
    $splitRight.BackColor = $clrSepLine
    $splitRight.Panel1.BackColor = $clrPanelBg
    $splitRight.Panel2.BackColor = $clrPanelBg

    # -------------------------------------------------------------------
    # Certificate grid (top-right)
    # -------------------------------------------------------------------

    $gridCerts = & $newGrid
    $gridCerts.AllowUserToOrderColumns = $true

    $colSubject = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colSubject.HeaderText = "Subject"; $colSubject.DataPropertyName = "Subject"; $colSubject.Width = 200
    $colIssuer = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colIssuer.HeaderText = "Issuer"; $colIssuer.DataPropertyName = "Issuer"; $colIssuer.Width = 180
    $colExpires = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colExpires.HeaderText = "Expires"; $colExpires.DataPropertyName = "Expires"; $colExpires.Width = 120
    $colThumb = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colThumb.HeaderText = "Thumbprint"; $colThumb.DataPropertyName = "Thumbprint"; $colThumb.Width = 150
    $colFriendly = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colFriendly.HeaderText = "Friendly Name"; $colFriendly.DataPropertyName = "FriendlyName"
    $colFriendly.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill

    $gridCerts.Columns.Add($colSubject) | Out-Null
    $gridCerts.Columns.Add($colIssuer) | Out-Null
    $gridCerts.Columns.Add($colExpires) | Out-Null
    $gridCerts.Columns.Add($colThumb) | Out-Null
    $gridCerts.Columns.Add($colFriendly) | Out-Null

    $gridCerts.DataSource = $certTable
    $splitRight.Panel1.Controls.Add($gridCerts)

    # Color-code Expires column
    $gridCerts.Add_CellFormatting(({
        param($s, $e)
        if ($e.RowIndex -lt 0 -or $e.ColumnIndex -ne 2) { return }
        $view = $certTable.DefaultView
        if ($e.RowIndex -ge $view.Count) { return }
        $expiryStr = [string]$view[$e.RowIndex]["_ExpiryDate"]
        if ([string]::IsNullOrEmpty($expiryStr)) { return }
        try {
            $expiry = [datetime]::Parse($expiryStr)
            $now = Get-Date
            if ($expiry -lt $now) {
                $e.CellStyle.ForeColor = $clrErrText
            } elseif ($expiry -lt $now.AddDays(30)) {
                $e.CellStyle.ForeColor = $clrWarnText
            }
        } catch { }
    }).GetNewClosure())

    # -------------------------------------------------------------------
    # Detail panel (bottom-right)
    # -------------------------------------------------------------------

    $txtDetail = New-Object System.Windows.Forms.TextBox
    $txtDetail.Dock = [System.Windows.Forms.DockStyle]::Fill
    $txtDetail.Multiline = $true; $txtDetail.ReadOnly = $true
    $txtDetail.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $txtDetail.Font = New-Object System.Drawing.Font("Consolas", 9.5)
    $txtDetail.BackColor = $clrDetailBg; $txtDetail.ForeColor = $clrText
    $txtDetail.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $txtDetail.WordWrap = $true
    $splitRight.Panel2.Controls.Add($txtDetail)

    $splitMain.Panel2.Controls.Add($splitRight)

    # -------------------------------------------------------------------
    # Load certificates on tree selection
    # -------------------------------------------------------------------

    $loadCerts = {
        param([hashtable]$StoreInfo)
        $certTable.Clear()
        $txtDetail.Text = ''
        $tabPage.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        [System.Windows.Forms.Application]::DoEvents()

        try {
            $store = New-Object System.Security.Cryptography.X509Certificates.X509Store(
                $StoreInfo.StoreName,
                $StoreInfo.Location
            )
            $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)

            foreach ($cert in $store.Certificates) {
                $row = $certTable.NewRow()
                $row["Subject"] = & $extractCN $cert.Subject
                $row["Issuer"] = & $extractCN $cert.Issuer
                $row["Expires"] = $cert.NotAfter.ToString("yyyy-MM-dd HH:mm")
                $row["Thumbprint"] = $cert.Thumbprint
                $row["FriendlyName"] = $cert.FriendlyName
                $row["_ExpiryDate"] = $cert.NotAfter.ToString("o")
                $row["_FullSubject"] = $cert.Subject
                $row["_FullIssuer"] = $cert.Issuer
                $row["_SerialNumber"] = $cert.SerialNumber
                $row["_SigAlgo"] = $cert.SignatureAlgorithm.FriendlyName
                $row["_ValidFrom"] = $cert.NotBefore.ToString("yyyy-MM-dd HH:mm")
                $row["_KeyUsage"] = & $getKeyUsage $cert
                $row["_SAN"] = & $getSAN $cert
                $row["_Template"] = & $getTemplate $cert
                $certTable.Rows.Add($row)
            }

            $store.Close()
            & $statusFunc "$($StoreInfo.Location)\$($StoreInfo.StoreName): $($certTable.Rows.Count) certificate(s)"
            Write-Log "Loaded $($certTable.Rows.Count) certificates from $($StoreInfo.Location)\$($StoreInfo.StoreName)"
        } catch {
            $errMsg = $_.Exception.Message
            if ($errMsg -like '*access*' -or $errMsg -like '*denied*' -or $errMsg -like '*crypto*') {
                & $statusFunc "Access denied: $($StoreInfo.Location)\$($StoreInfo.StoreName) - may require elevation"
            } else {
                & $statusFunc "Error: $errMsg"
            }
        }

        $tabPage.Cursor = [System.Windows.Forms.Cursors]::Default
    }.GetNewClosure()

    $treeCerts.Add_AfterSelect(({
        param($s, $e)
        $node = $e.Node
        if (-not $node -or -not ($node.Tag -is [hashtable])) {
            $certTable.Clear()
            $txtDetail.Text = ''
            return
        }
        & $loadCerts $node.Tag
    }).GetNewClosure())

    $btnRefresh.Add_Click(({
        $node = $treeCerts.SelectedNode
        if ($node -and ($node.Tag -is [hashtable])) {
            & $loadCerts $node.Tag
        }
    }).GetNewClosure())

    # -------------------------------------------------------------------
    # Detail on row select
    # -------------------------------------------------------------------

    $gridCerts.Add_SelectionChanged(({
        if ($gridCerts.CurrentRow -and $gridCerts.CurrentRow.Index -ge 0) {
            $view = $certTable.DefaultView
            $idx = $gridCerts.CurrentRow.Index
            if ($idx -lt $view.Count) {
                $r = $view[$idx]
                $lines = @(
                    "Subject: $([string]$r['_FullSubject'])",
                    "Issuer: $([string]$r['_FullIssuer'])",
                    "Serial Number: $([string]$r['_SerialNumber'])",
                    "Thumbprint: $([string]$r['Thumbprint'])",
                    "Friendly Name: $([string]$r['FriendlyName'])",
                    "",
                    "Valid From: $([string]$r['_ValidFrom'])",
                    "Valid To: $([string]$r['Expires'])",
                    "Signature Algorithm: $([string]$r['_SigAlgo'])",
                    ""
                )

                $keyUsage = [string]$r['_KeyUsage']
                if ($keyUsage) { $lines += "Key Usage: $keyUsage" }

                $san = [string]$r['_SAN']
                if ($san) { $lines += "Subject Alternative Names: $san" }

                $template = [string]$r['_Template']
                if ($template) { $lines += "Certificate Template: $template" }

                $txtDetail.Text = $lines -join "`r`n"
            }
        }
    }).GetNewClosure())

    # -------------------------------------------------------------------
    # Context menu
    # -------------------------------------------------------------------

    $ctxCerts = New-Object System.Windows.Forms.ContextMenuStrip
    if ($isDark -and $script:DarkRenderer) { $ctxCerts.Renderer = $script:DarkRenderer }
    $ctxCerts.BackColor = $clrPanelBg; $ctxCerts.ForeColor = $clrText

    $ctxCopyThumb = New-Object System.Windows.Forms.ToolStripMenuItem("Copy Thumbprint")
    $ctxCopyThumb.ForeColor = $clrText
    $ctxCopyThumb.Add_Click(({
        if ($gridCerts.CurrentRow -and $gridCerts.CurrentRow.Index -ge 0) {
            $view = $certTable.DefaultView
            $idx = $gridCerts.CurrentRow.Index
            if ($idx -lt $view.Count) {
                [System.Windows.Forms.Clipboard]::SetText([string]$view[$idx]["Thumbprint"])
            }
        }
    }).GetNewClosure())
    $ctxCerts.Items.Add($ctxCopyThumb) | Out-Null

    $ctxCopySubject = New-Object System.Windows.Forms.ToolStripMenuItem("Copy Subject")
    $ctxCopySubject.ForeColor = $clrText
    $ctxCopySubject.Add_Click(({
        if ($gridCerts.CurrentRow -and $gridCerts.CurrentRow.Index -ge 0) {
            $view = $certTable.DefaultView
            $idx = $gridCerts.CurrentRow.Index
            if ($idx -lt $view.Count) {
                [System.Windows.Forms.Clipboard]::SetText([string]$view[$idx]["_FullSubject"])
            }
        }
    }).GetNewClosure())
    $ctxCerts.Items.Add($ctxCopySubject) | Out-Null

    $ctxCopyDetail = New-Object System.Windows.Forms.ToolStripMenuItem("Copy Full Details")
    $ctxCopyDetail.ForeColor = $clrText
    $ctxCopyDetail.Add_Click(({
        if ($txtDetail.Text) { [System.Windows.Forms.Clipboard]::SetText($txtDetail.Text) }
    }).GetNewClosure())
    $ctxCerts.Items.Add($ctxCopyDetail) | Out-Null

    $ctxCerts.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null

    $ctxExportCsv = New-Object System.Windows.Forms.ToolStripMenuItem("Export to CSV...")
    $ctxExportCsv.ForeColor = $clrText
    $ctxExportCsv.Add_Click(({
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV Files (*.csv)|*.csv"
        $sfd.DefaultExt = "csv"
        $sfd.FileName = "Certificates-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Export-MmcIfCsv -DataTable $certTable -OutputPath $sfd.FileName
            & $statusFunc "Exported to $($sfd.FileName)"
        }
        $sfd.Dispose()
    }).GetNewClosure())
    $ctxCerts.Items.Add($ctxExportCsv) | Out-Null

    $ctxExportHtml = New-Object System.Windows.Forms.ToolStripMenuItem("Export to HTML...")
    $ctxExportHtml.ForeColor = $clrText
    $ctxExportHtml.Add_Click(({
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "HTML Files (*.html)|*.html"
        $sfd.DefaultExt = "html"
        $sfd.FileName = "Certificates-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Export-MmcIfHtml -DataTable $certTable -OutputPath $sfd.FileName -ReportTitle "Certificate Store Report"
            & $statusFunc "Exported to $($sfd.FileName)"
        }
        $sfd.Dispose()
    }).GetNewClosure())
    $ctxCerts.Items.Add($ctxExportHtml) | Out-Null

    $gridCerts.ContextMenuStrip = $ctxCerts

    # -------------------------------------------------------------------
    # Assemble layout
    # -------------------------------------------------------------------

    $tabPage.Controls.Add($splitMain)
    $splitMain.BringToFront()

    $tabPage.Controls.Add($pnlSep)
    $pnlSep.SendToBack()

    $tabPage.Controls.Add($pnlToolbar)
    $pnlToolbar.SendToBack()

    # Defer SplitterDistance and MinSize for both splits
    $mainSplitFlag = @{ Done = $false }
    $splitMain.Add_SizeChanged(({
        if (-not $mainSplitFlag.Done -and $splitMain.Width -gt 466) {
            $mainSplitFlag.Done = $true
            $splitMain.Panel1MinSize = 160
            $splitMain.Panel2MinSize = 300
            $splitMain.SplitterDistance = 200
        }
    }).GetNewClosure())

    $rightSplitFlag = @{ Done = $false }
    $splitRight.Add_SizeChanged(({
        if (-not $rightSplitFlag.Done -and $splitRight.Height -gt 186) {
            $rightSplitFlag.Done = $true
            $splitRight.Panel1MinSize = 100
            $splitRight.Panel2MinSize = 80
            $splitRight.SplitterDistance = [int]($splitRight.Height * 0.6)
        }
    }).GetNewClosure())
}
