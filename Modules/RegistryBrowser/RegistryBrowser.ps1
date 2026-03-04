<#
.SYNOPSIS
    Registry Browser module for MMC-Alt.

.DESCRIPTION
    Provides a regedit-like interface for browsing the Windows registry.
    Read-only for all hives except HKCU, which supports full read-write operations.
    Uses .NET Microsoft.Win32.RegistryKey API -- does not require mmc.exe or regedit.exe.
#>

function Initialize-RegistryBrowser {
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
    $clrTreeBg   = $theme.TreeBg
    $clrGridAlt  = $theme.GridAlt
    $clrGridLine = $theme.GridLine
    $clrHint     = $theme.Hint
    $clrInputBdr = $theme.InputBdr
    $clrErrText  = $theme.ErrText
    $isDark      = $theme.DarkMode

    # Registry hive definitions
    $hiveMap = [ordered]@{
        'HKCU' = @{ Key = [Microsoft.Win32.Registry]::CurrentUser;    ReadOnly = $false; Label = 'HKEY_CURRENT_USER' }
        'HKLM' = @{ Key = [Microsoft.Win32.Registry]::LocalMachine;   ReadOnly = $true;  Label = 'HKEY_LOCAL_MACHINE' }
        'HKCR' = @{ Key = [Microsoft.Win32.Registry]::ClassesRoot;    ReadOnly = $true;  Label = 'HKEY_CLASSES_ROOT' }
        'HKU'  = @{ Key = [Microsoft.Win32.Registry]::Users;          ReadOnly = $true;  Label = 'HKEY_USERS' }
        'HKCC' = @{ Key = [Microsoft.Win32.Registry]::CurrentConfig;  ReadOnly = $true;  Label = 'HKEY_CURRENT_CONFIG' }
    }

    # Shared mutable state across closures (hashtable = reference type)
    $regState = @{
        CurrentPath  = ''
        SearchText   = ''
        SearchKeys   = $true
        SearchValues = $true
        SearchData   = $true
        LastMatch    = $null
    }

    # Scriptblock references set later; accessible from closures via hashtable reference
    $regFuncs = @{
        ShowValueEditor  = $null
        SearchNext       = $null
        ShowSearchDialog = $null
    }

    # -----------------------------------------------------------------------
    # Helper: determine if a path is under HKCU (writable)
    # -----------------------------------------------------------------------
    $isWritable = {
        param([string]$Path)
        return ($Path -match '^HKCU(\\|$)')
    }

    # -----------------------------------------------------------------------
    # Helper: open a registry key from a full path string
    # -----------------------------------------------------------------------
    $openRegKey = {
        param([string]$FullPath, [bool]$Writable = $false)
        $parts = $FullPath -split '\\', 2
        $hiveName = $parts[0]
        $subPath = if ($parts.Count -gt 1) { $parts[1] } else { $null }

        if (-not $hiveMap.Contains($hiveName)) { return $null }
        $hive = $hiveMap[$hiveName]

        if (-not $subPath) { return $hive.Key }

        $actualWritable = $Writable -and (-not $hive.ReadOnly)
        try {
            return $hive.Key.OpenSubKey($subPath, $actualWritable)
        } catch {
            return $null
        }
    }.GetNewClosure()

    # -----------------------------------------------------------------------
    # Helper: get value type display name
    # -----------------------------------------------------------------------
    $getTypeName = {
        param([Microsoft.Win32.RegistryValueKind]$Kind)
        switch ($Kind) {
            'String'       { 'REG_SZ' }
            'ExpandString' { 'REG_EXPAND_SZ' }
            'Binary'       { 'REG_BINARY' }
            'DWord'        { 'REG_DWORD' }
            'QWord'        { 'REG_QWORD' }
            'MultiString'  { 'REG_MULTI_SZ' }
            'None'         { 'REG_NONE' }
            default        { $Kind.ToString() }
        }
    }

    # -----------------------------------------------------------------------
    # Helper: format value for display
    # -----------------------------------------------------------------------
    $formatValue = {
        param($Value, [Microsoft.Win32.RegistryValueKind]$Kind)
        if ($null -eq $Value) { return '(value not set)' }
        switch ($Kind) {
            'Binary' {
                if ($Value -is [byte[]]) {
                    return ($Value | ForEach-Object { '{0:X2}' -f $_ }) -join ' '
                }
                return $Value.ToString()
            }
            'MultiString' {
                if ($Value -is [string[]]) { return $Value -join ' | ' }
                return $Value.ToString()
            }
            'DWord'  { return '0x{0:X8} ({0})' -f [uint32]$Value }
            'QWord'  { return '0x{0:X16} ({0})' -f [uint64]$Value }
            default  { return $Value.ToString() }
        }
    }

    # -----------------------------------------------------------------------
    # Address bar panel (Dock:Top)
    # -----------------------------------------------------------------------

    $pnlAddress = New-Object System.Windows.Forms.Panel
    $pnlAddress.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlAddress.Height = 36
    $pnlAddress.BackColor = $clrPanelBg
    $pnlAddress.Padding = New-Object System.Windows.Forms.Padding(6, 4, 6, 4)

    $lblPath = New-Object System.Windows.Forms.Label
    $lblPath.Text = "Path:"
    $lblPath.AutoSize = $true
    $lblPath.Location = New-Object System.Drawing.Point(8, 9)
    $lblPath.ForeColor = $clrText
    $lblPath.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $pnlAddress.Controls.Add($lblPath)

    $txtPath = New-Object System.Windows.Forms.TextBox
    $txtPath.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $txtPath.Location = New-Object System.Drawing.Point(46, 5)
    $txtPath.Height = 24
    $txtPath.Font = New-Object System.Drawing.Font("Consolas", 9.5)
    $txtPath.BackColor = $clrDetailBg
    $txtPath.ForeColor = $clrText
    $txtPath.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $pnlAddress.Controls.Add($txtPath)

    $btnGo = New-Object System.Windows.Forms.Button
    $btnGo.Text = "Go"
    $btnGo.Size = New-Object System.Drawing.Size(50, 26)
    $btnGo.Location = New-Object System.Drawing.Point(300, 4)
    $btnGo.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnGo.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    & $btnStyle $btnGo $clrAccent

    # Set txtPath width to fill between label and button, anchored left+top+right
    $txtPath.Width = 248
    $txtPath.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right

    $pnlAddress.Controls.Add($btnGo)

    # Separator line below address bar
    $pnlAddrSep = New-Object System.Windows.Forms.Panel
    $pnlAddrSep.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlAddrSep.Height = 1
    $pnlAddrSep.BackColor = $clrSepLine

    # -----------------------------------------------------------------------
    # SplitContainer (Vertical: TreeView left, Values grid right)
    # -----------------------------------------------------------------------

    $splitMain = New-Object System.Windows.Forms.SplitContainer
    $splitMain.Dock = [System.Windows.Forms.DockStyle]::Fill
    $splitMain.Orientation = [System.Windows.Forms.Orientation]::Vertical
    $splitMain.SplitterWidth = 6
    $splitMain.BackColor = $clrSepLine
    $splitMain.Panel1.BackColor = $clrPanelBg
    $splitMain.Panel2.BackColor = $clrPanelBg
    $splitMain.Panel1MinSize = 100
    $splitMain.Panel2MinSize = 100

    # -----------------------------------------------------------------------
    # TreeView (left panel)
    # -----------------------------------------------------------------------

    $treeReg = New-Object System.Windows.Forms.TreeView
    $treeReg.Dock = [System.Windows.Forms.DockStyle]::Fill
    $treeReg.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
    $treeReg.BackColor = $clrTreeBg
    $treeReg.ForeColor = $clrText
    $treeReg.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $treeReg.FullRowSelect = $true
    $treeReg.HideSelection = $false
    $treeReg.ShowLines = $true
    $treeReg.ShowPlusMinus = $true
    $treeReg.ShowRootLines = $true
    $treeReg.PathSeparator = '\'
    if ($isDark) { $treeReg.LineColor = [System.Drawing.Color]::FromArgb(80, 80, 80) }
    & $dblBuffer $treeReg

    # Populate root nodes with dummy children for expand arrows
    foreach ($hiveName in $hiveMap.Keys) {
        $info = $hiveMap[$hiveName]
        $label = if ($info.ReadOnly) { "$hiveName (Read-Only)" } else { $hiveName }
        $node = New-Object System.Windows.Forms.TreeNode($label)
        $node.Tag = $hiveName
        $node.Name = $hiveName
        # Add dummy child so expand arrow shows
        $dummy = New-Object System.Windows.Forms.TreeNode("__dummy__")
        $node.Nodes.Add($dummy) | Out-Null
        $treeReg.Nodes.Add($node) | Out-Null
    }

    # BeforeExpand: lazy-load subkeys
    $treeReg.Add_BeforeExpand(({
        param($s, $e)
        $node = $e.Node
        # Only load if we have the dummy placeholder
        if ($node.Nodes.Count -eq 1 -and $node.Nodes[0].Text -eq '__dummy__') {
            $node.Nodes.Clear()
            $path = [string]$node.Tag

            $regKey = $null
            try {
                $regKey = & $openRegKey $path $false
                if ($null -eq $regKey) {
                    $errNode = New-Object System.Windows.Forms.TreeNode("(Access Denied)")
                    $errNode.ForeColor = $clrErrText
                    $node.Nodes.Add($errNode) | Out-Null
                    return
                }

                $subNames = $regKey.GetSubKeyNames() | Sort-Object
                foreach ($name in $subNames) {
                    $childPath = "$path\$name"
                    $childNode = New-Object System.Windows.Forms.TreeNode($name)
                    $childNode.Tag = $childPath
                    $childNode.Name = $name

                    # Check if this child has subkeys (for expand arrow)
                    $childKey = $null
                    try {
                        $childKey = $regKey.OpenSubKey($name, $false)
                        if ($null -ne $childKey -and $childKey.SubKeyCount -gt 0) {
                            $dummy = New-Object System.Windows.Forms.TreeNode("__dummy__")
                            $childNode.Nodes.Add($dummy) | Out-Null
                        }
                    } catch {
                        # Access denied on child -- still show node but with expand arrow
                        $dummy = New-Object System.Windows.Forms.TreeNode("__dummy__")
                        $childNode.Nodes.Add($dummy) | Out-Null
                    } finally {
                        if ($null -ne $childKey -and $childPath -ne $path) { $childKey.Dispose() }
                    }

                    $node.Nodes.Add($childNode) | Out-Null
                }

                if ($subNames.Count -eq 0) {
                    # No subkeys -- that's fine, node just won't be expandable next time
                }
            } catch [System.Security.SecurityException] {
                $errNode = New-Object System.Windows.Forms.TreeNode("(Access Denied)")
                $errNode.ForeColor = $clrErrText
                $node.Nodes.Add($errNode) | Out-Null
            } catch {
                $errNode = New-Object System.Windows.Forms.TreeNode("(Error: $($_.Exception.Message))")
                $errNode.ForeColor = $clrErrText
                $node.Nodes.Add($errNode) | Out-Null
            } finally {
                # Don't dispose hive root keys
                $parts = $path -split '\\', 2
                if ($null -ne $regKey -and $parts.Count -gt 1) { $regKey.Dispose() }
            }
        }
    }).GetNewClosure())

    $splitMain.Panel1.Controls.Add($treeReg)

    # -----------------------------------------------------------------------
    # Values DataGridView (right panel)
    # -----------------------------------------------------------------------

    $gridValues = & $newGrid
    $gridValues.AllowUserToOrderColumns = $true

    $colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colName.HeaderText = "Name"; $colName.DataPropertyName = "Name"; $colName.Width = 200
    $colType = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colType.HeaderText = "Type"; $colType.DataPropertyName = "Type"; $colType.Width = 120
    $colData = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colData.HeaderText = "Data"; $colData.DataPropertyName = "Data"; $colData.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill

    $gridValues.Columns.Add($colName) | Out-Null
    $gridValues.Columns.Add($colType) | Out-Null
    $gridValues.Columns.Add($colData) | Out-Null

    # DataTable for values (local; shared via closure capture)
    $valuesTable = New-Object System.Data.DataTable
    [void]$valuesTable.Columns.Add("Name", [string])
    [void]$valuesTable.Columns.Add("Type", [string])
    [void]$valuesTable.Columns.Add("Data", [string])
    $gridValues.DataSource = $valuesTable

    $splitMain.Panel2.Controls.Add($gridValues)

    # -----------------------------------------------------------------------
    # Populate values on tree selection
    # -----------------------------------------------------------------------

    $treeReg.Add_AfterSelect(({
        param($s, $e)
        $node = $e.Node
        if (-not $node -or -not $node.Tag) { return }
        $path = [string]$node.Tag
        $regState.CurrentPath = $path

        # Update address bar
        $txtPath.Text = $path

        # Update status bar
        & $statusFunc $path

        # Populate values grid
        $valuesTable.Clear()

        $regKey = $null
        try {
            $regKey = & $openRegKey $path $false
            if ($null -eq $regKey) {
                $row = $valuesTable.NewRow()
                $row["Name"] = "(Access Denied)"
                $row["Type"] = ""
                $row["Data"] = "Cannot read values from this key"
                $valuesTable.Rows.Add($row)
                return
            }

            # Default value
            $defaultVal = $regKey.GetValue('')
            $defaultKind = [Microsoft.Win32.RegistryValueKind]::String
            try { $defaultKind = $regKey.GetValueKind('') } catch { }
            $row = $valuesTable.NewRow()
            $row["Name"] = "(Default)"
            $row["Type"] = & $getTypeName $defaultKind
            $row["Data"] = if ($null -eq $defaultVal) { '(value not set)' } else { & $formatValue $defaultVal $defaultKind }
            $valuesTable.Rows.Add($row)

            # Named values
            $valueNames = $regKey.GetValueNames() | Where-Object { $_ -ne '' } | Sort-Object
            foreach ($vName in $valueNames) {
                try {
                    $val = $regKey.GetValue($vName, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
                    $kind = $regKey.GetValueKind($vName)
                    $row = $valuesTable.NewRow()
                    $row["Name"] = $vName
                    $row["Type"] = & $getTypeName $kind
                    $row["Data"] = & $formatValue $val $kind
                    $valuesTable.Rows.Add($row)
                } catch {
                    $row = $valuesTable.NewRow()
                    $row["Name"] = $vName
                    $row["Type"] = "ERROR"
                    $row["Data"] = $_.Exception.Message
                    $valuesTable.Rows.Add($row)
                }
            }

            # Update status with counts
            $subCount = $regKey.SubKeyCount
            $valCount = $regKey.ValueCount
            & $statusFunc "$path  |  $subCount subkey(s), $valCount value(s)"

        } catch [System.Security.SecurityException] {
            $row = $valuesTable.NewRow()
            $row["Name"] = "(Access Denied)"
            $row["Type"] = ""
            $row["Data"] = "Insufficient permissions to read this key"
            $valuesTable.Rows.Add($row)
        } catch {
            $row = $valuesTable.NewRow()
            $row["Name"] = "(Error)"
            $row["Type"] = ""
            $row["Data"] = $_.Exception.Message
            $valuesTable.Rows.Add($row)
        } finally {
            $parts = $path -split '\\', 2
            if ($null -ne $regKey -and $parts.Count -gt 1) { $regKey.Dispose() }
        }
    }).GetNewClosure())

    # -----------------------------------------------------------------------
    # Address bar navigation
    # -----------------------------------------------------------------------

    $navigateToPath = {
        param([string]$TargetPath)
        $TargetPath = $TargetPath.Trim().TrimEnd('\')
        if ([string]::IsNullOrWhiteSpace($TargetPath)) { return }

        # Normalize common aliases
        $TargetPath = $TargetPath -replace '^HKEY_CURRENT_USER',  'HKCU'
        $TargetPath = $TargetPath -replace '^HKEY_LOCAL_MACHINE', 'HKLM'
        $TargetPath = $TargetPath -replace '^HKEY_CLASSES_ROOT',  'HKCR'
        $TargetPath = $TargetPath -replace '^HKEY_USERS',         'HKU'
        $TargetPath = $TargetPath -replace '^HKEY_CURRENT_CONFIG','HKCC'

        $segments = $TargetPath -split '\\'
        $hiveName = $segments[0]

        # Find root node
        $rootNode = $null
        foreach ($n in $treeReg.Nodes) {
            if ($n.Name -eq $hiveName) { $rootNode = $n; break }
        }
        if (-not $rootNode) {
            [System.Windows.Forms.MessageBox]::Show("Invalid registry path: $TargetPath", "Navigation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }

        # Walk down the tree, expanding as we go
        $currentNode = $rootNode
        $treeReg.BeginUpdate()
        try {
            $rootNode.Expand()
            for ($i = 1; $i -lt $segments.Count; $i++) {
                $seg = $segments[$i]
                $found = $false
                foreach ($child in $currentNode.Nodes) {
                    if ($child.Name -eq $seg -or $child.Text -eq $seg) {
                        $child.Expand()
                        $currentNode = $child
                        $found = $true
                        break
                    }
                }
                if (-not $found) {
                    [System.Windows.Forms.MessageBox]::Show("Key not found: $seg`nPath: $TargetPath", "Navigation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                    break
                }
            }
        } finally {
            $treeReg.EndUpdate()
        }

        $treeReg.SelectedNode = $currentNode
        $currentNode.EnsureVisible()
    }.GetNewClosure()

    $txtPath.Add_KeyDown(({
        param($s, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $e.SuppressKeyPress = $true
            & $navigateToPath $txtPath.Text
        }
    }).GetNewClosure())

    $btnGo.Add_Click(({ & $navigateToPath $txtPath.Text }).GetNewClosure())

    # -----------------------------------------------------------------------
    # Context menu: TreeView
    # -----------------------------------------------------------------------

    $ctxTree = New-Object System.Windows.Forms.ContextMenuStrip
    if ($isDark -and $script:DarkRenderer) { $ctxTree.Renderer = $script:DarkRenderer }
    $ctxTree.BackColor = $clrPanelBg; $ctxTree.ForeColor = $clrText

    $ctxTreeCopyPath = New-Object System.Windows.Forms.ToolStripMenuItem("Copy Key Path")
    $ctxTreeCopyPath.ForeColor = $clrText
    $ctxTreeCopyPath.Add_Click(({
        if ($treeReg.SelectedNode -and $treeReg.SelectedNode.Tag) {
            [System.Windows.Forms.Clipboard]::SetText([string]$treeReg.SelectedNode.Tag)
        }
    }).GetNewClosure())
    $ctxTree.Items.Add($ctxTreeCopyPath) | Out-Null

    $ctxTreeRefresh = New-Object System.Windows.Forms.ToolStripMenuItem("Refresh")
    $ctxTreeRefresh.ForeColor = $clrText
    $ctxTreeRefresh.Add_Click(({
        $node = $treeReg.SelectedNode
        if (-not $node -or -not $node.Tag) { return }
        $node.Nodes.Clear()
        $dummy = New-Object System.Windows.Forms.TreeNode("__dummy__")
        $node.Nodes.Add($dummy) | Out-Null
        $node.Collapse()
        $node.Expand()
        # Re-trigger AfterSelect to refresh values
        $treeReg.SelectedNode = $null
        $treeReg.SelectedNode = $node
    }).GetNewClosure())
    $ctxTree.Items.Add($ctxTreeRefresh) | Out-Null

    $ctxTreeSep1 = New-Object System.Windows.Forms.ToolStripSeparator
    $ctxTree.Items.Add($ctxTreeSep1) | Out-Null

    $ctxTreeNewKey = New-Object System.Windows.Forms.ToolStripMenuItem("New Key")
    $ctxTreeNewKey.ForeColor = $clrText
    $ctxTreeNewKey.Add_Click(({
        $node = $treeReg.SelectedNode
        if (-not $node -or -not $node.Tag) { return }
        $path = [string]$node.Tag
        if (-not (& $isWritable $path)) { return }

        $name = [Microsoft.VisualBasic.Interaction]::InputBox("Enter new key name:", "New Key", "New Key #1")
        if ([string]::IsNullOrWhiteSpace($name)) { return }

        $regKey = $null
        try {
            $regKey = & $openRegKey $path $true
            if ($null -eq $regKey) {
                [System.Windows.Forms.MessageBox]::Show("Cannot open key for writing.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
                return
            }
            $newKey = $regKey.CreateSubKey($name)
            if ($newKey) { $newKey.Dispose() }

            # Refresh the parent node
            $node.Nodes.Clear()
            $dummy = New-Object System.Windows.Forms.TreeNode("__dummy__")
            $node.Nodes.Add($dummy) | Out-Null
            $node.Collapse()
            $node.Expand()

            Write-Log "Created registry key: $path\$name"
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to create key: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        } finally {
            $parts = $path -split '\\', 2
            if ($null -ne $regKey -and $parts.Count -gt 1) { $regKey.Dispose() }
        }
    }).GetNewClosure())
    $ctxTree.Items.Add($ctxTreeNewKey) | Out-Null

    $ctxTreeDeleteKey = New-Object System.Windows.Forms.ToolStripMenuItem("Delete Key")
    $ctxTreeDeleteKey.ForeColor = $clrText
    $ctxTreeDeleteKey.Add_Click(({
        $node = $treeReg.SelectedNode
        if (-not $node -or -not $node.Tag) { return }
        $path = [string]$node.Tag
        if (-not (& $isWritable $path)) { return }

        # Don't allow deleting root hive nodes
        $parts = $path -split '\\', 2
        if ($parts.Count -lt 2) { return }

        $result = [System.Windows.Forms.MessageBox]::Show("Delete key '$($node.Text)' and all its subkeys?`n`n$path", "Confirm Delete", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($result -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $parentPath = $path.Substring(0, $path.LastIndexOf('\'))
        $keyName = $path.Substring($path.LastIndexOf('\') + 1)

        $parentKey = $null
        try {
            $parentKey = & $openRegKey $parentPath $true
            if ($null -eq $parentKey) {
                [System.Windows.Forms.MessageBox]::Show("Cannot open parent key for writing.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
                return
            }
            $parentKey.DeleteSubKeyTree($keyName)

            Write-Log "Deleted registry key: $path"

            # Remove node from tree and select parent
            $parentNode = $node.Parent
            $node.Remove()
            if ($parentNode) { $treeReg.SelectedNode = $parentNode }

        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to delete key: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        } finally {
            $pathParts = $parentPath -split '\\', 2
            if ($null -ne $parentKey -and $pathParts.Count -gt 1) { $parentKey.Dispose() }
        }
    }).GetNewClosure())
    $ctxTree.Items.Add($ctxTreeDeleteKey) | Out-Null

    # Show/hide HKCU-only items based on selected node
    $ctxTree.Add_Opening(({
        param($s, $e)
        $node = $treeReg.SelectedNode
        $writable = $false
        if ($node -and $node.Tag) { $writable = & $isWritable ([string]$node.Tag) }
        $ctxTreeNewKey.Visible = $writable
        $ctxTreeDeleteKey.Visible = $writable
        $ctxTreeSep1.Visible = $writable
        # Don't allow deleting root nodes
        if ($writable -and $node) {
            $parts = ([string]$node.Tag) -split '\\', 2
            $ctxTreeDeleteKey.Enabled = ($parts.Count -ge 2)
        }
    }).GetNewClosure())

    $treeReg.ContextMenuStrip = $ctxTree

    # -----------------------------------------------------------------------
    # Context menu: Values grid
    # -----------------------------------------------------------------------

    $ctxValues = New-Object System.Windows.Forms.ContextMenuStrip
    if ($isDark -and $script:DarkRenderer) { $ctxValues.Renderer = $script:DarkRenderer }
    $ctxValues.BackColor = $clrPanelBg; $ctxValues.ForeColor = $clrText

    $ctxValCopyName = New-Object System.Windows.Forms.ToolStripMenuItem("Copy Value Name")
    $ctxValCopyName.ForeColor = $clrText
    $ctxValCopyName.Add_Click(({
        if ($gridValues.CurrentRow) {
            $name = $gridValues.CurrentRow.Cells[0].Value
            if ($name) { [System.Windows.Forms.Clipboard]::SetText([string]$name) }
        }
    }).GetNewClosure())
    $ctxValues.Items.Add($ctxValCopyName) | Out-Null

    $ctxValCopyData = New-Object System.Windows.Forms.ToolStripMenuItem("Copy Value Data")
    $ctxValCopyData.ForeColor = $clrText
    $ctxValCopyData.Add_Click(({
        if ($gridValues.CurrentRow) {
            $data = $gridValues.CurrentRow.Cells[2].Value
            if ($data) { [System.Windows.Forms.Clipboard]::SetText([string]$data) }
        }
    }).GetNewClosure())
    $ctxValues.Items.Add($ctxValCopyData) | Out-Null

    $ctxValSep1 = New-Object System.Windows.Forms.ToolStripSeparator
    $ctxValues.Items.Add($ctxValSep1) | Out-Null

    # New Value submenu
    $ctxValNew = New-Object System.Windows.Forms.ToolStripMenuItem("New")
    $ctxValNew.ForeColor = $clrText

    $valueTypes = @(
        @{ Label = 'String Value';       Kind = [Microsoft.Win32.RegistryValueKind]::String }
        @{ Label = 'DWORD (32-bit)';     Kind = [Microsoft.Win32.RegistryValueKind]::DWord }
        @{ Label = 'QWORD (64-bit)';     Kind = [Microsoft.Win32.RegistryValueKind]::QWord }
        @{ Label = 'Binary Value';       Kind = [Microsoft.Win32.RegistryValueKind]::Binary }
        @{ Label = 'Multi-String Value'; Kind = [Microsoft.Win32.RegistryValueKind]::MultiString }
        @{ Label = 'Expandable String';  Kind = [Microsoft.Win32.RegistryValueKind]::ExpandString }
    )

    foreach ($vt in $valueTypes) {
        $item = New-Object System.Windows.Forms.ToolStripMenuItem($vt.Label)
        $item.ForeColor = $clrText
        $item.Tag = $vt.Kind
        $item.Add_Click(({
            param($s, $e)
            $kind = $s.Tag
            & $regFuncs.ShowValueEditor '' $kind $null $false
        }).GetNewClosure())
        $ctxValNew.DropDownItems.Add($item) | Out-Null
    }
    $ctxValues.Items.Add($ctxValNew) | Out-Null

    $ctxValModify = New-Object System.Windows.Forms.ToolStripMenuItem("Modify...")
    $ctxValModify.ForeColor = $clrText
    $ctxValModify.Add_Click(({
        if (-not $gridValues.CurrentRow) { return }
        $name = [string]$gridValues.CurrentRow.Cells[0].Value
        if ($name -eq '(Default)') { $name = '' }
        $path = $regState.CurrentPath

        $regKey = $null
        try {
            $regKey = & $openRegKey $path $false
            if ($null -eq $regKey) { return }
            $kind = $regKey.GetValueKind($name)
            $val = $regKey.GetValue($name, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
            & $regFuncs.ShowValueEditor $name $kind $val $true
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Cannot read value: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        } finally {
            $parts = $path -split '\\', 2
            if ($null -ne $regKey -and $parts.Count -gt 1) { $regKey.Dispose() }
        }
    }).GetNewClosure())
    $ctxValues.Items.Add($ctxValModify) | Out-Null

    $ctxValDelete = New-Object System.Windows.Forms.ToolStripMenuItem("Delete")
    $ctxValDelete.ForeColor = $clrText
    $ctxValDelete.Add_Click(({
        if (-not $gridValues.CurrentRow) { return }
        $name = [string]$gridValues.CurrentRow.Cells[0].Value
        $actualName = if ($name -eq '(Default)') { '' } else { $name }
        $path = $regState.CurrentPath

        $result = [System.Windows.Forms.MessageBox]::Show("Delete value '$name'?", "Confirm Delete", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($result -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $regKey = $null
        try {
            $regKey = & $openRegKey $path $true
            if ($null -eq $regKey) {
                [System.Windows.Forms.MessageBox]::Show("Cannot open key for writing.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
                return
            }
            $regKey.DeleteValue($actualName, $false)
            Write-Log "Deleted registry value: $path\\$name"

            # Refresh values by re-navigating
            & $navigateToPath $path
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to delete value: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        } finally {
            $parts = $path -split '\\', 2
            if ($null -ne $regKey -and $parts.Count -gt 1) { $regKey.Dispose() }
        }
    }).GetNewClosure())
    $ctxValues.Items.Add($ctxValDelete) | Out-Null

    # Show/hide HKCU-only items
    $ctxValues.Add_Opening(({
        param($s, $e)
        $writable = & $isWritable $regState.CurrentPath
        $ctxValNew.Visible = $writable
        $ctxValModify.Visible = $writable
        $ctxValDelete.Visible = $writable
        $ctxValSep1.Visible = $writable
    }).GetNewClosure())

    $gridValues.ContextMenuStrip = $ctxValues

    # Double-click to modify (HKCU only)
    $gridValues.Add_CellDoubleClick(({
        param($s, $e)
        if ($e.RowIndex -lt 0) { return }
        if (-not (& $isWritable $regState.CurrentPath)) { return }
        $ctxValModify.PerformClick()
    }).GetNewClosure())

    # -----------------------------------------------------------------------
    # Value Editor Dialog
    # -----------------------------------------------------------------------

    # Load Microsoft.VisualBasic for InputBox
    try { [void][Microsoft.VisualBasic.Interaction] } catch {
        Add-Type -AssemblyName Microsoft.VisualBasic
    }

    $regFuncs.ShowValueEditor = {
        param([string]$ValueName, [Microsoft.Win32.RegistryValueKind]$Kind, $CurrentValue, [bool]$IsEdit)

        $path = $regState.CurrentPath
        if (-not (& $isWritable $path)) { return }

        $dlg = New-Object System.Windows.Forms.Form
        $dlg.Text = if ($IsEdit) { "Edit Value" } else { "New Value" }
        $dlg.Size = New-Object System.Drawing.Size(450, 320)
        $dlg.MinimumSize = $dlg.Size
        $dlg.StartPosition = "CenterParent"
        $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false; $dlg.ShowInTaskbar = $false
        $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
        $dlg.BackColor = $clrFormBg

        # Value name
        $lblName = New-Object System.Windows.Forms.Label
        $lblName.Text = "Value name:"; $lblName.AutoSize = $true
        $lblName.Location = New-Object System.Drawing.Point(16, 16); $lblName.ForeColor = $clrText
        $dlg.Controls.Add($lblName)

        $txtName = New-Object System.Windows.Forms.TextBox
        $txtName.SetBounds(16, 38, 400, 24)
        $txtName.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
        $txtName.BackColor = $clrDetailBg; $txtName.ForeColor = $clrText
        $txtName.Text = $ValueName
        if ($IsEdit) { $txtName.ReadOnly = $true }
        $dlg.Controls.Add($txtName)

        # Type label
        $lblType = New-Object System.Windows.Forms.Label
        $lblType.Text = "Type: $(& $getTypeName $Kind)"
        $lblType.AutoSize = $true
        $lblType.Location = New-Object System.Drawing.Point(16, 72); $lblType.ForeColor = $clrHint
        $dlg.Controls.Add($lblType)

        # Value data
        $lblData = New-Object System.Windows.Forms.Label
        $lblData.Text = "Value data:"; $lblData.AutoSize = $true
        $lblData.Location = New-Object System.Drawing.Point(16, 96); $lblData.ForeColor = $clrText
        $dlg.Controls.Add($lblData)

        $dataControl = $null

        switch ($Kind) {
            'DWord' {
                $numData = New-Object System.Windows.Forms.NumericUpDown
                $numData.SetBounds(16, 118, 200, 28)
                $numData.Minimum = 0; $numData.Maximum = [uint32]::MaxValue
                $numData.Font = New-Object System.Drawing.Font("Consolas", 10)
                $numData.BackColor = $clrDetailBg; $numData.ForeColor = $clrText
                if ($IsEdit -and $null -ne $CurrentValue) { $numData.Value = [Math]::Min([uint32]::MaxValue, [Math]::Max(0, [decimal]$CurrentValue)) }
                $dlg.Controls.Add($numData)
                $dataControl = $numData
            }
            'QWord' {
                $txtData = New-Object System.Windows.Forms.TextBox
                $txtData.SetBounds(16, 118, 400, 24)
                $txtData.Font = New-Object System.Drawing.Font("Consolas", 10)
                $txtData.BackColor = $clrDetailBg; $txtData.ForeColor = $clrText
                if ($IsEdit -and $null -ne $CurrentValue) { $txtData.Text = [string]$CurrentValue }
                else { $txtData.Text = '0' }
                $dlg.Controls.Add($txtData)
                $dataControl = $txtData
            }
            'Binary' {
                $txtData = New-Object System.Windows.Forms.TextBox
                $txtData.SetBounds(16, 118, 400, 80)
                $txtData.Multiline = $true
                $txtData.Font = New-Object System.Drawing.Font("Consolas", 10)
                $txtData.BackColor = $clrDetailBg; $txtData.ForeColor = $clrText
                $txtData.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
                if ($IsEdit -and $CurrentValue -is [byte[]]) {
                    $txtData.Text = ($CurrentValue | ForEach-Object { '{0:X2}' -f $_ }) -join ' '
                }
                $dlg.Controls.Add($txtData)
                $dataControl = $txtData
            }
            'MultiString' {
                $txtData = New-Object System.Windows.Forms.TextBox
                $txtData.SetBounds(16, 118, 400, 100)
                $txtData.Multiline = $true
                $txtData.Font = New-Object System.Drawing.Font("Consolas", 10)
                $txtData.BackColor = $clrDetailBg; $txtData.ForeColor = $clrText
                $txtData.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
                $txtData.AcceptsReturn = $true
                if ($IsEdit -and $CurrentValue -is [string[]]) {
                    $txtData.Text = $CurrentValue -join "`r`n"
                }
                $dlg.Controls.Add($txtData)
                $dataControl = $txtData
            }
            default {
                # String, ExpandString
                $txtData = New-Object System.Windows.Forms.TextBox
                $txtData.SetBounds(16, 118, 400, 24)
                $txtData.Font = New-Object System.Drawing.Font("Consolas", 10)
                $txtData.BackColor = $clrDetailBg; $txtData.ForeColor = $clrText
                if ($IsEdit -and $null -ne $CurrentValue) { $txtData.Text = [string]$CurrentValue }
                $dlg.Controls.Add($txtData)
                $dataControl = $txtData
            }
        }

        # OK / Cancel buttons
        $btnOk = New-Object System.Windows.Forms.Button; $btnOk.Text = "OK"
        $btnOk.SetBounds(236, 240, 85, 32); $btnOk.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        & $btnStyle $btnOk $clrAccent
        $dlg.Controls.Add($btnOk)

        $btnDlgCancel = New-Object System.Windows.Forms.Button; $btnDlgCancel.Text = "Cancel"
        $btnDlgCancel.SetBounds(330, 240, 85, 32); $btnDlgCancel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $btnDlgCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnDlgCancel.ForeColor = $clrText; $btnDlgCancel.BackColor = $clrFormBg
        $dlg.Controls.Add($btnDlgCancel)

        $btnOk.Add_Click(({
            $vName = $txtName.Text
            if (-not $IsEdit -and [string]::IsNullOrWhiteSpace($vName)) {
                [System.Windows.Forms.MessageBox]::Show("Value name cannot be empty.", "Validation", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                return
            }

            $regKey = $null
            try {
                $regKey = & $openRegKey $path $true
                if ($null -eq $regKey) {
                    [System.Windows.Forms.MessageBox]::Show("Cannot open key for writing.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
                    return
                }

                $writeValue = $null
                switch ($Kind) {
                    'DWord' {
                        $writeValue = [int]$dataControl.Value
                    }
                    'QWord' {
                        $writeValue = [long]$dataControl.Text
                    }
                    'Binary' {
                        $hexStr = $dataControl.Text.Trim()
                        if ($hexStr) {
                            $bytes = $hexStr -split '\s+' | ForEach-Object { [byte]("0x$_") }
                            $writeValue = [byte[]]$bytes
                        } else {
                            $writeValue = [byte[]]@()
                        }
                    }
                    'MultiString' {
                        $lines = $dataControl.Text -split "`r`n"
                        $writeValue = [string[]]$lines
                    }
                    default {
                        $writeValue = $dataControl.Text
                    }
                }

                $regKey.SetValue($vName, $writeValue, $Kind)
                $action = if ($IsEdit) { "Modified" } else { "Created" }
                Write-Log "$action registry value: $path\\$vName ($(& $getTypeName $Kind))"

                $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $dlg.Close()
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to write value: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
            } finally {
                $parts = $path -split '\\', 2
                if ($null -ne $regKey -and $parts.Count -gt 1) { $regKey.Dispose() }
            }
        }).GetNewClosure())
        $btnDlgCancel.Add_Click(({ $dlg.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $dlg.Close() }).GetNewClosure())
        $dlg.AcceptButton = $btnOk; $dlg.CancelButton = $btnDlgCancel

        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            # Refresh the values grid
            & $navigateToPath $path
        }
        $dlg.Dispose()
    }.GetNewClosure()

    # -----------------------------------------------------------------------
    # Search (Ctrl+F)
    # -----------------------------------------------------------------------

    $regFuncs.SearchNext = {
        param([System.Windows.Forms.TreeNode]$StartNode, [string]$SearchText, [bool]$InKeys, [bool]$InValues, [bool]$InData)

        # Simple depth-first search starting from the next sibling/child of StartNode
        # We'll search registry directly rather than relying on tree expansion
        $startPath = if ($StartNode -and $StartNode.Tag) { [string]$StartNode.Tag } else { 'HKCU' }

        # Search from the current key's children first, then move to next siblings
        $regKey = $null
        try {
            $regKey = & $openRegKey $startPath $false
            if ($null -eq $regKey) { return $false }

            # Check values of current key
            if ($InValues -or $InData) {
                $valueNames = @()
                try { $valueNames = $regKey.GetValueNames() } catch { }
                foreach ($vName in $valueNames) {
                    if ($InValues -and $vName -like "*$SearchText*") {
                        & $navigateToPath $startPath
                        & $statusFunc "Found value name: $vName in $startPath"
                        return $true
                    }
                    if ($InData) {
                        try {
                            $val = $regKey.GetValue($vName)
                            if ($null -ne $val -and $val.ToString() -like "*$SearchText*") {
                                & $navigateToPath $startPath
                                & $statusFunc "Found in data of '$vName' in $startPath"
                                return $true
                            }
                        } catch { }
                    }
                }
            }

            # Search subkeys
            $subNames = @()
            try { $subNames = $regKey.GetSubKeyNames() | Sort-Object } catch { }
            foreach ($subName in $subNames) {
                $childPath = "$startPath\$subName"

                # Check key name
                if ($InKeys -and $subName -like "*$SearchText*") {
                    & $navigateToPath $childPath
                    & $statusFunc "Found key: $childPath"
                    return $true
                }

                # Recurse into subkey (limited depth to avoid hanging)
                $childKey = $null
                try {
                    $childKey = $regKey.OpenSubKey($subName, $false)
                    if ($null -ne $childKey) {
                        # Check child values
                        if ($InValues -or $InData) {
                            $childVals = @()
                            try { $childVals = $childKey.GetValueNames() } catch { }
                            foreach ($cv in $childVals) {
                                if ($InValues -and $cv -like "*$SearchText*") {
                                    & $navigateToPath $childPath
                                    & $statusFunc "Found value name: $cv in $childPath"
                                    return $true
                                }
                                if ($InData) {
                                    try {
                                        $cvVal = $childKey.GetValue($cv)
                                        if ($null -ne $cvVal -and $cvVal.ToString() -like "*$SearchText*") {
                                            & $navigateToPath $childPath
                                            & $statusFunc "Found in data of '$cv' in $childPath"
                                            return $true
                                        }
                                    } catch { }
                                }
                            }
                        }
                    }
                } catch { } finally {
                    if ($null -ne $childKey) { $childKey.Dispose() }
                }
            }
        } catch { } finally {
            $parts = $startPath -split '\\', 2
            if ($null -ne $regKey -and $parts.Count -gt 1) { $regKey.Dispose() }
        }

        return $false
    }.GetNewClosure()

    $regFuncs.ShowSearchDialog = {
        $dlg = New-Object System.Windows.Forms.Form
        $dlg.Text = "Find"; $dlg.Size = New-Object System.Drawing.Size(400, 200)
        $dlg.MinimumSize = $dlg.Size; $dlg.MaximumSize = $dlg.Size
        $dlg.StartPosition = "CenterParent"
        $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false; $dlg.ShowInTaskbar = $false
        $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9.5); $dlg.BackColor = $clrFormBg

        $lblFind = New-Object System.Windows.Forms.Label
        $lblFind.Text = "Find what:"; $lblFind.AutoSize = $true
        $lblFind.Location = New-Object System.Drawing.Point(16, 18); $lblFind.ForeColor = $clrText
        $dlg.Controls.Add($lblFind)

        $txtFind = New-Object System.Windows.Forms.TextBox
        $txtFind.SetBounds(16, 40, 350, 24)
        $txtFind.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
        $txtFind.BackColor = $clrDetailBg; $txtFind.ForeColor = $clrText
        $txtFind.Text = $regState.SearchText
        $dlg.Controls.Add($txtFind)

        $chkKeys = New-Object System.Windows.Forms.CheckBox
        $chkKeys.Text = "Keys"; $chkKeys.AutoSize = $true
        $chkKeys.Location = New-Object System.Drawing.Point(16, 76); $chkKeys.Checked = $regState.SearchKeys
        $chkKeys.ForeColor = $clrText; $chkKeys.BackColor = $clrFormBg
        $dlg.Controls.Add($chkKeys)

        $chkValues = New-Object System.Windows.Forms.CheckBox
        $chkValues.Text = "Values"; $chkValues.AutoSize = $true
        $chkValues.Location = New-Object System.Drawing.Point(100, 76); $chkValues.Checked = $regState.SearchValues
        $chkValues.ForeColor = $clrText; $chkValues.BackColor = $clrFormBg
        $dlg.Controls.Add($chkValues)

        $chkData = New-Object System.Windows.Forms.CheckBox
        $chkData.Text = "Data"; $chkData.AutoSize = $true
        $chkData.Location = New-Object System.Drawing.Point(190, 76); $chkData.Checked = $regState.SearchData
        $chkData.ForeColor = $clrText; $chkData.BackColor = $clrFormBg
        $dlg.Controls.Add($chkData)

        $btnFind = New-Object System.Windows.Forms.Button; $btnFind.Text = "Find Next"
        $btnFind.SetBounds(200, 115, 80, 32); $btnFind.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        & $btnStyle $btnFind $clrAccent
        $dlg.Controls.Add($btnFind)

        $btnFindCancel = New-Object System.Windows.Forms.Button; $btnFindCancel.Text = "Cancel"
        $btnFindCancel.SetBounds(288, 115, 80, 32); $btnFindCancel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $btnFindCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnFindCancel.ForeColor = $clrText; $btnFindCancel.BackColor = $clrFormBg
        $dlg.Controls.Add($btnFindCancel)

        $btnFind.Add_Click(({
            $searchText = $txtFind.Text
            if ([string]::IsNullOrWhiteSpace($searchText)) { return }

            $regState.SearchText = $searchText
            $regState.SearchKeys = $chkKeys.Checked
            $regState.SearchValues = $chkValues.Checked
            $regState.SearchData = $chkData.Checked

            $dlg.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            & $statusFunc "Searching for '$searchText'..."

            # Start search from current node (or first root)
            $startNode = $treeReg.SelectedNode
            if (-not $startNode) { $startNode = $treeReg.Nodes[0] }

            $found = & $regFuncs.SearchNext $startNode $searchText $chkKeys.Checked $chkValues.Checked $chkData.Checked
            $dlg.Cursor = [System.Windows.Forms.Cursors]::Default

            if (-not $found) {
                & $statusFunc "Not found: '$searchText'"
                [System.Windows.Forms.MessageBox]::Show("Finished searching. No match found.", "Find", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            }
        }).GetNewClosure())

        $btnFindCancel.Add_Click(({ $dlg.Close() }).GetNewClosure())
        $dlg.AcceptButton = $btnFind; $dlg.CancelButton = $btnFindCancel
        $dlg.ShowDialog() | Out-Null; $dlg.Dispose()
    }.GetNewClosure()

    # -----------------------------------------------------------------------
    # Keyboard shortcuts
    # -----------------------------------------------------------------------

    $tabPage.Add_Enter(({
        param($s, $e)
        $form = $s.FindForm()
        if ($form) {
            $form.KeyPreview = $true
        }
    }).GetNewClosure())

    # We need to handle Ctrl+F at the form level via the tree/grid KeyDown
    $treeReg.Add_KeyDown(({
        param($s, $e)
        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::F) {
            $e.SuppressKeyPress = $true
            & $regFuncs.ShowSearchDialog
        }
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::F5) {
            $e.SuppressKeyPress = $true
            $ctxTreeRefresh.PerformClick()
        }
    }).GetNewClosure())

    $gridValues.Add_KeyDown(({
        param($s, $e)
        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::F) {
            $e.SuppressKeyPress = $true
            & $regFuncs.ShowSearchDialog
        }
    }).GetNewClosure())

    # -----------------------------------------------------------------------
    # Assemble layout (dock order matters)
    # -----------------------------------------------------------------------

    $tabPage.Controls.Add($splitMain)
    $splitMain.BringToFront()

    $tabPage.Controls.Add($pnlAddrSep)
    $pnlAddrSep.SendToBack()

    $tabPage.Controls.Add($pnlAddress)
    $pnlAddress.SendToBack()

    # Defer SplitContainer distance until it has valid size
    $regSplitFlag = @{ Done = $false }
    $splitMain.Add_SizeChanged(({
        if (-not $regSplitFlag.Done -and $splitMain.Width -gt ($splitMain.Panel1MinSize + $splitMain.Panel2MinSize + $splitMain.SplitterWidth)) {
            $regSplitFlag.Done = $true
            $splitMain.SplitterDistance = [Math]::Max($splitMain.Panel1MinSize, [int]($splitMain.Width * 0.3))
        }
    }).GetNewClosure())

    # Set initial address bar text
    $txtPath.Text = "HKCU"
}
