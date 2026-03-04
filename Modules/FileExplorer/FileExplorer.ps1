<#
.SYNOPSIS
    File Explorer module for MMC-Alt.

.DESCRIPTION
    Admin-oriented file browser that shows hidden files and file extensions by default,
    supports sorting by type, and can browse inside archives via 7-Zip (7z.exe).
    Does not require elevation.
#>

function Initialize-FileExplorer {
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
    $clrTreeBg   = $theme.TreeBg
    $clrGridAlt  = $theme.GridAlt
    $clrGridLine = $theme.GridLine
    $clrErrText  = $theme.ErrText
    $clrWarnText = $theme.WarnText
    $isDark      = $theme.DarkMode

    # -----------------------------------------------------------------------
    # 7-Zip detection
    # -----------------------------------------------------------------------

    $feState = @{
        SevenZipPath = $null
        CurrentPath = ''
        InArchive = $false
        ArchivePath = ''
        ArchiveSubPath = ''
        SplitDone = $false
    }
    $possiblePaths = @(
        "${env:ProgramFiles}\7-Zip\7z.exe"
        "${env:ProgramFiles(x86)}\7-Zip\7z.exe"
        "C:\Program Files\7-Zip\7z.exe"
    )
    foreach ($p in $possiblePaths) {
        if (Test-Path -LiteralPath $p) { $feState.SevenZipPath = $p; break }
    }

    # Archive extensions that 7-Zip can browse
    $archiveExts = @('.zip', '.7z', '.rar', '.tar', '.gz', '.bz2', '.xz', '.cab', '.iso', '.msi', '.wim')

    # Icon strings -- emoji are in the supplementary plane (above U+FFFF),
    # so [char] cannot hold them. ConvertFromUtf32 produces the surrogate pair.
    $iconFolder  = [char]::ConvertFromUtf32(0x1F4C1)
    $iconFile    = [char]::ConvertFromUtf32(0x1F4C4)
    $iconArchive = [char]::ConvertFromUtf32(0x1F4E6)

    $isArchive = {
        param([string]$FileName)
        $ext = [System.IO.Path]::GetExtension($FileName).ToLower()
        return ($ext -in $archiveExts)
    }.GetNewClosure()

    # -----------------------------------------------------------------------
    # Helper: format file size
    # -----------------------------------------------------------------------

    $formatSize = {
        param([long]$Bytes)
        if ($Bytes -lt 0) { return '' }
        if ($Bytes -ge 1GB) { return '{0:N2} GB' -f ($Bytes / 1GB) }
        if ($Bytes -ge 1MB) { return '{0:N2} MB' -f ($Bytes / 1MB) }
        if ($Bytes -ge 1KB) { return '{0:N2} KB' -f ($Bytes / 1KB) }
        return "$Bytes bytes"
    }

    # -----------------------------------------------------------------------
    # Address bar (Dock:Top)
    # -----------------------------------------------------------------------

    $pnlAddress = New-Object System.Windows.Forms.Panel
    $pnlAddress.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlAddress.Height = 36
    $pnlAddress.BackColor = $clrPanelBg
    $pnlAddress.Padding = New-Object System.Windows.Forms.Padding(6, 4, 6, 4)

    $btnUp = New-Object System.Windows.Forms.Button
    $btnUp.Text = [char]0x2191  # up arrow
    $btnUp.Size = New-Object System.Drawing.Size(30, 26)
    $btnUp.Location = New-Object System.Drawing.Point(6, 4)
    $btnUp.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    & $btnStyle $btnUp $clrAccent
    $pnlAddress.Controls.Add($btnUp)

    $txtFEPath = New-Object System.Windows.Forms.TextBox
    $txtFEPath.Location = New-Object System.Drawing.Point(42, 5)
    $txtFEPath.Height = 24
    $txtFEPath.Width = 300
    $txtFEPath.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $txtFEPath.Font = New-Object System.Drawing.Font("Consolas", 9.5)
    $txtFEPath.BackColor = $clrDetailBg; $txtFEPath.ForeColor = $clrText
    $txtFEPath.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $pnlAddress.Controls.Add($txtFEPath)

    $btnFEGo = New-Object System.Windows.Forms.Button
    $btnFEGo.Text = "Go"
    $btnFEGo.Size = New-Object System.Drawing.Size(50, 26)
    $btnFEGo.Location = New-Object System.Drawing.Point(350, 4)
    $btnFEGo.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnFEGo.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    & $btnStyle $btnFEGo $clrAccent
    $pnlAddress.Controls.Add($btnFEGo)

    $pnlAddrSep = New-Object System.Windows.Forms.Panel
    $pnlAddrSep.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlAddrSep.Height = 1; $pnlAddrSep.BackColor = $clrSepLine

    # -----------------------------------------------------------------------
    # SplitContainer (Directory tree left, Files grid right)
    # -----------------------------------------------------------------------

    $splitFE = New-Object System.Windows.Forms.SplitContainer
    $splitFE.Dock = [System.Windows.Forms.DockStyle]::Fill
    $splitFE.Orientation = [System.Windows.Forms.Orientation]::Vertical
    $splitFE.SplitterWidth = 6
    $splitFE.BackColor = $clrSepLine
    $splitFE.Panel1.BackColor = $clrPanelBg
    $splitFE.Panel2.BackColor = $clrPanelBg
    $splitFE.Panel1MinSize = 100
    $splitFE.Panel2MinSize = 100

    # -----------------------------------------------------------------------
    # Directory TreeView (left)
    # -----------------------------------------------------------------------

    $treeDirs = New-Object System.Windows.Forms.TreeView
    $treeDirs.Dock = [System.Windows.Forms.DockStyle]::Fill
    $treeDirs.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
    $treeDirs.BackColor = $clrTreeBg; $treeDirs.ForeColor = $clrText
    $treeDirs.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $treeDirs.FullRowSelect = $true; $treeDirs.HideSelection = $false
    $treeDirs.ShowLines = $true; $treeDirs.ShowPlusMinus = $true; $treeDirs.ShowRootLines = $true
    $treeDirs.PathSeparator = '\'
    if ($isDark) { $treeDirs.LineColor = [System.Drawing.Color]::FromArgb(80, 80, 80) }
    & $dblBuffer $treeDirs

    # Populate drive roots
    $drives = [System.IO.DriveInfo]::GetDrives() | Where-Object { $_.IsReady }
    foreach ($drive in $drives) {
        $label = if ($drive.VolumeLabel) { "$($drive.Name.TrimEnd('\')) ($($drive.VolumeLabel))" } else { $drive.Name.TrimEnd('\') }
        $driveNode = New-Object System.Windows.Forms.TreeNode($label)
        $driveNode.Tag = $drive.Name
        $driveNode.Name = $drive.Name
        # Add dummy for expand arrow
        $dummy = New-Object System.Windows.Forms.TreeNode("__dummy__")
        $driveNode.Nodes.Add($dummy) | Out-Null
        $treeDirs.Nodes.Add($driveNode) | Out-Null
    }

    # Quick-access nodes
    $specialDirs = @(
        @{ Name = "Desktop";   Path = [Environment]::GetFolderPath('Desktop') }
        @{ Name = "Documents"; Path = [Environment]::GetFolderPath('MyDocuments') }
        @{ Name = "Downloads"; Path = (Join-Path $env:USERPROFILE 'Downloads') }
        @{ Name = "Temp";      Path = $env:TEMP }
    )
    foreach ($sd in $specialDirs) {
        if (Test-Path -LiteralPath $sd.Path) {
            $node = New-Object System.Windows.Forms.TreeNode($sd.Name)
            $node.Tag = $sd.Path
            $node.Name = $sd.Name
            $node.ForeColor = $clrAccent
            $dummy = New-Object System.Windows.Forms.TreeNode("__dummy__")
            $node.Nodes.Add($dummy) | Out-Null
            $treeDirs.Nodes.Add($node) | Out-Null
        }
    }

    # BeforeExpand: lazy-load subdirectories
    $treeDirs.Add_BeforeExpand(({
        param($s, $e)
        $node = $e.Node
        if ($node.Nodes.Count -eq 1 -and $node.Nodes[0].Text -eq '__dummy__') {
            $node.Nodes.Clear()
            $dirPath = [string]$node.Tag

            try {
                $subDirs = [System.IO.Directory]::GetDirectories($dirPath) | Sort-Object
                foreach ($subDir in $subDirs) {
                    $name = [System.IO.Path]::GetFileName($subDir)
                    $childNode = New-Object System.Windows.Forms.TreeNode($name)
                    $childNode.Tag = $subDir
                    $childNode.Name = $name

                    # Check for subdirectories (for expand arrow)
                    try {
                        $hasChildren = ([System.IO.Directory]::GetDirectories($subDir)).Count -gt 0
                        if ($hasChildren) {
                            $dummy = New-Object System.Windows.Forms.TreeNode("__dummy__")
                            $childNode.Nodes.Add($dummy) | Out-Null
                        }
                    } catch {
                        # Access denied -- still show expand arrow
                        $dummy = New-Object System.Windows.Forms.TreeNode("__dummy__")
                        $childNode.Nodes.Add($dummy) | Out-Null
                    }

                    # Highlight hidden directories
                    try {
                        $attrs = [System.IO.File]::GetAttributes($subDir)
                        if ($attrs -band [System.IO.FileAttributes]::Hidden) {
                            $childNode.ForeColor = $clrHint
                        }
                    } catch { }

                    $node.Nodes.Add($childNode) | Out-Null
                }
            } catch [System.UnauthorizedAccessException] {
                $errNode = New-Object System.Windows.Forms.TreeNode("(Access Denied)")
                $errNode.ForeColor = $clrErrText
                $node.Nodes.Add($errNode) | Out-Null
            } catch {
                $errNode = New-Object System.Windows.Forms.TreeNode("(Error: $($_.Exception.Message))")
                $errNode.ForeColor = $clrErrText
                $node.Nodes.Add($errNode) | Out-Null
            }
        }
    }).GetNewClosure())

    $splitFE.Panel1.Controls.Add($treeDirs)

    # -----------------------------------------------------------------------
    # Files DataGridView (right)
    # -----------------------------------------------------------------------

    $gridFiles = & $newGrid
    $gridFiles.AllowUserToOrderColumns = $true
    $gridFiles.AllowUserToResizeColumns = $true

    $colIcon = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colIcon.HeaderText = ""; $colIcon.DataPropertyName = "Icon"; $colIcon.Width = 30
    $colIcon.DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
    $colFName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colFName.HeaderText = "Name"; $colFName.DataPropertyName = "Name"; $colFName.Width = 280
    $colFExt = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colFExt.HeaderText = "Type"; $colFExt.DataPropertyName = "Type"; $colFExt.Width = 80
    $colFSize = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colFSize.HeaderText = "Size"; $colFSize.DataPropertyName = "Size"; $colFSize.Width = 90
    $colFSize.DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleRight
    $colFDate = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colFDate.HeaderText = "Modified"; $colFDate.DataPropertyName = "Modified"; $colFDate.Width = 140
    $colFAttr = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colFAttr.HeaderText = "Attributes"; $colFAttr.DataPropertyName = "Attributes"; $colFAttr.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill

    $gridFiles.Columns.Add($colIcon) | Out-Null
    $gridFiles.Columns.Add($colFName) | Out-Null
    $gridFiles.Columns.Add($colFExt) | Out-Null
    $gridFiles.Columns.Add($colFSize) | Out-Null
    $gridFiles.Columns.Add($colFDate) | Out-Null
    $gridFiles.Columns.Add($colFAttr) | Out-Null

    $filesTable = New-Object System.Data.DataTable
    [void]$filesTable.Columns.Add("Icon", [string])
    [void]$filesTable.Columns.Add("Name", [string])
    [void]$filesTable.Columns.Add("Type", [string])
    [void]$filesTable.Columns.Add("Size", [string])
    [void]$filesTable.Columns.Add("Modified", [string])
    [void]$filesTable.Columns.Add("Attributes", [string])
    [void]$filesTable.Columns.Add("_FullPath", [string])
    [void]$filesTable.Columns.Add("_IsDir", [bool])
    [void]$filesTable.Columns.Add("_SortSize", [long])
    $gridFiles.DataSource = $filesTable

    # Color hidden files and directories
    $gridFiles.Add_CellFormatting(({
        param($s, $e)
        if ($e.RowIndex -lt 0 -or $e.ColumnIndex -ne 1) { return }
        $view = $filesTable.DefaultView
        if ($e.RowIndex -ge $view.Count) { return }
        $attrs = [string]$view[$e.RowIndex]["Attributes"]
        if ($attrs -like '*H*') { $e.CellStyle.ForeColor = $clrHint }
        $isDir = [bool]$view[$e.RowIndex]["_IsDir"]
        if ($isDir) { $e.CellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold) }
    }).GetNewClosure())

    $splitFE.Panel2.Controls.Add($gridFiles)

    # -----------------------------------------------------------------------
    # Load directory contents
    # -----------------------------------------------------------------------

    $loadDirectory = {
        param([string]$DirPath)
        $filesTable.Clear()
        $feState.CurrentPath = $DirPath
        $feState.InArchive = $false
        $txtFEPath.Text = $DirPath

        & $statusFunc "Loading $DirPath..."
        [System.Windows.Forms.Application]::DoEvents()

        $fileCount = 0
        $dirCount = 0

        try {
            # Directories first
            $dirs = @()
            try { $dirs = [System.IO.Directory]::GetDirectories($DirPath) | Sort-Object } catch { }
            foreach ($d in $dirs) {
                $dirCount++
                $name = [System.IO.Path]::GetFileName($d)
                $attrs = ''
                $modified = ''
                try {
                    $di = New-Object System.IO.DirectoryInfo($d)
                    $attrList = @()
                    if ($di.Attributes -band [System.IO.FileAttributes]::Hidden) { $attrList += 'H' }
                    if ($di.Attributes -band [System.IO.FileAttributes]::System) { $attrList += 'S' }
                    if ($di.Attributes -band [System.IO.FileAttributes]::ReadOnly) { $attrList += 'R' }
                    $attrs = $attrList -join ''
                    $modified = $di.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
                } catch { }

                $row = $filesTable.NewRow()
                $row["Icon"] = $iconFolder
                $row["Name"] = $name
                $row["Type"] = "<DIR>"
                $row["Size"] = ""
                $row["Modified"] = $modified
                $row["Attributes"] = $attrs
                $row["_FullPath"] = $d
                $row["_IsDir"] = $true
                $row["_SortSize"] = -1
                $filesTable.Rows.Add($row)
            }

            # Files (sorted by type then name)
            $fileInfos = @()
            try {
                $dirInfo = New-Object System.IO.DirectoryInfo($DirPath)
                $fileInfos = @($dirInfo.GetFiles() | Sort-Object Extension, Name)
            } catch { $fileInfos = @() }
            foreach ($fi in $fileInfos) {
                $fileCount++
                $name = $fi.Name
                $ext = $fi.Extension.ToLower()
                $size = ''
                $sortSize = [long]0
                $modified = ''
                $attrs = ''

                try {
                    $size = & $formatSize $fi.Length
                    $sortSize = $fi.Length
                    $modified = $fi.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
                    $attrList = @()
                    if ($fi.Attributes -band [System.IO.FileAttributes]::Hidden) { $attrList += 'H' }
                    if ($fi.Attributes -band [System.IO.FileAttributes]::System) { $attrList += 'S' }
                    if ($fi.Attributes -band [System.IO.FileAttributes]::ReadOnly) { $attrList += 'R' }
                    if ($fi.Attributes -band [System.IO.FileAttributes]::Archive) { $attrList += 'A' }
                    $attrs = $attrList -join ''
                } catch { }

                $icon = if (& $isArchive $name) { $iconArchive } else { $iconFile }

                $row = $filesTable.NewRow()
                $row["Icon"] = $icon
                $row["Name"] = $name
                $row["Type"] = if ($ext) { $ext } else { "" }
                $row["Size"] = $size
                $row["Modified"] = $modified
                $row["Attributes"] = $attrs
                $row["_FullPath"] = $fi.FullName
                $row["_IsDir"] = $false
                $row["_SortSize"] = $sortSize
                $filesTable.Rows.Add($row)
            }

            & $statusFunc "$DirPath  |  $dirCount folder(s), $fileCount file(s)"

        } catch [System.UnauthorizedAccessException] {
            $row = $filesTable.NewRow()
            $row["Icon"] = ""; $row["Name"] = "(Access Denied)"; $row["Type"] = ""; $row["Size"] = ""
            $row["Modified"] = ""; $row["Attributes"] = ""; $row["_FullPath"] = ""; $row["_IsDir"] = $false; $row["_SortSize"] = 0
            $filesTable.Rows.Add($row)
            & $statusFunc "Access denied: $DirPath"
        } catch {
            & $statusFunc "Error: $($_.Exception.Message)"
        }
    }.GetNewClosure()

    # -----------------------------------------------------------------------
    # Load archive contents via 7-Zip
    # -----------------------------------------------------------------------

    $loadArchive = {
        param([string]$ArchivePath, [string]$SubPath)

        if (-not $feState.SevenZipPath) {
            [System.Windows.Forms.MessageBox]::Show("7-Zip not found. Install 7-Zip to browse archives.", "7-Zip Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            return
        }

        $filesTable.Clear()
        $feState.InArchive = $true
        $feState.ArchivePath = $ArchivePath
        $feState.ArchiveSubPath = $SubPath

        $displayPath = if ($SubPath) { "$ArchivePath\$SubPath" } else { "$ArchivePath\" }
        $txtFEPath.Text = $displayPath
        $feState.CurrentPath = $displayPath

        & $statusFunc "Listing archive: $ArchivePath..."
        [System.Windows.Forms.Application]::DoEvents()

        try {
            $output = & $feState.SevenZipPath l "$ArchivePath" -slt 2>&1

            # Parse 7z technical listing format
            $entries = @()
            $currentEntry = $null
            foreach ($line in $output) {
                $lineStr = [string]$line
                if ($lineStr -match '^Path = (.+)$') {
                    if ($currentEntry -and $currentEntry.Path) { $entries += $currentEntry }
                    $currentEntry = @{ Path = $Matches[1]; Size = 0; Modified = ''; IsDir = $false; Attrs = '' }
                } elseif ($currentEntry) {
                    if ($lineStr -match '^Size = (\d+)$') { $currentEntry.Size = [long]$Matches[1] }
                    if ($lineStr -match '^Modified = (.+)$') { $currentEntry.Modified = $Matches[1].Trim() }
                    if ($lineStr -match '^Folder = \+$') { $currentEntry.IsDir = $true }
                    if ($lineStr -match '^Attributes = (.+)$') { $currentEntry.Attrs = $Matches[1].Trim() }
                }
            }
            if ($currentEntry -and $currentEntry.Path) { $entries += $currentEntry }

            # Filter to current subdirectory level
            $targetPrefix = if ($SubPath) { "$SubPath\" } else { '' }
            $dirs = @{}
            $filesInDir = @()

            foreach ($entry in $entries) {
                $entryPath = $entry.Path -replace '/', '\'

                # Skip entries not under our subpath
                if ($targetPrefix -and -not $entryPath.StartsWith($targetPrefix, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
                if (-not $targetPrefix -and $entryPath -eq ([System.IO.Path]::GetFileName($ArchivePath))) { continue }

                $relativePath = if ($targetPrefix) { $entryPath.Substring($targetPrefix.Length) } else { $entryPath }
                if ([string]::IsNullOrEmpty($relativePath)) { continue }

                # Check if this is a direct child or nested
                $slashIdx = $relativePath.IndexOf('\')
                if ($slashIdx -ge 0) {
                    # Nested: just record the immediate directory
                    $dirName = $relativePath.Substring(0, $slashIdx)
                    $dirs[$dirName] = $true
                } elseif ($entry.IsDir) {
                    $dirs[$relativePath] = $true
                } else {
                    $filesInDir += @{
                        Name = $relativePath
                        Size = $entry.Size
                        Modified = $entry.Modified
                        Attrs = $entry.Attrs
                    }
                }
            }

            # Add directory entries
            foreach ($dirName in ($dirs.Keys | Sort-Object)) {
                $row = $filesTable.NewRow()
                $row["Icon"] = $iconFolder
                $row["Name"] = $dirName
                $row["Type"] = "<DIR>"
                $row["Size"] = ""
                $row["Modified"] = ""
                $row["Attributes"] = ""
                $row["_FullPath"] = if ($SubPath) { "$SubPath\$dirName" } else { $dirName }
                $row["_IsDir"] = $true
                $row["_SortSize"] = -1
                $filesTable.Rows.Add($row)
            }

            # Add file entries
            $filesInDir = $filesInDir | Sort-Object { [System.IO.Path]::GetExtension($_.Name).ToLower() }, { $_.Name }
            foreach ($f in $filesInDir) {
                $ext = [System.IO.Path]::GetExtension($f.Name).ToLower()
                $row = $filesTable.NewRow()
                $row["Icon"] = $iconFile
                $row["Name"] = $f.Name
                $row["Type"] = if ($ext) { $ext } else { "" }
                $row["Size"] = & $formatSize $f.Size
                $row["Modified"] = $f.Modified
                $row["Attributes"] = $f.Attrs
                $row["_FullPath"] = if ($SubPath) { "$SubPath\$($f.Name)" } else { $f.Name }
                $row["_IsDir"] = $false
                $row["_SortSize"] = $f.Size
                $filesTable.Rows.Add($row)
            }

            & $statusFunc "Archive: $ArchivePath  |  $($dirs.Count) folder(s), $($filesInDir.Count) file(s)"

        } catch {
            & $statusFunc "Error reading archive: $($_.Exception.Message)"
            $row = $filesTable.NewRow()
            $row["Icon"] = ""; $row["Name"] = "(Error reading archive)"; $row["Type"] = ""; $row["Size"] = ""
            $row["Modified"] = ""; $row["Attributes"] = $_.Exception.Message; $row["_FullPath"] = ""; $row["_IsDir"] = $false; $row["_SortSize"] = 0
            $filesTable.Rows.Add($row)
        }
    }.GetNewClosure()

    # -----------------------------------------------------------------------
    # Tree selection: load directory
    # -----------------------------------------------------------------------

    $treeDirs.Add_AfterSelect(({
        param($s, $e)
        $node = $e.Node
        if (-not $node -or -not $node.Tag) { return }
        & $loadDirectory ([string]$node.Tag)
    }).GetNewClosure())

    # -----------------------------------------------------------------------
    # Grid double-click: navigate into directory or archive
    # -----------------------------------------------------------------------

    $gridFiles.Add_CellDoubleClick(({
        param($s, $e)
        if ($e.RowIndex -lt 0) { return }
        $view = $filesTable.DefaultView
        if ($e.RowIndex -ge $view.Count) { return }

        $rowView = $view[$e.RowIndex]
        $fullPath = [string]$rowView["_FullPath"]
        $isDir = [bool]$rowView["_IsDir"]
        $name = [string]$rowView["Name"]

        if ($feState.InArchive) {
            # Inside an archive: navigate into subdirectory
            if ($isDir) {
                & $loadArchive $feState.ArchivePath $fullPath
            }
            return
        }

        if ($isDir) {
            # Navigate into directory
            & $loadDirectory $fullPath
            # Also expand tree to this path
        } elseif (& $isArchive $name) {
            # Open archive
            & $loadArchive $fullPath ''
        }
    }).GetNewClosure())

    # -----------------------------------------------------------------------
    # Navigation: address bar + Up button
    # -----------------------------------------------------------------------

    $navigateFE = {
        param([string]$Path)
        $Path = $Path.Trim()
        if ([string]::IsNullOrWhiteSpace($Path)) { return }

        if (Test-Path -LiteralPath $Path -PathType Container) {
            & $loadDirectory $Path
        } elseif (Test-Path -LiteralPath $Path -PathType Leaf) {
            if (& $isArchive $Path) {
                & $loadArchive $Path ''
            } else {
                # Navigate to parent and highlight the file
                $parent = [System.IO.Path]::GetDirectoryName($Path)
                if ($parent) { & $loadDirectory $parent }
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Path not found: $Path", "Navigation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        }
    }.GetNewClosure()

    $txtFEPath.Add_KeyDown(({
        param($s, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $e.SuppressKeyPress = $true
            & $navigateFE $txtFEPath.Text
        }
    }).GetNewClosure())

    $btnFEGo.Add_Click(({ & $navigateFE $txtFEPath.Text }).GetNewClosure())

    $btnUp.Add_Click(({
        if ($feState.InArchive) {
            if ($feState.ArchiveSubPath) {
                # Go up within archive
                $parentSub = [System.IO.Path]::GetDirectoryName($feState.ArchiveSubPath)
                if ($parentSub) {
                    & $loadArchive $feState.ArchivePath $parentSub
                } else {
                    & $loadArchive $feState.ArchivePath ''
                }
            } else {
                # Exit archive, go to containing directory
                $parentDir = [System.IO.Path]::GetDirectoryName($feState.ArchivePath)
                if ($parentDir) { & $loadDirectory $parentDir }
            }
        } else {
            $current = $feState.CurrentPath
            if ($current) {
                $parent = [System.IO.Path]::GetDirectoryName($current)
                if ($parent) { & $loadDirectory $parent }
            }
        }
    }).GetNewClosure())

    # -----------------------------------------------------------------------
    # Context menu: Files grid
    # -----------------------------------------------------------------------

    $ctxFiles = New-Object System.Windows.Forms.ContextMenuStrip
    if ($isDark -and $script:DarkRenderer) { $ctxFiles.Renderer = $script:DarkRenderer }
    $ctxFiles.BackColor = $clrPanelBg; $ctxFiles.ForeColor = $clrText

    $ctxCopyPath = New-Object System.Windows.Forms.ToolStripMenuItem("Copy Full Path")
    $ctxCopyPath.ForeColor = $clrText
    $ctxCopyPath.Add_Click(({
        if ($gridFiles.CurrentRow) {
            $fp = $gridFiles.CurrentRow.Cells["_FullPath"].Value
            if (-not $fp) {
                $view = $filesTable.DefaultView
                $idx = $gridFiles.CurrentRow.Index
                if ($idx -lt $view.Count) { $fp = [string]$view[$idx]["_FullPath"] }
            }
            if ($fp) { [System.Windows.Forms.Clipboard]::SetText([string]$fp) }
        }
    }).GetNewClosure())
    $ctxFiles.Items.Add($ctxCopyPath) | Out-Null

    $ctxOpenLocation = New-Object System.Windows.Forms.ToolStripMenuItem("Open in Windows Explorer")
    $ctxOpenLocation.ForeColor = $clrText
    $ctxOpenLocation.Add_Click(({
        if (-not $feState.InArchive -and $feState.CurrentPath) {
            Start-Process explorer.exe -ArgumentList @("`"$($feState.CurrentPath)`"")
        }
    }).GetNewClosure())
    $ctxFiles.Items.Add($ctxOpenLocation) | Out-Null

    $gridFiles.ContextMenuStrip = $ctxFiles

    # -----------------------------------------------------------------------
    # Assemble layout
    # -----------------------------------------------------------------------

    $tabPage.Controls.Add($splitFE)
    $splitFE.BringToFront()

    $tabPage.Controls.Add($pnlAddrSep)
    $pnlAddrSep.SendToBack()

    $tabPage.Controls.Add($pnlAddress)
    $pnlAddress.SendToBack()

    # Defer SplitterDistance until the control has a valid size
    $splitFE.Add_SizeChanged(({
        if (-not $feState.SplitDone -and $splitFE.Width -gt ($splitFE.Panel1MinSize + $splitFE.Panel2MinSize + $splitFE.SplitterWidth)) {
            $feState.SplitDone = $true
            $splitFE.SplitterDistance = [Math]::Max($splitFE.Panel1MinSize, [int]($splitFE.Width * 0.25))
        }
    }).GetNewClosure())

    # Default to user profile
    $defaultPath = $env:USERPROFILE
    if (Test-Path -LiteralPath $defaultPath) { & $loadDirectory $defaultPath }
}
