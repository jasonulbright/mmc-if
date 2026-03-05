<#
.SYNOPSIS
    Device Info module for MMC-If.

.DESCRIPTION
    View system hardware and device information via WMI/CIM classes.
    Does not require Device Manager (mmc.exe) or elevation.
#>

function Initialize-DeviceInfo {
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
    $clrGridAlt  = $theme.GridAlt
    $clrGridLine = $theme.GridLine
    $isDark      = $theme.DarkMode

    # -----------------------------------------------------------------------
    # Category definitions (ordered sections to query)
    # -----------------------------------------------------------------------

    $categories = [ordered]@{
        'System' = @{
            Classes = @(
                @{ Class = 'Win32_ComputerSystem'; Props = @('Name','Domain','Manufacturer','Model','SystemType','TotalPhysicalMemory','NumberOfProcessors','NumberOfLogicalProcessors','UserName') }
                @{ Class = 'Win32_OperatingSystem'; Props = @('Caption','Version','BuildNumber','OSArchitecture','InstallDate','LastBootUpTime','FreePhysicalMemory','TotalVisibleMemorySize','WindowsDirectory') }
            )
        }
        'BIOS / Firmware' = @{
            Classes = @(
                @{ Class = 'Win32_BIOS'; Props = @('Manufacturer','Name','Version','SMBIOSBIOSVersion','SerialNumber','ReleaseDate') }
                @{ Class = 'Win32_BaseBoard'; Props = @('Manufacturer','Product','SerialNumber','Version') }
                @{ Class = 'Win32_SystemEnclosure'; Props = @('Manufacturer','ChassisTypes','SerialNumber','SMBIOSAssetTag') }
            )
        }
        'Processor' = @{
            Classes = @(
                @{ Class = 'Win32_Processor'; Props = @('Name','Manufacturer','NumberOfCores','NumberOfLogicalProcessors','MaxClockSpeed','CurrentClockSpeed','L2CacheSize','L3CacheSize','Architecture','SocketDesignation','ProcessorId') }
            )
        }
        'Memory' = @{
            Classes = @(
                @{ Class = 'Win32_PhysicalMemory'; Props = @('BankLabel','Capacity','Speed','MemoryType','FormFactor','Manufacturer','PartNumber','SerialNumber','DeviceLocator') }
            )
        }
        'Storage' = @{
            Classes = @(
                @{ Class = 'Win32_DiskDrive'; Props = @('Model','Size','MediaType','InterfaceType','Partitions','SerialNumber','FirmwareRevision','DeviceID') }
                @{ Class = 'Win32_LogicalDisk'; Props = @('DeviceID','VolumeName','FileSystem','Size','FreeSpace','DriveType','VolumeSerialNumber') }
            )
        }
        'Network' = @{
            Classes = @(
                @{ Class = 'Win32_NetworkAdapter'; Props = @('Name','MACAddress','NetConnectionID','Speed','AdapterType','Manufacturer','NetConnectionStatus','NetEnabled'); Filter = "NetConnectionID IS NOT NULL" }
                @{ Class = 'Win32_NetworkAdapterConfiguration'; Props = @('Description','IPAddress','IPSubnet','DefaultIPGateway','DNSServerSearchOrder','DHCPEnabled','DHCPServer','MACAddress'); Filter = "IPEnabled = TRUE" }
            )
        }
        'Display' = @{
            Classes = @(
                @{ Class = 'Win32_VideoController'; Props = @('Name','AdapterRAM','DriverVersion','DriverDate','VideoModeDescription','CurrentHorizontalResolution','CurrentVerticalResolution','CurrentRefreshRate','VideoProcessor') }
                @{ Class = 'Win32_DesktopMonitor'; Props = @('Name','MonitorManufacturer','MonitorType','ScreenHeight','ScreenWidth') }
            )
        }
        'Devices (PnP)' = @{
            Classes = @(
                @{ Class = 'Win32_PnPEntity'; Props = @('Name','DeviceID','Manufacturer','Status','PNPClass','Service','ConfigManagerErrorCode'); Filter = "ConfigManagerErrorCode <> 0" ; Label = 'Problem Devices' }
            )
        }
    }

    # -----------------------------------------------------------------------
    # Helper: format WMI values for display
    # -----------------------------------------------------------------------

    $formatWmiValue = {
        param([string]$PropName, $Value)
        if ($null -eq $Value) { return '' }

        # Size-related properties: convert bytes to human-readable
        if ($PropName -in @('TotalPhysicalMemory', 'Size', 'FreeSpace', 'Capacity', 'AdapterRAM', 'TotalVisibleMemorySize', 'FreePhysicalMemory')) {
            $bytes = [double]$Value
            # TotalVisibleMemorySize and FreePhysicalMemory are in KB
            if ($PropName -in @('TotalVisibleMemorySize', 'FreePhysicalMemory')) { $bytes = $bytes * 1024 }
            if ($bytes -ge 1TB) { return '{0:N2} TB' -f ($bytes / 1TB) }
            if ($bytes -ge 1GB) { return '{0:N2} GB' -f ($bytes / 1GB) }
            if ($bytes -ge 1MB) { return '{0:N2} MB' -f ($bytes / 1MB) }
            if ($bytes -ge 1KB) { return '{0:N2} KB' -f ($bytes / 1KB) }
            return "$bytes bytes"
        }

        # Speed: MHz
        if ($PropName -in @('MaxClockSpeed', 'CurrentClockSpeed', 'Speed')) {
            return "$Value MHz"
        }

        # Cache: KB
        if ($PropName -in @('L2CacheSize', 'L3CacheSize')) {
            $kb = [int]$Value
            if ($kb -ge 1024) { return '{0:N0} MB' -f ($kb / 1024) }
            return "$kb KB"
        }

        # Network speed: bps to human
        if ($PropName -eq 'Speed' -and $Value -gt 0) {
            $bps = [double]$Value
            if ($bps -ge 1000000000) { return '{0:N0} Gbps' -f ($bps / 1000000000) }
            if ($bps -ge 1000000) { return '{0:N0} Mbps' -f ($bps / 1000000) }
            return "$Value bps"
        }

        # DriveType enum
        if ($PropName -eq 'DriveType') {
            switch ([int]$Value) {
                0 { return 'Unknown' }
                1 { return 'No Root Directory' }
                2 { return 'Removable' }
                3 { return 'Local Disk' }
                4 { return 'Network Drive' }
                5 { return 'Compact Disc' }
                6 { return 'RAM Disk' }
                default { return "$Value" }
            }
        }

        # ChassisTypes array
        if ($PropName -eq 'ChassisTypes' -and $Value -is [array]) {
            $names = foreach ($ct in $Value) {
                switch ([int]$ct) {
                    1  { 'Other' }     2  { 'Unknown' }    3  { 'Desktop' }
                    4  { 'Low Profile Desktop' } 5  { 'Pizza Box' } 6  { 'Mini Tower' }
                    7  { 'Tower' }     8  { 'Portable' }   9  { 'Laptop' }
                    10 { 'Notebook' }  11 { 'Hand Held' }  12 { 'Docking Station' }
                    13 { 'All in One' } 14 { 'Sub Notebook' } 15 { 'Space-Saving' }
                    16 { 'Lunch Box' } 17 { 'Main System Chassis' }
                    23 { 'Rack Mount' } 24 { 'Sealed-Case PC' }
                    30 { 'Tablet' }   31 { 'Convertible' } 32 { 'Detachable' }
                    default { "$ct" }
                }
            }
            return $names -join ', '
        }

        # NetConnectionStatus enum
        if ($PropName -eq 'NetConnectionStatus') {
            switch ([int]$Value) {
                0 { return 'Disconnected' }
                1 { return 'Connecting' }
                2 { return 'Connected' }
                3 { return 'Disconnecting' }
                7 { return 'Media Disconnected' }
                default { return "$Value" }
            }
        }

        # Arrays (IPAddress, etc.)
        if ($Value -is [array]) { return $Value -join ', ' }

        # Dates
        if ($PropName -in @('InstallDate', 'LastBootUpTime', 'ReleaseDate', 'DriverDate') -and $Value -is [datetime]) {
            return $Value.ToString('yyyy-MM-dd HH:mm:ss')
        }

        $str = $Value.ToString()

        # Flag common SMBIOS placeholder values
        $placeholders = @(
            'System Serial Number', 'System Product Name', 'System manufacturer',
            'Default string', 'To Be Filled By O.E.M.', 'Not Specified',
            'None', 'No Asset Tag', 'Type1ProductConfigId', 'All',
            'To be filled by O.E.M.', 'Chassis Serial Number', 'Base Board Serial Number'
        )
        if ($str -in $placeholders) {
            return "$str (not set by manufacturer)"
        }

        return $str
    }.GetNewClosure()

    # -----------------------------------------------------------------------
    # Toolbar (Dock:Top)
    # -----------------------------------------------------------------------

    $pnlToolbar = New-Object System.Windows.Forms.Panel
    $pnlToolbar.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlToolbar.Height = 40
    $pnlToolbar.BackColor = $clrPanelBg
    $pnlToolbar.Padding = New-Object System.Windows.Forms.Padding(6, 4, 6, 4)

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = "Refresh All"
    $btnRefresh.Location = New-Object System.Drawing.Point(8, 5)
    $btnRefresh.Size = New-Object System.Drawing.Size(100, 28)
    $btnRefresh.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    & $btnStyle $btnRefresh $clrAccent
    $pnlToolbar.Controls.Add($btnRefresh)

    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Text = "Hardware and device information via WMI (read-only, no elevation required)"
    $lblInfo.AutoSize = $true
    $lblInfo.Location = New-Object System.Drawing.Point(120, 10)
    $lblInfo.ForeColor = $clrHint
    $lblInfo.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
    $pnlToolbar.Controls.Add($lblInfo)

    $pnlToolbarSep = New-Object System.Windows.Forms.Panel
    $pnlToolbarSep.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlToolbarSep.Height = 1; $pnlToolbarSep.BackColor = $clrSepLine

    # -----------------------------------------------------------------------
    # SplitContainer (Category tree left, Properties grid right)
    # -----------------------------------------------------------------------

    $splitDev = New-Object System.Windows.Forms.SplitContainer
    $splitDev.Dock = [System.Windows.Forms.DockStyle]::Fill
    $splitDev.Orientation = [System.Windows.Forms.Orientation]::Vertical
    $splitDev.SplitterWidth = 6
    $splitDev.BackColor = $clrSepLine
    $splitDev.Panel1.BackColor = $clrPanelBg
    $splitDev.Panel2.BackColor = $clrPanelBg
    $splitDev.Panel1MinSize = 100
    $splitDev.Panel2MinSize = 100

    # -----------------------------------------------------------------------
    # Category TreeView (left)
    # -----------------------------------------------------------------------

    $treeCat = New-Object System.Windows.Forms.TreeView
    $treeCat.Dock = [System.Windows.Forms.DockStyle]::Fill
    $treeCat.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
    $treeCat.BackColor = if ($isDark) { [System.Drawing.Color]::FromArgb(38, 38, 38) } else { [System.Drawing.Color]::White }
    $treeCat.ForeColor = $clrText
    $treeCat.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $treeCat.FullRowSelect = $true
    $treeCat.HideSelection = $false
    $treeCat.ShowLines = $true
    $treeCat.ShowPlusMinus = $true
    $treeCat.ShowRootLines = $true
    if ($isDark) { $treeCat.LineColor = [System.Drawing.Color]::FromArgb(80, 80, 80) }
    & $dblBuffer $treeCat

    # Build category nodes
    $rootNode = New-Object System.Windows.Forms.TreeNode("This Computer")
    $rootNode.Tag = '__root__'
    foreach ($catName in $categories.Keys) {
        $catNode = New-Object System.Windows.Forms.TreeNode($catName)
        $catNode.Tag = $catName
        $catDef = $categories[$catName]
        foreach ($classDef in $catDef.Classes) {
            $label = if ($classDef.Label) { $classDef.Label } else { $classDef.Class }
            $childNode = New-Object System.Windows.Forms.TreeNode($label)
            $childNode.Tag = @{ Category = $catName; ClassDef = $classDef }
            $catNode.Nodes.Add($childNode) | Out-Null
        }
        $rootNode.Nodes.Add($catNode) | Out-Null
    }
    $treeCat.Nodes.Add($rootNode) | Out-Null
    $rootNode.Expand()

    $splitDev.Panel1.Controls.Add($treeCat)

    # -----------------------------------------------------------------------
    # Properties grid (right)
    # -----------------------------------------------------------------------

    $gridProps = & $newGrid
    $gridProps.AllowUserToOrderColumns = $true

    $colProp = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colProp.HeaderText = "Property"; $colProp.DataPropertyName = "Property"; $colProp.Width = 220
    $colVal = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colVal.HeaderText = "Value"; $colVal.DataPropertyName = "Value"; $colVal.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill

    $gridProps.Columns.Add($colProp) | Out-Null
    $gridProps.Columns.Add($colVal) | Out-Null

    $devPropsTable = New-Object System.Data.DataTable
    [void]$devPropsTable.Columns.Add("Property", [string])
    [void]$devPropsTable.Columns.Add("Value", [string])
    $gridProps.DataSource = $devPropsTable

    $splitDev.Panel2.Controls.Add($gridProps)

    # -----------------------------------------------------------------------
    # Load data for a WMI class
    # -----------------------------------------------------------------------

    $loadClassData = {
        param([hashtable]$ClassDef)
        $devPropsTable.Clear()
        $className = $ClassDef.Class
        $props = $ClassDef.Props
        $filter = $ClassDef.Filter

        & $statusFunc "Querying $className..."
        [System.Windows.Forms.Application]::DoEvents()

        try {
            $cimParams = @{ ClassName = $className; ErrorAction = 'Stop' }
            if ($filter) { $cimParams['Filter'] = $filter }
            $instances = @(Get-CimInstance @cimParams)

            if ($instances.Count -eq 0) {
                $row = $devPropsTable.NewRow()
                $row["Property"] = "(No results)"
                $row["Value"] = if ($filter) { "Filter: $filter" } else { "No instances found" }
                $devPropsTable.Rows.Add($row)
                & $statusFunc "$($className): no results"
                return
            }

            $instNum = 0
            foreach ($inst in $instances) {
                $instNum++
                if ($instances.Count -gt 1) {
                    $row = $devPropsTable.NewRow()
                    $row["Property"] = "--- Instance $instNum ---"
                    $row["Value"] = ""
                    $devPropsTable.Rows.Add($row)
                }

                foreach ($propName in $props) {
                    $val = $inst.$propName
                    $formatted = & $formatWmiValue $propName $val
                    $row = $devPropsTable.NewRow()
                    $row["Property"] = $propName
                    $row["Value"] = $formatted
                    $devPropsTable.Rows.Add($row)
                }
            }

            & $statusFunc "$($className): $($instances.Count) instance(s)"
        } catch {
            $row = $devPropsTable.NewRow()
            $row["Property"] = "(Error)"
            $row["Value"] = $_.Exception.Message
            $devPropsTable.Rows.Add($row)
            & $statusFunc "Error querying $className"
        }
    }.GetNewClosure()

    # Load all properties for a category
    $loadCategory = {
        param([string]$CatName)
        $devPropsTable.Clear()
        $catDef = $categories[$CatName]

        & $statusFunc "Loading $CatName..."
        [System.Windows.Forms.Application]::DoEvents()

        foreach ($classDef in $catDef.Classes) {
            $className = $classDef.Class
            $label = if ($classDef.Label) { $classDef.Label } else { $className }

            # Section header
            $row = $devPropsTable.NewRow()
            $row["Property"] = "=== $label ==="
            $row["Value"] = ""
            $devPropsTable.Rows.Add($row)

            try {
                $cimParams = @{ ClassName = $className; ErrorAction = 'Stop' }
                if ($classDef.Filter) { $cimParams['Filter'] = $classDef.Filter }
                $instances = @(Get-CimInstance @cimParams)

                if ($instances.Count -eq 0) {
                    $row = $devPropsTable.NewRow()
                    $row["Property"] = "(No results)"
                    $row["Value"] = ""
                    $devPropsTable.Rows.Add($row)
                    continue
                }

                $instNum = 0
                foreach ($inst in $instances) {
                    $instNum++
                    if ($instances.Count -gt 1) {
                        $row = $devPropsTable.NewRow()
                        $row["Property"] = "--- #$instNum ---"
                        $row["Value"] = ""
                        $devPropsTable.Rows.Add($row)
                    }

                    foreach ($propName in $classDef.Props) {
                        $val = $inst.$propName
                        $formatted = & $formatWmiValue $propName $val
                        $row = $devPropsTable.NewRow()
                        $row["Property"] = $propName
                        $row["Value"] = $formatted
                        $devPropsTable.Rows.Add($row)
                    }
                }
            } catch {
                $row = $devPropsTable.NewRow()
                $row["Property"] = "(Error)"
                $row["Value"] = $_.Exception.Message
                $devPropsTable.Rows.Add($row)
            }
        }

        & $statusFunc "$CatName loaded"
    }.GetNewClosure()

    # -----------------------------------------------------------------------
    # Tree selection handler
    # -----------------------------------------------------------------------

    $treeCat.Add_AfterSelect(({
        param($s, $e)
        $node = $e.Node
        if (-not $node -or -not $node.Tag) { return }

        $tabPage.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            if ($node.Tag -is [string]) {
                $tagStr = [string]$node.Tag
                if ($tagStr -eq '__root__') {
                    # Show summary of all categories
                    & $loadCategory 'System'
                } else {
                    & $loadCategory $tagStr
                }
            } elseif ($node.Tag -is [hashtable]) {
                $info = [hashtable]$node.Tag
                & $loadClassData $info.ClassDef
            }
        } finally {
            $tabPage.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }).GetNewClosure())

    # -----------------------------------------------------------------------
    # Refresh All
    # -----------------------------------------------------------------------

    $btnRefresh.Add_Click(({
        $sel = $treeCat.SelectedNode
        if ($sel) {
            $treeCat.SelectedNode = $null
            $treeCat.SelectedNode = $sel
        } else {
            $treeCat.SelectedNode = $treeCat.Nodes[0]
        }
    }).GetNewClosure())

    # -----------------------------------------------------------------------
    # Context menu: Copy
    # -----------------------------------------------------------------------

    $ctxProps = New-Object System.Windows.Forms.ContextMenuStrip
    if ($isDark -and $script:DarkRenderer) { $ctxProps.Renderer = $script:DarkRenderer }
    $ctxProps.BackColor = $clrPanelBg; $ctxProps.ForeColor = $clrText

    $ctxCopyVal = New-Object System.Windows.Forms.ToolStripMenuItem("Copy Value")
    $ctxCopyVal.ForeColor = $clrText
    $ctxCopyVal.Add_Click(({
        if ($gridProps.CurrentRow) {
            $val = $gridProps.CurrentRow.Cells[1].Value
            if ($val) { [System.Windows.Forms.Clipboard]::SetText([string]$val) }
        }
    }).GetNewClosure())
    $ctxProps.Items.Add($ctxCopyVal) | Out-Null

    $ctxCopyAll = New-Object System.Windows.Forms.ToolStripMenuItem("Copy All Properties")
    $ctxCopyAll.ForeColor = $clrText
    $ctxCopyAll.Add_Click(({
        $lines = @()
        foreach ($row in $devPropsTable.Rows) {
            $lines += "{0}: {1}" -f $row["Property"], $row["Value"]
        }
        if ($lines.Count -gt 0) { [System.Windows.Forms.Clipboard]::SetText($lines -join "`r`n") }
    }).GetNewClosure())
    $ctxProps.Items.Add($ctxCopyAll) | Out-Null

    $gridProps.ContextMenuStrip = $ctxProps

    # -----------------------------------------------------------------------
    # Assemble layout
    # -----------------------------------------------------------------------

    $tabPage.Controls.Add($splitDev)
    $splitDev.BringToFront()

    $tabPage.Controls.Add($pnlToolbarSep)
    $pnlToolbarSep.SendToBack()

    $tabPage.Controls.Add($pnlToolbar)
    $pnlToolbar.SendToBack()

    # Defer SplitterDistance until the control has a valid size
    $devSplitFlag = @{ Done = $false }
    $splitDev.Add_SizeChanged(({
        if (-not $devSplitFlag.Done -and $splitDev.Width -gt ($splitDev.Panel1MinSize + $splitDev.Panel2MinSize + $splitDev.SplitterWidth)) {
            $devSplitFlag.Done = $true
            $splitDev.SplitterDistance = [Math]::Max($splitDev.Panel1MinSize, [int]($splitDev.Width * 0.25))
        }
    }).GetNewClosure())

    # Auto-load System category on first display
    $treeCat.SelectedNode = $treeCat.Nodes[0]
}
