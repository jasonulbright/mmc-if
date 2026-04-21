<#
.SYNOPSIS
    Device Info module for MMC-If (WPF).

.DESCRIPTION
    View hardware and device information via CIM/WMI classes. Read-only.
    No mmc.exe, no elevation required.
#>

function New-DeviceInfoView {
    param([hashtable]$Context)

    $setStatus = $Context.SetStatus
    $log       = $Context.Log

    $xamlPath = Join-Path $PSScriptRoot 'DeviceInfo.xaml'
    $xamlRaw  = Get-Content -LiteralPath $xamlPath -Raw
    $reader   = New-Object System.Xml.XmlNodeReader ([xml]$xamlRaw)
    $view     = [Windows.Markup.XamlReader]::Load($reader)

    $btnRefresh   = $view.FindName('btnRefresh')
    $treeCat      = $view.FindName('treeCat')
    $gridProps    = $view.FindName('gridProps')
    $mnuCopyValue = $view.FindName('mnuCopyValue')
    $mnuCopyAll   = $view.FindName('mnuCopyAll')

    # -----------------------------------------------------------------------
    # Category definitions
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
                @{ Class = 'Win32_PnPEntity'; Props = @('Name','DeviceID','Manufacturer','Status','PNPClass','Service','ConfigManagerErrorCode'); Filter = "ConfigManagerErrorCode <> 0"; Label = 'Problem Devices' }
            )
        }
    }

    # -----------------------------------------------------------------------
    # Value formatting
    # -----------------------------------------------------------------------
    $formatWmiValue = {
        param([string]$PropName, $Value)
        if ($null -eq $Value) { return '' }

        if ($PropName -in @('TotalPhysicalMemory','Size','FreeSpace','Capacity','AdapterRAM','TotalVisibleMemorySize','FreePhysicalMemory')) {
            $bytes = [double]$Value
            if ($PropName -in @('TotalVisibleMemorySize','FreePhysicalMemory')) { $bytes *= 1024 }
            if ($bytes -ge 1TB) { return '{0:N2} TB' -f ($bytes / 1TB) }
            if ($bytes -ge 1GB) { return '{0:N2} GB' -f ($bytes / 1GB) }
            if ($bytes -ge 1MB) { return '{0:N2} MB' -f ($bytes / 1MB) }
            if ($bytes -ge 1KB) { return '{0:N2} KB' -f ($bytes / 1KB) }
            return "$bytes bytes"
        }

        if ($PropName -in @('MaxClockSpeed','CurrentClockSpeed')) { return "$Value MHz" }

        if ($PropName -in @('L2CacheSize','L3CacheSize')) {
            $kb = [int]$Value
            if ($kb -ge 1024) { return '{0:N0} MB' -f ($kb / 1024) }
            return "$kb KB"
        }

        if ($PropName -eq 'Speed' -and $Value -gt 0) {
            $bps = [double]$Value
            if ($bps -ge 1000000000) { return '{0:N0} Gbps' -f ($bps / 1000000000) }
            if ($bps -ge 1000000)    { return '{0:N0} Mbps' -f ($bps / 1000000) }
            return "$Value bps"
        }

        if ($PropName -eq 'DriveType') {
            switch ([int]$Value) {
                0 { return 'Unknown' } 1 { return 'No Root Directory' }
                2 { return 'Removable' } 3 { return 'Local Disk' }
                4 { return 'Network Drive' } 5 { return 'Compact Disc' } 6 { return 'RAM Disk' }
                default { return "$Value" }
            }
        }

        if ($PropName -eq 'ChassisTypes' -and $Value -is [array]) {
            $names = foreach ($ct in $Value) {
                switch ([int]$ct) {
                    1 { 'Other' } 2 { 'Unknown' } 3 { 'Desktop' } 4 { 'Low Profile Desktop' }
                    5 { 'Pizza Box' } 6 { 'Mini Tower' } 7 { 'Tower' } 8 { 'Portable' }
                    9 { 'Laptop' } 10 { 'Notebook' } 11 { 'Hand Held' } 12 { 'Docking Station' }
                    13 { 'All in One' } 14 { 'Sub Notebook' } 15 { 'Space-Saving' } 16 { 'Lunch Box' }
                    17 { 'Main System Chassis' } 23 { 'Rack Mount' } 24 { 'Sealed-Case PC' }
                    30 { 'Tablet' } 31 { 'Convertible' } 32 { 'Detachable' }
                    default { "$ct" }
                }
            }
            return $names -join ', '
        }

        if ($PropName -eq 'NetConnectionStatus') {
            switch ([int]$Value) {
                0 { return 'Disconnected' } 1 { return 'Connecting' }
                2 { return 'Connected' } 3 { return 'Disconnecting' }
                7 { return 'Media Disconnected' }
                default { return "$Value" }
            }
        }

        if ($Value -is [array]) { return $Value -join ', ' }

        if ($PropName -in @('InstallDate','LastBootUpTime','ReleaseDate','DriverDate') -and $Value -is [datetime]) {
            return $Value.ToString('yyyy-MM-dd HH:mm:ss')
        }

        $str = $Value.ToString()
        $placeholders = @(
            'System Serial Number','System Product Name','System manufacturer',
            'Default string','To Be Filled By O.E.M.','Not Specified','None',
            'No Asset Tag','Type1ProductConfigId','All','To be filled by O.E.M.',
            'Chassis Serial Number','Base Board Serial Number'
        )
        if ($str -in $placeholders) { return "$str (not set by manufacturer)" }
        return $str
    }

    # -----------------------------------------------------------------------
    # State
    # -----------------------------------------------------------------------
    $state = @{
        Rows = New-Object System.Collections.ObjectModel.ObservableCollection[object]
    }
    $gridProps.ItemsSource = $state.Rows

    # -----------------------------------------------------------------------
    # Populate tree
    # -----------------------------------------------------------------------
    $rootNode = New-Object System.Windows.Controls.TreeViewItem
    $rootNode.Header = 'This Computer'
    $rootNode.Tag = '__root__'
    $rootNode.IsExpanded = $true
    foreach ($catName in $categories.Keys) {
        $catNode = New-Object System.Windows.Controls.TreeViewItem
        $catNode.Header = $catName
        $catNode.Tag = $catName
        foreach ($classDef in $categories[$catName].Classes) {
            $label = if ($classDef.Label) { $classDef.Label } else { $classDef.Class }
            $childNode = New-Object System.Windows.Controls.TreeViewItem
            $childNode.Header = $label
            $childNode.Tag = @{ Category = $catName; ClassDef = $classDef }
            [void]$catNode.Items.Add($childNode)
        }
        [void]$rootNode.Items.Add($catNode)
    }
    [void]$treeCat.Items.Add($rootNode)

    # -----------------------------------------------------------------------
    # Loaders
    # -----------------------------------------------------------------------
    $addRow = {
        param([string]$Prop, [string]$Val)
        $state.Rows.Add([pscustomobject]@{ Property = $Prop; Value = $Val })
    }.GetNewClosure()

    $loadClassData = {
        param([hashtable]$ClassDef)
        $state.Rows.Clear()
        $className = $ClassDef.Class
        & $setStatus "Querying $className..."
        try {
            $cimParams = @{ ClassName = $className; ErrorAction = 'Stop' }
            if ($ClassDef.Filter) { $cimParams['Filter'] = $ClassDef.Filter }
            $instances = @(Get-CimInstance @cimParams)
            if ($instances.Count -eq 0) {
                & $addRow '(No results)' ($(if ($ClassDef.Filter) { "Filter: $($ClassDef.Filter)" } else { 'No instances' }))
                & $setStatus "$($className): no results"
                return
            }
            $n = 0
            foreach ($inst in $instances) {
                $n++
                if ($instances.Count -gt 1) { & $addRow "--- Instance $n ---" '' }
                foreach ($propName in $ClassDef.Props) {
                    & $addRow $propName (& $formatWmiValue $propName $inst.$propName)
                }
            }
            & $setStatus "$($className): $($instances.Count) instance(s)"
        }
        catch {
            & $addRow '(Error)' $_.Exception.Message
            & $setStatus "Error querying $className"
        }
    }.GetNewClosure()

    $loadCategory = {
        param([string]$CatName)
        $state.Rows.Clear()
        $catDef = $categories[$CatName]
        & $setStatus "Loading $CatName..."
        foreach ($classDef in $catDef.Classes) {
            $label = if ($classDef.Label) { $classDef.Label } else { $classDef.Class }
            & $addRow "=== $label ===" ''
            try {
                $cimParams = @{ ClassName = $classDef.Class; ErrorAction = 'Stop' }
                if ($classDef.Filter) { $cimParams['Filter'] = $classDef.Filter }
                $instances = @(Get-CimInstance @cimParams)
                if ($instances.Count -eq 0) {
                    & $addRow '(No results)' ''
                    continue
                }
                $n = 0
                foreach ($inst in $instances) {
                    $n++
                    if ($instances.Count -gt 1) { & $addRow "--- #$n ---" '' }
                    foreach ($propName in $classDef.Props) {
                        & $addRow $propName (& $formatWmiValue $propName $inst.$propName)
                    }
                }
            }
            catch {
                & $addRow '(Error)' $_.Exception.Message
            }
        }
        & $setStatus "$CatName loaded"
    }.GetNewClosure()

    # -----------------------------------------------------------------------
    # Selection
    # -----------------------------------------------------------------------
    $treeCat.Add_SelectedItemChanged({
        $node = $treeCat.SelectedItem
        if (-not $node -or $null -eq $node.Tag) { return }
        $view.Cursor = [System.Windows.Input.Cursors]::Wait
        try {
            if ($node.Tag -is [string]) {
                if ([string]$node.Tag -eq '__root__') { & $loadCategory 'System' }
                else { & $loadCategory ([string]$node.Tag) }
            }
            elseif ($node.Tag -is [hashtable]) {
                & $loadClassData $node.Tag.ClassDef
            }
        }
        finally {
            $view.Cursor = [System.Windows.Input.Cursors]::Arrow
        }
    }.GetNewClosure())

    $btnRefresh.Add_Click({
        $sel = $treeCat.SelectedItem
        if ($sel) {
            $tag = $sel.Tag
            $sel.IsSelected = $false
            $sel.IsSelected = $true
            # SelectedItemChanged re-fires; nothing to do
        }
        else {
            $rootNode.IsSelected = $true
        }
    }.GetNewClosure())

    $mnuCopyValue.Add_Click({
        $sel = $gridProps.SelectedItem
        if ($sel -and $sel.Value) { [System.Windows.Clipboard]::SetText([string]$sel.Value) }
    }.GetNewClosure())

    $mnuCopyAll.Add_Click({
        $lines = foreach ($r in $state.Rows) { '{0}: {1}' -f $r.Property, $r.Value }
        if ($lines.Count -gt 0) { [System.Windows.Clipboard]::SetText($lines -join "`r`n") }
    }.GetNewClosure())

    # Auto-select root on first display
    $view.Add_Loaded({ if (-not $rootNode.IsSelected) { $rootNode.IsSelected = $true } }.GetNewClosure())

    return $view
}
