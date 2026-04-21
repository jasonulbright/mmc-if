<#
.SYNOPSIS
    Disks module for MMC-If (WPF).

.DESCRIPTION
    View physical disks, partitions, and volumes via Get-Disk / Get-Partition /
    Get-Volume. Read-only - no partition or volume operations.
#>

function New-DisksView {
    param([hashtable]$Context)

    $setStatus = $Context.SetStatus
    $log       = $Context.Log

    $xamlPath = Join-Path $PSScriptRoot 'Disks.xaml'
    $xamlRaw  = Get-Content -LiteralPath $xamlPath -Raw
    $reader   = New-Object System.Xml.XmlNodeReader ([xml]$xamlRaw)
    $view     = [Windows.Markup.XamlReader]::Load($reader)

    $btnRefresh     = $view.FindName('btnRefresh')
    $gridDisks      = $view.FindName('gridDisks')
    $gridPartitions = $view.FindName('gridPartitions')
    $gridVolumes    = $view.FindName('gridVolumes')

    $state = @{
        Disks      = New-Object System.Collections.ObjectModel.ObservableCollection[object]
        Partitions = New-Object System.Collections.ObjectModel.ObservableCollection[object]
        Volumes    = New-Object System.Collections.ObjectModel.ObservableCollection[object]
    }
    $gridDisks.ItemsSource      = $state.Disks
    $gridPartitions.ItemsSource = $state.Partitions
    $gridVolumes.ItemsSource    = $state.Volumes

    $formatSize = {
        param($Bytes)
        if ($null -eq $Bytes -or $Bytes -le 0) { return '' }
        $b = [double]$Bytes
        if ($b -ge 1TB) { return '{0:N2} TB' -f ($b / 1TB) }
        if ($b -ge 1GB) { return '{0:N2} GB' -f ($b / 1GB) }
        if ($b -ge 1MB) { return '{0:N2} MB' -f ($b / 1MB) }
        return '{0:N0} KB' -f ($b / 1KB)
    }

    $loadAll = {
        & $setStatus 'Loading disk info...'
        $view.Cursor = [System.Windows.Input.Cursors]::Wait
        $state.Disks.Clear()
        $state.Partitions.Clear()
        $state.Volumes.Clear()

        try {
            $disks = Get-Disk -ErrorAction Stop | Sort-Object Number
            foreach ($d in $disks) {
                $state.Disks.Add([pscustomobject]@{
                    Number         = [int]$d.Number
                    FriendlyName   = $d.FriendlyName
                    SizeText       = & $formatSize $d.Size
                    PartitionStyle = [string]$d.PartitionStyle
                    BusType        = [string]$d.BusType
                    HealthStatus   = [string]$d.HealthStatus
                    SerialNumber   = if ($d.SerialNumber) { $d.SerialNumber.Trim() } else { '' }
                })
            }
        }
        catch {
            & $log "Get-Disk failed: $($_.Exception.Message)" 'WARN'
        }

        try {
            $parts = Get-Partition -ErrorAction Stop | Sort-Object DiskNumber, PartitionNumber
            foreach ($p in $parts) {
                $state.Partitions.Add([pscustomobject]@{
                    DiskNumber      = [int]$p.DiskNumber
                    PartitionNumber = [int]$p.PartitionNumber
                    DriveLetter     = if ($p.DriveLetter) { "$($p.DriveLetter):" } else { '' }
                    SizeText        = & $formatSize $p.Size
                    Type            = [string]$p.Type
                    Offset          = if ($p.Offset) { '{0:N0}' -f [long]$p.Offset } else { '' }
                    Guid            = if ($p.Guid) { [string]$p.Guid } else { '' }
                })
            }
        }
        catch {
            & $log "Get-Partition failed: $($_.Exception.Message)" 'WARN'
        }

        try {
            $vols = Get-Volume -ErrorAction Stop | Sort-Object DriveLetter
            foreach ($v in $vols) {
                $total = if ($v.Size) { [double]$v.Size } else { 0 }
                $free = if ($v.SizeRemaining) { [double]$v.SizeRemaining } else { 0 }
                $freePct = if ($total -gt 0) { '{0:N1} %' -f ($free * 100 / $total) } else { '' }
                $state.Volumes.Add([pscustomobject]@{
                    DriveLetter     = if ($v.DriveLetter) { "$($v.DriveLetter):" } else { '' }
                    FileSystemLabel = if ($v.FileSystemLabel) { $v.FileSystemLabel } else { '' }
                    FileSystem      = [string]$v.FileSystem
                    SizeText        = & $formatSize $v.Size
                    FreeText        = & $formatSize $v.SizeRemaining
                    FreePct         = $freePct
                    HealthStatus    = [string]$v.HealthStatus
                    DriveType       = [string]$v.DriveType
                })
            }
        }
        catch {
            & $log "Get-Volume failed: $($_.Exception.Message)" 'WARN'
        }

        & $setStatus "$($state.Disks.Count) disk(s), $($state.Partitions.Count) partition(s), $($state.Volumes.Count) volume(s)"
        $view.Cursor = [System.Windows.Input.Cursors]::Arrow
    }.GetNewClosure()

    $btnRefresh.Add_Click({ & $loadAll }.GetNewClosure())
    $view.Add_Loaded({ if ($state.Disks.Count -eq 0) { & $loadAll } }.GetNewClosure())

    return $view
}
