<#
.SYNOPSIS
    Networking module for MMC-If (WPF).

.DESCRIPTION
    View network adapters, IP configuration, DNS servers, gateways. Provides
    ping and tracert actions. Read-only - no adapter enable/disable, no DHCP
    release/renew (those require admin; this tool targets non-admins).
#>

function New-NetworkingView {
    param([hashtable]$Context)

    $setStatus = $Context.SetStatus
    $log       = $Context.Log

    $xamlPath = Join-Path $PSScriptRoot 'Networking.xaml'
    $xamlRaw  = Get-Content -LiteralPath $xamlPath -Raw
    $reader   = New-Object System.Xml.XmlNodeReader ([xml]$xamlRaw)
    $view     = [Windows.Markup.XamlReader]::Load($reader)

    $btnRefresh   = $view.FindName('btnRefresh')
    $txtHost      = $view.FindName('txtHost')
    $btnPing      = $view.FindName('btnPing')
    $btnTracert   = $view.FindName('btnTracert')
    $gridAdapters = $view.FindName('gridAdapters')
    $txtOutput    = $view.FindName('txtOutput')
    $mnuCopyIp    = $view.FindName('mnuCopyIp')
    $mnuCopyMac   = $view.FindName('mnuCopyMac')
    $mnuCopyAll   = $view.FindName('mnuCopyAll')

    $state = @{
        Rows = New-Object System.Collections.ObjectModel.ObservableCollection[object]
    }
    $gridAdapters.ItemsSource = $state.Rows

    $formatSpeed = {
        param($Speed)
        if ($null -eq $Speed -or $Speed -le 0) { return '' }
        $bps = [double]$Speed
        if ($bps -ge 1000000000) { return '{0:N0} Gbps' -f ($bps / 1000000000) }
        if ($bps -ge 1000000)    { return '{0:N0} Mbps' -f ($bps / 1000000) }
        return "$Speed bps"
    }

    $loadAdapters = {
        & $setStatus 'Loading network adapters...'
        $view.Cursor = [System.Windows.Input.Cursors]::Wait
        $state.Rows.Clear()
        try {
            $adapters = Get-NetAdapter -ErrorAction Stop | Sort-Object Name
            foreach ($a in $adapters) {
                $ipv4 = ''
                $ipv6 = ''
                $gw = ''
                $dns = ''
                try {
                    $ipInfo = Get-NetIPAddress -InterfaceIndex $a.InterfaceIndex -ErrorAction Stop
                    $v4 = @($ipInfo | Where-Object { $_.AddressFamily -eq 'IPv4' })
                    $v6 = @($ipInfo | Where-Object { $_.AddressFamily -eq 'IPv6' })
                    if ($v4.Count -gt 0) { $ipv4 = ($v4 | ForEach-Object { "$($_.IPAddress)/$($_.PrefixLength)" }) -join ', ' }
                    if ($v6.Count -gt 0) { $ipv6 = ($v6 | ForEach-Object { "$($_.IPAddress)/$($_.PrefixLength)" }) -join ', ' }
                } catch { }
                try {
                    $route = Get-NetRoute -InterfaceIndex $a.InterfaceIndex -DestinationPrefix '0.0.0.0/0' -ErrorAction Stop
                    if ($route) { $gw = ($route | ForEach-Object { $_.NextHop }) -join ', ' }
                } catch { }
                try {
                    $dnsCfg = Get-DnsClientServerAddress -InterfaceIndex $a.InterfaceIndex -ErrorAction Stop
                    $serverList = @()
                    foreach ($d in $dnsCfg) { $serverList += $d.ServerAddresses }
                    $dns = ($serverList | Where-Object { $_ }) -join ', '
                } catch { }

                $state.Rows.Add([pscustomobject]@{
                    Name                 = $a.Name
                    InterfaceDescription = $a.InterfaceDescription
                    Status               = $a.Status
                    LinkSpeed            = & $formatSpeed $a.LinkSpeed
                    MacAddress           = $a.MacAddress
                    IPv4                 = $ipv4
                    IPv6                 = $ipv6
                    Gateway              = $gw
                    Dns                  = $dns
                    InterfaceIndex       = $a.InterfaceIndex
                    MediaType            = [string]$a.MediaType
                    AdapterType          = [string]$a.AdapterType
                })
            }
            & $setStatus "$($adapters.Count) adapter(s)"
            & $log "Loaded $($adapters.Count) network adapters"
        }
        catch {
            & $setStatus "Error loading adapters: $($_.Exception.Message)"
        }
        $view.Cursor = [System.Windows.Input.Cursors]::Arrow
    }.GetNewClosure()

    $runNetCommand = {
        param([string]$Command, [string]$Target)
        if ([string]::IsNullOrWhiteSpace($Target)) {
            [System.Windows.MessageBox]::Show('Enter a host or IP first.', 'Input required',
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
            return
        }
        & $setStatus "$Command $Target..."
        $view.Cursor = [System.Windows.Input.Cursors]::Wait
        $txtOutput.Clear()
        try {
            $output = & $Command $Target 2>&1
            $txtOutput.Text = ($output | Out-String)
            & $setStatus "$Command $Target done"
            & $log "$Command $Target"
        }
        catch {
            $txtOutput.Text = "Error: $($_.Exception.Message)"
            & $setStatus "$Command failed"
        }
        $view.Cursor = [System.Windows.Input.Cursors]::Arrow
    }.GetNewClosure()

    $btnRefresh.Add_Click({ & $loadAdapters }.GetNewClosure())
    $btnPing.Add_Click({
        & $runNetCommand 'ping' $txtHost.Text
    }.GetNewClosure())
    $btnTracert.Add_Click({
        & $runNetCommand 'tracert' $txtHost.Text
    }.GetNewClosure())

    $gridAdapters.Add_SelectionChanged({
        $a = $gridAdapters.SelectedItem
        if (-not $a) { return }
        $lines = @(
            "Adapter: $($a.Name)",
            "Interface: $($a.InterfaceDescription)",
            "Status: $($a.Status)  |  Speed: $($a.LinkSpeed)  |  MAC: $($a.MacAddress)",
            "Media: $($a.MediaType)  |  Type: $($a.AdapterType)  |  IfIndex: $($a.InterfaceIndex)",
            '',
            "IPv4: $($a.IPv4)",
            "IPv6: $($a.IPv6)",
            "Gateway: $($a.Gateway)",
            "DNS: $($a.Dns)"
        )
        $txtOutput.Text = $lines -join "`r`n"
    }.GetNewClosure())

    $mnuCopyIp.Add_Click({
        $a = $gridAdapters.SelectedItem
        if ($a -and $a.IPv4) { [System.Windows.Clipboard]::SetText([string]$a.IPv4) }
    }.GetNewClosure())
    $mnuCopyMac.Add_Click({
        $a = $gridAdapters.SelectedItem
        if ($a -and $a.MacAddress) { [System.Windows.Clipboard]::SetText([string]$a.MacAddress) }
    }.GetNewClosure())
    $mnuCopyAll.Add_Click({
        if ($txtOutput.Text) { [System.Windows.Clipboard]::SetText($txtOutput.Text) }
    }.GetNewClosure())

    # Populate synchronously before the shell attaches the view to the visual
    # tree. Other modules rely on Add_Loaded, but Networking's CIM chain
    # (Get-NetAdapter + Get-NetIPAddress + Get-NetRoute + Get-DnsClientServerAddress)
    # was reliably leaving the grid blank until the user forced a refresh via
    # the Ping action. Doing it synchronously here guarantees the grid is
    # populated the moment the module is shown; the Loaded handler stays as
    # an idempotent fallback for edge cases (e.g. module reloaded on theme swap).
    & $loadAdapters

    $view.Add_Loaded({ if ($state.Rows.Count -eq 0) { & $loadAdapters } }.GetNewClosure())

    return $view
}
