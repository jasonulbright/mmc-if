<#
.SYNOPSIS
    Certificate Store Browser module for MMC-If (WPF).

.DESCRIPTION
    Browse user and machine certificate stores via .NET X509Store.
    Read-only. No elevation required (though some LocalMachine stores may
    be unreadable without admin, those fail gracefully).
#>

function New-CertificateStoreView {
    param([hashtable]$Context)

    $setStatus = $Context.SetStatus
    $log       = $Context.Log

    $xamlPath = Join-Path $PSScriptRoot 'CertificateStore.xaml'
    $xamlRaw  = Get-Content -LiteralPath $xamlPath -Raw
    $reader   = New-Object System.Xml.XmlNodeReader ([xml]$xamlRaw)
    $view     = [Windows.Markup.XamlReader]::Load($reader)

    $btnRefresh    = $view.FindName('btnRefresh')
    $treeCerts     = $view.FindName('treeCerts')
    $gridCerts     = $view.FindName('gridCerts')
    $txtDetail     = $view.FindName('txtDetail')
    $mnuCopyThumb  = $view.FindName('mnuCopyThumb')
    $mnuCopySubject = $view.FindName('mnuCopySubject')
    $mnuCopyDetail = $view.FindName('mnuCopyDetail')

    $state = @{
        Rows = New-Object System.Collections.ObjectModel.ObservableCollection[object]
    }
    $gridCerts.ItemsSource = $state.Rows

    $extractCN = {
        param([string]$Dn)
        if ([string]::IsNullOrEmpty($Dn)) { return '' }
        if ($Dn -match '^CN=([^,]+)') { return $Matches[1] }
        return $Dn
    }

    $getSan = {
        param($cert)
        $ext = $cert.Extensions | Where-Object { $_.Oid.Value -eq '2.5.29.17' }
        if ($ext) { return $ext.Format($false) }
        ''
    }

    $getTemplate = {
        param($cert)
        $ext = $cert.Extensions | Where-Object { $_.Oid.Value -eq '1.3.6.1.4.1.311.21.7' }
        if (-not $ext) {
            $ext = $cert.Extensions | Where-Object { $_.Oid.Value -eq '1.3.6.1.4.1.311.20.2' }
        }
        if ($ext) { return $ext.Format($false) }
        ''
    }

    $getKeyUsage = {
        param($cert)
        $parts = @()
        $ku = $cert.Extensions | Where-Object { $_ -is [System.Security.Cryptography.X509Certificates.X509KeyUsageExtension] }
        if ($ku) { $parts += $ku.KeyUsages.ToString() }
        $eku = $cert.Extensions | Where-Object { $_ -is [System.Security.Cryptography.X509Certificates.X509EnhancedKeyUsageExtension] }
        if ($eku) { $parts += ($eku.EnhancedKeyUsages | ForEach-Object { $_.FriendlyName }) -join ', ' }
        $parts -join '; '
    }

    # -----------------------------------------------------------------------
    # Build store tree
    # -----------------------------------------------------------------------
    $storeDefs = [ordered]@{
        'CurrentUser' = @(
            @{ Name = 'My';                  Label = 'Personal' }
            @{ Name = 'Root';                Label = 'Trusted Root CA' }
            @{ Name = 'CertificateAuthority'; Label = 'Intermediate CA' }
            @{ Name = 'TrustedPeople';       Label = 'Trusted People' }
            @{ Name = 'TrustedPublisher';    Label = 'Trusted Publishers' }
            @{ Name = 'Disallowed';          Label = 'Disallowed' }
        )
        'LocalMachine' = @(
            @{ Name = 'My';                  Label = 'Personal' }
            @{ Name = 'Root';                Label = 'Trusted Root CA' }
            @{ Name = 'CertificateAuthority'; Label = 'Intermediate CA' }
            @{ Name = 'TrustedPeople';       Label = 'Trusted People' }
            @{ Name = 'TrustedPublisher';    Label = 'Trusted Publishers' }
            @{ Name = 'Disallowed';          Label = 'Disallowed' }
            @{ Name = 'Remote Desktop';      Label = 'Remote Desktop' }
            @{ Name = 'WebHosting';          Label = 'Web Hosting' }
        )
    }

    foreach ($locationName in $storeDefs.Keys) {
        $locNode = New-Object System.Windows.Controls.TreeViewItem
        $locNode.Header = $locationName
        $locNode.Tag = $null
        $locNode.IsExpanded = $true
        $loc = [System.Security.Cryptography.X509Certificates.StoreLocation]::$locationName
        foreach ($def in $storeDefs[$locationName]) {
            $n = New-Object System.Windows.Controls.TreeViewItem
            $n.Header = $def.Label
            $n.Tag = @{ Location = $loc; StoreName = $def.Name }
            [void]$locNode.Items.Add($n)
        }
        [void]$treeCerts.Items.Add($locNode)
    }

    # -----------------------------------------------------------------------
    # Load certificates for selected store
    # -----------------------------------------------------------------------
    $loadStore = {
        param([hashtable]$StoreInfo)
        $state.Rows.Clear()
        $txtDetail.Clear()
        $view.Cursor = [System.Windows.Input.Cursors]::Wait

        try {
            $store = New-Object System.Security.Cryptography.X509Certificates.X509Store($StoreInfo.StoreName, $StoreInfo.Location)
            $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
            foreach ($cert in $store.Certificates) {
                $state.Rows.Add([pscustomobject]@{
                    Subject      = & $extractCN $cert.Subject
                    Issuer       = & $extractCN $cert.Issuer
                    Expires      = $cert.NotAfter.ToString('yyyy-MM-dd HH:mm')
                    Thumbprint   = $cert.Thumbprint
                    FriendlyName = $cert.FriendlyName
                    ExpiryDate   = $cert.NotAfter
                    FullSubject  = $cert.Subject
                    FullIssuer   = $cert.Issuer
                    SerialNumber = $cert.SerialNumber
                    SigAlgo      = $cert.SignatureAlgorithm.FriendlyName
                    ValidFrom    = $cert.NotBefore.ToString('yyyy-MM-dd HH:mm')
                    KeyUsage     = & $getKeyUsage $cert
                    SAN          = & $getSan $cert
                    Template     = & $getTemplate $cert
                })
            }
            $store.Close()
            & $setStatus "$($StoreInfo.Location)\$($StoreInfo.StoreName): $($state.Rows.Count) certificate(s)"
            & $log "Loaded $($state.Rows.Count) certs from $($StoreInfo.Location)\$($StoreInfo.StoreName)"
        }
        catch {
            $err = $_.Exception.Message
            if ($err -like '*access*' -or $err -like '*denied*' -or $err -like '*crypto*') {
                & $setStatus "Access denied: $($StoreInfo.Location)\$($StoreInfo.StoreName) - may require elevation"
            }
            else {
                & $setStatus "Error: $err"
            }
        }
        $view.Cursor = [System.Windows.Input.Cursors]::Arrow
    }.GetNewClosure()

    $treeCerts.Add_SelectedItemChanged({
        $sel = $treeCerts.SelectedItem
        if (-not $sel -or -not ($sel.Tag -is [hashtable])) {
            $state.Rows.Clear()
            $txtDetail.Clear()
            return
        }
        & $loadStore $sel.Tag
    }.GetNewClosure())

    $btnRefresh.Add_Click({
        $sel = $treeCerts.SelectedItem
        if ($sel -and ($sel.Tag -is [hashtable])) { & $loadStore $sel.Tag }
    }.GetNewClosure())

    $gridCerts.Add_SelectionChanged({
        $r = $gridCerts.SelectedItem
        if (-not $r) { $txtDetail.Clear(); return }
        $lines = @(
            "Subject: $($r.FullSubject)",
            "Issuer: $($r.FullIssuer)",
            "Serial Number: $($r.SerialNumber)",
            "Thumbprint: $($r.Thumbprint)",
            "Friendly Name: $($r.FriendlyName)",
            '',
            "Valid From: $($r.ValidFrom)",
            "Valid To: $($r.Expires)",
            "Signature Algorithm: $($r.SigAlgo)",
            ''
        )
        if ($r.KeyUsage) { $lines += "Key Usage: $($r.KeyUsage)" }
        if ($r.SAN)      { $lines += "Subject Alternative Names: $($r.SAN)" }
        if ($r.Template) { $lines += "Certificate Template: $($r.Template)" }
        $txtDetail.Text = $lines -join "`r`n"
    }.GetNewClosure())

    $mnuCopyThumb.Add_Click({
        $r = $gridCerts.SelectedItem
        if ($r) { [System.Windows.Clipboard]::SetText([string]$r.Thumbprint) }
    }.GetNewClosure())
    $mnuCopySubject.Add_Click({
        $r = $gridCerts.SelectedItem
        if ($r) { [System.Windows.Clipboard]::SetText([string]$r.FullSubject) }
    }.GetNewClosure())
    $mnuCopyDetail.Add_Click({
        if ($txtDetail.Text) { [System.Windows.Clipboard]::SetText($txtDetail.Text) }
    }.GetNewClosure())

    return $view
}
