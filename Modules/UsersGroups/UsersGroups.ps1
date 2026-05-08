<#
.SYNOPSIS
    Local Users & Groups module for MMC-If (WPF).

.DESCRIPTION
    View local user accounts and groups via Get-LocalUser / Get-LocalGroup.
    Read-only. No lusrmgr.msc required.
#>

function New-UsersGroupsView {
    param([hashtable]$Context)

    $setStatus = $Context.SetStatus
    $log       = $Context.Log

    $xamlPath = Join-Path $PSScriptRoot 'UsersGroups.xaml'
    $xamlRaw  = Get-Content -LiteralPath $xamlPath -Raw
    $reader   = New-Object System.Xml.XmlNodeReader ([xml]$xamlRaw)
    $view     = [Windows.Markup.XamlReader]::Load($reader)

    $cboView    = $view.FindName('cboView')
    $txtFilter  = $view.FindName('txtFilter')
    $btnRefresh = $view.FindName('btnRefresh')
    $gridUsers  = $view.FindName('gridUsers')
    $gridGroups = $view.FindName('gridGroups')
    $txtDetail  = $view.FindName('txtDetail')

    $mnuUsrCopyName   = $view.FindName('mnuUsrCopyName')
    $mnuUsrCopySid    = $view.FindName('mnuUsrCopySid')
    $mnuUsrCopyDetail = $view.FindName('mnuUsrCopyDetail')
    $mnuGrpCopyName   = $view.FindName('mnuGrpCopyName')
    $mnuGrpCopySid    = $view.FindName('mnuGrpCopySid')
    $mnuGrpCopyDetail = $view.FindName('mnuGrpCopyDetail')

    $state = @{
        AllUsers      = New-Object System.Collections.ObjectModel.ObservableCollection[object]
        AllGroups     = New-Object System.Collections.ObjectModel.ObservableCollection[object]
        Mode          = 'Users'
    }
    $gridUsers.ItemsSource = $state.AllUsers
    $gridGroups.ItemsSource = $state.AllGroups

    $applyFilter = {
        $ft = $txtFilter.Text
        if ($state.Mode -eq 'Users') {
            if ([string]::IsNullOrWhiteSpace($ft)) {
                $gridUsers.ItemsSource = $state.AllUsers
                return
            }
            $needle = $ft.ToLowerInvariant()
            $f = New-Object System.Collections.ObjectModel.ObservableCollection[object]
            foreach ($u in $state.AllUsers) {
                $hay = "$($u.Name) $($u.FullName) $($u.Description)".ToLowerInvariant()
                if ($hay.Contains($needle)) { [void]$f.Add($u) }
            }
            $gridUsers.ItemsSource = $f
        }
        else {
            if ([string]::IsNullOrWhiteSpace($ft)) {
                $gridGroups.ItemsSource = $state.AllGroups
                return
            }
            $needle = $ft.ToLowerInvariant()
            $f = New-Object System.Collections.ObjectModel.ObservableCollection[object]
            foreach ($g in $state.AllGroups) {
                $hay = "$($g.Name) $($g.Description)".ToLowerInvariant()
                if ($hay.Contains($needle)) { [void]$f.Add($g) }
            }
            $gridGroups.ItemsSource = $f
        }
    }.GetNewClosure()

    $loadUsers = {
        & $setStatus 'Loading local users...'
        $view.Cursor = [System.Windows.Input.Cursors]::Wait
        $state.AllUsers.Clear()
        $txtDetail.Clear()
        try {
            $users = Get-LocalUser -ErrorAction Stop | Sort-Object Name
            foreach ($u in $users) {
                $state.AllUsers.Add([pscustomobject]@{
                    Name            = $u.Name
                    FullName        = if ($u.FullName) { $u.FullName } else { '' }
                    EnabledStr      = if ($u.Enabled) { 'Yes' } else { 'No' }
                    EnabledBool     = [bool]$u.Enabled
                    PasswordLastSet = if ($u.PasswordLastSet) { $u.PasswordLastSet.ToString('yyyy-MM-dd HH:mm') } else { '(never)' }
                    LastLogon       = if ($u.LastLogon) { $u.LastLogon.ToString('yyyy-MM-dd HH:mm') } else { '(never)' }
                    Description     = if ($u.Description) { $u.Description } else { '' }
                    SID             = $u.SID.ToString()
                    Source          = if ($u.PrincipalSource) { $u.PrincipalSource.ToString() } else { 'Local' }
                })
            }
            & $setStatus "$($users.Count) local users"
            & $log "Loaded $($users.Count) local users"
        }
        catch {
            & $setStatus "Error loading users: $($_.Exception.Message)"
        }
        & $applyFilter
        $view.Cursor = [System.Windows.Input.Cursors]::Arrow
    }.GetNewClosure()

    $loadGroups = {
        & $setStatus 'Loading local groups...'
        $view.Cursor = [System.Windows.Input.Cursors]::Wait
        $state.AllGroups.Clear()
        $txtDetail.Clear()
        try {
            $groups = Get-LocalGroup -ErrorAction Stop | Sort-Object Name
            foreach ($g in $groups) {
                $memberCount = ''
                try { $memberCount = @(Get-LocalGroupMember -Group $g.Name -ErrorAction Stop).Count.ToString() }
                catch { $memberCount = '(error)' }
                $state.AllGroups.Add([pscustomobject]@{
                    Name        = $g.Name
                    Description = if ($g.Description) { $g.Description } else { '' }
                    MemberCount = $memberCount
                    SID         = $g.SID.ToString()
                })
            }
            & $setStatus "$($groups.Count) local groups"
            & $log "Loaded $($groups.Count) local groups"
        }
        catch {
            & $setStatus "Error loading groups: $($_.Exception.Message)"
        }
        & $applyFilter
        $view.Cursor = [System.Windows.Input.Cursors]::Arrow
    }.GetNewClosure()

    $switchView = {
        $mode = $cboView.SelectedItem.Content
        $state.Mode = $mode
        $txtFilter.Text = ''
        $txtDetail.Clear()
        if ($mode -eq 'Users') {
            $gridUsers.Visibility = 'Visible'
            $gridGroups.Visibility = 'Collapsed'
            & $loadUsers
        }
        else {
            $gridUsers.Visibility = 'Collapsed'
            $gridGroups.Visibility = 'Visible'
            & $loadGroups
        }
    }.GetNewClosure()

    $cboView.Add_SelectionChanged({ & $switchView }.GetNewClosure())
    $btnRefresh.Add_Click({ & $switchView }.GetNewClosure())
    $txtFilter.Add_TextChanged({ & $applyFilter }.GetNewClosure())

    $gridUsers.Add_SelectionChanged({
        $u = $gridUsers.SelectedItem
        if (-not $u) { $txtDetail.Clear(); return }
        $lines = @(
            "User: $($u.Name)",
            "Full Name: $($u.FullName)",
            "SID: $($u.SID)",
            "Enabled: $($u.EnabledStr)  |  Source: $($u.Source)",
            "Password Last Set: $($u.PasswordLastSet)",
            "Last Logon: $($u.LastLogon)",
            '',
            'Description:',
            $u.Description,
            ''
        )
        try {
            $output = & net user $u.Name 2>&1
            $inGroups = $false
            $groupLines = @()
            foreach ($line in $output) {
                $s = [string]$line
                if ($s -match 'Local Group Memberships') {
                    $inGroups = $true
                    $parts = ($s -split '\*' | Select-Object -Skip 1) | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                    $groupLines += $parts
                }
                elseif ($s -match 'Global Group memberships') { $inGroups = $false }
                elseif ($inGroups -and $s.Trim().StartsWith('*')) {
                    $parts = ($s -split '\*' | Select-Object -Skip 1) | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                    $groupLines += $parts
                }
            }
            $lines += 'Group Memberships:'
            if ($groupLines.Count -gt 0) { foreach ($g in $groupLines) { $lines += "  $g" } }
            else { $lines += '  (none)' }
        }
        catch { $lines += 'Group Memberships: (could not query)' }

        $txtDetail.Text = $lines -join "`r`n"
    }.GetNewClosure())

    $gridGroups.Add_SelectionChanged({
        $g = $gridGroups.SelectedItem
        if (-not $g) { $txtDetail.Clear(); return }
        $lines = @(
            "Group: $($g.Name)",
            "SID: $($g.SID)",
            "Description: $($g.Description)",
            ''
        )
        try {
            $members = @(Get-LocalGroupMember -Group $g.Name -ErrorAction Stop)
            $lines += "Members ($($members.Count)):"
            if ($members.Count -gt 0) {
                foreach ($m in $members) {
                    $src = if ($m.PrincipalSource) { $m.PrincipalSource.ToString() } else { '' }
                    $lines += "  $($m.Name)  ($($m.ObjectClass), $src)"
                }
            }
            else { $lines += '  (empty)' }
        }
        catch {
            $lines += "Members: Error - $($_.Exception.Message)"
            $lines += '  (common with orphaned SIDs)'
        }
        $txtDetail.Text = $lines -join "`r`n"
    }.GetNewClosure())

    $mnuUsrCopyName.Add_Click({
        $u = $gridUsers.SelectedItem
        if ($u) { [System.Windows.Clipboard]::SetText([string]$u.Name) }
    }.GetNewClosure())
    $mnuUsrCopySid.Add_Click({
        $u = $gridUsers.SelectedItem
        if ($u) { [System.Windows.Clipboard]::SetText([string]$u.SID) }
    }.GetNewClosure())
    $mnuUsrCopyDetail.Add_Click({
        if ($txtDetail.Text) { [System.Windows.Clipboard]::SetText($txtDetail.Text) }
    }.GetNewClosure())
    $mnuGrpCopyName.Add_Click({
        $g = $gridGroups.SelectedItem
        if ($g) { [System.Windows.Clipboard]::SetText([string]$g.Name) }
    }.GetNewClosure())
    $mnuGrpCopySid.Add_Click({
        $g = $gridGroups.SelectedItem
        if ($g) { [System.Windows.Clipboard]::SetText([string]$g.SID) }
    }.GetNewClosure())
    $mnuGrpCopyDetail.Add_Click({
        if ($txtDetail.Text) { [System.Windows.Clipboard]::SetText($txtDetail.Text) }
    }.GetNewClosure())

    $view.Add_Loaded({
        if ($state.AllUsers.Count -eq 0 -and $state.AllGroups.Count -eq 0) {
            & $loadUsers
        }
    }.GetNewClosure())

    return $view
}
