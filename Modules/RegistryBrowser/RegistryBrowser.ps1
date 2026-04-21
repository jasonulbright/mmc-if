<#
.SYNOPSIS
    Registry Browser module for MMC-If (WPF).

.DESCRIPTION
    Regedit-like interface for browsing the Windows registry.
    Read-only for HKLM/HKCR/HKU/HKCC; read-write for HKCU (no UAC elevation).
    Uses Microsoft.Win32.RegistryKey API; does not require regedit.exe / mmc.exe.

    Pure .reg format helpers live in RegistryIO.ps1 (dot-sourced below) so
    they can be unit-tested without spinning up the WPF UserControl.
#>

. (Join-Path $PSScriptRoot 'RegistryIO.ps1')

function New-RegistryBrowserView {
    param([hashtable]$Context)

    $setStatus = $Context.SetStatus
    $log       = $Context.Log
    $window    = $Context.Window
    $prefs     = $Context.Prefs
    $savePrefs = $Context.SavePrefs

    # -----------------------------------------------------------------------
    # XAML load
    # -----------------------------------------------------------------------
    $xamlPath = Join-Path $PSScriptRoot 'RegistryBrowser.xaml'
    $xamlRaw  = Get-Content -LiteralPath $xamlPath -Raw
    $reader   = New-Object System.Xml.XmlNodeReader ([xml]$xamlRaw)
    $view     = [Windows.Markup.XamlReader]::Load($reader)

    $txtPath           = $view.FindName('txtPath')
    $btnGo             = $view.FindName('btnGo')
    $btnFind           = $view.FindName('btnFind')
    $btnFavorites      = $view.FindName('btnFavorites')
    $favMenu           = $view.FindName('favMenu')
    $mnuFavAdd         = $view.FindName('mnuFavAdd')
    $mnuFavManage      = $view.FindName('mnuFavManage')
    $mnuFavSep         = $view.FindName('mnuFavSep')
    $treeReg           = $view.FindName('treeReg')
    $gridValues        = $view.FindName('gridValues')
    $mnuTreeCopyPath   = $view.FindName('mnuTreeCopyPath')
    $mnuTreeRefresh    = $view.FindName('mnuTreeRefresh')
    $mnuTreeSep        = $view.FindName('mnuTreeSep')
    $mnuTreeNewKey     = $view.FindName('mnuTreeNewKey')
    $mnuTreeDeleteKey  = $view.FindName('mnuTreeDeleteKey')
    $mnuTreeExport     = $view.FindName('mnuTreeExport')
    $mnuTreeImport     = $view.FindName('mnuTreeImport')
    $mnuTreeAddFav     = $view.FindName('mnuTreeAddFav')
    $mnuValCopyName    = $view.FindName('mnuValCopyName')
    $mnuValCopyData    = $view.FindName('mnuValCopyData')
    $mnuValSep         = $view.FindName('mnuValSep')
    $mnuValNew         = $view.FindName('mnuValNew')
    $mnuValModify      = $view.FindName('mnuValModify')
    $mnuValDelete      = $view.FindName('mnuValDelete')

    # -----------------------------------------------------------------------
    # Dialog functions captured as scriptblocks.
    # Calling them by name from inside event-handler closures is fragile -
    # the scriptblock's session state binding doesn't always resolve the
    # function. Capturing as a local scriptblock variable makes the closure
    # snapshot the reference and invoke via `& $sb` with no name lookup.
    # See feedback_ps_wpf_handler_rules.md.
    # -----------------------------------------------------------------------
    $showInputDialogSb          = ${function:Show-InputDialog}
    $showRegistryValueEditorSb  = ${function:Show-RegistryValueEditor}
    $showRegistrySearchDialogSb = ${function:Show-RegistrySearchDialog}
    $showRegImportPreviewSb     = ${function:Show-RegImportPreview}
    $showFavoritesManagerSb     = ${function:Show-FavoritesManager}

    # -----------------------------------------------------------------------
    # Hive map
    # -----------------------------------------------------------------------
    $hiveMap = [ordered]@{
        'HKCU' = @{ Key = [Microsoft.Win32.Registry]::CurrentUser;   ReadOnly = $false }
        'HKLM' = @{ Key = [Microsoft.Win32.Registry]::LocalMachine;  ReadOnly = $true  }
        'HKCR' = @{ Key = [Microsoft.Win32.Registry]::ClassesRoot;   ReadOnly = $true  }
        'HKU'  = @{ Key = [Microsoft.Win32.Registry]::Users;         ReadOnly = $true  }
        'HKCC' = @{ Key = [Microsoft.Win32.Registry]::CurrentConfig; ReadOnly = $true  }
    }

    $state = @{
        CurrentPath = ''
        Values      = New-Object System.Collections.ObjectModel.ObservableCollection[object]
        SearchText  = ''
        SearchKeys  = $true
        SearchVals  = $true
        SearchData  = $true
    }
    $gridValues.ItemsSource = $state.Values

    # -----------------------------------------------------------------------
    # Helpers
    # -----------------------------------------------------------------------
    $isWritable = { param([string]$Path) $Path -match '^HKCU(\\|$)' }

    $openRegKey = {
        param([string]$FullPath, [bool]$Writable = $false)
        $parts = $FullPath -split '\\', 2
        $hiveName = $parts[0]
        $subPath = if ($parts.Count -gt 1) { $parts[1] } else { $null }
        if (-not $hiveMap.Contains($hiveName)) { return $null }
        $hive = $hiveMap[$hiveName]
        if (-not $subPath) { return $hive.Key }
        $actualWritable = $Writable -and (-not $hive.ReadOnly)
        try { $hive.Key.OpenSubKey($subPath, $actualWritable) }
        catch { $null }
    }.GetNewClosure()

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

    $formatValue = {
        param($Value, [Microsoft.Win32.RegistryValueKind]$Kind)
        if ($null -eq $Value) { return '(value not set)' }
        switch ($Kind) {
            'Binary' {
                if ($Value -is [byte[]]) { return ($Value | ForEach-Object { '{0:X2}' -f $_ }) -join ' ' }
                return $Value.ToString()
            }
            'MultiString' {
                if ($Value -is [string[]]) { return $Value -join ' | ' }
                return $Value.ToString()
            }
            'DWord' { return '0x{0:X8} ({0})' -f [uint32]$Value }
            'QWord' { return '0x{0:X16} ({0})' -f [uint64]$Value }
            default { return $Value.ToString() }
        }
    }

    # -----------------------------------------------------------------------
    # .reg format helpers: pure functions live in RegistryIO.ps1
    # (ConvertTo-RegValueLine, ConvertFrom-RegFileText, ConvertTo-RegLongHivePath).
    # -----------------------------------------------------------------------

    # -----------------------------------------------------------------------
    # Recursive .reg export
    # -----------------------------------------------------------------------
    # Returns a single .reg file string (CRLF line endings, minus the BOM + header
    # which the caller prepends once). Skips subkeys that fail to open so an export
    # can still succeed in partial form.
    $exportSubtreeLines = {
        param([string]$Path, [System.Collections.Generic.List[string]]$Lines)
        $regKey = & $openRegKey $Path $false
        if ($null -eq $regKey) { return }
        try {
            $longPath = ConvertTo-RegLongHivePath -Path $Path
            $Lines.Add("[$longPath]")

            # Default value first if it exists
            $defaultVal = $regKey.GetValue('')
            if ($null -ne $defaultVal) {
                try {
                    $defaultKind = $regKey.GetValueKind('')
                    $Lines.Add((ConvertTo-RegValueLine -Name '' -Kind $defaultKind -Value $defaultVal))
                } catch { }
            }

            # Named values
            $valueNames = $regKey.GetValueNames() | Where-Object { $_ -ne '' } | Sort-Object
            foreach ($vn in $valueNames) {
                try {
                    $kind = $regKey.GetValueKind($vn)
                    $val = $regKey.GetValue($vn, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
                    $Lines.Add((ConvertTo-RegValueLine -Name $vn -Kind $kind -Value $val))
                }
                catch { }
            }
            $Lines.Add('')

            # Recurse into subkeys
            $subNames = @()
            try { $subNames = $regKey.GetSubKeyNames() | Sort-Object } catch { }
            foreach ($sn in $subNames) {
                & $exportSubtreeLines "$Path\$sn" $Lines
            }
        }
        finally {
            $parts = $Path -split '\\', 2
            if ($parts.Count -gt 1) { $regKey.Dispose() }
        }
    }.GetNewClosure()

    # -----------------------------------------------------------------------
    # .reg file parser (thin wrapper: reads file bytes, decodes, delegates to
    # ConvertFrom-RegFileText in RegistryIO.ps1). Returns list of operations.
    # -----------------------------------------------------------------------
    $parseRegFile = {
        param([string]$FilePath)
        $rawBytes = [System.IO.File]::ReadAllBytes($FilePath)
        $text = ''
        if ($rawBytes.Count -ge 2 -and $rawBytes[0] -eq 0xFF -and $rawBytes[1] -eq 0xFE) {
            $text = [System.Text.Encoding]::Unicode.GetString($rawBytes, 2, $rawBytes.Count - 2)
        }
        elseif ($rawBytes.Count -ge 3 -and $rawBytes[0] -eq 0xEF -and $rawBytes[1] -eq 0xBB -and $rawBytes[2] -eq 0xBF) {
            $text = [System.Text.Encoding]::UTF8.GetString($rawBytes, 3, $rawBytes.Count - 3)
        }
        elseif ($rawBytes.Count -ge 2 -and $rawBytes[1] -eq 0) {
            $text = [System.Text.Encoding]::Unicode.GetString($rawBytes)
        }
        else {
            $text = [System.Text.Encoding]::Default.GetString($rawBytes)
        }
        ConvertFrom-RegFileText -Text $text
    }

    # -----------------------------------------------------------------------
    # Apply parsed operations. HKCU-only enforcement is assumed done by caller;
    # this function re-validates as defense in depth and refuses anything else.
    # -----------------------------------------------------------------------
    $applyRegOperations = {
        param([array]$Ops)
        $applied = 0; $failed = 0
        foreach ($op in $Ops) {
            $path = [string]$op.Path
            if (-not (& $isWritable $path)) {
                & $log "SKIPPED (outside HKCU): $($op.Op) $path" 'WARN'
                $failed++
                continue
            }
            try {
                switch ($op.Op) {
                    'CreateKey' {
                        $parts = $path -split '\\', 2
                        if ($parts.Count -lt 2) { throw "Refusing to create hive root: $path" }
                        $hive = $hiveMap[$parts[0]].Key
                        $k = $hive.CreateSubKey($parts[1])
                        if ($k) { $k.Dispose() }
                        $applied++
                    }
                    'DeleteKey' {
                        $parts = $path -split '\\', 2
                        if ($parts.Count -lt 2) { throw "Refusing to delete hive root: $path" }
                        $hive = $hiveMap[$parts[0]].Key
                        try { $hive.DeleteSubKeyTree($parts[1], $false) } catch { }
                        $applied++
                    }
                    'SetValue' {
                        $parts = $path -split '\\', 2
                        if ($parts.Count -lt 2) { throw "Cannot set value on hive root: $path" }
                        $hive = $hiveMap[$parts[0]].Key
                        $k = $hive.CreateSubKey($parts[1])
                        try { $k.SetValue($op.Name, $op.Value, $op.Kind) }
                        finally { if ($k) { $k.Dispose() } }
                        $applied++
                    }
                    'DeleteValue' {
                        $k = & $openRegKey $path $true
                        if ($null -eq $k) { throw "Cannot open for write: $path" }
                        try { $k.DeleteValue($op.Name, $false) } catch { }
                        finally {
                            $parts = $path -split '\\', 2
                            if ($parts.Count -gt 1) { $k.Dispose() }
                        }
                        $applied++
                    }
                }
            }
            catch {
                & $log "FAILED: $($op.Op) $path - $($_.Exception.Message)" 'ERROR'
                $failed++
            }
        }
        return @{ Applied = $applied; Failed = $failed }
    }.GetNewClosure()

    # Favorites helpers defined later, after $navigateToPath. Indirection hashtable
    # so the click handler assigned at definition time can invoke navigate even
    # though the closure snapshot happened before $navigateToPath existed.
    $favRefs = @{ ClickHandler = $null }

    # -----------------------------------------------------------------------
    # Tree node helpers
    # -----------------------------------------------------------------------
    $makeTreeNode = {
        param([string]$Label, [string]$Path, [bool]$HasChildren)
        $node = New-Object System.Windows.Controls.TreeViewItem
        $node.Header = $Label
        $node.Tag = $Path
        if ($HasChildren) {
            $dummy = New-Object System.Windows.Controls.TreeViewItem
            $dummy.Header = '(loading)'
            $dummy.Tag = '__dummy__'
            [void]$node.Items.Add($dummy)
        }
        $node
    }

    # Hashtable indirection so $loadChildren can reach $expandHandler even though
    # the handler is assigned after the closure is created (GetNewClosure snapshots
    # values, but the hashtable reference is stable).
    $handlers = @{ Expand = $null }

    $loadChildren = {
        param([System.Windows.Controls.TreeViewItem]$Node)
        $path = [string]$Node.Tag
        $Node.Items.Clear()
        $regKey = $null
        try {
            $regKey = & $openRegKey $path $false
            if ($null -eq $regKey) {
                $err = New-Object System.Windows.Controls.TreeViewItem
                $err.Header = '(Access Denied)'
                $err.Foreground = [System.Windows.Media.Brushes]::IndianRed
                [void]$Node.Items.Add($err)
                return
            }
            $subNames = @()
            try { $subNames = $regKey.GetSubKeyNames() | Sort-Object } catch { }
            foreach ($name in $subNames) {
                $childPath = "$path\$name"
                $hasChildren = $false
                $childKey = $null
                try {
                    $childKey = $regKey.OpenSubKey($name, $false)
                    if ($null -ne $childKey -and $childKey.SubKeyCount -gt 0) { $hasChildren = $true }
                } catch { $hasChildren = $true }  # show expand arrow on access-denied so user sees it's there
                finally { if ($null -ne $childKey) { $childKey.Dispose() } }

                $child = & $makeTreeNode $name $childPath $hasChildren
                $child.Add_Expanded($handlers.Expand)
                [void]$Node.Items.Add($child)
            }
        }
        catch {
            $err = New-Object System.Windows.Controls.TreeViewItem
            $err.Header = "(Error: $($_.Exception.Message))"
            $err.Foreground = [System.Windows.Media.Brushes]::IndianRed
            [void]$Node.Items.Add($err)
        }
        finally {
            $parts = $path -split '\\', 2
            if ($null -ne $regKey -and $parts.Count -gt 1) { $regKey.Dispose() }
        }
    }.GetNewClosure()

    $handlers.Expand = {
        param($sender, $e)
        $node = $sender
        if ($node.Items.Count -eq 1 -and $node.Items[0].Tag -eq '__dummy__') {
            & $loadChildren $node
        }
    }.GetNewClosure()

    # Populate root hives
    foreach ($hiveName in $hiveMap.Keys) {
        $info = $hiveMap[$hiveName]
        $label = if ($info.ReadOnly) { "$hiveName (Read-Only)" } else { $hiveName }
        $root = & $makeTreeNode $label $hiveName $true
        $root.Add_Expanded($handlers.Expand)
        [void]$treeReg.Items.Add($root)
    }

    # -----------------------------------------------------------------------
    # Populate values grid on selection
    # -----------------------------------------------------------------------
    $populateValues = {
        param([string]$Path)
        $state.CurrentPath = $Path
        $txtPath.Text = $Path
        $txtPath.CaretIndex = $Path.Length
        & $setStatus $Path
        $state.Values.Clear()

        $regKey = $null
        try {
            $regKey = & $openRegKey $Path $false
            if ($null -eq $regKey) {
                $state.Values.Add([pscustomobject]@{
                    Name = '(Access Denied)'; TypeName = ''; DisplayData = 'Cannot read values from this key'
                })
                return
            }
            $defaultVal = $regKey.GetValue('')
            $defaultKind = [Microsoft.Win32.RegistryValueKind]::String
            try { $defaultKind = $regKey.GetValueKind('') } catch { }
            $state.Values.Add([pscustomobject]@{
                Name        = '(Default)'
                TypeName    = & $getTypeName $defaultKind
                DisplayData = if ($null -eq $defaultVal) { '(value not set)' } else { & $formatValue $defaultVal $defaultKind }
            })

            $valueNames = $regKey.GetValueNames() | Where-Object { $_ -ne '' } | Sort-Object
            foreach ($vName in $valueNames) {
                try {
                    $val = $regKey.GetValue($vName, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
                    $kind = $regKey.GetValueKind($vName)
                    $state.Values.Add([pscustomobject]@{
                        Name        = $vName
                        TypeName    = & $getTypeName $kind
                        DisplayData = & $formatValue $val $kind
                    })
                }
                catch {
                    $state.Values.Add([pscustomobject]@{
                        Name = $vName; TypeName = 'ERROR'; DisplayData = $_.Exception.Message
                    })
                }
            }
            & $setStatus "$Path  |  $($regKey.SubKeyCount) subkey(s), $($regKey.ValueCount) value(s)"
        }
        catch [System.Security.SecurityException] {
            $state.Values.Add([pscustomobject]@{
                Name = '(Access Denied)'; TypeName = ''; DisplayData = 'Insufficient permissions'
            })
        }
        catch {
            $state.Values.Add([pscustomobject]@{
                Name = '(Error)'; TypeName = ''; DisplayData = $_.Exception.Message
            })
        }
        finally {
            $parts = $Path -split '\\', 2
            if ($null -ne $regKey -and $parts.Count -gt 1) { $regKey.Dispose() }
        }
    }.GetNewClosure()

    $treeReg.Add_SelectedItemChanged({
        $sel = $treeReg.SelectedItem
        if ($sel -and $sel.Tag -and $sel.Tag -ne '__dummy__') {
            & $populateValues ([string]$sel.Tag)
        }
    }.GetNewClosure())

    # -----------------------------------------------------------------------
    # Navigate-to-path
    # -----------------------------------------------------------------------
    $navigateToPath = {
        param([string]$TargetPath)
        $TargetPath = $TargetPath.Trim().TrimEnd('\')
        if ([string]::IsNullOrWhiteSpace($TargetPath)) { return }
        $TargetPath = $TargetPath `
            -replace '^HKEY_CURRENT_USER',  'HKCU' `
            -replace '^HKEY_LOCAL_MACHINE', 'HKLM' `
            -replace '^HKEY_CLASSES_ROOT',  'HKCR' `
            -replace '^HKEY_USERS',         'HKU'  `
            -replace '^HKEY_CURRENT_CONFIG','HKCC'
        $segments = $TargetPath -split '\\'
        $hiveName = $segments[0]

        $rootNode = $null
        foreach ($n in $treeReg.Items) {
            if (([string]$n.Tag) -eq $hiveName) { $rootNode = $n; break }
        }
        if (-not $rootNode) {
            [System.Windows.MessageBox]::Show("Invalid registry path: $TargetPath", 'Navigation',
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
            return
        }

        $currentNode = $rootNode
        $rootNode.IsExpanded = $true

        for ($i = 1; $i -lt $segments.Count; $i++) {
            $seg = $segments[$i]
            # Force child load if we still only have dummy
            if ($currentNode.Items.Count -eq 1 -and $currentNode.Items[0].Tag -eq '__dummy__') {
                & $loadChildren $currentNode
            }
            $found = $null
            foreach ($c in $currentNode.Items) {
                if ($c.Header -eq $seg) { $found = $c; break }
            }
            if (-not $found) {
                [System.Windows.MessageBox]::Show("Key not found: $seg`n`nPath: $TargetPath", 'Navigation',
                    [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
                return
            }
            $found.IsExpanded = $true
            $currentNode = $found
        }
        $currentNode.IsSelected = $true
        $currentNode.BringIntoView()
    }.GetNewClosure()

    $btnGo.Add_Click({ & $navigateToPath $txtPath.Text }.GetNewClosure())
    $txtPath.Add_KeyDown({
        if ($_.Key -eq [System.Windows.Input.Key]::Return) {
            $_.Handled = $true
            & $navigateToPath $txtPath.Text
        }
    }.GetNewClosure())

    # -----------------------------------------------------------------------
    # Favorites (load / save / rebuild menu / click handler)
    # -----------------------------------------------------------------------
    $getFavorites = {
        $favs = @()
        if ($null -ne $prefs.RegistryFavorites) {
            foreach ($f in $prefs.RegistryFavorites) {
                if ($f.Name -and $f.Path) {
                    $favs += [pscustomobject]@{ Name = [string]$f.Name; Path = [string]$f.Path }
                }
            }
        }
        ,$favs
    }.GetNewClosure()

    $setFavorites = {
        param([array]$Favs)
        $prefs.RegistryFavorites = @($Favs | ForEach-Object { @{ Name = $_.Name; Path = $_.Path } })
        if ($savePrefs) { & $savePrefs -Prefs $prefs }
    }.GetNewClosure()

    # Click handler for dynamic favorite menu items. Reads target from sender.Tag.
    $favRefs.ClickHandler = {
        param($sender, $e)
        $path = [string]$sender.Tag
        if ($path) { & $navigateToPath $path }
    }.GetNewClosure()

    $rebuildFavoritesMenu = {
        # Clear any items added after $mnuFavSep (dynamic entries from previous build).
        $toRemove = @()
        $sepSeen = $false
        foreach ($item in $favMenu.Items) {
            if ($sepSeen) { $toRemove += $item }
            if ($item -eq $mnuFavSep) { $sepSeen = $true }
        }
        foreach ($item in $toRemove) { [void]$favMenu.Items.Remove($item) }

        $favs = & $getFavorites
        if ($favs.Count -eq 0) {
            $empty = New-Object System.Windows.Controls.MenuItem
            $empty.Header = '(no favorites)'
            $empty.IsEnabled = $false
            [void]$favMenu.Items.Add($empty)
            return
        }
        foreach ($fav in $favs) {
            $item = New-Object System.Windows.Controls.MenuItem
            $item.Header = $fav.Name
            $item.ToolTip = $fav.Path
            $item.Tag = $fav.Path
            $item.Add_Click($favRefs.ClickHandler)
            [void]$favMenu.Items.Add($item)
        }
    }.GetNewClosure()

    $addCurrentToFavorites = {
        $path = $state.CurrentPath
        if ([string]::IsNullOrWhiteSpace($path)) {
            [System.Windows.MessageBox]::Show('Select a key first.', 'Favorites',
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
            return
        }
        $defaultName = ($path -split '\\')[-1]
        if (-not $defaultName) { $defaultName = $path }
        $name = & $showInputDialogSb -Title 'Add to Favorites' -Prompt "Favorite name for:`n$path" -Default $defaultName -Owner $window
        if ([string]::IsNullOrWhiteSpace($name)) { return }

        $favs = @(& $getFavorites)
        if ($favs | Where-Object { $_.Name -eq $name }) {
            $res = [System.Windows.MessageBox]::Show("A favorite named '$name' already exists. Replace it?",
                'Favorites', [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
            if ($res -ne [System.Windows.MessageBoxResult]::Yes) { return }
            $favs = @($favs | Where-Object { $_.Name -ne $name })
        }
        $favs += [pscustomobject]@{ Name = $name; Path = $path }
        & $setFavorites $favs
        & $rebuildFavoritesMenu
        & $log "Added favorite '$name' -> $path"
    }.GetNewClosure()

    $btnFavorites.Add_Click({
        $btnFavorites.ContextMenu.PlacementTarget = $btnFavorites
        $btnFavorites.ContextMenu.IsOpen = $true
    }.GetNewClosure())

    $mnuFavAdd.Add_Click({ & $addCurrentToFavorites }.GetNewClosure())

    $mnuFavManage.Add_Click({
        $current = @(& $getFavorites)
        $updated = & $showFavoritesManagerSb -Owner $window -Favorites $current
        if ($null -ne $updated) {
            & $setFavorites $updated
            & $rebuildFavoritesMenu
            & $log "Favorites updated ($(($updated | Measure-Object).Count) entries)"
        }
    }.GetNewClosure())

    # -----------------------------------------------------------------------
    # Tree context menu handlers
    # -----------------------------------------------------------------------
    $treeReg.ContextMenu.Add_Opened({
        $sel = $treeReg.SelectedItem
        $writable = $false
        $canDelete = $false
        $hasValidPath = $false
        if ($sel -and $sel.Tag -and $sel.Tag -ne '__dummy__') {
            $writable = & $isWritable ([string]$sel.Tag)
            $canDelete = $writable -and ((([string]$sel.Tag) -split '\\', 2).Count -ge 2)
            $hasValidPath = $true
        }
        $mnuTreeNewKey.Visibility    = if ($writable)  { 'Visible' } else { 'Collapsed' }
        $mnuTreeDeleteKey.Visibility = if ($writable)  { 'Visible' } else { 'Collapsed' }
        $mnuTreeSep.Visibility       = if ($writable)  { 'Visible' } else { 'Collapsed' }
        $mnuTreeDeleteKey.IsEnabled  = $canDelete
        # Export: available on any selected key (read-only hives OK)
        $mnuTreeExport.IsEnabled = $hasValidPath
        # Import: HKCU subtrees only
        $mnuTreeImport.Visibility = if ($writable) { 'Visible' } else { 'Collapsed' }
        # Add to Favorites: needs a real path
        $mnuTreeAddFav.IsEnabled = $hasValidPath
    }.GetNewClosure())

    $mnuTreeCopyPath.Add_Click({
        $sel = $treeReg.SelectedItem
        if ($sel -and $sel.Tag -and $sel.Tag -ne '__dummy__') {
            [System.Windows.Clipboard]::SetText([string]$sel.Tag)
        }
    }.GetNewClosure())

    $refreshNode = {
        param($Node)
        if (-not $Node) { return }
        $Node.Items.Clear()
        $dummy = New-Object System.Windows.Controls.TreeViewItem
        $dummy.Header = '(loading)'
        $dummy.Tag = '__dummy__'
        [void]$Node.Items.Add($dummy)
        $Node.IsExpanded = $false
        $Node.IsExpanded = $true
        if ($Node.Tag -and $Node.Tag -ne '__dummy__') {
            & $populateValues ([string]$Node.Tag)
        }
    }.GetNewClosure()

    $mnuTreeRefresh.Add_Click({ & $refreshNode $treeReg.SelectedItem }.GetNewClosure())

    $mnuTreeNewKey.Add_Click({
        $sel = $treeReg.SelectedItem
        if (-not $sel -or -not $sel.Tag) { return }
        $path = [string]$sel.Tag
        if (-not (& $isWritable $path)) { return }

        $name = & $showInputDialogSb -Title 'New Key' -Prompt 'Enter new key name:' -Default 'New Key #1' -Owner $window
        if ([string]::IsNullOrWhiteSpace($name)) { return }

        $regKey = $null
        try {
            $regKey = & $openRegKey $path $true
            if ($null -eq $regKey) {
                [System.Windows.MessageBox]::Show('Cannot open key for writing.', 'Error',
                    [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
                return
            }
            $newKey = $regKey.CreateSubKey($name)
            if ($newKey) { $newKey.Dispose() }
            & $log "Created registry key: $path\$name"
            & $refreshNode $sel
        }
        catch {
            [System.Windows.MessageBox]::Show("Failed to create key: $($_.Exception.Message)", 'Error',
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
        }
        finally {
            $parts = $path -split '\\', 2
            if ($null -ne $regKey -and $parts.Count -gt 1) { $regKey.Dispose() }
        }
    }.GetNewClosure())

    $mnuTreeDeleteKey.Add_Click({
        $sel = $treeReg.SelectedItem
        if (-not $sel -or -not $sel.Tag) { return }
        $path = [string]$sel.Tag
        if (-not (& $isWritable $path)) { return }
        $parts = $path -split '\\', 2
        if ($parts.Count -lt 2) { return }

        $res = [System.Windows.MessageBox]::Show(
            "Delete key '$($sel.Header)' and all its subkeys?`n`n$path",
            'Confirm Delete',
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning)
        if ($res -ne [System.Windows.MessageBoxResult]::Yes) { return }

        $parentPath = $path.Substring(0, $path.LastIndexOf('\'))
        $keyName = $path.Substring($path.LastIndexOf('\') + 1)
        $parentKey = $null
        try {
            $parentKey = & $openRegKey $parentPath $true
            if ($null -eq $parentKey) {
                [System.Windows.MessageBox]::Show('Cannot open parent for writing.', 'Error',
                    [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
                return
            }
            $parentKey.DeleteSubKeyTree($keyName)
            & $log "Deleted registry key: $path"
            $parent = $sel.Parent
            if ($parent -is [System.Windows.Controls.TreeViewItem]) {
                [void]$parent.Items.Remove($sel)
                $parent.IsSelected = $true
            }
        }
        catch {
            [System.Windows.MessageBox]::Show("Failed to delete: $($_.Exception.Message)", 'Error',
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
        }
        finally {
            $ppParts = $parentPath -split '\\', 2
            if ($null -ne $parentKey -and $ppParts.Count -gt 1) { $parentKey.Dispose() }
        }
    }.GetNewClosure())

    # -----------------------------------------------------------------------
    # Export / Import / Add-to-Favorites tree handlers
    # -----------------------------------------------------------------------
    $mnuTreeExport.Add_Click({
        $sel = $treeReg.SelectedItem
        if (-not $sel -or -not $sel.Tag -or $sel.Tag -eq '__dummy__') { return }
        $path = [string]$sel.Tag

        $dlg = New-Object Microsoft.Win32.SaveFileDialog
        $dlg.Filter = 'Registration Files (*.reg)|*.reg|All Files (*.*)|*.*'
        $dlg.DefaultExt = '.reg'
        $safeName = ($path -replace '[\\\/:*?"<>|]', '_')
        $dlg.FileName = "$safeName.reg"
        if ($window) { $ok = $dlg.ShowDialog($window) } else { $ok = $dlg.ShowDialog() }
        if (-not $ok) { return }

        & $setStatus "Exporting $path ..."
        try {
            $lines = New-Object System.Collections.Generic.List[string]
            & $exportSubtreeLines $path $lines
            $body = ($lines -join "`r`n")
            $fullText = "Windows Registry Editor Version 5.00`r`n`r`n" + $body
            if (-not $fullText.EndsWith("`r`n")) { $fullText += "`r`n" }
            # UTF-16 LE with BOM (standard .reg encoding)
            $enc = New-Object System.Text.UnicodeEncoding($false, $true)
            [System.IO.File]::WriteAllText($dlg.FileName, $fullText, $enc)
            & $log "Exported $path -> $($dlg.FileName) ($($lines.Count) lines)"
            & $setStatus "Exported to $($dlg.FileName)"
        }
        catch {
            [System.Windows.MessageBox]::Show("Export failed: $($_.Exception.Message)", 'Export',
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
            & $log "Export failed: $($_.Exception.Message)" 'ERROR'
            & $setStatus "Export failed"
        }
    }.GetNewClosure())

    $mnuTreeImport.Add_Click({
        $sel = $treeReg.SelectedItem
        if (-not $sel -or -not $sel.Tag) { return }
        $path = [string]$sel.Tag
        if (-not (& $isWritable $path)) {
            [System.Windows.MessageBox]::Show('Import is only allowed into HKCU.', 'Import',
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
            return
        }

        $dlg = New-Object Microsoft.Win32.OpenFileDialog
        $dlg.Filter = 'Registration Files (*.reg)|*.reg|All Files (*.*)|*.*'
        $dlg.Multiselect = $false
        if ($window) { $ok = $dlg.ShowDialog($window) } else { $ok = $dlg.ShowDialog() }
        if (-not $ok) { return }

        & $setStatus "Parsing $($dlg.FileName) ..."
        $parsed = $null
        try { $parsed = & $parseRegFile $dlg.FileName }
        catch {
            [System.Windows.MessageBox]::Show("Parse failed: $($_.Exception.Message)", 'Import',
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
            return
        }
        if (-not $parsed -or $parsed.Ops.Count -eq 0) {
            [System.Windows.MessageBox]::Show('No operations found in file.', 'Import',
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
            return
        }

        # Enforce HKCU-only BEFORE prompting
        $nonHkcu = @($parsed.Ops | Where-Object { -not (& $isWritable $_.Path) })
        if ($nonHkcu.Count -gt 0) {
            $sample = ($nonHkcu | Select-Object -First 5 | ForEach-Object { "  $($_.Op) $($_.Path)" }) -join "`r`n"
            [System.Windows.MessageBox]::Show(
                "Import refused. File contains $($nonHkcu.Count) operation(s) outside HKEY_CURRENT_USER:`r`n`r`n$sample$(if ($nonHkcu.Count -gt 5) { "`r`n  ..." })",
                'Import',
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Stop) | Out-Null
            & $log "Import refused ($($nonHkcu.Count) non-HKCU ops): $($dlg.FileName)" 'WARN'
            return
        }

        $confirmed = & $showRegImportPreviewSb -Owner $window -FilePath $dlg.FileName -Parsed $parsed
        if (-not $confirmed) { return }

        & $setStatus 'Applying import ...'
        $result = & $applyRegOperations $parsed.Ops
        & $log "Import complete: $($result.Applied) applied, $($result.Failed) failed ($($dlg.FileName))"
        & $setStatus "Imported: $($result.Applied) applied, $($result.Failed) failed"
        # Refresh the currently selected node
        & $refreshNode $treeReg.SelectedItem
    }.GetNewClosure())

    $mnuTreeAddFav.Add_Click({
        $sel = $treeReg.SelectedItem
        if ($sel -and $sel.Tag -and $sel.Tag -ne '__dummy__') {
            $state.CurrentPath = [string]$sel.Tag
        }
        & $addCurrentToFavorites
    }.GetNewClosure())

    # -----------------------------------------------------------------------
    # Values context menu handlers
    # -----------------------------------------------------------------------
    $gridValues.ContextMenu.Add_Opened({
        $writable = & $isWritable $state.CurrentPath
        $mnuValNew.Visibility    = if ($writable) { 'Visible' } else { 'Collapsed' }
        $mnuValModify.Visibility = if ($writable) { 'Visible' } else { 'Collapsed' }
        $mnuValDelete.Visibility = if ($writable) { 'Visible' } else { 'Collapsed' }
        $mnuValSep.Visibility    = if ($writable) { 'Visible' } else { 'Collapsed' }
    }.GetNewClosure())

    $mnuValCopyName.Add_Click({
        $sel = $gridValues.SelectedItem
        if ($sel -and $sel.Name) { [System.Windows.Clipboard]::SetText([string]$sel.Name) }
    }.GetNewClosure())
    $mnuValCopyData.Add_Click({
        $sel = $gridValues.SelectedItem
        if ($sel -and $sel.DisplayData) { [System.Windows.Clipboard]::SetText([string]$sel.DisplayData) }
    }.GetNewClosure())

    $openValueEditor = {
        param([string]$ValueName, [Microsoft.Win32.RegistryValueKind]$Kind, $CurrentValue, [bool]$IsEdit)
        $path = $state.CurrentPath
        if (-not (& $isWritable $path)) { return }

        $result = & $showRegistryValueEditorSb -Owner $window -ValueName $ValueName -Kind $Kind `
            -CurrentValue $CurrentValue -IsEdit $IsEdit -TypeNameFn $getTypeName
        if (-not $result) { return }

        $regKey = $null
        try {
            $regKey = & $openRegKey $path $true
            if ($null -eq $regKey) {
                [System.Windows.MessageBox]::Show('Cannot open key for writing.', 'Error',
                    [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
                return
            }
            $regKey.SetValue($result.Name, $result.Value, $Kind)
            $action = if ($IsEdit) { 'Modified' } else { 'Created' }
            & $log "$action registry value: $path\$($result.Name) ($(& $getTypeName $Kind))"
            & $populateValues $path
        }
        catch {
            [System.Windows.MessageBox]::Show("Failed to write value: $($_.Exception.Message)", 'Error',
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
        }
        finally {
            $parts = $path -split '\\', 2
            if ($null -ne $regKey -and $parts.Count -gt 1) { $regKey.Dispose() }
        }
    }.GetNewClosure()

    foreach ($mnuPair in @(
        @{ Menu = $view.FindName('mnuValNewString');       Kind = 'String' }
        @{ Menu = $view.FindName('mnuValNewDWord');        Kind = 'DWord' }
        @{ Menu = $view.FindName('mnuValNewQWord');        Kind = 'QWord' }
        @{ Menu = $view.FindName('mnuValNewBinary');       Kind = 'Binary' }
        @{ Menu = $view.FindName('mnuValNewMultiString');  Kind = 'MultiString' }
        @{ Menu = $view.FindName('mnuValNewExpandString'); Kind = 'ExpandString' }
    )) {
        $menuItem = $mnuPair.Menu
        $kindName = $mnuPair.Kind
        $handler = [scriptblock]::Create("& `$openValueEditor '' ([Microsoft.Win32.RegistryValueKind]::$kindName) `$null `$false")
        $menuItem.Add_Click($handler.GetNewClosure())
    }

    $mnuValModify.Add_Click({
        $sel = $gridValues.SelectedItem
        if (-not $sel) { return }
        $name = [string]$sel.Name
        if ($name -eq '(Default)') { $name = '' }
        $path = $state.CurrentPath
        $regKey = $null
        try {
            $regKey = & $openRegKey $path $false
            if ($null -eq $regKey) { return }
            $kind = $regKey.GetValueKind($name)
            $val = $regKey.GetValue($name, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
            & $openValueEditor $name $kind $val $true
        }
        catch {
            [System.Windows.MessageBox]::Show("Cannot read value: $($_.Exception.Message)", 'Error',
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
        }
        finally {
            $parts = $path -split '\\', 2
            if ($null -ne $regKey -and $parts.Count -gt 1) { $regKey.Dispose() }
        }
    }.GetNewClosure())

    $mnuValDelete.Add_Click({
        $sel = $gridValues.SelectedItem
        if (-not $sel) { return }
        $name = [string]$sel.Name
        $actualName = if ($name -eq '(Default)') { '' } else { $name }
        $path = $state.CurrentPath
        $res = [System.Windows.MessageBox]::Show("Delete value '$name'?", 'Confirm Delete',
            [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
        if ($res -ne [System.Windows.MessageBoxResult]::Yes) { return }

        $regKey = $null
        try {
            $regKey = & $openRegKey $path $true
            if ($null -eq $regKey) {
                [System.Windows.MessageBox]::Show('Cannot open key for writing.', 'Error',
                    [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
                return
            }
            $regKey.DeleteValue($actualName, $false)
            & $log "Deleted registry value: $path\$name"
            & $populateValues $path
        }
        catch {
            [System.Windows.MessageBox]::Show("Failed to delete: $($_.Exception.Message)", 'Error',
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
        }
        finally {
            $parts = $path -split '\\', 2
            if ($null -ne $regKey -and $parts.Count -gt 1) { $regKey.Dispose() }
        }
    }.GetNewClosure())

    # Double-click to modify (HKCU only)
    $gridValues.Add_MouseDoubleClick({
        if (-not (& $isWritable $state.CurrentPath)) { return }
        $mnuValModify.RaiseEvent(([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.MenuItem]::ClickEvent)))
    }.GetNewClosure())

    # -----------------------------------------------------------------------
    # Find dialog + keyboard shortcuts
    # -----------------------------------------------------------------------
    $searchRegistry = {
        param([string]$StartPath, [string]$Needle, [bool]$InKeys, [bool]$InValues, [bool]$InData)
        $regKey = $null
        try {
            $regKey = & $openRegKey $StartPath $false
            if ($null -eq $regKey) { return $null }
            if ($InValues -or $InData) {
                $vnames = @()
                try { $vnames = $regKey.GetValueNames() } catch { }
                foreach ($vn in $vnames) {
                    if ($InValues -and $vn -like "*$Needle*") { return @{ Path = $StartPath; Hit = "value name: $vn" } }
                    if ($InData) {
                        try {
                            $v = $regKey.GetValue($vn)
                            if ($null -ne $v -and $v.ToString() -like "*$Needle*") {
                                return @{ Path = $StartPath; Hit = "value data of '$vn'" }
                            }
                        } catch { }
                    }
                }
            }
            $subs = @()
            try { $subs = $regKey.GetSubKeyNames() | Sort-Object } catch { }
            foreach ($s in $subs) {
                $childPath = "$StartPath\$s"
                if ($InKeys -and $s -like "*$Needle*") { return @{ Path = $childPath; Hit = "key: $s" } }
                # One-level descent through child values only (full recursion is too slow)
                if ($InValues -or $InData) {
                    $ck = $null
                    try {
                        $ck = $regKey.OpenSubKey($s, $false)
                        if ($null -ne $ck) {
                            $cvs = @()
                            try { $cvs = $ck.GetValueNames() } catch { }
                            foreach ($cv in $cvs) {
                                if ($InValues -and $cv -like "*$Needle*") { return @{ Path = $childPath; Hit = "value name: $cv" } }
                                if ($InData) {
                                    try {
                                        $cvv = $ck.GetValue($cv)
                                        if ($null -ne $cvv -and $cvv.ToString() -like "*$Needle*") {
                                            return @{ Path = $childPath; Hit = "value data of '$cv'" }
                                        }
                                    } catch { }
                                }
                            }
                        }
                    } catch { }
                    finally { if ($null -ne $ck) { $ck.Dispose() } }
                }
            }
        }
        catch { }
        finally {
            $parts = $StartPath -split '\\', 2
            if ($null -ne $regKey -and $parts.Count -gt 1) { $regKey.Dispose() }
        }
        return $null
    }.GetNewClosure()

    $showFindDialog = {
        $sel = $treeReg.SelectedItem
        $startPath = if ($sel -and $sel.Tag -and $sel.Tag -ne '__dummy__') { [string]$sel.Tag } else { 'HKCU' }
        $result = & $showRegistrySearchDialogSb -Owner $window -StartPath $startPath -State $state
        if (-not $result) { return }
        $state.SearchText = $result.Text
        $state.SearchKeys = $result.InKeys
        $state.SearchVals = $result.InValues
        $state.SearchData = $result.InData
        & $setStatus "Searching for '$($result.Text)' under $startPath..."
        $hit = & $searchRegistry $startPath $result.Text $result.InKeys $result.InValues $result.InData
        if ($hit) {
            & $navigateToPath $hit.Path
            & $setStatus "Found $($hit.Hit) in $($hit.Path)"
        }
        else {
            & $setStatus "No matches for '$($result.Text)' under $startPath"
            [System.Windows.MessageBox]::Show("No match found.", 'Find',
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
        }
    }.GetNewClosure()

    $btnFind.Add_Click({ & $showFindDialog }.GetNewClosure())

    $view.Add_KeyDown({
        if ($_.Key -eq [System.Windows.Input.Key]::F5) {
            $_.Handled = $true
            & $refreshNode $treeReg.SelectedItem
            return
        }
        if ($_.KeyboardDevice.Modifiers -band [System.Windows.Input.ModifierKeys]::Control) {
            if ($_.Key -eq [System.Windows.Input.Key]::F) {
                $_.Handled = $true
                & $showFindDialog
            }
        }
    }.GetNewClosure())

    # Initial path
    $txtPath.Text = 'HKCU'

    # Populate favorites menu now that all dependencies (navigate, save, etc.) exist
    & $rebuildFavoritesMenu

    return $view
}

# =============================================================================
# Dialog: input box (replaces Microsoft.VisualBasic.Interaction.InputBox)
# =============================================================================
function Show-InputDialog {
    param(
        [string]$Title = 'Input',
        [string]$Prompt = 'Value:',
        [string]$Default = '',
        [System.Windows.Window]$Owner
    )

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="Input"
    Width="440" Height="190"
    WindowStartupLocation="CenterOwner"
    TitleCharacterCasing="Normal"
    ShowIconOnTitleBar="False"
    ResizeMode="NoResize"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml"/>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml"/>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml"/>
            </ResourceDictionary.MergedDictionaries>
        </ResourceDictionary>
    </Window.Resources>
    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Grid.Row="0" x:Name="lblPrompt" Margin="0,0,0,8"/>
        <TextBox Grid.Row="1" x:Name="txtInput" Padding="6,4,6,4"/>
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,14,0,0">
            <Button Controls:ControlsHelper.ContentCharacterCasing="Normal" x:Name="btnOK"
                    Content="OK" MinWidth="90" Height="32" Margin="0,0,8,0" IsDefault="True"
                    Style="{DynamicResource MahApps.Styles.Button.Square.Accent}"/>
            <Button Controls:ControlsHelper.ContentCharacterCasing="Normal" x:Name="btnCancel"
                    Content="Cancel" MinWidth="90" Height="32" IsCancel="True"
                    Style="{DynamicResource MahApps.Styles.Button.Square}"/>
        </StackPanel>
    </Grid>
</Controls:MetroWindow>
'@

    [xml]$dlgXml = $dlgXaml
    $dlgReader = New-Object System.Xml.XmlNodeReader $dlgXml
    $dlg = [System.Windows.Markup.XamlReader]::Load($dlgReader)
    $dlg.Title = $Title

    if ($Owner) {
        $currentTheme = [ControlzEx.Theming.ThemeManager]::Current.DetectTheme($Owner)
        if ($currentTheme) { [ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($dlg, $currentTheme) | Out-Null }
        $dlg.Owner = $Owner
    }

    $lblPrompt = $dlg.FindName('lblPrompt')
    $txtInput  = $dlg.FindName('txtInput')
    $btnOK     = $dlg.FindName('btnOK')
    $btnCancel = $dlg.FindName('btnCancel')

    $lblPrompt.Text = $Prompt
    $txtInput.Text  = $Default

    $script:__dlgResult = $null
    $btnOK.Add_Click({
        $script:__dlgResult = $txtInput.Text
        $dlg.DialogResult = $true
        $dlg.Close()
    }.GetNewClosure())

    $dlg.Add_Loaded({ $txtInput.Focus() | Out-Null; $txtInput.SelectAll() }.GetNewClosure())
    [void]$dlg.ShowDialog()
    return $script:__dlgResult
}

# =============================================================================
# Dialog: registry value editor
# =============================================================================
function Show-RegistryValueEditor {
    param(
        [System.Windows.Window]$Owner,
        [string]$ValueName,
        [Microsoft.Win32.RegistryValueKind]$Kind,
        $CurrentValue,
        [bool]$IsEdit,
        [scriptblock]$TypeNameFn
    )

    $multiline = ($Kind -eq 'Binary' -or $Kind -eq 'MultiString')

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="Value"
    Width="520" Height="380"
    MinWidth="420" MinHeight="300"
    WindowStartupLocation="CenterOwner"
    TitleCharacterCasing="Normal"
    ShowIconOnTitleBar="False"
    ResizeMode="CanResize"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml"/>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml"/>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml"/>
            </ResourceDictionary.MergedDictionaries>
        </ResourceDictionary>
    </Window.Resources>
    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Grid.Row="0" Text="Value name:" Margin="0,0,0,4"/>
        <TextBox   Grid.Row="1" x:Name="txtName" Padding="6,4,6,4"/>
        <TextBlock Grid.Row="2" x:Name="lblType" Margin="0,10,0,4"
                   Foreground="{DynamicResource MahApps.Brushes.Gray5}"/>
        <TextBlock Grid.Row="3" Text="Value data:" Margin="0,6,0,4"/>
        <TextBox   Grid.Row="4" x:Name="txtData" Padding="6,4,6,4"
                   FontFamily="Cascadia Code, Consolas, Courier New"/>
        <StackPanel Grid.Row="5" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,14,0,0">
            <Button Controls:ControlsHelper.ContentCharacterCasing="Normal" x:Name="btnOK"
                    Content="OK" MinWidth="90" Height="32" Margin="0,0,8,0" IsDefault="True"
                    Style="{DynamicResource MahApps.Styles.Button.Square.Accent}"/>
            <Button Controls:ControlsHelper.ContentCharacterCasing="Normal" x:Name="btnCancel"
                    Content="Cancel" MinWidth="90" Height="32" IsCancel="True"
                    Style="{DynamicResource MahApps.Styles.Button.Square}"/>
        </StackPanel>
    </Grid>
</Controls:MetroWindow>
'@

    [xml]$dlgXml = $dlgXaml
    $dlgReader = New-Object System.Xml.XmlNodeReader $dlgXml
    $dlg = [System.Windows.Markup.XamlReader]::Load($dlgReader)
    $dlg.Title = if ($IsEdit) { 'Edit Value' } else { 'New Value' }

    if ($Owner) {
        $currentTheme = [ControlzEx.Theming.ThemeManager]::Current.DetectTheme($Owner)
        if ($currentTheme) { [ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($dlg, $currentTheme) | Out-Null }
        $dlg.Owner = $Owner
    }

    $txtName = $dlg.FindName('txtName')
    $lblType = $dlg.FindName('lblType')
    $txtData = $dlg.FindName('txtData')
    $btnOK   = $dlg.FindName('btnOK')

    $txtName.Text = $ValueName
    if ($IsEdit) { $txtName.IsReadOnly = $true }
    $lblType.Text = 'Type: ' + (& $TypeNameFn $Kind)

    if ($multiline) {
        $txtData.AcceptsReturn = $true
        $txtData.TextWrapping  = if ($Kind -eq 'Binary') { 'Wrap' } else { 'NoWrap' }
        $txtData.VerticalScrollBarVisibility = 'Auto'
    }

    if ($IsEdit -and $null -ne $CurrentValue) {
        switch ($Kind) {
            'Binary'      { if ($CurrentValue -is [byte[]])   { $txtData.Text = ($CurrentValue | ForEach-Object { '{0:X2}' -f $_ }) -join ' ' } }
            'MultiString' { if ($CurrentValue -is [string[]]) { $txtData.Text = $CurrentValue -join "`r`n" } }
            'DWord'       { $txtData.Text = ([uint32]$CurrentValue).ToString() }
            'QWord'       { $txtData.Text = ([uint64]$CurrentValue).ToString() }
            default       { $txtData.Text = [string]$CurrentValue }
        }
    }

    $script:__regResult = $null
    $btnOK.Add_Click({
        $vName = $txtName.Text
        if (-not $IsEdit -and [string]::IsNullOrWhiteSpace($vName)) {
            [System.Windows.MessageBox]::Show('Value name cannot be empty.', 'Validation',
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
            return
        }
        try {
            $val = switch ($Kind) {
                'DWord' { [int]$txtData.Text }
                'QWord' { [long]$txtData.Text }
                'Binary' {
                    $hex = $txtData.Text.Trim()
                    if ($hex) {
                        [byte[]]($hex -split '\s+' | ForEach-Object { [byte]("0x$_") })
                    } else { [byte[]]@() }
                }
                'MultiString' { [string[]]($txtData.Text -split "`r`n") }
                default { $txtData.Text }
            }
            $script:__regResult = @{ Name = $vName; Value = $val }
            $dlg.DialogResult = $true
            $dlg.Close()
        }
        catch {
            [System.Windows.MessageBox]::Show("Invalid value: $($_.Exception.Message)", 'Validation',
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
        }
    }.GetNewClosure())

    [void]$dlg.ShowDialog()
    return $script:__regResult
}

# =============================================================================
# Dialog: find
# =============================================================================
function Show-RegistrySearchDialog {
    param(
        [System.Windows.Window]$Owner,
        [string]$StartPath,
        [hashtable]$State
    )

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="Find"
    Width="460" Height="260"
    WindowStartupLocation="CenterOwner"
    TitleCharacterCasing="Normal"
    ShowIconOnTitleBar="False"
    ResizeMode="NoResize"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml"/>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml"/>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml"/>
            </ResourceDictionary.MergedDictionaries>
        </ResourceDictionary>
    </Window.Resources>
    <StackPanel Margin="16">
        <TextBlock x:Name="lblStart" Margin="0,0,0,10"
                   Foreground="{DynamicResource MahApps.Brushes.Gray5}"/>
        <TextBlock Text="Find what:"/>
        <TextBox   x:Name="txtFind" Padding="6,4,6,4" Margin="0,4,0,10"/>
        <StackPanel Orientation="Horizontal" Margin="0,0,0,14">
            <CheckBox x:Name="chkKeys" Content="Keys"   Margin="0,0,14,0"/>
            <CheckBox x:Name="chkVals" Content="Values" Margin="0,0,14,0"/>
            <CheckBox x:Name="chkData" Content="Data"/>
        </StackPanel>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Controls:ControlsHelper.ContentCharacterCasing="Normal" x:Name="btnFind"
                    Content="Find Next" MinWidth="110" Height="32" Margin="0,0,8,0" IsDefault="True"
                    Style="{DynamicResource MahApps.Styles.Button.Square.Accent}"/>
            <Button Controls:ControlsHelper.ContentCharacterCasing="Normal" x:Name="btnCancel"
                    Content="Cancel" MinWidth="90" Height="32" IsCancel="True"
                    Style="{DynamicResource MahApps.Styles.Button.Square}"/>
        </StackPanel>
    </StackPanel>
</Controls:MetroWindow>
'@

    [xml]$dlgXml = $dlgXaml
    $dlgReader = New-Object System.Xml.XmlNodeReader $dlgXml
    $dlg = [System.Windows.Markup.XamlReader]::Load($dlgReader)

    if ($Owner) {
        $currentTheme = [ControlzEx.Theming.ThemeManager]::Current.DetectTheme($Owner)
        if ($currentTheme) { [ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($dlg, $currentTheme) | Out-Null }
        $dlg.Owner = $Owner
    }

    $lblStart = $dlg.FindName('lblStart')
    $txtFind  = $dlg.FindName('txtFind')
    $chkKeys  = $dlg.FindName('chkKeys')
    $chkVals  = $dlg.FindName('chkVals')
    $chkData  = $dlg.FindName('chkData')
    $btnFind  = $dlg.FindName('btnFind')

    $lblStart.Text    = "Search under: $StartPath"
    $txtFind.Text     = $State.SearchText
    $chkKeys.IsChecked = $State.SearchKeys
    $chkVals.IsChecked = $State.SearchVals
    $chkData.IsChecked = $State.SearchData

    $script:__findResult = $null
    $btnFind.Add_Click({
        if ([string]::IsNullOrWhiteSpace($txtFind.Text)) { return }
        $script:__findResult = @{
            Text     = $txtFind.Text
            InKeys   = [bool]$chkKeys.IsChecked
            InValues = [bool]$chkVals.IsChecked
            InData   = [bool]$chkData.IsChecked
        }
        $dlg.DialogResult = $true
        $dlg.Close()
    }.GetNewClosure())

    $dlg.Add_Loaded({ $txtFind.Focus() | Out-Null }.GetNewClosure())
    [void]$dlg.ShowDialog()
    return $script:__findResult
}

# =============================================================================
# Dialog: .reg import preview
# Returns $true if the operator confirms apply, $false (or $null) otherwise.
# =============================================================================
function Show-RegImportPreview {
    param(
        [System.Windows.Window]$Owner,
        [string]$FilePath,
        [hashtable]$Parsed
    )

    $dlg = New-Object System.Windows.Window
    $dlg.Title = 'Import Preview'
    $dlg.Width = 720
    $dlg.Height = 480
    $dlg.WindowStartupLocation = 'CenterOwner'
    if ($Owner) { $dlg.Owner = $Owner }

    $root = New-Object System.Windows.Controls.Grid
    $root.Margin = '14'
    foreach ($h in @('Auto', 'Auto', '*', 'Auto', 'Auto')) {
        $rd = New-Object System.Windows.Controls.RowDefinition; $rd.Height = $h
        [void]$root.RowDefinitions.Add($rd)
    }

    $lblFile = New-Object System.Windows.Controls.TextBlock
    $lblFile.Text = "File: $FilePath"
    $lblFile.FontFamily = 'Cascadia Code, Consolas'
    $lblFile.Margin = '0,0,0,6'
    [System.Windows.Controls.Grid]::SetRow($lblFile, 0); [void]$root.Children.Add($lblFile)

    $opCount = $Parsed.Ops.Count
    $warnCount = $Parsed.Warnings.Count
    $lblSummary = New-Object System.Windows.Controls.TextBlock
    $lblSummary.Text = "$opCount operation(s), $warnCount warning(s). All keys are within HKEY_CURRENT_USER."
    $lblSummary.Margin = '0,0,0,8'
    $lblSummary.Foreground = [System.Windows.Media.Brushes]::Gray
    [System.Windows.Controls.Grid]::SetRow($lblSummary, 1); [void]$root.Children.Add($lblSummary)

    $list = New-Object System.Windows.Controls.DataGrid
    $list.AutoGenerateColumns = $false
    $list.CanUserAddRows = $false
    $list.CanUserDeleteRows = $false
    $list.IsReadOnly = $true
    $list.HeadersVisibility = 'Column'
    $list.GridLinesVisibility = 'Horizontal'
    $list.RowHeaderWidth = 0
    $list.FontFamily = 'Cascadia Code, Consolas'

    $colOp = New-Object System.Windows.Controls.DataGridTextColumn
    $colOp.Header = 'Operation'; $colOp.Width = 110
    $colOp.Binding = New-Object System.Windows.Data.Binding 'Op'
    [void]$list.Columns.Add($colOp)
    $colPath = New-Object System.Windows.Controls.DataGridTextColumn
    $colPath.Header = 'Path'; $colPath.Width = 340
    $colPath.Binding = New-Object System.Windows.Data.Binding 'Path'
    [void]$list.Columns.Add($colPath)
    $colName = New-Object System.Windows.Controls.DataGridTextColumn
    $colName.Header = 'Name'; $colName.Width = 120
    $colName.Binding = New-Object System.Windows.Data.Binding 'Name'
    [void]$list.Columns.Add($colName)
    $colKind = New-Object System.Windows.Controls.DataGridTextColumn
    $colKind.Header = 'Type'; $colKind.Width = '*'
    $colKind.Binding = New-Object System.Windows.Data.Binding 'Kind'
    [void]$list.Columns.Add($colKind)

    $rows = foreach ($op in $Parsed.Ops) {
        [pscustomobject]@{
            Op   = $op.Op
            Path = $op.Path
            Name = if ($op.ContainsKey('Name')) { if ([string]::IsNullOrEmpty($op.Name)) { '(Default)' } else { $op.Name } } else { '' }
            Kind = if ($op.ContainsKey('Kind')) { $op.Kind.ToString() } else { '' }
        }
    }
    $list.ItemsSource = @($rows)
    [System.Windows.Controls.Grid]::SetRow($list, 2); [void]$root.Children.Add($list)

    if ($warnCount -gt 0) {
        $warnBox = New-Object System.Windows.Controls.TextBox
        $warnBox.IsReadOnly = $true
        $warnBox.Height = 80
        $warnBox.Margin = '0,8,0,0'
        $warnBox.VerticalScrollBarVisibility = 'Auto'
        $warnBox.TextWrapping = 'Wrap'
        $warnBox.FontFamily = 'Cascadia Code, Consolas'
        $warnBox.Text = ($Parsed.Warnings -join "`r`n")
        [System.Windows.Controls.Grid]::SetRow($warnBox, 3); [void]$root.Children.Add($warnBox)
    }

    $buttons = New-Object System.Windows.Controls.StackPanel
    $buttons.Orientation = 'Horizontal'; $buttons.HorizontalAlignment = 'Right'
    $buttons.Margin = '0,14,0,0'
    [System.Windows.Controls.Grid]::SetRow($buttons, 4); [void]$root.Children.Add($buttons)

    $btnApply = New-Object System.Windows.Controls.Button
    $btnApply.Content = 'Apply'; $btnApply.IsDefault = $true
    $btnApply.MinWidth = 90; $btnApply.Margin = '0,0,6,0'
    [void]$buttons.Children.Add($btnApply)

    $btnCancel = New-Object System.Windows.Controls.Button
    $btnCancel.Content = 'Cancel'; $btnCancel.IsCancel = $true
    $btnCancel.MinWidth = 80
    [void]$buttons.Children.Add($btnCancel)

    $script:__regImportConfirm = $false
    $btnApply.Add_Click({
        $script:__regImportConfirm = $true
        $dlg.DialogResult = $true
        $dlg.Close()
    }.GetNewClosure())

    $dlg.Content = $root
    [void]$dlg.ShowDialog()
    return $script:__regImportConfirm
}

# =============================================================================
# Dialog: favorites manager
# Takes current favorites array; returns updated array (or $null if cancelled).
# =============================================================================
function Show-FavoritesManager {
    param(
        [System.Windows.Window]$Owner,
        [array]$Favorites
    )

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="Manage Favorites"
    Width="580" Height="400"
    MinWidth="500" MinHeight="320"
    WindowStartupLocation="CenterOwner"
    TitleCharacterCasing="Normal"
    ShowIconOnTitleBar="False"
    ResizeMode="CanResize"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml"/>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml"/>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml"/>
            </ResourceDictionary.MergedDictionaries>
        </ResourceDictionary>
    </Window.Resources>
    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Grid.Row="0" Text="Select a favorite to remove, then click Remove. Click OK to save changes."
                   Foreground="{DynamicResource MahApps.Brushes.Gray5}" Margin="0,0,0,8"/>
        <DataGrid Grid.Row="1" x:Name="gridFavs"
                  AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False"
                  IsReadOnly="True" SelectionMode="Extended"
                  HeadersVisibility="Column" GridLinesVisibility="Horizontal"
                  RowHeaderWidth="0" BorderThickness="0" ColumnHeaderHeight="30">
            <DataGrid.Columns>
                <DataGridTextColumn Header="Name" Width="180" Binding="{Binding Name}"/>
                <DataGridTextColumn Header="Path" Width="*"   Binding="{Binding Path}"/>
            </DataGrid.Columns>
        </DataGrid>
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,14,0,0">
            <Button Controls:ControlsHelper.ContentCharacterCasing="Normal" x:Name="btnRemove"
                    Content="Remove" MinWidth="100" Height="32" Margin="0,0,16,0"
                    Style="{DynamicResource MahApps.Styles.Button.Square}"/>
            <Button Controls:ControlsHelper.ContentCharacterCasing="Normal" x:Name="btnOK"
                    Content="OK" MinWidth="90" Height="32" Margin="0,0,8,0" IsDefault="True"
                    Style="{DynamicResource MahApps.Styles.Button.Square.Accent}"/>
            <Button Controls:ControlsHelper.ContentCharacterCasing="Normal" x:Name="btnCancel"
                    Content="Cancel" MinWidth="90" Height="32" IsCancel="True"
                    Style="{DynamicResource MahApps.Styles.Button.Square}"/>
        </StackPanel>
    </Grid>
</Controls:MetroWindow>
'@

    [xml]$dlgXml = $dlgXaml
    $dlgReader = New-Object System.Xml.XmlNodeReader $dlgXml
    $dlg = [System.Windows.Markup.XamlReader]::Load($dlgReader)

    if ($Owner) {
        $currentTheme = [ControlzEx.Theming.ThemeManager]::Current.DetectTheme($Owner)
        if ($currentTheme) { [ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($dlg, $currentTheme) | Out-Null }
        $dlg.Owner = $Owner
    }

    $gridFavs  = $dlg.FindName('gridFavs')
    $btnRemove = $dlg.FindName('btnRemove')
    $btnOK     = $dlg.FindName('btnOK')

    $items = New-Object System.Collections.ObjectModel.ObservableCollection[object]
    foreach ($f in $Favorites) {
        $items.Add([pscustomobject]@{ Name = $f.Name; Path = $f.Path })
    }
    $gridFavs.ItemsSource = $items

    $btnRemove.Add_Click({
        $selected = @($gridFavs.SelectedItems)
        foreach ($sel in $selected) { [void]$items.Remove($sel) }
    }.GetNewClosure())

    $script:__favResult = $null
    $btnOK.Add_Click({
        $script:__favResult = @($items | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Path = $_.Path } })
        $dlg.DialogResult = $true
        $dlg.Close()
    }.GetNewClosure())

    [void]$dlg.ShowDialog()
    return $script:__favResult
}
