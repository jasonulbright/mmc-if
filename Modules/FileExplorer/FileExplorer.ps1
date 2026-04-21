<#
.SYNOPSIS
    File Explorer module for MMC-If (WPF).

.DESCRIPTION
    Admin-oriented file browser: hidden files shown, extensions shown, sorted by type.
    Read-only navigation; no file operations. Archive browsing via 7-Zip when installed.
#>

function New-FileExplorerView {
    param([hashtable]$Context)

    $setStatus = $Context.SetStatus
    $log       = $Context.Log

    $xamlPath = Join-Path $PSScriptRoot 'FileExplorer.xaml'
    $xamlRaw  = Get-Content -LiteralPath $xamlPath -Raw
    $reader   = New-Object System.Xml.XmlNodeReader ([xml]$xamlRaw)
    $view     = [Windows.Markup.XamlReader]::Load($reader)

    $btnUp           = $view.FindName('btnUp')
    $btnGo           = $view.FindName('btnGo')
    $txtPath         = $view.FindName('txtPath')
    $treeDirs        = $view.FindName('treeDirs')
    $gridFiles       = $view.FindName('gridFiles')
    $mnuCopyPath     = $view.FindName('mnuCopyPath')
    $mnuOpenExplorer = $view.FindName('mnuOpenExplorer')

    # 7-Zip detection
    $state = @{
        SevenZipPath   = $null
        CurrentPath    = ''
        InArchive      = $false
        ArchivePath    = ''
        ArchiveSubPath = ''
        Rows           = New-Object System.Collections.ObjectModel.ObservableCollection[object]
    }
    foreach ($p in @("${env:ProgramFiles}\7-Zip\7z.exe", "${env:ProgramFiles(x86)}\7-Zip\7z.exe", 'C:\Program Files\7-Zip\7z.exe')) {
        if (Test-Path -LiteralPath $p) { $state.SevenZipPath = $p; break }
    }
    $gridFiles.ItemsSource = $state.Rows

    $archiveExts = @('.zip', '.7z', '.rar', '.tar', '.gz', '.bz2', '.xz', '.cab', '.iso', '.msi', '.wim')
    $iconFolder  = [char]::ConvertFromUtf32(0x1F4C1)
    $iconFile    = [char]::ConvertFromUtf32(0x1F4C4)
    $iconArchive = [char]::ConvertFromUtf32(0x1F4E6)

    $isArchive = {
        param([string]$FileName)
        [System.IO.Path]::GetExtension($FileName).ToLower() -in $archiveExts
    }.GetNewClosure()

    $formatSize = {
        param([long]$Bytes)
        if ($Bytes -lt 0) { return '' }
        if ($Bytes -ge 1GB) { return '{0:N2} GB' -f ($Bytes / 1GB) }
        if ($Bytes -ge 1MB) { return '{0:N2} MB' -f ($Bytes / 1MB) }
        if ($Bytes -ge 1KB) { return '{0:N2} KB' -f ($Bytes / 1KB) }
        return "$Bytes bytes"
    }

    # -----------------------------------------------------------------------
    # Tree population
    # -----------------------------------------------------------------------
    $makeTreeNode = {
        param([string]$Label, [string]$Path, [bool]$HasChildren, [string]$Color = $null)
        $n = New-Object System.Windows.Controls.TreeViewItem
        $n.Header = $Label
        $n.Tag = $Path
        if ($Color) {
            try { $n.Foreground = (New-Object System.Windows.Media.BrushConverter).ConvertFromString($Color) } catch { }
        }
        if ($HasChildren) {
            $d = New-Object System.Windows.Controls.TreeViewItem
            $d.Header = '(loading)'; $d.Tag = '__dummy__'
            [void]$n.Items.Add($d)
        }
        $n
    }

    # Hashtable indirection for lazy-load handler (see RegistryBrowser for rationale)
    $handlers = @{ Expand = $null }
    $loadSubdirs = {
        param([System.Windows.Controls.TreeViewItem]$Node)
        $Node.Items.Clear()
        $dirPath = [string]$Node.Tag
        try {
            $subDirs = [System.IO.Directory]::GetDirectories($dirPath) | Sort-Object
            foreach ($s in $subDirs) {
                $name = [System.IO.Path]::GetFileName($s)
                $hasKids = $false
                try { $hasKids = ([System.IO.Directory]::GetDirectories($s)).Count -gt 0 }
                catch { $hasKids = $true }  # show arrow even if access-denied

                $color = $null
                try {
                    $attrs = [System.IO.File]::GetAttributes($s)
                    if ($attrs -band [System.IO.FileAttributes]::Hidden) { $color = '#888888' }
                } catch { }

                $child = & $makeTreeNode $name $s $hasKids $color
                $child.Add_Expanded($handlers.Expand)
                [void]$Node.Items.Add($child)
            }
        }
        catch [System.UnauthorizedAccessException] {
            $err = New-Object System.Windows.Controls.TreeViewItem
            $err.Header = '(Access Denied)'
            $err.Foreground = [System.Windows.Media.Brushes]::IndianRed
            [void]$Node.Items.Add($err)
        }
        catch {
            $err = New-Object System.Windows.Controls.TreeViewItem
            $err.Header = "(Error: $($_.Exception.Message))"
            $err.Foreground = [System.Windows.Media.Brushes]::IndianRed
            [void]$Node.Items.Add($err)
        }
    }.GetNewClosure()

    $handlers.Expand = {
        param($sender, $e)
        $node = $sender
        if ($node.Items.Count -eq 1 -and $node.Items[0].Tag -eq '__dummy__') {
            & $loadSubdirs $node
        }
    }.GetNewClosure()

    # Populate roots
    foreach ($drive in ([System.IO.DriveInfo]::GetDrives() | Where-Object { $_.IsReady })) {
        $label = if ($drive.VolumeLabel) { "$($drive.Name.TrimEnd('\')) ($($drive.VolumeLabel))" } else { $drive.Name.TrimEnd('\') }
        $root = & $makeTreeNode $label $drive.Name $true
        $root.Add_Expanded($handlers.Expand)
        [void]$treeDirs.Items.Add($root)
    }
    foreach ($sd in @(
        @{ Name = 'Desktop';   Path = [Environment]::GetFolderPath('Desktop') }
        @{ Name = 'Documents'; Path = [Environment]::GetFolderPath('MyDocuments') }
        @{ Name = 'Downloads'; Path = (Join-Path $env:USERPROFILE 'Downloads') }
        @{ Name = 'Temp';      Path = $env:TEMP }
    )) {
        if (Test-Path -LiteralPath $sd.Path) {
            $node = & $makeTreeNode $sd.Name $sd.Path $true '#2A8AF6'
            $node.Add_Expanded($handlers.Expand)
            [void]$treeDirs.Items.Add($node)
        }
    }

    # -----------------------------------------------------------------------
    # Directory load
    # -----------------------------------------------------------------------
    $addFileRow = {
        param($Icon, $Name, $Type, $Size, $Modified, $Attributes, $FullPath, $IsDir)
        $state.Rows.Add([pscustomobject]@{
            Icon = $Icon; Name = $Name; Type = $Type; Size = $Size
            Modified = $Modified; Attributes = $Attributes
            FullPath = $FullPath; IsDir = $IsDir
        })
    }.GetNewClosure()

    $loadDirectory = {
        param([string]$DirPath)
        $state.Rows.Clear()
        $state.CurrentPath = $DirPath
        $state.InArchive = $false
        $txtPath.Text = $DirPath
        & $setStatus "Loading $DirPath..."

        $fileCount = 0
        $dirCount = 0
        try {
            $dirs = @()
            try { $dirs = [System.IO.Directory]::GetDirectories($DirPath) | Sort-Object } catch { }
            foreach ($d in $dirs) {
                $dirCount++
                $name = [System.IO.Path]::GetFileName($d)
                $attrs = ''; $modified = ''
                try {
                    $di = New-Object System.IO.DirectoryInfo($d)
                    $list = @()
                    if ($di.Attributes -band [System.IO.FileAttributes]::Hidden) { $list += 'H' }
                    if ($di.Attributes -band [System.IO.FileAttributes]::System) { $list += 'S' }
                    if ($di.Attributes -band [System.IO.FileAttributes]::ReadOnly) { $list += 'R' }
                    $attrs = $list -join ''
                    $modified = $di.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
                } catch { }
                & $addFileRow $iconFolder $name '<DIR>' '' $modified $attrs $d $true
            }

            $files = @()
            try {
                $di = New-Object System.IO.DirectoryInfo($DirPath)
                $files = @($di.GetFiles() | Sort-Object Extension, Name)
            } catch { $files = @() }
            foreach ($fi in $files) {
                $fileCount++
                $ext = $fi.Extension.ToLower()
                $size = ''; $modified = ''; $attrs = ''
                try {
                    $size = & $formatSize $fi.Length
                    $modified = $fi.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
                    $list = @()
                    if ($fi.Attributes -band [System.IO.FileAttributes]::Hidden)   { $list += 'H' }
                    if ($fi.Attributes -band [System.IO.FileAttributes]::System)   { $list += 'S' }
                    if ($fi.Attributes -band [System.IO.FileAttributes]::ReadOnly) { $list += 'R' }
                    if ($fi.Attributes -band [System.IO.FileAttributes]::Archive)  { $list += 'A' }
                    $attrs = $list -join ''
                } catch { }
                $icon = if (& $isArchive $fi.Name) { $iconArchive } else { $iconFile }
                $type = if ($ext) { $ext } else { '' }
                & $addFileRow $icon $fi.Name $type $size $modified $attrs $fi.FullName $false
            }

            & $setStatus "$DirPath  |  $dirCount folder(s), $fileCount file(s)"
        }
        catch [System.UnauthorizedAccessException] {
            & $addFileRow '' '(Access Denied)' '' '' '' '' '' $false
            & $setStatus "Access denied: $DirPath"
        }
        catch {
            & $setStatus "Error: $($_.Exception.Message)"
        }
    }.GetNewClosure()

    # -----------------------------------------------------------------------
    # Archive browse via 7-Zip
    # -----------------------------------------------------------------------
    $loadArchive = {
        param([string]$ArchivePath, [string]$SubPath)

        if (-not $state.SevenZipPath) {
            [System.Windows.MessageBox]::Show(
                '7-Zip not found. Install 7-Zip to browse archives.',
                '7-Zip Required',
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information) | Out-Null
            return
        }

        $state.Rows.Clear()
        $state.InArchive = $true
        $state.ArchivePath = $ArchivePath
        $state.ArchiveSubPath = $SubPath
        $displayPath = if ($SubPath) { "$ArchivePath\$SubPath" } else { "$ArchivePath\" }
        $txtPath.Text = $displayPath
        $state.CurrentPath = $displayPath
        & $setStatus "Listing archive: $ArchivePath..."

        try {
            $output = & $state.SevenZipPath l $ArchivePath -slt 2>&1

            $entries = @()
            $cur = $null
            foreach ($line in $output) {
                $s = [string]$line
                if ($s -match '^Path = (.+)$') {
                    if ($cur -and $cur.Path) { $entries += $cur }
                    $cur = @{ Path = $Matches[1]; Size = 0; Modified = ''; IsDir = $false; Attrs = '' }
                }
                elseif ($cur) {
                    if ($s -match '^Size = (\d+)$')      { $cur.Size = [long]$Matches[1] }
                    if ($s -match '^Modified = (.+)$')   { $cur.Modified = $Matches[1].Trim() }
                    if ($s -match '^Folder = \+$')       { $cur.IsDir = $true }
                    if ($s -match '^Attributes = (.+)$') { $cur.Attrs = $Matches[1].Trim() }
                }
            }
            if ($cur -and $cur.Path) { $entries += $cur }

            $targetPrefix = if ($SubPath) { "$SubPath\" } else { '' }
            $dirs = @{}
            $filesInDir = @()

            foreach ($entry in $entries) {
                $p = $entry.Path -replace '/', '\'
                if ($targetPrefix -and -not $p.StartsWith($targetPrefix, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
                if (-not $targetPrefix -and $p -eq ([System.IO.Path]::GetFileName($ArchivePath))) { continue }
                $rel = if ($targetPrefix) { $p.Substring($targetPrefix.Length) } else { $p }
                if ([string]::IsNullOrEmpty($rel)) { continue }
                $slash = $rel.IndexOf('\')
                if ($slash -ge 0) {
                    $dirs[$rel.Substring(0, $slash)] = $true
                }
                elseif ($entry.IsDir) {
                    $dirs[$rel] = $true
                }
                else {
                    $filesInDir += @{ Name = $rel; Size = $entry.Size; Modified = $entry.Modified; Attrs = $entry.Attrs }
                }
            }

            foreach ($d in ($dirs.Keys | Sort-Object)) {
                $full = if ($SubPath) { "$SubPath\$d" } else { $d }
                & $addFileRow $iconFolder $d '<DIR>' '' '' '' $full $true
            }
            $filesInDir = $filesInDir | Sort-Object { [System.IO.Path]::GetExtension($_.Name).ToLower() }, { $_.Name }
            foreach ($f in $filesInDir) {
                $ext = [System.IO.Path]::GetExtension($f.Name).ToLower()
                $full = if ($SubPath) { "$SubPath\$($f.Name)" } else { $f.Name }
                & $addFileRow $iconFile $f.Name (if ($ext) { $ext } else { '' }) (& $formatSize $f.Size) $f.Modified $f.Attrs $full $false
            }

            & $setStatus "Archive: $ArchivePath  |  $($dirs.Count) folder(s), $($filesInDir.Count) file(s)"
        }
        catch {
            & $setStatus "Error reading archive: $($_.Exception.Message)"
            & $addFileRow '' '(Error reading archive)' '' '' '' $_.Exception.Message '' $false
        }
    }.GetNewClosure()

    # -----------------------------------------------------------------------
    # Navigation
    # -----------------------------------------------------------------------
    $navigate = {
        param([string]$Path)
        $Path = $Path.Trim()
        if ([string]::IsNullOrWhiteSpace($Path)) { return }
        if (Test-Path -LiteralPath $Path -PathType Container) {
            & $loadDirectory $Path
        }
        elseif (Test-Path -LiteralPath $Path -PathType Leaf) {
            if (& $isArchive $Path) { & $loadArchive $Path '' }
            else {
                $parent = [System.IO.Path]::GetDirectoryName($Path)
                if ($parent) { & $loadDirectory $parent }
            }
        }
        else {
            [System.Windows.MessageBox]::Show("Path not found: $Path", 'Navigation',
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
        }
    }.GetNewClosure()

    $treeDirs.Add_SelectedItemChanged({
        $sel = $treeDirs.SelectedItem
        if ($sel -and $sel.Tag -and $sel.Tag -ne '__dummy__') {
            & $loadDirectory ([string]$sel.Tag)
        }
    }.GetNewClosure())

    $gridFiles.Add_MouseDoubleClick({
        $sel = $gridFiles.SelectedItem
        if (-not $sel) { return }
        $full = [string]$sel.FullPath
        if ($state.InArchive) {
            if ($sel.IsDir) { & $loadArchive $state.ArchivePath $full }
            return
        }
        if ($sel.IsDir) { & $loadDirectory $full }
        elseif (& $isArchive $sel.Name) { & $loadArchive $full '' }
    }.GetNewClosure())

    $btnGo.Add_Click({ & $navigate $txtPath.Text }.GetNewClosure())
    $txtPath.Add_KeyDown({
        if ($_.Key -eq [System.Windows.Input.Key]::Return) {
            $_.Handled = $true
            & $navigate $txtPath.Text
        }
    }.GetNewClosure())

    $btnUp.Add_Click({
        if ($state.InArchive) {
            if ($state.ArchiveSubPath) {
                $parentSub = [System.IO.Path]::GetDirectoryName($state.ArchiveSubPath)
                if (-not $parentSub) { $parentSub = '' }
                & $loadArchive $state.ArchivePath $parentSub
            }
            else {
                $pd = [System.IO.Path]::GetDirectoryName($state.ArchivePath)
                if ($pd) { & $loadDirectory $pd }
            }
        }
        else {
            if ($state.CurrentPath) {
                $parent = [System.IO.Path]::GetDirectoryName($state.CurrentPath)
                if ($parent) { & $loadDirectory $parent }
            }
        }
    }.GetNewClosure())

    $mnuCopyPath.Add_Click({
        $sel = $gridFiles.SelectedItem
        if ($sel -and $sel.FullPath) {
            [System.Windows.Clipboard]::SetText([string]$sel.FullPath)
        }
    }.GetNewClosure())

    $mnuOpenExplorer.Add_Click({
        if (-not $state.InArchive -and $state.CurrentPath) {
            Start-Process explorer.exe -ArgumentList @("`"$($state.CurrentPath)`"")
        }
    }.GetNewClosure())

    # Default to user profile
    $view.Add_Loaded({
        if (-not $state.CurrentPath) {
            $def = $env:USERPROFILE
            if (Test-Path -LiteralPath $def) { & $loadDirectory $def }
        }
    }.GetNewClosure())

    return $view
}
