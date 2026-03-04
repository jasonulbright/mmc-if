<#
.SYNOPSIS
    Plugin-based WinForms shell providing alternatives to Windows tools blocked by mmc.exe restrictions.

.DESCRIPTION
    MMC-Alt is a modular application that loads plugin modules from the Modules/ folder.
    Each module provides an alternative UI for a Windows administration tool that is
    normally hosted in mmc.exe (regedit, Event Viewer, Device Manager, etc.).

    The underlying .NET APIs and WMI classes work without elevation even when mmc.exe
    is blocked by endpoint protection policies.

.EXAMPLE
    .\start-mmcalt.ps1

.NOTES
    Requirements:
      - PowerShell 5.1
      - .NET Framework 4.8+
      - Windows Forms (System.Windows.Forms)

    ScriptName : start-mmcalt.ps1
    Purpose    : Plugin shell for MMC alternative tools
    Version    : 1.0.0
    Updated    : 2026-03-04
#>

param()

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()
try { [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false) } catch { }

$moduleRoot = Join-Path $PSScriptRoot "Module"
Import-Module (Join-Path $moduleRoot "MmcAltCommon.psd1") -Force -DisableNameChecking

# Initialize tool logging
$toolLogFolder = Join-Path $PSScriptRoot "Logs"
if (-not (Test-Path -LiteralPath $toolLogFolder)) {
    New-Item -ItemType Directory -Path $toolLogFolder -Force | Out-Null
}
$toolLogPath = Join-Path $toolLogFolder ("MmcAlt-{0}.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
Initialize-Logging -LogPath $toolLogPath

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Set-ModernButtonStyle {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.Button]$Button,
        [Parameter(Mandatory)][System.Drawing.Color]$BackColor
    )
    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.FlatAppearance.BorderSize = 0
    $Button.BackColor = $BackColor
    $Button.ForeColor = [System.Drawing.Color]::White
    $Button.UseVisualStyleBackColor = $false
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
    $hover = [System.Drawing.Color]::FromArgb([Math]::Max(0, $BackColor.R - 18), [Math]::Max(0, $BackColor.G - 18), [Math]::Max(0, $BackColor.B - 18))
    $down  = [System.Drawing.Color]::FromArgb([Math]::Max(0, $BackColor.R - 36), [Math]::Max(0, $BackColor.G - 36), [Math]::Max(0, $BackColor.B - 36))
    $Button.FlatAppearance.MouseOverBackColor = $hover
    $Button.FlatAppearance.MouseDownBackColor = $down
}

function Enable-DoubleBuffer {
    param([Parameter(Mandatory)][System.Windows.Forms.Control]$Control)
    $prop = $Control.GetType().GetProperty("DoubleBuffered", [System.Reflection.BindingFlags] "Instance,NonPublic")
    if ($prop) { $prop.SetValue($Control, $true, $null) | Out-Null }
}

function Add-LogLine {
    param([Parameter(Mandatory)][System.Windows.Forms.TextBox]$TextBox, [Parameter(Mandatory)][string]$Message)
    $ts = (Get-Date).ToString("HH:mm:ss")
    $line = "{0}  {1}" -f $ts, $Message
    if ([string]::IsNullOrWhiteSpace($TextBox.Text)) { $TextBox.Text = $line }
    else { $TextBox.AppendText([Environment]::NewLine + $line) }
    $TextBox.SelectionStart = $TextBox.TextLength
    $TextBox.ScrollToCaret()
}

function New-ThemedGrid {
    param([switch]$MultiSelect)

    $g = New-Object System.Windows.Forms.DataGridView
    $g.Dock = [System.Windows.Forms.DockStyle]::Fill
    $g.ReadOnly = $true
    $g.AllowUserToAddRows = $false
    $g.AllowUserToDeleteRows = $false
    $g.AllowUserToResizeRows = $false
    $g.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $g.MultiSelect = [bool]$MultiSelect
    $g.AutoGenerateColumns = $false
    $g.RowHeadersVisible = $false
    $g.BackgroundColor = $clrPanelBg
    $g.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $g.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
    $g.GridColor = $clrGridLine
    $g.ColumnHeadersDefaultCellStyle.BackColor = $clrAccent
    $g.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
    $g.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $g.ColumnHeadersDefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(4)
    $g.ColumnHeadersDefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::False
    $g.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $g.ColumnHeadersHeight = 32
    $g.ColumnHeadersBorderStyle = [System.Windows.Forms.DataGridViewHeaderBorderStyle]::Single
    $g.EnableHeadersVisualStyles = $false
    $g.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $g.DefaultCellStyle.ForeColor = $clrGridText
    $g.DefaultCellStyle.BackColor = $clrPanelBg
    $g.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(2)
    $selBg = if ($script:Prefs.DarkMode) { [System.Drawing.Color]::FromArgb(38, 79, 120) } else { [System.Drawing.Color]::FromArgb(0, 120, 215) }
    $g.DefaultCellStyle.SelectionBackColor = $selBg
    $g.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
    $g.RowTemplate.Height = 26
    $g.AlternatingRowsDefaultCellStyle.BackColor = $clrGridAlt
    Enable-DoubleBuffer -Control $g
    return $g
}

function Save-WindowState {
    $statePath = Join-Path $PSScriptRoot "MmcAlt.windowstate.json"
    $state = @{
        X = $form.Location.X; Y = $form.Location.Y
        Width = $form.Size.Width; Height = $form.Size.Height
        Maximized = ($form.WindowState -eq [System.Windows.Forms.FormWindowState]::Maximized)
        ActiveTab = $tabMain.SelectedIndex
    }
    $state | ConvertTo-Json | Set-Content -LiteralPath $statePath -Encoding UTF8
}

function Restore-WindowState {
    $statePath = Join-Path $PSScriptRoot "MmcAlt.windowstate.json"
    if (-not (Test-Path -LiteralPath $statePath)) { return }
    try {
        $state = Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json
        if ($state.Maximized) { $form.WindowState = [System.Windows.Forms.FormWindowState]::Maximized }
        else {
            $screen = [System.Windows.Forms.Screen]::FromPoint((New-Object System.Drawing.Point($state.X, $state.Y)))
            $bounds = $screen.WorkingArea
            $x = [Math]::Max($bounds.X, [Math]::Min($state.X, $bounds.Right - 200))
            $y = [Math]::Max($bounds.Y, [Math]::Min($state.Y, $bounds.Bottom - 100))
            $form.Location = New-Object System.Drawing.Point($x, $y)
            $form.Size = New-Object System.Drawing.Size([Math]::Max($form.MinimumSize.Width, $state.Width), [Math]::Max($form.MinimumSize.Height, $state.Height))
        }
        if ($null -ne $state.ActiveTab -and $state.ActiveTab -ge 0 -and $state.ActiveTab -lt $tabMain.TabCount) {
            $tabMain.SelectedIndex = [int]$state.ActiveTab
        }
    } catch { }
}

# ---------------------------------------------------------------------------
# Preferences
# ---------------------------------------------------------------------------

function Get-MmcAltPreferences {
    $prefsPath = Join-Path $PSScriptRoot "MmcAlt.prefs.json"
    $defaults = @{ DarkMode = $false }
    if (Test-Path -LiteralPath $prefsPath) {
        try {
            $loaded = Get-Content -LiteralPath $prefsPath -Raw | ConvertFrom-Json
            if ($null -ne $loaded.DarkMode) { $defaults.DarkMode = [bool]$loaded.DarkMode }
        } catch { }
    }
    return $defaults
}

function Save-MmcAltPreferences {
    param([hashtable]$Prefs)
    $prefsPath = Join-Path $PSScriptRoot "MmcAlt.prefs.json"
    $Prefs | ConvertTo-Json | Set-Content -LiteralPath $prefsPath -Encoding UTF8
}

$script:Prefs = Get-MmcAltPreferences

# ---------------------------------------------------------------------------
# Colors (theme-aware)
# ---------------------------------------------------------------------------

$clrAccent = [System.Drawing.Color]::FromArgb(0, 120, 212)

if ($script:Prefs.DarkMode) {
    $clrFormBg   = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $clrPanelBg  = [System.Drawing.Color]::FromArgb(40, 40, 40)
    $clrHint     = [System.Drawing.Color]::FromArgb(140, 140, 140)
    $clrSubtitle = [System.Drawing.Color]::FromArgb(180, 200, 220)
    $clrGridAlt  = [System.Drawing.Color]::FromArgb(48, 48, 48)
    $clrGridLine = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $clrDetailBg = [System.Drawing.Color]::FromArgb(45, 45, 45)
    $clrSepLine  = [System.Drawing.Color]::FromArgb(55, 55, 55)
    $clrLogBg    = [System.Drawing.Color]::FromArgb(35, 35, 35)
    $clrLogFg    = [System.Drawing.Color]::FromArgb(200, 200, 200)
    $clrText     = [System.Drawing.Color]::FromArgb(220, 220, 220)
    $clrGridText = [System.Drawing.Color]::FromArgb(220, 220, 220)
    $clrErrText  = [System.Drawing.Color]::FromArgb(255, 100, 100)
    $clrWarnText = [System.Drawing.Color]::FromArgb(255, 200, 80)
    $clrOkText   = [System.Drawing.Color]::FromArgb(80, 200, 80)
    $clrInfoText = [System.Drawing.Color]::FromArgb(100, 180, 255)
    $clrTreeBg   = [System.Drawing.Color]::FromArgb(38, 38, 38)
    $clrInputBdr = [System.Drawing.Color]::FromArgb(70, 70, 70)
} else {
    $clrFormBg   = [System.Drawing.Color]::FromArgb(245, 246, 248)
    $clrPanelBg  = [System.Drawing.Color]::White
    $clrHint     = [System.Drawing.Color]::FromArgb(140, 140, 140)
    $clrSubtitle = [System.Drawing.Color]::FromArgb(220, 230, 245)
    $clrGridAlt  = [System.Drawing.Color]::FromArgb(248, 250, 252)
    $clrGridLine = [System.Drawing.Color]::FromArgb(230, 230, 230)
    $clrDetailBg = [System.Drawing.Color]::FromArgb(250, 250, 250)
    $clrSepLine  = [System.Drawing.Color]::FromArgb(218, 220, 224)
    $clrLogBg    = [System.Drawing.Color]::White
    $clrLogFg    = [System.Drawing.Color]::Black
    $clrText     = [System.Drawing.Color]::Black
    $clrGridText = [System.Drawing.Color]::Black
    $clrErrText  = [System.Drawing.Color]::FromArgb(180, 0, 0)
    $clrWarnText = [System.Drawing.Color]::FromArgb(180, 120, 0)
    $clrOkText   = [System.Drawing.Color]::FromArgb(34, 139, 34)
    $clrInfoText = [System.Drawing.Color]::FromArgb(0, 100, 180)
    $clrTreeBg   = [System.Drawing.Color]::White
    $clrInputBdr = [System.Drawing.Color]::FromArgb(200, 200, 200)
}

# Build theme hashtable for plugins
$script:ThemeColors = @{
    FormBg   = $clrFormBg;   PanelBg  = $clrPanelBg;  Accent   = $clrAccent
    Hint     = $clrHint;     Subtitle = $clrSubtitle;  GridAlt  = $clrGridAlt
    GridLine = $clrGridLine; DetailBg = $clrDetailBg;  SepLine  = $clrSepLine
    LogBg    = $clrLogBg;    LogFg    = $clrLogFg;     Text     = $clrText
    GridText = $clrGridText; ErrText  = $clrErrText;   WarnText = $clrWarnText
    OkText   = $clrOkText;   InfoText = $clrInfoText;  TreeBg   = $clrTreeBg
    InputBdr = $clrInputBdr; DarkMode = $script:Prefs.DarkMode
}

# Dark mode ToolStrip renderer
if ($script:Prefs.DarkMode) {
    if (-not ('DarkToolStripRenderer' -as [type])) {
        $rendererCs = @(
            'using System.Drawing;', 'using System.Windows.Forms;',
            'public class DarkToolStripRenderer : ToolStripProfessionalRenderer {',
            '    private Color _bg;',
            '    public DarkToolStripRenderer(Color bg) : base() { _bg = bg; }',
            '    protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e) { }',
            '    protected override void OnRenderToolStripBackground(ToolStripRenderEventArgs e) { using (var b = new SolidBrush(_bg)) { e.Graphics.FillRectangle(b, e.AffectedBounds); } }',
            '    protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e) { if (e.Item.Selected || e.Item.Pressed) { using (var b = new SolidBrush(Color.FromArgb(60, 60, 60))) { e.Graphics.FillRectangle(b, new Rectangle(Point.Empty, e.Item.Size)); } } }',
            '    protected override void OnRenderSeparator(ToolStripSeparatorRenderEventArgs e) { int y = e.Item.Height / 2; using (var p = new Pen(Color.FromArgb(70, 70, 70))) { e.Graphics.DrawLine(p, 0, y, e.Item.Width, y); } }',
            '    protected override void OnRenderImageMargin(ToolStripRenderEventArgs e) { using (var b = new SolidBrush(_bg)) { e.Graphics.FillRectangle(b, e.AffectedBounds); } }',
            '}'
        ) -join "`r`n"
        Add-Type -ReferencedAssemblies System.Windows.Forms, System.Drawing -TypeDefinition $rendererCs
    }
    $script:DarkRenderer = New-Object DarkToolStripRenderer($clrPanelBg)
}

# ---------------------------------------------------------------------------
# Preferences dialog
# ---------------------------------------------------------------------------

function Show-PreferencesDialog {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Preferences"; $dlg.Size = New-Object System.Drawing.Size(440, 200)
    $dlg.MinimumSize = $dlg.Size; $dlg.MaximumSize = $dlg.Size
    $dlg.StartPosition = "CenterParent"; $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false; $dlg.ShowInTaskbar = $false
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9.5); $dlg.BackColor = $clrFormBg

    $grpApp = New-Object System.Windows.Forms.GroupBox
    $grpApp.Text = "Appearance"; $grpApp.SetBounds(16, 12, 392, 60)
    $grpApp.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $grpApp.ForeColor = $clrText; $grpApp.BackColor = $clrFormBg
    if ($script:Prefs.DarkMode) { $grpApp.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $grpApp.ForeColor = $clrSepLine }
    $dlg.Controls.Add($grpApp)

    $chkDark = New-Object System.Windows.Forms.CheckBox
    $chkDark.Text = "Enable dark mode (requires restart)"; $chkDark.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $chkDark.AutoSize = $true; $chkDark.Location = New-Object System.Drawing.Point(14, 24)
    $chkDark.Checked = $script:Prefs.DarkMode; $chkDark.ForeColor = $clrText; $chkDark.BackColor = $clrFormBg
    if ($script:Prefs.DarkMode) { $chkDark.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $chkDark.ForeColor = [System.Drawing.Color]::FromArgb(170, 170, 170) }
    $grpApp.Controls.Add($chkDark)

    $btnSave = New-Object System.Windows.Forms.Button; $btnSave.Text = "Save"; $btnSave.SetBounds(220, 110, 90, 32)
    $btnSave.Font = New-Object System.Drawing.Font("Segoe UI", 9); Set-ModernButtonStyle -Button $btnSave -BackColor $clrAccent
    $dlg.Controls.Add($btnSave)
    $btnCancel = New-Object System.Windows.Forms.Button; $btnCancel.Text = "Cancel"; $btnCancel.SetBounds(318, 110, 90, 32)
    $btnCancel.Font = New-Object System.Drawing.Font("Segoe UI", 9); $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCancel.ForeColor = $clrText; $btnCancel.BackColor = $clrFormBg
    $dlg.Controls.Add($btnCancel)

    $btnSave.Add_Click({
        $needsRestart = ($chkDark.Checked -ne $script:Prefs.DarkMode)
        $script:Prefs.DarkMode = $chkDark.Checked
        Save-MmcAltPreferences -Prefs $script:Prefs
        $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK; $dlg.Close()
        if ($needsRestart) {
            $result = [System.Windows.Forms.MessageBox]::Show("Theme change requires a restart. Restart now?", "Restart Required", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                Save-WindowState
                Start-Process powershell.exe -ArgumentList @('-ExecutionPolicy', 'Bypass', '-File', "`"$($MyInvocation.ScriptName)`"")
                $form.Close()
            }
        }
    })
    $btnCancel.Add_Click({ $dlg.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $dlg.Close() })
    $dlg.AcceptButton = $btnSave; $dlg.CancelButton = $btnCancel
    $dlg.ShowDialog($form) | Out-Null; $dlg.Dispose()
}

# ---------------------------------------------------------------------------
# Form
# ---------------------------------------------------------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text = "MMC-Alt"; $form.Size = New-Object System.Drawing.Size(1360, 900)
$form.MinimumSize = New-Object System.Drawing.Size(1000, 650); $form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9.5); $form.BackColor = $clrFormBg

# ---------------------------------------------------------------------------
# StatusStrip (Dock:Bottom)
# ---------------------------------------------------------------------------

$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.BackColor = $clrPanelBg; $statusStrip.ForeColor = $clrText; $statusStrip.SizingGrip = $false
if ($script:Prefs.DarkMode -and $script:DarkRenderer) { $statusStrip.Renderer = $script:DarkRenderer }
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel; $statusLabel.Text = "Ready"
$statusLabel.Spring = $true; $statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$statusStrip.Items.Add($statusLabel) | Out-Null
$form.Controls.Add($statusStrip)

# ---------------------------------------------------------------------------
# MenuStrip (Dock:Top)
# ---------------------------------------------------------------------------

$menuStrip = New-Object System.Windows.Forms.MenuStrip
$menuStrip.BackColor = $clrPanelBg; $menuStrip.ForeColor = $clrText
if ($script:Prefs.DarkMode -and $script:DarkRenderer) { $menuStrip.Renderer = $script:DarkRenderer }

$menuFile = New-Object System.Windows.Forms.ToolStripMenuItem("&File")
$menuFile.ForeColor = $clrText
$menuFilePrefs = New-Object System.Windows.Forms.ToolStripMenuItem("&Preferences...")
$menuFilePrefs.ForeColor = $clrText
$menuFilePrefs.Add_Click({ Show-PreferencesDialog })
$menuFile.DropDownItems.Add($menuFilePrefs) | Out-Null
$menuFile.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null
$menuFileExit = New-Object System.Windows.Forms.ToolStripMenuItem("E&xit")
$menuFileExit.ForeColor = $clrText
$menuFileExit.Add_Click({ $form.Close() })
$menuFile.DropDownItems.Add($menuFileExit) | Out-Null
$menuStrip.Items.Add($menuFile) | Out-Null

$menuHelp = New-Object System.Windows.Forms.ToolStripMenuItem("&Help")
$menuHelp.ForeColor = $clrText
$menuHelpAbout = New-Object System.Windows.Forms.ToolStripMenuItem("&About MMC-Alt...")
$menuHelpAbout.ForeColor = $clrText
$menuHelpAbout.Add_Click({
    $aboutLines = @("MMC-Alt v1.0.0", "", "Plugin-based alternative to MMC-hosted Windows tools.", "")
    # List loaded modules
    if ($script:LoadedModules.Count -gt 0) {
        $aboutLines += "Loaded modules:"
        foreach ($mod in $script:LoadedModules) {
            $aboutLines += "  - $($mod.Name) v$($mod.Version)"
        }
    } else {
        $aboutLines += "No modules loaded."
    }
    $aboutLines += "", "Copyright (c) 2026 Jason Ulbright", "MIT License"
    [System.Windows.Forms.MessageBox]::Show(($aboutLines -join "`r`n"), "About MMC-Alt", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
})
$menuHelp.DropDownItems.Add($menuHelpAbout) | Out-Null
$menuStrip.Items.Add($menuHelp) | Out-Null

$form.Controls.Add($menuStrip)
$form.MainMenuStrip = $menuStrip
$menuStrip.SendToBack()

# ---------------------------------------------------------------------------
# TabControl (Dock:Fill)
# ---------------------------------------------------------------------------

$tabMain = New-Object System.Windows.Forms.TabControl
$tabMain.Dock = [System.Windows.Forms.DockStyle]::Fill
$tabMain.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$tabMain.SizeMode = [System.Windows.Forms.TabSizeMode]::Fixed
$tabMain.ItemSize = New-Object System.Drawing.Size(130, 30)
$tabMain.DrawMode = [System.Windows.Forms.TabDrawMode]::OwnerDrawFixed

$tabMain.Add_DrawItem({
    param($s, $e)
    $e.Graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
    $tab = $s.TabPages[$e.Index]
    $sel = ($s.SelectedIndex -eq $e.Index)
    $bg = if ($script:Prefs.DarkMode) {
        if ($sel) { $clrAccent } else { $clrPanelBg }
    } else {
        if ($sel) { $clrAccent } else { [System.Drawing.Color]::FromArgb(240, 240, 240) }
    }
    $fg = if ($sel) { [System.Drawing.Color]::White } else { $clrText }
    $bb = New-Object System.Drawing.SolidBrush($bg)
    $e.Graphics.FillRectangle($bb, $e.Bounds)
    $ft = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $sf = New-Object System.Drawing.StringFormat
    $sf.Alignment = [System.Drawing.StringAlignment]::Near
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Far
    $sf.FormatFlags = [System.Drawing.StringFormatFlags]::NoWrap
    $tr = New-Object System.Drawing.RectangleF(($e.Bounds.X + 8), $e.Bounds.Y, ($e.Bounds.Width - 12), ($e.Bounds.Height - 3))
    $tb = New-Object System.Drawing.SolidBrush($fg)
    $e.Graphics.DrawString($tab.Text, $ft, $tb, $tr, $sf)
    $bb.Dispose(); $tb.Dispose(); $ft.Dispose(); $sf.Dispose()
})

$form.Controls.Add($tabMain)
$tabMain.BringToFront()

# ---------------------------------------------------------------------------
# Plugin loader
# ---------------------------------------------------------------------------

$script:LoadedModules = @()
$modulesDir = Join-Path $PSScriptRoot "Modules"

if (Test-Path -LiteralPath $modulesDir) {
    $pluginFolders = Get-ChildItem -LiteralPath $modulesDir -Directory -ErrorAction SilentlyContinue
    foreach ($folder in $pluginFolders) {
        $manifestPath = Join-Path $folder.FullName "module.json"
        if (-not (Test-Path -LiteralPath $manifestPath)) { continue }

        try {
            $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json

            # Validate required fields
            if (-not $manifest.Name -or -not $manifest.EntryScript -or -not $manifest.InitFunction) {
                Write-Log "Skipping module in $($folder.Name): missing required fields in module.json" -Level WARN
                continue
            }

            $entryPath = Join-Path $folder.FullName $manifest.EntryScript
            if (-not (Test-Path -LiteralPath $entryPath)) {
                Write-Log "Skipping module $($manifest.Name): entry script not found at $entryPath" -Level WARN
                continue
            }

            # Create tab for this module
            $tabLabel = if ($manifest.TabLabel) { $manifest.TabLabel } else { $manifest.Name }
            $tabPage = New-Object System.Windows.Forms.TabPage
            $tabPage.Text = $tabLabel
            $tabPage.BackColor = $clrFormBg
            $tabMain.TabPages.Add($tabPage)

            # Dot-source the module script
            . $entryPath

            # Build context to pass to the module
            $moduleContext = @{
                TabPage        = $tabPage
                Theme          = $script:ThemeColors
                Prefs          = $script:Prefs
                LogFunction    = ${function:Add-LogLine}
                StatusFunction = { param([string]$Text) $statusLabel.Text = $Text }
                NewGridFunction = ${function:New-ThemedGrid}
                ButtonStyleFunction = ${function:Set-ModernButtonStyle}
                DoubleBufferFunction = ${function:Enable-DoubleBuffer}
                ScriptRoot     = $PSScriptRoot
                ModuleRoot     = $folder.FullName
            }

            # Call the module init function
            & $manifest.InitFunction $moduleContext

            $script:LoadedModules += @{
                Name = $manifest.Name
                Version = if ($manifest.Version) { $manifest.Version } else { '0.0.0' }
                Description = $manifest.Description
                TabPage = $tabPage
            }

            Write-Log "Loaded module: $($manifest.Name) v$(if ($manifest.Version) { $manifest.Version } else { '0.0.0' })"

        } catch {
            Write-Log "Failed to load module from $($folder.Name): $_" -Level ERROR
        }
    }
}

if ($script:LoadedModules.Count -eq 0) {
    $tabEmpty = New-Object System.Windows.Forms.TabPage
    $tabEmpty.Text = "No Modules"
    $tabEmpty.BackColor = $clrFormBg
    $lblEmpty = New-Object System.Windows.Forms.Label
    $lblEmpty.Text = "No modules found in the Modules/ folder."
    $lblEmpty.Dock = [System.Windows.Forms.DockStyle]::Fill
    $lblEmpty.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $lblEmpty.Font = New-Object System.Drawing.Font("Segoe UI", 12)
    $lblEmpty.ForeColor = $clrHint
    $tabEmpty.Controls.Add($lblEmpty)
    $tabMain.TabPages.Add($tabEmpty)
}

$statusLabel.Text = "Ready - $($script:LoadedModules.Count) module(s) loaded"

# ---------------------------------------------------------------------------
# Form lifecycle
# ---------------------------------------------------------------------------

$form.Add_FormClosing({ Save-WindowState })
$form.Add_Shown({ Restore-WindowState })

[System.Windows.Forms.Application]::Run($form)
