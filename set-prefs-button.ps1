<#
.SYNOPSIS
    Enables or disables the Preferences button in the MMC-If sidebar.

.DESCRIPTION
    The Preferences button is not shipped in 1.0 because no built-in module
    surfaces user-configurable preferences (Registry favorites are managed
    inside the Registry module, and the theme toggle lives in the title bar).
    If you are adding a module that needs a preferences UI, run this script
    with -State On to enable three things:

      1. MainWindow.xaml: OPTIONS sidebar header + btnPreferences button.
      2. start-mmcif.ps1: Get-Control lookup for $btnPreferences.
      3. start-mmcif.ps1: Add_Click handler that opens the existing
         Show-MmcIfPreferencesDialog function.

    Run with -State Off to reverse all three changes and return the files
    to their shipped state. Both directions are idempotent.

.PARAMETER State
    On to enable the button, Off to disable it.

.EXAMPLE
    .\set-prefs-button.ps1 -State On

.EXAMPLE
    .\set-prefs-button.ps1 -State Off
#>

param(
    [Parameter(Mandatory)]
    [ValidateSet('On', 'Off')]
    [string]$State
)

$ErrorActionPreference = 'Stop'

$xamlPath = Join-Path $PSScriptRoot 'MainWindow.xaml'
$ps1Path  = Join-Path $PSScriptRoot 'start-mmcif.ps1'

if (-not (Test-Path -LiteralPath $xamlPath)) { throw "MainWindow.xaml not found at $xamlPath" }
if (-not (Test-Path -LiteralPath $ps1Path))  { throw "start-mmcif.ps1 not found at $ps1Path" }

$utf8NoBom = New-Object System.Text.UTF8Encoding $false

function Save-File {
    param([string]$Path, [string]$Content)
    [System.IO.File]::WriteAllText($Path, $Content, $utf8NoBom)
}

$xamlAnchorBare = "                    </StackPanel>`r`n                </ScrollViewer>`r`n"

$xamlAnchorWithButton = "                        <TextBlock Text=`"OPTIONS`" Style=`"{StaticResource SectionHeaderStyle}`"/>`r`n                        <Button Controls:ControlsHelper.ContentCharacterCasing=`"Normal`" x:Name=`"btnPreferences`"   Content=`"Preferences`"     Style=`"{StaticResource OptionsButtonStyle}`"/>`r`n`r`n                    </StackPanel>`r`n                </ScrollViewer>`r`n"

$ps1LookupBare = "`$toggleTheme      = Get-Control 'toggleTheme'"
$ps1LookupWith = "`$toggleTheme      = Get-Control 'toggleTheme'`r`n`$btnPreferences   = Get-Control 'btnPreferences'"

$ps1HandlerAnchor = '$script:ShowMmcIfPreferencesDialogSb = ${function:Show-MmcIfPreferencesDialog}'
$ps1HandlerBlock  = @'


$btnPreferences.Add_Click({
    [void](& $script:ShowMmcIfPreferencesDialogSb -Owner $window)
    if ($toggleTheme) { $toggleTheme.IsOn = [bool]$script:Prefs.DarkMode }
}.GetNewClosure())
'@

$enable = ($State -eq 'On')

# --- MainWindow.xaml ---
$xaml = [System.IO.File]::ReadAllText($xamlPath)
$hasButton = $xaml.Contains('btnPreferences')

if ($enable) {
    if ($hasButton) {
        Write-Host 'MainWindow.xaml: Preferences button already present, skipping'
    }
    elseif (-not $xaml.Contains($xamlAnchorBare)) {
        throw 'MainWindow.xaml: could not find sidebar closing StackPanel/ScrollViewer anchor'
    }
    else {
        $xaml = $xaml.Replace($xamlAnchorBare, $xamlAnchorWithButton)
        Save-File -Path $xamlPath -Content $xaml
        Write-Host 'MainWindow.xaml: inserted OPTIONS section header and Preferences button'
    }
}
else {
    if (-not $hasButton) {
        Write-Host 'MainWindow.xaml: Preferences button already absent, skipping'
    }
    elseif (-not $xaml.Contains($xamlAnchorWithButton)) {
        throw 'MainWindow.xaml: button present but does not match the expected block - manual removal required'
    }
    else {
        $xaml = $xaml.Replace($xamlAnchorWithButton, $xamlAnchorBare)
        Save-File -Path $xamlPath -Content $xaml
        Write-Host 'MainWindow.xaml: removed OPTIONS section and Preferences button'
    }
}

# --- start-mmcif.ps1 ---
$ps1 = [System.IO.File]::ReadAllText($ps1Path)
$ps1Changed = $false

$hasLookup  = $ps1 -match '\$btnPreferences\s*=\s*Get-Control'
$hasHandler = $ps1 -match '\$btnPreferences\.Add_Click'

if ($enable) {
    if ($hasLookup) {
        Write-Host 'start-mmcif.ps1: $btnPreferences lookup already present, skipping'
    }
    elseif (-not $ps1.Contains($ps1LookupBare)) {
        throw 'start-mmcif.ps1: could not find $toggleTheme Get-Control anchor'
    }
    else {
        $ps1 = $ps1.Replace($ps1LookupBare, $ps1LookupWith)
        $ps1Changed = $true
        Write-Host 'start-mmcif.ps1: inserted $btnPreferences Get-Control lookup'
    }

    if ($hasHandler) {
        Write-Host 'start-mmcif.ps1: btnPreferences click handler already present, skipping'
    }
    elseif (-not $ps1.Contains($ps1HandlerAnchor)) {
        throw 'start-mmcif.ps1: could not find ShowMmcIfPreferencesDialogSb anchor'
    }
    else {
        $ps1 = $ps1.Replace($ps1HandlerAnchor, $ps1HandlerAnchor + $ps1HandlerBlock)
        $ps1Changed = $true
        Write-Host 'start-mmcif.ps1: inserted btnPreferences.Add_Click handler'
    }
}
else {
    if (-not $hasHandler) {
        Write-Host 'start-mmcif.ps1: btnPreferences click handler already absent, skipping'
    }
    elseif (-not $ps1.Contains($ps1HandlerAnchor + $ps1HandlerBlock)) {
        throw 'start-mmcif.ps1: click handler present but does not match the expected block - manual removal required'
    }
    else {
        $ps1 = $ps1.Replace($ps1HandlerAnchor + $ps1HandlerBlock, $ps1HandlerAnchor)
        $ps1Changed = $true
        Write-Host 'start-mmcif.ps1: removed btnPreferences.Add_Click handler'
    }

    if (-not $hasLookup) {
        Write-Host 'start-mmcif.ps1: $btnPreferences lookup already absent, skipping'
    }
    elseif (-not $ps1.Contains($ps1LookupWith)) {
        throw 'start-mmcif.ps1: $btnPreferences lookup present but does not match the expected block - manual removal required'
    }
    else {
        $ps1 = $ps1.Replace($ps1LookupWith, $ps1LookupBare)
        $ps1Changed = $true
        Write-Host 'start-mmcif.ps1: removed $btnPreferences Get-Control lookup'
    }
}

if ($ps1Changed) {
    Save-File -Path $ps1Path -Content $ps1
}

Write-Host ''
if ($enable) {
    Write-Host 'Done. Preferences button enabled. Restart MMC-If to pick up changes.'
}
else {
    Write-Host 'Done. Preferences button disabled. Restart MMC-If to pick up changes.'
}
