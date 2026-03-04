<#
.SYNOPSIS
    Shared module for MMC-Alt plugin shell.

.DESCRIPTION
    Provides:
      - Structured logging (Initialize-Logging, Write-Log)
      - Registry helper functions (Get-RegistryHive, Format-RegistryValue, Get-RegistryValueTypeName)
      - Export to CSV and HTML (Export-MmcAltCsv, Export-MmcAltHtml)

.EXAMPLE
    Import-Module "$PSScriptRoot\Module\MmcAltCommon.psd1" -Force
    Initialize-Logging -LogPath "C:\projects\mmcalt\Logs\mmcalt.log"
#>

# ---------------------------------------------------------------------------
# Module-scoped state
# ---------------------------------------------------------------------------

$script:__LogPath = $null

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

function Initialize-Logging {
    param([string]$LogPath)

    $script:__LogPath = $LogPath

    if ($LogPath) {
        $parentDir = Split-Path -Path $LogPath -Parent
        if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
            New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
        }

        $header = "[{0}] [INFO ] === Log initialized ===" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Set-Content -LiteralPath $LogPath -Value $header -Encoding UTF8
    }
}

function Write-Log {
    <#
    .SYNOPSIS
        Writes a timestamped, severity-tagged log message.
    #>
    param(
        [AllowEmptyString()]
        [Parameter(Mandatory, Position = 0)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO',

        [switch]$Quiet
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $formatted = "[{0}] [{1,-5}] {2}" -f $timestamp, $Level, $Message

    if (-not $Quiet) {
        Write-Host $formatted
        if ($Level -eq 'ERROR') {
            $host.UI.WriteErrorLine($formatted)
        }
    }

    if ($script:__LogPath) {
        Add-Content -LiteralPath $script:__LogPath -Value $formatted -Encoding UTF8 -ErrorAction SilentlyContinue
    }
}

# ---------------------------------------------------------------------------
# Registry Helpers
# ---------------------------------------------------------------------------

function Get-RegistryHive {
    <#
    .SYNOPSIS
        Returns the .NET RegistryKey for a hive name string.
    #>
    param(
        [Parameter(Mandatory)]
        [ValidateSet('HKCU', 'HKLM', 'HKCR', 'HKU', 'HKCC')]
        [string]$HiveName
    )

    switch ($HiveName) {
        'HKCU' { return [Microsoft.Win32.Registry]::CurrentUser }
        'HKLM' { return [Microsoft.Win32.Registry]::LocalMachine }
        'HKCR' { return [Microsoft.Win32.Registry]::ClassesRoot }
        'HKU'  { return [Microsoft.Win32.Registry]::Users }
        'HKCC' { return [Microsoft.Win32.Registry]::CurrentConfig }
    }
}

function Get-RegistryValueTypeName {
    <#
    .SYNOPSIS
        Converts a RegistryValueKind enum value to a display string.
    #>
    param(
        [Parameter(Mandatory)]
        [Microsoft.Win32.RegistryValueKind]$Kind
    )

    switch ($Kind) {
        'String'       { return 'REG_SZ' }
        'ExpandString' { return 'REG_EXPAND_SZ' }
        'Binary'       { return 'REG_BINARY' }
        'DWord'        { return 'REG_DWORD' }
        'QWord'        { return 'REG_QWORD' }
        'MultiString'  { return 'REG_MULTI_SZ' }
        'None'         { return 'REG_NONE' }
        'Unknown'      { return 'REG_UNKNOWN' }
        default        { return $Kind.ToString() }
    }
}

function Format-RegistryValue {
    <#
    .SYNOPSIS
        Formats a registry value for display based on its type.
    #>
    param(
        [AllowNull()]
        $Value,

        [Microsoft.Win32.RegistryValueKind]$Kind
    )

    if ($null -eq $Value) { return '(value not set)' }

    switch ($Kind) {
        'Binary' {
            if ($Value -is [byte[]]) {
                return ($Value | ForEach-Object { '{0:X2}' -f $_ }) -join ' '
            }
            return $Value.ToString()
        }
        'MultiString' {
            if ($Value -is [string[]]) {
                return $Value -join ' | '
            }
            return $Value.ToString()
        }
        'DWord' {
            return '0x{0:X8} ({0})' -f [uint32]$Value
        }
        'QWord' {
            return '0x{0:X16} ({0})' -f [uint64]$Value
        }
        'ExpandString' {
            return $Value.ToString()
        }
        default {
            return $Value.ToString()
        }
    }
}

# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------

function Export-MmcAltCsv {
    <#
    .SYNOPSIS
        Exports a DataTable to CSV.
    #>
    param(
        [Parameter(Mandatory)][System.Data.DataTable]$DataTable,
        [Parameter(Mandatory)][string]$OutputPath
    )

    $parentDir = Split-Path -Path $OutputPath -Parent
    if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
        New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
    }

    $rows = @()
    foreach ($row in $DataTable.Rows) {
        $obj = [ordered]@{}
        foreach ($col in $DataTable.Columns) {
            $obj[$col.ColumnName] = $row[$col.ColumnName]
        }
        $rows += [PSCustomObject]$obj
    }

    $rows | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Log "Exported CSV to $OutputPath"
}

function Export-MmcAltHtml {
    <#
    .SYNOPSIS
        Exports a DataTable to a self-contained HTML report.
    #>
    param(
        [Parameter(Mandatory)][System.Data.DataTable]$DataTable,
        [Parameter(Mandatory)][string]$OutputPath,
        [string]$ReportTitle = 'MMC-Alt Report'
    )

    $parentDir = Split-Path -Path $OutputPath -Parent
    if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
        New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
    }

    $css = @(
        '<style>',
        'body { font-family: "Segoe UI", Arial, sans-serif; margin: 20px; background: #fafafa; }',
        'h1 { color: #0078D4; margin-bottom: 4px; }',
        '.summary { color: #666; margin-bottom: 12px; font-size: 0.9em; }',
        'table { border-collapse: collapse; width: 100%; margin-top: 12px; }',
        'th { background: #0078D4; color: #fff; padding: 8px 12px; text-align: left; }',
        'td { padding: 6px 12px; border-bottom: 1px solid #e0e0e0; }',
        'tr:nth-child(even) { background: #f5f5f5; }',
        '</style>'
    ) -join "`r`n"

    $headerRow = ($DataTable.Columns | ForEach-Object { "<th>$($_.ColumnName)</th>" }) -join ''
    $bodyRows = foreach ($row in $DataTable.Rows) {
        $cells = foreach ($col in $DataTable.Columns) {
            $val = [string]$row[$col.ColumnName]
            "<td>$val</td>"
        }
        "<tr>$($cells -join '')</tr>"
    }

    $html = @(
        '<!DOCTYPE html>',
        '<html><head><meta charset="utf-8"><title>' + $ReportTitle + '</title>',
        $css,
        '</head><body>',
        "<h1>$ReportTitle</h1>",
        "<div class='summary'>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Rows: $($DataTable.Rows.Count)</div>",
        "<table><thead><tr>$headerRow</tr></thead>",
        "<tbody>$($bodyRows -join "`r`n")</tbody></table>",
        '</body></html>'
    ) -join "`r`n"

    Set-Content -LiteralPath $OutputPath -Value $html -Encoding UTF8
    Write-Log "Exported HTML to $OutputPath"
}
