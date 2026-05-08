<#
.SYNOPSIS
    Pure .reg file format I/O helpers for the Registry Browser module.

.DESCRIPTION
    No WPF, no registry access - just text formatting and parsing. Safe to
    dot-source from unit tests without spinning up a UserControl.

    The UserControl (RegistryBrowser.ps1) wraps these in closures and calls
    them from its event handlers. Tests call them directly.
#>

$script:RegHiveLongName = @{
    'HKCU' = 'HKEY_CURRENT_USER'
    'HKLM' = 'HKEY_LOCAL_MACHINE'
    'HKCR' = 'HKEY_CLASSES_ROOT'
    'HKU'  = 'HKEY_USERS'
    'HKCC' = 'HKEY_CURRENT_CONFIG'
}

$script:RegHiveShortName = @{}
foreach ($k in $script:RegHiveLongName.Keys) {
    $script:RegHiveShortName[$script:RegHiveLongName[$k]] = $k
}

function ConvertTo-RegEscapedString {
    <#
    .SYNOPSIS
        Escapes a string for the "name"=value literal form in a .reg file.
        Backslash becomes \\, quote becomes \".
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Value
    )
    if ($null -eq $Value) { return '' }
    $Value.Replace('\', '\\').Replace('"', '\"')
}

function ConvertFrom-RegEscapedString {
    <#
    .SYNOPSIS
        Inverse of ConvertTo-RegEscapedString.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Value
    )
    if ($null -eq $Value) { return '' }
    $sb = New-Object System.Text.StringBuilder
    $i = 0
    while ($i -lt $Value.Length) {
        $c = $Value[$i]
        if ($c -eq '\' -and $i + 1 -lt $Value.Length) {
            $n = $Value[$i + 1]
            if ($n -eq '\') { [void]$sb.Append('\'); $i += 2; continue }
            if ($n -eq '"') { [void]$sb.Append('"'); $i += 2; continue }
        }
        [void]$sb.Append($c); $i++
    }
    $sb.ToString()
}

function ConvertTo-RegHexBytes {
    <#
    .SYNOPSIS
        Converts a byte array to the comma-separated hex form used in .reg
        files, with line continuation (backslash + CRLF + two spaces) at
        column 77 to match regedit output style.

    .PARAMETER LeadingWidth
        Width of the prefix already on the first line (e.g., '"name"=hex:'
        is 14 characters). Used for accurate wrap-column math.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [AllowNull()][byte[]]$Bytes,
        [int]$LeadingWidth = 0
    )
    if ($null -eq $Bytes -or $Bytes.Count -eq 0) { return '' }
    $parts = foreach ($b in $Bytes) { '{0:x2}' -f $b }
    $line = New-Object System.Text.StringBuilder
    $col = $LeadingWidth
    for ($i = 0; $i -lt $parts.Count; $i++) {
        $seg = $parts[$i]
        $sep = if ($i -lt $parts.Count - 1) { ',' } else { '' }
        $needed = $seg.Length + $sep.Length
        if ($col + $needed -gt 77) {
            [void]$line.Append("\`r`n  ")
            $col = 2
        }
        [void]$line.Append($seg).Append($sep)
        $col += $needed
    }
    $line.ToString()
}

function ConvertTo-RegValueLine {
    <#
    .SYNOPSIS
        Formats one registry value as a single .reg file line (no trailing CRLF).

    .DESCRIPTION
        Supports the standard types (String, ExpandString, Binary, DWord,
        QWord, MultiString, None) and emits hex(N): for any other RegistryValueKind.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Name,
        [Parameter(Mandatory)][Microsoft.Win32.RegistryValueKind]$Kind,
        [AllowNull()]$Value
    )
    $namePart = if ([string]::IsNullOrEmpty($Name)) { '@' } else { '"' + (ConvertTo-RegEscapedString $Name) + '"' }
    $prefix = "$namePart="
    $enc = [System.Text.Encoding]::Unicode

    switch ($Kind) {
        'String' {
            $s = if ($null -eq $Value) { '' } else { [string]$Value }
            return $prefix + '"' + (ConvertTo-RegEscapedString $s) + '"'
        }
        'DWord' {
            $u = [uint32]$Value
            return $prefix + ('dword:{0:x8}' -f $u)
        }
        'QWord' {
            $bytes = [BitConverter]::GetBytes([uint64]$Value)
            return $prefix + 'hex(b):' + (ConvertTo-RegHexBytes $bytes ($prefix.Length + 'hex(b):'.Length))
        }
        'Binary' {
            $bytes = if ($Value -is [byte[]]) { $Value } else { [byte[]]@() }
            return $prefix + 'hex:' + (ConvertTo-RegHexBytes $bytes ($prefix.Length + 'hex:'.Length))
        }
        'ExpandString' {
            $s = if ($null -eq $Value) { '' } else { [string]$Value }
            $bytes = $enc.GetBytes($s) + [byte[]]@(0, 0)
            return $prefix + 'hex(2):' + (ConvertTo-RegHexBytes $bytes ($prefix.Length + 'hex(2):'.Length))
        }
        'MultiString' {
            $arr = if ($Value -is [string[]]) { $Value } else { [string[]]@() }
            $list = New-Object System.Collections.Generic.List[byte]
            foreach ($s in $arr) {
                $b = $enc.GetBytes([string]$s)
                $list.AddRange($b)
                $list.Add(0); $list.Add(0)
            }
            $list.Add(0); $list.Add(0)
            return $prefix + 'hex(7):' + (ConvertTo-RegHexBytes $list.ToArray() ($prefix.Length + 'hex(7):'.Length))
        }
        'None' {
            $bytes = if ($Value -is [byte[]]) { $Value } else { [byte[]]@() }
            return $prefix + 'hex(0):' + (ConvertTo-RegHexBytes $bytes ($prefix.Length + 'hex(0):'.Length))
        }
        default {
            $kindNum = [int]$Kind
            $bytes = if ($Value -is [byte[]]) { $Value } else { [byte[]]@() }
            return $prefix + ('hex({0:x}):' -f $kindNum) + (ConvertTo-RegHexBytes $bytes ($prefix.Length + 8))
        }
    }
}

function ConvertTo-RegLongHivePath {
    <#
    .SYNOPSIS
        Expands short hive prefix (HKCU) to .reg-file long form (HKEY_CURRENT_USER).
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][string]$Path)
    $parts = $Path -split '\\', 2
    $h = $parts[0]
    if ($script:RegHiveLongName.Contains($h)) {
        if ($parts.Count -gt 1) { return "$($script:RegHiveLongName[$h])\$($parts[1])" }
        return $script:RegHiveLongName[$h]
    }
    $Path
}

function ConvertTo-RegShortHivePath {
    <#
    .SYNOPSIS
        Collapses .reg-file long-form (HKEY_CURRENT_USER) back to short (HKCU).
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][string]$Path)
    foreach ($long in $script:RegHiveShortName.Keys) {
        if ($Path -match "^$([regex]::Escape($long))(\\|$)") {
            $short = $script:RegHiveShortName[$long]
            return $short + $Path.Substring($long.Length)
        }
    }
    $Path
}

function Test-RegPathIsHkcu {
    <#
    .SYNOPSIS
        Returns $true if the supplied short-form path is within HKCU.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param([Parameter(Mandatory)][string]$Path)
    $Path -match '^HKCU(\\|$)'
}

function ConvertFrom-RegFileText {
    <#
    .SYNOPSIS
        Parses .reg file text into an operations list.

    .DESCRIPTION
        Accepts the text content of a .reg file (caller is responsible for
        stripping BOM and decoding from UTF-16 LE). Returns a hashtable:
            Header    = first non-empty line (for signature check)
            Warnings  = list of parse warning strings
            Ops       = list of operation hashtables:
                          @{ Op='CreateKey';   Path='HKCU\...' }
                          @{ Op='DeleteKey';   Path='HKCU\...' }
                          @{ Op='SetValue';    Path; Name; Kind; Value }
                          @{ Op='DeleteValue'; Path; Name }
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Text)

    $result = @{
        Ops      = New-Object System.Collections.Generic.List[hashtable]
        Warnings = New-Object System.Collections.Generic.List[string]
        Header   = ''
    }

    $rawLines = ($Text -replace "`r`n", "`n") -split "`n"

    # Join continuation lines (backslash at EOL).
    $joined = New-Object System.Collections.Generic.List[string]
    $buffer = $null
    foreach ($ln in $rawLines) {
        if ($null -ne $buffer) { $buffer += $ln.TrimStart() }
        else { $buffer = $ln }
        if ($buffer.EndsWith('\')) { $buffer = $buffer.Substring(0, $buffer.Length - 1) }
        else { $joined.Add($buffer); $buffer = $null }
    }
    if ($null -ne $buffer) { $joined.Add($buffer) }

    foreach ($ln in $joined) {
        $t = $ln.Trim()
        if ($t) { $result.Header = $t; break }
    }
    if ($result.Header -notlike 'Windows Registry Editor Version*' -and
        $result.Header -notlike 'REGEDIT4*') {
        $result.Warnings.Add("File header is not a recognized .reg signature: '$($result.Header)'")
    }

    $currentKey = $null
    foreach ($ln in $joined) {
        $trim = $ln.Trim()
        if (-not $trim) { continue }
        if ($trim.StartsWith(';')) { continue }
        if ($trim -like 'Windows Registry Editor*') { continue }
        if ($trim -like 'REGEDIT4*') { continue }

        # Key header: [path] or [-path]
        if ($trim.StartsWith('[') -and $trim.EndsWith(']')) {
            $inside = $trim.Substring(1, $trim.Length - 2)
            if ($inside.StartsWith('-')) {
                $kp = ConvertTo-RegShortHivePath $inside.Substring(1).Trim()
                [void]$result.Ops.Add(@{ Op = 'DeleteKey'; Path = $kp })
                $currentKey = $null
            } else {
                $kp = ConvertTo-RegShortHivePath $inside.Trim()
                [void]$result.Ops.Add(@{ Op = 'CreateKey'; Path = $kp })
                $currentKey = $kp
            }
            continue
        }

        if ($null -eq $currentKey) {
            $result.Warnings.Add("Value line with no preceding key: $trim")
            continue
        }

        $eqIdx = -1
        $name = $null
        if ($trim.StartsWith('@')) {
            $eqIdx = $trim.IndexOf('=')
            $name = ''
        }
        elseif ($trim.StartsWith('"')) {
            $i = 1
            while ($i -lt $trim.Length) {
                if ($trim[$i] -eq '\' -and $i + 1 -lt $trim.Length) { $i += 2; continue }
                if ($trim[$i] -eq '"') { break }
                $i++
            }
            if ($i -ge $trim.Length) {
                $result.Warnings.Add("Unterminated value name: $trim")
                continue
            }
            $name = ConvertFrom-RegEscapedString $trim.Substring(1, $i - 1)
            $rest = $trim.Substring($i + 1).TrimStart()
            if (-not $rest.StartsWith('=')) {
                $result.Warnings.Add("Expected '=' after value name: $trim")
                continue
            }
            $eqIdx = $trim.IndexOf('=', $i + 1)
        }
        else {
            $result.Warnings.Add("Unrecognized line: $trim")
            continue
        }

        if ($eqIdx -lt 0) { $result.Warnings.Add("Missing '=': $trim"); continue }
        $data = $trim.Substring($eqIdx + 1).Trim()

        if ($data -eq '-') {
            [void]$result.Ops.Add(@{ Op = 'DeleteValue'; Path = $currentKey; Name = $name })
            continue
        }

        if ($data.StartsWith('"') -and $data.EndsWith('"')) {
            $s = ConvertFrom-RegEscapedString $data.Substring(1, $data.Length - 2)
            [void]$result.Ops.Add(@{
                Op = 'SetValue'; Path = $currentKey; Name = $name
                Kind = [Microsoft.Win32.RegistryValueKind]::String; Value = $s
            })
            continue
        }

        if ($data -match '^dword:([0-9a-fA-F]{1,8})$') {
            $u = [Convert]::ToUInt32($matches[1], 16)
            [void]$result.Ops.Add(@{
                Op = 'SetValue'; Path = $currentKey; Name = $name
                Kind = [Microsoft.Win32.RegistryValueKind]::DWord; Value = [int]$u
            })
            continue
        }

        if ($data -match '^hex(?:\(([0-9a-fA-F]+)\))?:(.*)$') {
            $kindHex = $matches[1]
            $bytesCsv = $matches[2].Trim()
            $bytes = if ($bytesCsv) {
                [byte[]]@($bytesCsv -split ',' | ForEach-Object {
                    $h = $_.Trim()
                    if ($h) { [byte]([Convert]::ToByte($h, 16)) }
                })
            } else { [byte[]]@() }
            $kindInt = if ($kindHex) { [Convert]::ToInt32($kindHex, 16) } else { 3 }
            $kind = [Microsoft.Win32.RegistryValueKind]$kindInt
            # if/elseif instead of switch so the array-typed branches (3, 7,
            # default) emit a single typed object rather than an enumerated
            # pipeline that gets recollected as Object[].
            $val = $null
            if ($kindInt -eq 1 -or $kindInt -eq 2) {
                $val = [System.Text.Encoding]::Unicode.GetString($bytes).TrimEnd([char]0)
            }
            elseif ($kindInt -eq 3) {
                $val = [byte[]]$bytes
            }
            elseif ($kindInt -eq 4) {
                $val = [int][BitConverter]::ToUInt32($bytes, 0)
            }
            elseif ($kindInt -eq 7) {
                $s = [System.Text.Encoding]::Unicode.GetString($bytes).TrimEnd([char]0)
                $val = if ($s) { [string[]]@($s -split [char]0) } else { [string[]]@() }
            }
            elseif ($kindInt -eq 11) {
                $val = [long][BitConverter]::ToInt64($bytes, 0)
            }
            else {
                $val = [byte[]]$bytes
            }
            [void]$result.Ops.Add(@{
                Op = 'SetValue'; Path = $currentKey; Name = $name
                Kind = $kind; Value = $val
            })
            continue
        }

        $result.Warnings.Add("Unrecognized value format for '$name': $data")
    }

    return $result
}
