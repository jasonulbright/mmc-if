#Requires -Modules Pester

<#
.SYNOPSIS
    Pester tests for MMC-Alt shell and plugin modules.
#>

# Paths available at discovery time (not inside BeforeAll)
$ProjectRoot = Split-Path -Path $PSScriptRoot -Parent
$ModulesDir  = Join-Path $ProjectRoot 'Modules'

BeforeAll {
    $script:ProjectRoot = Split-Path -Path $PSScriptRoot -Parent
    $script:ModulePath  = Join-Path $script:ProjectRoot 'Module\MmcAltCommon.psd1'
    $script:ModulesDir  = Join-Path $script:ProjectRoot 'Modules'
    Import-Module $script:ModulePath -Force
}

# =========================================================================
# Parse Validation
# =========================================================================

Describe 'Script Parse Validation' {

    It 'start-mmcalt.ps1 parses without errors' {
        $tokens = $null; $errors = $null
        $path = Join-Path $script:ProjectRoot 'start-mmcalt.ps1'
        [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors) | Out-Null
        $errors.Count | Should -Be 0
    }

    # Use discovery-time variable, iterate only module entry scripts
    $entryScripts = Get-ChildItem -Path $ModulesDir -Recurse -Filter '*.ps1' |
        Where-Object { $_.DirectoryName -like "$ModulesDir*" -and $_.Name -ne 'MmcAlt.Tests.ps1' }

    foreach ($s in $entryScripts) {
        It "$($s.Name) parses without errors" -TestCases @{ ScriptPath = $s.FullName } {
            $tokens = $null; $errors = $null
            [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$errors) | Out-Null
            $errors.Count | Should -Be 0
        }
    }
}

# =========================================================================
# Plugin Manifest Validation
# =========================================================================

Describe 'Plugin Manifest Validation' {

    $manifests = Get-ChildItem -Path $ModulesDir -Recurse -Filter 'module.json'

    It 'At least one plugin module exists' {
        @($manifests).Count | Should -BeGreaterThan 0
    }

    foreach ($mj in $manifests) {
        $moduleName = $mj.Directory.Name
        $mjPath = $mj.FullName
        $mjDir = $mj.DirectoryName
        $tc = @{ MjPath = $mjPath; MjDir = $mjDir; ModName = $moduleName }

        It "<ModName>: has all required manifest fields" -TestCases $tc {
            $json = Get-Content $MjPath -Raw | ConvertFrom-Json
            $json.Name | Should -Not -BeNullOrEmpty
            $json.TabLabel | Should -Not -BeNullOrEmpty
            $json.EntryScript | Should -Not -BeNullOrEmpty
            $json.InitFunction | Should -Not -BeNullOrEmpty
            $json.Version | Should -Not -BeNullOrEmpty
        }

        It "<ModName>: entry script exists" -TestCases $tc {
            $json = Get-Content $MjPath -Raw | ConvertFrom-Json
            $scriptPath = Join-Path $MjDir $json.EntryScript
            $scriptPath | Should -Exist
        }

        It "<ModName>: init function defined with [hashtable] param" -TestCases $tc {
            $json = Get-Content $MjPath -Raw | ConvertFrom-Json
            $scriptPath = Join-Path $MjDir $json.EntryScript
            $tokens = $null; $errors = $null
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$errors)
            $funcAst = $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $args[0].Name -eq $json.InitFunction }, $true)
            $funcAst.Count | Should -BeGreaterThan 0
            $paramBlock = $funcAst[0].Body.ParamBlock
            $paramBlock | Should -Not -BeNullOrEmpty
            $paramBlock.Parameters[0].StaticType.Name | Should -Be 'hashtable'
        }
    }
}

# =========================================================================
# Shared Module: MmcAltCommon
# =========================================================================

Describe 'MmcAltCommon Module' {

    It 'Module loads successfully' {
        { Import-Module $script:ModulePath -Force -ErrorAction Stop } | Should -Not -Throw
    }

    It 'Exports expected functions' {
        $exported = (Get-Module MmcAltCommon).ExportedFunctions.Keys
        $expected = @('Initialize-Logging', 'Write-Log', 'Get-RegistryHive', 'Format-RegistryValue', 'Get-RegistryValueTypeName', 'Export-MmcAltCsv', 'Export-MmcAltHtml')
        foreach ($fn in $expected) {
            $exported | Should -Contain $fn
        }
    }
}

# =========================================================================
# Registry Helpers
# =========================================================================

Describe 'Get-RegistryHive' {

    It 'Returns CurrentUser for HKCU' {
        (Get-RegistryHive -HiveName 'HKCU') | Should -Be ([Microsoft.Win32.Registry]::CurrentUser)
    }

    It 'Returns LocalMachine for HKLM' {
        (Get-RegistryHive -HiveName 'HKLM') | Should -Be ([Microsoft.Win32.Registry]::LocalMachine)
    }

    It 'Returns ClassesRoot for HKCR' {
        (Get-RegistryHive -HiveName 'HKCR') | Should -Be ([Microsoft.Win32.Registry]::ClassesRoot)
    }

    It 'Returns Users for HKU' {
        (Get-RegistryHive -HiveName 'HKU') | Should -Be ([Microsoft.Win32.Registry]::Users)
    }

    It 'Returns CurrentConfig for HKCC' {
        (Get-RegistryHive -HiveName 'HKCC') | Should -Be ([Microsoft.Win32.Registry]::CurrentConfig)
    }
}

Describe 'Get-RegistryValueTypeName' {

    It 'Converts String to REG_SZ' {
        Get-RegistryValueTypeName -Kind ([Microsoft.Win32.RegistryValueKind]::String) | Should -Be 'REG_SZ'
    }

    It 'Converts ExpandString to REG_EXPAND_SZ' {
        Get-RegistryValueTypeName -Kind ([Microsoft.Win32.RegistryValueKind]::ExpandString) | Should -Be 'REG_EXPAND_SZ'
    }

    It 'Converts Binary to REG_BINARY' {
        Get-RegistryValueTypeName -Kind ([Microsoft.Win32.RegistryValueKind]::Binary) | Should -Be 'REG_BINARY'
    }

    It 'Converts DWord to REG_DWORD' {
        Get-RegistryValueTypeName -Kind ([Microsoft.Win32.RegistryValueKind]::DWord) | Should -Be 'REG_DWORD'
    }

    It 'Converts QWord to REG_QWORD' {
        Get-RegistryValueTypeName -Kind ([Microsoft.Win32.RegistryValueKind]::QWord) | Should -Be 'REG_QWORD'
    }

    It 'Converts MultiString to REG_MULTI_SZ' {
        Get-RegistryValueTypeName -Kind ([Microsoft.Win32.RegistryValueKind]::MultiString) | Should -Be 'REG_MULTI_SZ'
    }

    It 'Converts None to REG_NONE' {
        Get-RegistryValueTypeName -Kind ([Microsoft.Win32.RegistryValueKind]::None) | Should -Be 'REG_NONE'
    }

    It 'Converts Unknown to REG_UNKNOWN' {
        Get-RegistryValueTypeName -Kind ([Microsoft.Win32.RegistryValueKind]::Unknown) | Should -Be 'REG_UNKNOWN'
    }
}

Describe 'Format-RegistryValue' {

    It 'Returns "(value not set)" for null value' {
        Format-RegistryValue -Value $null -Kind ([Microsoft.Win32.RegistryValueKind]::String) | Should -Be '(value not set)'
    }

    It 'Formats string values' {
        Format-RegistryValue -Value 'hello' -Kind ([Microsoft.Win32.RegistryValueKind]::String) | Should -Be 'hello'
    }

    It 'Formats DWORD as hex and decimal' {
        Format-RegistryValue -Value 255 -Kind ([Microsoft.Win32.RegistryValueKind]::DWord) | Should -Be '0x000000FF (255)'
    }

    It 'Formats QWORD as hex and decimal' {
        Format-RegistryValue -Value 1024 -Kind ([Microsoft.Win32.RegistryValueKind]::QWord) | Should -Be '0x0000000000000400 (1024)'
    }

    It 'Formats binary as hex string' {
        $bytes = [byte[]]@(0x01, 0x0A, 0xFF)
        Format-RegistryValue -Value $bytes -Kind ([Microsoft.Win32.RegistryValueKind]::Binary) | Should -Be '01 0A FF'
    }

    It 'Formats multi-string with pipe separator' {
        $strings = [string[]]@('one', 'two', 'three')
        Format-RegistryValue -Value $strings -Kind ([Microsoft.Win32.RegistryValueKind]::MultiString) | Should -Be 'one | two | three'
    }

    It 'Formats expand string as-is' {
        Format-RegistryValue -Value '%SystemRoot%\test' -Kind ([Microsoft.Win32.RegistryValueKind]::ExpandString) | Should -Be '%SystemRoot%\test'
    }
}

# =========================================================================
# Export Functions
# =========================================================================

Describe 'Export-MmcAltCsv' {

    BeforeAll {
        $script:TestDir = Join-Path $env:TEMP "mmcalt-test-$(Get-Random)"
        New-Item -ItemType Directory -Path $script:TestDir -Force | Out-Null
        Initialize-Logging -LogPath (Join-Path $script:TestDir 'test.log')
    }

    AfterAll {
        Remove-Item -Path $script:TestDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'Exports DataTable to CSV file' {
        $dt = New-Object System.Data.DataTable
        [void]$dt.Columns.Add("Name", [string])
        [void]$dt.Columns.Add("Value", [string])
        $row = $dt.NewRow(); $row["Name"] = "Test"; $row["Value"] = "123"; $dt.Rows.Add($row)
        $row = $dt.NewRow(); $row["Name"] = "Other"; $row["Value"] = "456"; $dt.Rows.Add($row)

        $csvPath = Join-Path $script:TestDir 'test.csv'
        Export-MmcAltCsv -DataTable $dt -OutputPath $csvPath

        $csvPath | Should -Exist
        $imported = Import-Csv $csvPath
        $imported.Count | Should -Be 2
        $imported[0].Name | Should -Be 'Test'
        $imported[0].Value | Should -Be '123'
    }

    It 'Creates parent directories if missing' {
        $dt = New-Object System.Data.DataTable
        [void]$dt.Columns.Add("Col", [string])
        $row = $dt.NewRow(); $row["Col"] = "data"; $dt.Rows.Add($row)

        $nestedPath = Join-Path $script:TestDir 'sub\deep\test.csv'
        Export-MmcAltCsv -DataTable $dt -OutputPath $nestedPath
        $nestedPath | Should -Exist
    }
}

Describe 'Export-MmcAltHtml' {

    BeforeAll {
        $script:TestDir = Join-Path $env:TEMP "mmcalt-test-html-$(Get-Random)"
        New-Item -ItemType Directory -Path $script:TestDir -Force | Out-Null
        Initialize-Logging -LogPath (Join-Path $script:TestDir 'test.log')
    }

    AfterAll {
        Remove-Item -Path $script:TestDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'Exports DataTable to HTML file' {
        $dt = New-Object System.Data.DataTable
        [void]$dt.Columns.Add("Name", [string])
        [void]$dt.Columns.Add("Count", [string])
        $row = $dt.NewRow(); $row["Name"] = "Item1"; $row["Count"] = "5"; $dt.Rows.Add($row)

        $htmlPath = Join-Path $script:TestDir 'report.html'
        Export-MmcAltHtml -DataTable $dt -OutputPath $htmlPath -ReportTitle 'Test Report'

        $htmlPath | Should -Exist
        $html = Get-Content $htmlPath -Raw
        $html | Should -Match 'Test Report'
        $html | Should -Match '<th>Name</th>'
        $html | Should -Match '<td>Item1</td>'
    }
}

# =========================================================================
# Logging
# =========================================================================

Describe 'Logging' {

    BeforeAll {
        $script:TestDir = Join-Path $env:TEMP "mmcalt-test-log-$(Get-Random)"
        New-Item -ItemType Directory -Path $script:TestDir -Force | Out-Null
    }

    AfterAll {
        Remove-Item -Path $script:TestDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'Initialize-Logging creates log file with header' {
        $logPath = Join-Path $script:TestDir 'init.log'
        Initialize-Logging -LogPath $logPath
        $logPath | Should -Exist
        $content = Get-Content $logPath -Raw
        $content | Should -Match 'Log initialized'
    }

    It 'Write-Log appends to log file' {
        $logPath = Join-Path $script:TestDir 'write.log'
        Initialize-Logging -LogPath $logPath
        Write-Log -Message 'Test message' -Quiet
        $content = Get-Content $logPath -Raw
        $content | Should -Match 'Test message'
    }

    It 'Write-Log includes severity level' {
        $logPath = Join-Path $script:TestDir 'level.log'
        Initialize-Logging -LogPath $logPath
        Write-Log -Message 'Warning test' -Level 'WARN' -Quiet
        $content = Get-Content $logPath -Raw
        $content | Should -Match '\[WARN \]'
    }

    It 'Write-Log accepts empty string message' {
        $logPath = Join-Path $script:TestDir 'empty.log'
        Initialize-Logging -LogPath $logPath
        { Write-Log -Message '' -Quiet } | Should -Not -Throw
    }
}

# =========================================================================
# Shell Script Validation
# =========================================================================

Describe 'Shell Script' {

    It 'Defines required helper functions' {
        $content = Get-Content (Join-Path $script:ProjectRoot 'start-mmcalt.ps1') -Raw
        $content | Should -Match 'function\s+New-ThemedGrid'
        $content | Should -Match 'function\s+Set-ModernButtonStyle'
        $content | Should -Match 'function\s+Enable-DoubleBuffer'
        $content | Should -Match 'function\s+Add-LogLine'
    }

    It 'Loads WinForms and Drawing assemblies' {
        $content = Get-Content (Join-Path $script:ProjectRoot 'start-mmcalt.ps1') -Raw
        $content | Should -Match 'System\.Windows\.Forms'
        $content | Should -Match 'System\.Drawing'
    }

    It 'Scans Modules directory for plugins' {
        $content = Get-Content (Join-Path $script:ProjectRoot 'start-mmcalt.ps1') -Raw
        $content | Should -Match 'module\.json'
    }
}

# =========================================================================
# No $script: Variables Inside .GetNewClosure() Blocks
# =========================================================================

Describe 'Closure Safety: no script-scope inside .GetNewClosure()' {

    $closureScripts = Get-ChildItem -Path $ModulesDir -Recurse -Filter '*.ps1'

    foreach ($s in $closureScripts) {
        $sPath = $s.FullName
        $sName = $s.Name

        It "$sName has no script-scope references inside .GetNewClosure() blocks" -TestCases @{ FilePath = $sPath } {
            $content = Get-Content $FilePath -Raw

            # Find all .GetNewClosure() scriptblocks and check for $script: inside them
            $pattern = '\(\{(?<body>[\s\S]*?)\}\.GetNewClosure\(\)\)'
            $regexMatches = [regex]::Matches($content, $pattern)

            $violations = @()
            foreach ($m in $regexMatches) {
                $body = $m.Groups['body'].Value
                if ($body -match '\$script:') {
                    # Allow $script:DarkRenderer (evaluated during init, not deferred)
                    $cleaned = $body -replace '\$script:DarkRenderer', ''
                    if ($cleaned -match '\$script:') {
                        $violations += "Found script-scope var inside .GetNewClosure()"
                    }
                }
            }

            $violations.Count | Should -Be 0 -Because '.GetNewClosure() creates a new module scope where script-scope resolves incorrectly'
        }
    }
}
