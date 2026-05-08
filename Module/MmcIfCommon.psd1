@{
    RootModule        = 'MmcIfCommon.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'c3d4e5f6-7890-abcd-ef01-234567890abc'
    Author            = 'Jason Ulbright'
    Description       = 'Shared module for MMC-If plugin shell: logging, registry helpers, export.'
    PowerShellVersion = '5.1'

    FunctionsToExport = @(
        # Logging
        'Initialize-Logging'
        'Write-Log'

        # Registry Helpers
        'Get-RegistryHive'
        'Format-RegistryValue'
        'Get-RegistryValueTypeName'

        # Export
        'Export-MmcIfCsv'
        'Export-MmcIfHtml'
    )

    CmdletsToExport   = @()
    VariablesToExport  = @()
    AliasesToExport    = @()
}
