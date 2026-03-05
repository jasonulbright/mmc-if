# Changelog

All notable changes to MMC-If are documented in this file.

## [1.2.0] - 2026-03-04

### Added
- **Services Manager module** -- view, start, stop, restart Windows services via `Win32_Service`; color-coded status (Running/Stopped/Paused), dependency detail panel, filter, CSV/HTML export
- **Certificate Store Browser module** -- browse CurrentUser and LocalMachine certificate stores via .NET `X509Store`; TreeView navigation, expiry color-coding, SAN/template extraction, thumbprint copy
- **Local Users & Groups module** -- view local users and groups via `Get-LocalUser`/`Get-LocalGroup`; toggle between Users/Groups views, disabled user dimming, group membership detail, orphaned SID handling
- **Module enable/disable** -- Preferences dialog now lists all discovered modules with checkboxes; disabled modules are skipped at startup; changes require restart
- 7 prerequisite Pester tests for new modules (Win32_Service, X509Store, Get-LocalUser/Get-LocalGroup/Get-LocalGroupMember)

### Fixed
- Dark mode restart now captures script path at function scope (`$scriptFile`) instead of relying on `$PSCommandPath` which is unavailable inside event handler scriptblocks
- CertificateStore SplitContainer MinSize deferred to SizeChanged handler to prevent "SplitterDistance must be between Panel1MinSize and Width - Panel2MinSize" error on nested splits

### Changed
- Renamed from MMC-Alt to MMC-If

## [1.1.0] - 2026-03-04

### Added
- **Event Log Viewer module** -- browse Application, System, Security, and Setup logs via `Get-WinEvent`; filter by level, time range, and text; color-coded severity rows; detail panel for full event messages
- **Device Info module** -- WMI-based system information viewer with categories for System, BIOS, Processor, Memory, Storage, Network, Display, and PnP Devices; formatted values for sizes, speeds, and enums
- **File Explorer module** -- alternative file browser showing hidden files and extensions by default, sorted by type; drive roots and quick-access folders; 7-Zip integration for archive browsing when installed
- SMBIOS placeholder detection in Device Info -- annotates known firmware placeholder strings (e.g., "System Serial Number") with "(not set by manufacturer)"

### Fixed
- Closure/scope isolation for all module event handlers -- converted `$script:` variables to local reference types with `.GetNewClosure()` so UI events work reliably after initialization
- File Explorer not displaying files -- replaced scriptblock-based `Sort-Object` (incompatible with closures) with property-based sorting
- File Explorer emoji crash (`Cannot convert value "128193" to type "System.Char"`) -- use `ConvertFromUtf32()` for supplementary-plane Unicode characters
- Event Log Viewer "no events" message now includes the active time filter to guide users toward expanding the range

## [1.0.0] - 2026-03-04

### Added
- **Plugin shell** -- modular app framework that discovers and loads plugin modules from `Modules/` subfolders; each module gets its own tab, theme colors, and logging access
- **Registry Browser module** -- read-only registry browser for HKLM, HKCR, HKU, HKCC with full read-write for HKCU; TreeView with lazy-loaded subkeys, values DataGridView with Name/Type/Data columns, address bar with path navigation
- **HKCU write operations** -- create/delete keys, create/modify/delete values with type-aware editor dialog (String, DWORD, QWORD, Binary, MultiString, ExpandString)
- **Context menus** -- Copy Key Path and Refresh for all hives; New Key, Delete Key, New Value, Modify, Delete for HKCU
- **Search** -- Ctrl+F to find keys, value names, or value data with depth-first tree walk
- Dark/light theme, window state persistence, preferences dialog
