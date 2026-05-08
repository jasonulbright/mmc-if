# Changelog

## [1.0.0] - 2026-05-02

MMC-If is a MahApps.Metro WPF shell that loads module plugins on
demand. Each module replaces one blocked `mmc.exe` snap-in using the
PowerShell cmdlets and .NET APIs the snap-in was wrapping anyway. No
admin rights, no MECM console, no `mmc.exe`. Extract the zip and run
`start-mmcif.ps1`.

### Features

- **Plugin shell** — MahApps.Metro WPF shell with a sidebar of module
  buttons, shared log drawer, status bar, and live dark/light theme
  toggle.
- **Eleven built-in modules** — Registry, Event Logs, Device Info,
  Files, Services, Certificates, Users & Groups, Task Scheduler,
  Networking, Disks, Group Policy.
- **Registry Browser** — HKCU read-write (no UAC), HKLM/HKCR/HKU/HKCC
  read-only; regedit-style TreeView, values grid, typed value editor,
  Ctrl+F search, `.reg` export (any hive, recursive), HKCU-only
  `.reg` import with preview, persistent Favorites.
- **Themed dialogs throughout modules** — every confirm / message
  dialog routes through a brand-themed `Show-MmcIfThemedMessage`
  helper; no raw system MessageBoxes.
- **Status by content, not color** — error states communicated by
  the message text on the affected node, never by red foreground or
  red row coloring (per brand WCAG SC 1.4.1 rule).
- **Module plugin architecture** — drop a subfolder under `Modules/`
  with a `module.json`, `.ps1` entry, and `.xaml` UserControl; the
  shell auto-discovers and loads on demand.
- **Shared Context** — every module receives prefs, save-prefs,
  set-status, log, owner-window references, and themed-dialog
  callbacks through a shared `$Context` hashtable.
- **Title-bar drag fallback** — native `WM_NCHITTEST` hook + managed
  `DragMove` for the main window and the Preferences modal so the
  title bar drags reliably under any host.
- **Persistent window state** — size, position, maximized, theme,
  and registry favorites survive across sessions.

### Stack

- PowerShell 5.1 + .NET Framework 4.7.2+
- WPF + MahApps.Metro (vendored DLLs in `Lib\`)

### License

MIT License.
