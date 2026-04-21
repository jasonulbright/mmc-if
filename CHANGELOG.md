# Changelog

## [1.0.0] - 2026-04-17

Initial public release.

### Features
- **Plugin shell** — MahApps.Metro WPF shell with a sidebar of module buttons, shared log drawer, status bar, and live dark/light theme toggle
- **Eleven built-in modules** — Registry, Event Logs, Device Info, Files, Services, Certificates, Users & Groups, Task Scheduler, Networking, Disks, Group Policy
- **Registry Browser** — HKCU read-write (no UAC), HKLM/HKCR/HKU/HKCC read-only; regedit-style TreeView, values grid, typed value editor, Ctrl+F search, `.reg` export (any hive, recursive), HKCU-only `.reg` import with preview, persistent Favorites
- **Module plugin architecture** — drop a subfolder under `Modules/` with a `module.json`, `.ps1` entry, and `.xaml` UserControl; the shell auto-discovers and loads on demand
- **Shared Context** — every module receives prefs, save-prefs, set-status, log, and owner-window references through a shared Context hashtable
- **Persistent window state** — size, position, maximized, theme, and registry favorites survive across sessions

### License
MIT License.
