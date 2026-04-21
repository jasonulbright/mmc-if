# Changelog

## [1.0.1] - 2026-04-21

- Minor UI adjustments for WCAG AA contrast compliance
  - Section header foreground now `MahApps.Brushes.Gray1` (theme-adaptive, AAA on both themes) instead of static `Gray2` which failed AA on both themes
  - Version-stamp, module subtitle, and Preferences dialog field labels now use `Gray1` instead of `Gray5` (static `#B9B9B9`), which was invisible on the light theme at 1.96:1
  - LOG OUTPUT label foreground now applied per theme in `ApplyTheme`: `#B0B0B0` on dark (7.07:1), `#595959` on light (7.46:1). Previously a single hardcoded `#B0B0B0` which was invisible on white in light theme

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
