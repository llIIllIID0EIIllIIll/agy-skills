---
name: windows
description: >
  REQUIRED for customization and system management of Windows 11 desktop, system config,
  terminal, appearance, and utilities. Use when configuring Windows Terminal, PowerShell profiles,
  window rules/tiling (Snap Layouts, FancyZones, GlazeWM), desktop appearance (dark mode, wallpaper,
  mica/acrylic, accent colors), keyboard shortcuts, package management (winget), screen capture,
  registry tweaks, or Windows services. Triggers: Windows 11, PowerShell profile, winget,
  Windows Terminal, settings.json, Snap Layouts, FancyZones, GlazeWM, dark mode, accent color,
  wallpaper, night light, virtual desktops, registry tweaks, taskbar customization.
---

# Windows 11 Customization & Management

Manage and customize Windows 11 systems — terminal environments, tiling, themes, system shortcuts, package management, and registry configurations.

This skill is for end-user customization and automation on Windows 11 systems.

---

## When This Skill MUST Be Used

**ALWAYS invoke this skill for user requests involving ANY of these:**

- Configuring Windows Terminal (`settings.json`) or PowerShell profiles (`$PROFILE`)
- Window layout, tiling, Snap Layouts, FancyZones, GlazeWM, or Virtual Desktops
- Theming, Dark Mode / Light Mode, accent colors, Mica/Acrylic effects, wallpaper
- Screen capture, screen recording, OCR Text Extractor, Clipboard History
- Windows Package Manager (`winget`), Scoop, or Chocolatey package operations
- Taskbar customization, Start Menu customization, classic context menus
- Registry tweaks (`HKCU` / `HKLM`), system policies, Developer Mode, long path limits
- Network, RDP, shared folders, and Windows Services management

---

## Topic Guides

Deep-dive instructions live alongside this file. Refer to the matching guide:

| Guide | Purpose & Scope |
| :--- | :--- |
| [`terminal.md`](terminal.md) | Windows Terminal (`settings.json`), PowerShell 7 profile, PSReadLine, Oh My Posh, Starship |
| [`tiling.md`](tiling.md) | Snap Layouts (`Win+Z`), Virtual Desktops, PowerToys FancyZones, GlazeWM tiling |
| [`theming.md`](theming.md) | Dark/Light mode, Accent colors, Transparency, Wallpaper via PowerShell, Night Light |
| [`capture.md`](capture.md) | Snipping Tool (`Win+Shift+S`), Screen Recording (`Win+Alt+R`), OCR, Clipboard History (`Win+V`) |
| [`registry.md`](registry.md) | Safe registry customization, Taskbar alignment, Classic context menu, Developer Mode |
| [`packages.md`](packages.md) | Software installation, updates, and removal via `winget`, Scoop, and Chocolatey |

---

## Critical Safety Rules

1. **Registry Safety**:
   - Always backup a registry key before modifying it:
     ```powershell
     reg export "HKCU\Software\Target\Key" "C:\backup.reg" /y
     ```
   - Prefer user-level `HKCU:` over machine-level `HKLM:` whenever possible to avoid requiring unnecessary Administrator elevation.
2. **Elevated Permissions (UAC)**:
   - When modifying system-wide settings (`HKLM:`, Windows Services, Program Files), execute commands from an elevated Administrator PowerShell prompt or use:
     ```powershell
     Start-Process powershell -Verb RunAs -ArgumentList "-NoProfile -Command <command>"
     ```
3. **Restarting Explorer Safely**:
   - Many desktop and taskbar changes require restarting `explorer.exe`:
     ```powershell
     Stop-Process -Name explorer -Force
     ```
     Windows automatically restarts Explorer within 2 seconds.

---

## Command Discovery & Quick Tools

| Task | Command |
| :--- | :--- |
| Search / Install software | `winget search <app>` / `winget install <id>` |
| Query system details | `Get-ComputerInfo` or `systeminfo` |
| View active services | `Get-Service | Where-Object Status -eq 'Running'` |
| Check listening ports | `Get-NetTCPConnection -State Listen | Select-Object LocalPort, OwningProcess` |
| Open Settings page | `Start-Process "ms-settings:<page>"` (e.g. `ms-settings:personalization`) |
| Open Windows Terminal settings | `code "$env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json"` |
