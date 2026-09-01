# Antigravity (AGY) Skills for Windows

A curated collection of production-ready skills for **Google Antigravity (AGY)** on Windows, providing deep capabilities for system troubleshooting, Windows 11 optimization, terminal configurations, and crash analysis.

---

## 🚀 Quick Install (1-Line Command)

Open **PowerShell** on any Windows machine and run:

```powershell
irm https://raw.githubusercontent.com/llIIllIID0EIIllIIll/agy-skills/main/install.ps1 | iex
```

The installer automatically places the skills into Antigravity's global discovery directories:
- `~/.gemini/antigravity-cli/skills/`
- `~/.gemini/config/skills/`
- `~/.agents/skills/`

---

## 📦 Offline / Manual Installation

If you prefer to clone the repository or install offline:

1. Clone or download this repository:
   ```powershell
   git clone https://github.com/llIIllIID0EIIllIIll/agy-skills.git
   cd agy-skills
   ```
2. Run the installer:
   ```powershell
   .\install.ps1
   ```
   *Or simply double-click `install-skills.bat`.*

---

## 🛠️ Included Skills

### 1. `diagnose-crash`
> **Root cause analysis for Windows application crashes, freezes, and BSODs.**

- **Triggers:** Crash, unhandled exception, `0xC0000005`, Access Violation, Event ID 1000, WER, minidump, `.dmp`, BSOD, BugCheck.
- **Capabilities:**
  - Automated inspection of Windows Event Logs (`Application` & `System`).
  - Memory and pagefile starvation verification.
  - Crash dump location discovery (`%LOCALAPPDATA%\CrashDumps`, `C:\Windows\Minidump`).
  - Dump symbolization and backtrace analysis using CDB / WinDbg (`!analyze -v`).
  - NTSTATUS error code reference guide.

### 2. `windows`
> **Comprehensive Windows 11 system management, terminal customization, and desktop tuning.**

- **Triggers:** Windows 11, PowerShell profile, `winget`, Windows Terminal, `settings.json`, Snap Layouts, FancyZones, GlazeWM, Dark Mode, accent color, wallpaper, registry tweaks.
- **Sub-guides:**
  - `terminal.md`: Windows Terminal, PowerShell 7, PSReadLine, Oh My Posh, Starship.
  - `tiling.md`: Snap Layouts (`Win+Z`), PowerToys FancyZones, GlazeWM.
  - `theming.md`: Dark/Light theme, Mica/Acrylic effects, wallpaper automation.
  - `capture.md`: Snipping Tool, screen recording, OCR Text Extractor.
  - `registry.md`: Safe registry modifications, taskbar alignment, classic context menu.
  - `packages.md`: Package management via `winget`, Scoop, and Chocolatey.

---

## 🔍 Verifying Installation

Launch the Antigravity CLI and run:

```text
agy
```

Inside the interactive prompt, type `/skills` to verify that all installed skills are listed.

---

## 📄 License

MIT License. Inspired by Omarchy skills collection.
