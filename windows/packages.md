# Package Management on Windows

Read this before installing, updating, or managing software applications on Windows.

---

## 1. Windows Package Manager (`winget`)

`winget` is the official, native package manager for Windows 11.

### Key Operations:

```powershell
# Search for an application
winget search <name>

# Install software (Exact ID, accept agreements silently)
winget install --id <Package.ID> --exact --accept-package-agreements --accept-source-agreements

# List installed packages
winget list

# Check for updates
winget upgrade

# Upgrade all installed software
winget upgrade --all --include-unknown

# Uninstall an application
winget uninstall --id <Package.ID>
```

### Essential Developer Software IDs:

| Category | Software | Winget Command |
| :--- | :--- | :--- |
| **Terminal & Shell** | PowerShell 7 | `winget install Microsoft.PowerShell --exact` |
| **Terminal & Shell** | Windows Terminal | `winget install Microsoft.WindowsTerminal --exact` |
| **Code Editor** | VS Code | `winget install Microsoft.VisualStudioCode --exact` |
| **Version Control** | Git for Windows | `winget install Git.Git --exact` |
| **Utilities** | Microsoft PowerToys | `winget install Microsoft.PowerToys --exact` |
| **Archiver** | 7-Zip | `winget install 7zip.7zip --exact` |
| **Browser** | Google Chrome | `winget install Google.Chrome --exact` |
| **Browser** | Mozilla Firefox | `winget install Mozilla.Firefox --exact` |
| **Font** | Cascadia Code Nerd Font | `winget install RomanKogan.CascadiaCodeNerdFont --exact` |

---

## 2. Scoop (Alternative: Developer-friendly user-space manager)

Scoop installs programs into `%USERPROFILE%\scoop`, requiring no Administrator elevation and no clutter in Program Files:

```powershell
# Install Scoop (in PowerShell)
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
irm get.scoop.op | iex

# Add buckets
scoop bucket add extras

# Install tools
scoop install git neovim ripgrep fzf
```
