# Windows 11 Registry Tweaks & System Preferences

Read this before modifying Windows 11 taskbar behavior, context menus, developer options, or system policies.

---

## 1. Safety Rules for Registry Edits

1. **Always export before modifying**:
   ```powershell
   reg export "HKCU\Software\Target\Key" "C:\backup.reg" /y
   ```
2. **Apply changes**: Most Explorer-related tweaks take effect after restarting Explorer:
   ```powershell
   Stop-Process -Name explorer -Force
   ```

---

## 2. Common Windows 11 Tweaks

### Restore Classic Full Right-Click Context Menu
Removes the "Show more options" extra click in File Explorer:
```powershell
# Restore Windows 10 style full context menu
$regPath = "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32"
New-Item -Path $regPath -Force | Out-Null
Set-ItemProperty -Path $regPath -Name "(Default)" -Value ""
Stop-Process -Name explorer -Force

# To revert back to modern Windows 11 menu:
# Remove-Item -Path "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}" -Recurse -Force
```

### Taskbar Alignment (Left vs Center)
```powershell
$advPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"

# Align Left (Windows 10 style)
Set-ItemProperty -Path $advPath -Name "TaskbarAl" -Value 0 -Type DWord

# Align Center (Windows 11 default)
# Set-ItemProperty -Path $advPath -Name "TaskbarAl" -Value 1 -Type DWord
```

### Show Hidden Files & File Extensions
```powershell
# Show file extensions (.txt, .exe, etc.)
Set-ItemProperty -Path $advPath -Name "HideFileExt" -Value 0 -Type DWord

# Show hidden files
Set-ItemProperty -Path $advPath -Name "Hidden" -Value 1 -Type DWord
Stop-Process -Name explorer -Force
```

### Disable Web Search in Start Menu (Bing Search in Start)
Keeps Start menu search purely local and fast:
```powershell
$searchPath = "HKCU:\Software\Policies\Microsoft\Windows\Explorer"
if (!(Test-Path $searchPath)) { New-Item -Path $searchPath -Force | Out-Null }
Set-ItemProperty -Path $searchPath -Name "DisableSearchBoxSuggestions" -Value 1 -Type DWord
```

### Enable Long Paths (>260 characters support)
Crucial for Node.js (`node_modules`), Python, and Git repositories with deep folder nesting:
```powershell
# Requires Administrator elevation
Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem" -Name "LongPathsEnabled" -Value 1 -Type DWord
```

### Enable Developer Mode (Sideloading, Symlinks without admin)
```powershell
# Requires Administrator elevation
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock" -Name "AllowDevelopmentWithoutDevLicense" -Value 1 -Type DWord
```
