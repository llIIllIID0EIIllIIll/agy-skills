# Windows Terminal & PowerShell Customization

Read this before configuring Windows Terminal, PowerShell profiles, fonts, keybindings, or prompt themes.

---

## 1. Windows Terminal Configuration

The master configuration file is `settings.json`.

### Location:
* **Packaged (Microsoft Store / Default)**:
  `$env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json`
* **Unpackaged / Portable**:
  `$env:LOCALAPPDATA\Microsoft\Windows Terminal\settings.json`

### Quick Inspection & Backup:
```powershell
$wtPath = "$env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json"
if (Test-Path $wtPath) {
    Copy-Item $wtPath "$wtPath.bak"
    Get-Content $wtPath | ConvertFrom-Json | Select-Object -ExpandProperty profiles
}
```

### Essential Settings in `settings.json`:
- **Default Profile**: Set `"defaultProfile": "{...guid...}"` to PowerShell 7 or Windows PowerShell.
- **Acrylic / Opacity**:
  ```json
  "opacity": 85,
  "useAcrylic": true
  ```
- **Font**:
  ```json
  "font": {
      "face": "Cascadia Code NF",
      "size": 11.5
  }
  ```
- **Color Scheme**: Set `"colorScheme": "One Half Dark"` (or `"Catppuccin Mocha"`, `"Tokyo Night"`).

---

## 2. PowerShell Profile Customization

### Profile Paths:
- Current User, Current Host: `$PROFILE`
  - PowerShell 7+: `C:\Users\<user>\Documents\PowerShell\Microsoft.PowerShell_profile.ps1`
  - Windows PowerShell 5.1: `C:\Users\<user>\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1`

### Creating / Editing the Profile:
```powershell
if (!(Test-Path $PROFILE)) {
    New-Item -Type File -Path $PROFILE -Force
}
notepad $PROFILE
```

### Recommended Additions to `$PROFILE`:

```powershell
# 1. Enhance PSReadLine (Autocomplete, Fish-like history, Emacs/Vi mode)
Set-PSReadLineOption -PredictionSource HistoryAndPlugin
Set-PSReadLineOption -PredictionViewStyle ListView
Set-PSReadLineOption -EditMode Windows
Set-PSReadLineKeyHandler -Key Tab -Function Complete

# 2. Useful Aliases
Set-Alias -Name ll -Value Get-ChildItem
Set-Alias -Name g -Value git
Set-Alias -Name grep -Value Select-String

# 3. Fast Directory Navigation
function .. { Set-Location .. }
function ... { Set-Location ..\.. }

# 4. Prompt Theme (if Oh My Posh or Starship is installed)
# oh-my-posh init pwsh --config "$env:POSH_THEMES_PATH\jandedobbeleer.omp.json" | Invoke-Expression
# Invoke-Expression (&starship init powershell)
```

---

## 3. Essential Command-line Tools via Winget

```powershell
# Install PowerShell 7 (Latest modern cross-platform PowerShell)
winget install --id Microsoft.PowerShell --exact --accept-package-agreements

# Install Nerd Fonts (for terminal glyphs and icons)
winget install --id RomanKogan.CascadiaCodeNerdFont --exact

# Install Starship prompt
winget install --id Starship.Starship --exact
```
