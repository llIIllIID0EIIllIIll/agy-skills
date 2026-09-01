# Theming, Colors, and Desktop Appearance

Read this before changing themes, Dark/Light mode, transparency effects, accent colors, wallpaper, or night light.

---

## 1. Dark Mode / Light Mode Automation

Windows 11 separates app theme and system (taskbar/shell) theme.

### Switch to Dark Mode via PowerShell / Registry:
```powershell
$themeReg = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"

# Apps Dark Mode (0 = Dark, 1 = Light)
Set-ItemProperty -Path $themeReg -Name "AppsUseLightTheme" -Value 0 -Type DWord

# System / Taskbar Dark Mode (0 = Dark, 1 = Light)
Set-ItemProperty -Path $themeReg -Name "SystemUsesLightTheme" -Value 0 -Type DWord
```

### Switch to Light Mode:
```powershell
Set-ItemProperty -Path $themeReg -Name "AppsUseLightTheme" -Value 1 -Type DWord
Set-ItemProperty -Path $themeReg -Name "SystemUsesLightTheme" -Value 1 -Type DWord
```

---

## 2. Transparency (Mica / Acrylic) & Accent Color

```powershell
# Enable Transparency effects
Set-ItemProperty -Path $themeReg -Name "EnableTransparency" -Value 1 -Type DWord

# Show accent color on Title bars and window borders
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\DWM" -Name "ColorPrevalence" -Value 1 -Type DWord

# Show accent color on Start and Taskbar
Set-ItemProperty -Path $themeReg -Name "ColorPrevalence" -Value 1 -Type DWord
```

---

## 3. Setting Desktop Wallpaper Programmatically

You can change the desktop wallpaper instantly without opening GUI settings using Windows API:

```powershell
function Set-WallPaper {
    param([Parameter(Mandatory=$true)][string]$Path)
    $resolvedPath = (Resolve-Path $Path).Path
    $code = @"
using System;
using System.Runtime.InteropServices;
public class Wallpaper {
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
}
"@
    Add-Type -TypeDefinition $code -ErrorAction SilentlyContinue
    [Wallpaper]::SystemParametersInfo(20, 0, $resolvedPath, 3) | Out-Null
}

# Example usage:
# Set-WallPaper -Path "C:\path\to\wallpaper.jpg"
```

---

## 4. Night Light (Blue Light Reduction)

- Open Night Light settings page directly:
  ```powershell
  Start-Process "ms-settings:nightlight"
  ```
- Fast toggle: Click **Action Center** (`Win + A`) $\rightarrow$ Click the **Night light** quick tile.
