# Window Management & Tiling on Windows 11

Read this before configuring window layouts, shortcuts, Virtual Desktops, FancyZones, or tiling window managers.

---

## 1. Native Windows 11 Tiling & Snap

Windows 11 includes powerful built-in tiling capabilities.

### Keyboard Shortcuts:
| Shortcut | Action |
| :--- | :--- |
| **`Win + Z`** | Open **Snap Layouts** menu on the active window |
| **`Win + Left / Right`** | Snap active window to left / right half of screen |
| **`Win + Up`** | Maximize active window (or snap to top quadrant if already snapped) |
| **`Win + Down`** | Restore or minimize active window |
| **`Win + Home`** | Minimize all background windows except the active window |
| **`Win + D`** | Show desktop (toggle all minimized) |

### Virtual Desktops:
| Shortcut | Action |
| :--- | :--- |
| **`Win + Ctrl + D`** | Create a new Virtual Desktop |
| **`Win + Ctrl + Left / Right`** | Switch between Virtual Desktops |
| **`Win + Ctrl + F4`** | Close current Virtual Desktop |
| **`Win + Tab`** | Open Task View (visual grid of all windows and desktops) |

---

## 2. PowerToys FancyZones (Advanced Grids & Regions)

For users who want multi-monitor custom zones, column layouts, or priority grids:

### Installation:
```powershell
winget install --id Microsoft.PowerToys --exact --accept-package-agreements
```

### Key Behaviors:
- **`Win + Shift + ` `~`** (or `Win + Shift + Z`): Open FancyZones Editor to create custom grid layouts.
- **`Shift + Drag`**: Hold `Shift` while dragging any window to snap it into a designated zone.
- **`Win + Ctrl + Alt + Left/Right`**: Move window between custom zones using keyboard.

---

## 3. Tiling Window Managers (i3 / Hyprland style)

For users seeking automatic dynamic tiling on Windows 11:

### GlazeWM (Recommended, modern & lightweight):
* Inspired by i3 / bspwm / Hyprland.
* Written in Rust, highly configurable YAML config.
* Supports workspaces, gaps, borders, bar integration, and keyboard-driven navigation.

```powershell
# Install GlazeWM via winget
winget install --id GlazeWM.GlazeWM --exact
```

* Config location: `%USERPROFILE%\.glazewm\config.yaml`

### Komorebi:
* Another powerful tiling window manager for Windows with dynamic tiling algorithms and multi-monitor focus.
```powershell
winget install --id LGUG2Z.komorebi --exact
```
