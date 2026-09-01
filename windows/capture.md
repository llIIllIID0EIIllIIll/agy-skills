# Screen Capture, Recording, and OCR

Read this before capturing screenshots, recording the screen, extracting text from images, or sharing files.

---

## 1. Screenshots (Snipping Tool)

Windows 11 features a built-in Snipping Tool with hotkeys:

| Shortcut | Description |
| :--- | :--- |
| **`Win + Shift + S`** | Interactive snipping toolbar (Rectangular, Freeform, Window, Full Screen). Automatically copies snip to clipboard. |
| **`PrtScn`** | Standard print screen (can be mapped in Settings $\rightarrow$ Accessibility $\rightarrow$ Keyboard to trigger `Win+Shift+S`). |
| **`Win + PrtScn`** | Captures entire screen and saves PNG directly into `Pictures\Screenshots`. |
| **`Alt + PrtScn`** | Captures only the currently active window to clipboard. |

---

## 2. Screen Recording

### Built-in Options:
1. **Snipping Tool Screen Recording**:
   - Press **`Win + Shift + S`** $\rightarrow$ Click the **Video Camera icon** at the top $\rightarrow$ Select region $\rightarrow$ Click **Start**.
2. **Xbox Game Bar**:
   - **`Win + Alt + R`**: Start / stop instant screen recording of the active application.
   - Files land in `Videos\Captures`.

---

## 3. Text Extraction (OCR)

To copy unselectable text directly from images, videos, or PDFs on screen:

### PowerToys Text Extractor (Recommended):
- Shortcut: **`Win + Shift + T`**
- Select any box of text on screen $\rightarrow$ Extracted plain text is copied to your clipboard immediately.
- Powered by local Windows OCR (no internet connection required).

---

## 4. Clipboard History

Never lose copied items:

- **`Win + V`**: Open **Clipboard History**.
- Allows pasting older copied items, pinning frequently used snippets, and pasting as plain text (`Ctrl + Shift + V` in modern apps).
- Enable via Settings: `ms-settings:clipboard` or press `Win + V` and click **Turn on**.

---

## 5. Sharing Files (Host ↔ Guest)

In this virtualized environment:
- **Shared Folder**:
  - Host directory: `~/Windows`
  - Windows guest location: **`Z:\`** or `%USERPROFILE%\Desktop\Shared`
  - Copying files to `Z:\` makes them immediately available on the Linux host.
