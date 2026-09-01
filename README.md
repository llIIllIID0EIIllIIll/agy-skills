# Antigravity (AGY) Skills for Windows

A curated collection of production-ready skills for **Google Antigravity (AGY)** on Windows, providing deep capabilities for system troubleshooting, Windows 11 optimization, Vietnamese administrative document processing, and Office automation (.docx, .xlsx, .pptx).

---

## 🚀 Quick Install (1-Line Command)

Open **PowerShell** on any Windows machine and run:

```powershell
irm https://raw.githubusercontent.com/llIIllIID0EIIllIIll/agy-skills/main/install.ps1 | iex
```

The installer automatically places the skills into Antigravity's discovery directories:
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

### 1. `vn-officecli` ⭐ *(Mới - Chuẩn Thể thức Việt Nam)*
> **Soạn thảo, tạo mới và chuẩn hóa Tờ trình & Văn bản hành chính Việt Nam theo Nghị định 30/2020/NĐ-CP và tiêu chuẩn doanh nghiệp 2026. Tích hợp bộ công cụ tự động hóa Office (.docx, .xlsx, .pptx) bằng OfficeCLI.**

- **Triggers:** Tờ trình, to trinh, văn bản hành chính, thể thức văn bản, nghị định 30, công văn, báo cáo, quyết định, tạo file word, tạo file excel, tạo slide powerpoint, officecli.
- **Tính năng nổi bật:**
  - **Tạo Tờ trình chuẩn thể thức tự động:** Khổ A4, lề chuẩn (Trên 25mm, Dưới 20mm, Trái 30mm, Phải 15mm), font Times New Roman 13pt, dãn dòng 1.3x, thụt đầu dòng 1.25cm.
  - **Bố cục Header & Chữ ký chống xô lệch:** Tự động dùng bảng 2 cột ẩn viền (`border.all=none`) chuẩn xác tuyệt đối.
  - **Chuẩn cấu trúc 3 phần Tờ trình:** Sự cần thiết & căn cứ thực tiễn -> Nội dung đề xuất & phương án triển khai -> Kiến nghị phê duyệt.
  - **Bộ tài liệu hướng dẫn toàn diện:**
    - [`references/to-trinh-standard.md`](vn-officecli/references/to-trinh-standard.md): Quy định chi tiết 14 thành phần thể thức hành chính Việt Nam.
    - [`references/docx-guide.md`](vn-officecli/references/docx-guide.md): Cẩm nang soạn thảo Word .docx với officecli.
    - [`references/xlsx-guide.md`](vn-officecli/references/xlsx-guide.md): Cẩm nang bảng tính Excel, công thức, bảng tổng hợp.
    - [`references/pptx-guide.md`](vn-officecli/references/pptx-guide.md): Cẩm nang thiết kế slide PowerPoint, pitch deck.
  - **Script hỗ trợ sẵn sàng:**
    - [`scripts/New-VnToTrinh.ps1`](vn-officecli/scripts/New-VnToTrinh.ps1): Script sinh nhanh file Word Tờ trình mẫu hoàn chỉnh.
    - [`scripts/Format-VnDocx.ps1`](vn-officecli/scripts/Format-VnDocx.ps1): Script chuẩn hóa file Word có sẵn sang chuẩn thể thức VN.

### 2. `diagnose-crash`
> **Chẩn đoán nguyên nhân sự cố ứng dụng bị crash, freeze, unhandled exception và màn hình xanh (BSOD) trên Windows.**

- **Triggers:** Crash, exception, `0xC0000005`, Access Violation, Event ID 1000, WER, minidump, `.dmp`, BSOD, BugCheck.
- **Khả năng:**
  - Tự động tra cứu Windows Event Logs (`Application` & `System`).
  - Kiểm tra tài nguyên RAM, ảo hóa bộ nhớ và cạn kiệt pagefile (Commit Limit).
  - Tự động phát hiện crash dump (`%LOCALAPPDATA%\CrashDumps`, `C:\Windows\Minidump`).
  - Phân tích và trích xuất backtrace bằng CDB / WinDbg (`!analyze -v`).
  - Bảng tra cứu mã lỗi ngoại lệ NTSTATUS phổ biến.

### 3. `windows`
> **Quản trị và tùy biến toàn diện hệ thống Windows 11, terminal, tiling, theme và registry.**

- **Triggers:** Windows 11, PowerShell profile, `winget`, Windows Terminal, `settings.json`, Snap Layouts, FancyZones, GlazeWM, Dark Mode, accent color, wallpaper, registry tweaks.
- **Tài liệu hướng dẫn con:**
  - `terminal.md`: Windows Terminal, PowerShell 7, PSReadLine, Oh My Posh, Starship.
  - `tiling.md`: Snap Layouts (`Win+Z`), PowerToys FancyZones, GlazeWM.
  - `theming.md`: Dark/Light theme, Mica/Acrylic effects, wallpaper automation.
  - `capture.md`: Snipping Tool, quay video màn hình, OCR Text Extractor.
  - `registry.md`: Tinh chỉnh Registry an toàn, căn giữa taskbar, menu chuột phải cổ điển.
  - `packages.md`: Quản lý cài đặt/gỡ bỏ phần mềm bằng `winget`, Scoop, Chocolatey.

---

## 🔍 Kiểm tra Kỹ năng sau khi cài đặt

Khởi động Antigravity CLI và gõ:

```text
agy
```

Trong dấu nhắc lệnh tương tác, gõ `/skills` để kiểm tra danh sách toàn bộ kỹ năng đã sẵn sàng.

---

## 📄 License

MIT License.
