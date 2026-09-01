# Antigravity (AGY) Skills for Windows

A curated collection of production-ready skills for **Google Antigravity (AGY)** on Windows, providing deep capabilities for system troubleshooting, Windows 11 optimization, Vietnamese administrative document processing, and Office automation (.docx, .xlsx, .pptx).

---

## 🚀 Quick Install (1-Line Command)

Open **PowerShell** on any Windows machine and run:

```powershell
irm https://raw.githubusercontent.com/llIIllIID0EIIllIIll/agy-skills/main/install.ps1 | iex
```

The zero-config installer automatically:
- Checks and auto-installs the **OfficeCLI engine** if missing.
- Configures environment `PATH`.
- Installs all 3 skills into Antigravity discovery paths (`~/.gemini/antigravity-cli/skills/`, `~/.gemini/config/skills/`, `~/.agents/skills/`).

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

### 1. `vn-officecli` ⭐ *(Chuẩn Thể thức Doanh nghiệp Hoàn Cầu 2025/2026 & NĐ 30/2020)*
> **Soạn thảo, tạo mới và chuẩn hóa văn bản hành chính Việt Nam tuân thủ nghiêm ngặt Quy định số 01/2025/QyĐ-PC (Tập đoàn Hoàn Cầu) và Nghị định 30/2020/NĐ-CP. Tích hợp trọn bộ công cụ tự động hóa Office (.docx, .xlsx, .pptx) bằng OfficeCLI.**

- **Triggers:** Tờ trình, to trinh, văn bản hành chính, thể thức văn bản, nghị định 30, quy định 01/2025, hoàn cầu, công văn, báo cáo, quyết định, biên bản họp, giấy ủy quyền, tạo file word, tạo file excel, tạo slide powerpoint, officecli.
- **Tính năng nổi bật:**
  - **Chuẩn hóa thông số kỹ thuật A4:** Lề trên 25mm, dưới 20mm, trái 30mm, phải 15mm; font Times New Roman 13pt; giãn dòng 1.3x; thụt đầu dòng 1.25cm.
  - **Chuẩn 14 ô thành phần thể thức:** Ma trận kiểu chữ, cỡ chữ và vị trí chính xác tuyệt đối theo Phụ lục 02 & 03.
  - **Hệ thống 33 mã viết tắt loại văn bản:** Chuẩn hóa toàn bộ tên loại văn bản (`TTr`, `CV`, `QĐ`, `QyĐ`, `QC`, `QT`, `BC`, `BB`, `TB`, `KH`, `NQ`, `UQ`...).
  - **Tờ trình Doanh nghiệp chuyên sâu (Mẫu 06 - Trang 30):**
    - Bố cục 3 phần chuẩn (Sự cần thiết -> Đề xuất phương án -> Kiến nghị phê duyệt).
    - Hỗ trợ ghi chú `(Trình lần thứ ...)`.
    - Tích hợp khối **Ý kiến thẩm định** và **Phê duyệt** của Lãnh đạo cấp cao ở chân trang.
  - **9 Mẫu biểu hành chính chuẩn Doanh nghiệp (Phụ lục 05):**
    - Công văn (`CV`), Nghị quyết HĐTV (`NQ`), Quyết định HĐTV & TGĐ (`QĐ`), Thông báo (`TB`), Tờ trình (`TTr`), Báo cáo (`BC`), Biên bản họp (`BB`), Giấy ủy quyền (`UQ`).
  - **Quy chuẩn Lưu đồ Quy trình BPMN 2.0 (Phụ lục 04):** Hướng dẫn vẽ lưu đồ và lập bảng mô tả quy trình 5 cột chuẩn.
  - **Bộ tài liệu tham chiếu & Script thực thi:**
    - [`references/to-trinh-standard.md`](vn-officecli/references/to-trinh-standard.md): Chi tiết 14 ô thể thức và quy định kỹ thuật A4.
    - [`references/administrative-templates.md`](vn-officecli/references/administrative-templates.md): Chi tiết 9 mẫu biểu hành chính chuẩn.
    - [`references/docx-guide.md`](vn-officecli/references/docx-guide.md): Cẩm nang soạn thảo Word .docx với officecli.
    - [`references/xlsx-guide.md`](vn-officecli/references/xlsx-guide.md): Cẩm nang bảng tính Excel, công thức, dashboard.
    - [`references/pptx-guide.md`](vn-officecli/references/pptx-guide.md): Cẩm nang thiết kế slide PowerPoint, pitch deck.
    - [`scripts/New-VnToTrinh.ps1`](vn-officecli/scripts/New-VnToTrinh.ps1): Script sinh file Word Tờ trình mẫu Hoàn Cầu.
    - [`scripts/New-VnCongVan.ps1`](vn-officecli/scripts/New-VnCongVan.ps1): Script sinh file Word Công văn chuẩn.
    - [`scripts/Format-VnDocx.ps1`](vn-officecli/scripts/Format-VnDocx.ps1): Script chuẩn hóa file Word có sẵn.

### 2. `diagnose-crash`
> **Chẩn đoán nguyên nhân sự cố ứng dụng bị crash, freeze, unhandled exception và màn hình xanh (BSOD) trên Windows.**

- **Triggers:** Crash, exception, `0xC0000005`, Access Violation, Event ID 1000, WER, minidump, `.dmp`, BSOD, BugCheck.
- **Khả năng:**
  - Tự động tra cứu Windows Event Logs (`Application` & `System`).
  - Kiểm tra tài nguyên RAM, ảo hóa bộ nhớ và cạn kiệt pagefile.
  - Tự động phát hiện crash dump (`%LOCALAPPDATA%\CrashDumps`, `C:\Windows\Minidump`).
  - Phân tích và trích xuất backtrace bằng CDB / WinDbg (`!analyze -v`).

### 3. `windows`
> **Quản trị và tùy biến toàn diện hệ thống Windows 11, terminal, tiling, theme và registry.**

- **Triggers:** Windows 11, PowerShell profile, `winget`, Windows Terminal, `settings.json`, Snap Layouts, FancyZones, GlazeWM, Dark Mode, accent color, wallpaper, registry tweaks.
- **Tài liệu hướng dẫn:**
  - `terminal.md`: Windows Terminal, PowerShell 7, Oh My Posh, Starship.
  - `tiling.md`: Snap Layouts (`Win+Z`), PowerToys FancyZones, GlazeWM.
  - `theming.md`: Dark/Light theme, Mica/Acrylic effects, wallpaper automation.
  - `capture.md`: Snipping Tool, quay video màn hình, OCR Text Extractor.
  - `registry.md`: Tinh chỉnh Registry an toàn.
  - `packages.md`: Quản lý phần mềm qua `winget`, Scoop, Chocolatey.

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
