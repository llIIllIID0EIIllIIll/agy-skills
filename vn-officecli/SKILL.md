---
name: vn-officecli
description: >-
  Soạn thảo, tạo mới và chuẩn hóa văn bản hành chính Việt Nam (Tờ trình, Công văn, Quyết định, Báo cáo) theo Nghị định 30/2020/NĐ-CP và chuẩn thể thức doanh nghiệp 2026. Đồng thời cung cấp đầy đủ công cụ tự động hóa tài liệu văn phòng Office (.docx, .xlsx, .pptx) bằng officecli.
  Kích hoạt khi người dùng yêu cầu: "tờ trình", "làm tờ trình", "xuất tờ trình", "văn bản hành chính", "thể thức văn bản", "nghị định 30", "công văn", "báo cáo", "quyết định", "tạo file word", "tạo file excel", "tạo slide powerpoint", "officecli", hoặc xử lý file .docx/.xlsx/.pptx.
---

# VN-OfficeCLI: Kỹ năng Xử lý Văn phòng & Thể thức Văn bản Việt Nam

Bộ kỹ năng chuyên sâu hỗ trợ Agent AI tạo mới, chuyển đổi và chuẩn hóa các văn bản hành chính theo đúng thể thức Việt Nam (Nghị định 30/2020/NĐ-CP và tiêu chuẩn doanh nghiệp cập nhật 2026), kết hợp cùng sức mạnh toàn diện của công cụ **OfficeCLI** cho bộ 3 tài liệu: Word (`.docx`), Excel (`.xlsx`) và PowerPoint (`.pptx`).

---

## 1. Bản đồ Kỹ năng & Hướng dẫn Chuyên sâu

Tùy theo loại công việc cần thực hiện, xem tài liệu chi tiết tương ứng:

| Mục tiêu công việc | Tài liệu hướng dẫn | Công cụ / Script hỗ trợ |
| :--- | :--- | :--- |
| **Văn bản hành chính & Tờ trình** | [`references/to-trinh-standard.md`](references/to-trinh-standard.md) | [`scripts/New-VnToTrinh.ps1`](scripts/New-VnToTrinh.ps1) |
| **Chuẩn hóa văn bản Word có sẵn** | [`references/to-trinh-standard.md`](references/to-trinh-standard.md) | [`scripts/Format-VnDocx.ps1`](scripts/Format-VnDocx.ps1) |
| **Soạn thảo văn bản Word (.docx)** | [`references/docx-guide.md`](references/docx-guide.md) | `officecli create / add / set` |
| **Bảng tính & Báo cáo Excel (.xlsx)**| [`references/xlsx-guide.md`](references/xlsx-guide.md) | `officecli create / set / import` |
| **Slide Thuyết trình PowerPoint (.pptx)**| [`references/pptx-guide.md`](references/pptx-guide.md) | `officecli create / add shape` |

---

## 2. Quy trình Soạn thảo Tờ trình Chuẩn Thể thức Việt Nam

Khi người dùng yêu cầu làm hoặc xuất **Tờ trình** (hoặc công văn, báo cáo):

### Bước 1: Thu thập / Xác nhận các trường thông tin cốt lõi
- **Trích yếu nội dung:** V/v phê duyệt vấn đề gì?
- **Đơn vị trình:** Phòng ban / đơn vị trực tiếp đề xuất.
- **Cơ quan cấp trên / Công ty:** Tên doanh nghiệp chủ quản.
- **Kính gửi:** Ban Lãnh đạo / Hội đồng Quản trị / Ban Tổng Giám đốc.
- **Người ký & Chức vụ:** Họ tên và chức danh.

### Bước 2: Tạo văn bản bằng Script tự động
Chạy script PowerShell đã tích hợp sẵn để sinh toàn bộ khung văn bản chuẩn:

```powershell
& "Z:\skills\vn-officecli\scripts\New-VnToTrinh.ps1" `
    -OutputPath "ToTrinh_DeXuat.docx" `
    -Title "V/v phê duyệt kế hoạch triển khai công nghệ năm 2026" `
    -Recipient "Hội đồng Quản trị / Ban Tổng Giám đốc" `
    -Organization "PHÒNG CÔNG NGHỆ THÔNG TIN" `
    -ParentOrg "TẬP ĐOÀN CÔNG NGHỆ ABC" `
    -DocNumber "12/TTr-CNTT" `
    -SignerTitle "TRƯỞNG PHÒNG" `
    -SignerName "Nguyễn Văn A"
```

### Bước 3: Đảm bảo Cấu trúc 3 Phần chuẩn của Tờ trình
1. **Phần I: SỰ CẦN THIẾT VÀ CĂN CỨ THỰC TIỄN** (Bối cảnh, thực trạng, mục tiêu).
2. **Phần II: NỘI DUNG ĐỀ XUẤT VÀ PHƯƠNG ÁN TRIỂN KHAI** (Chi tiết giải pháp, ngân sách, nhân sự, tiến độ, đánh giá rủi ro).
3. **Phần III: KIẾN NGHỊ VÀ ĐỀ XUẤT PHÊ DUYỆT** (Tóm tắt các điểm cần Lãnh đạo ký duyệt).

---

## 3. Quy chuẩn Thể thức Hành chính Cốt lõi (Bắt buộc tuân thủ)

- **Khổ giấy:** A4 đứng (210mm x 297mm).
- **Căn lề (Margins):**
  - **Trên (Top):** 25mm (1417 twips)
  - **Dưới (Bottom):** 20mm (1134 twips)
  - **Trái (Left):** 30mm (1701 twips) — Chừa khoảng đóng gáy hồ sơ
  - **Phải (Right):** 15mm (850 twips)
- **Typography:**
  - Font chữ: **Times New Roman**.
  - Cỡ chữ nội dung: **13pt**.
  - Giãn dòng: **1.3x**, căn đều hai bên (`alignment=both` / `justify`).
  - Thụt đầu dòng đoạn văn: **1.25cm** (`firstLineIndent=709` twips).
  - Khoảng cách đoạn: Before 3pt, After 6pt.
- **Kỹ thuật chống vỡ trang:**
  - Header (Quốc hiệu vs Tên cơ quan) và Footer (Nơi nhận vs Chữ ký) **luôn đặt trong Bảng 2 cột ẩn viền** (`border.all=none`). Không dùng phím Tab hay Space để căn chỉnh.

---

## 4. Chuẩn hóa File Word Hiện có

Nếu người dùng cung cấp một file Word có sẵn và yêu cầu căn chỉnh theo chuẩn thể thức:

```powershell
& "Z:\skills\vn-officecli\scripts\Format-VnDocx.ps1" -FilePath "C:\duong-dan\tai-lieu.docx"
```

---

## 5. Lệnh OfficeCLI Thường dùng

- **Kiểm tra công cụ:** `officecli --version`
- **Cài đặt nhanh trên máy mới:** `irm https://raw.githubusercontent.com/iOfficeAI/OfficeCli/main/install.ps1 | iex`
- **Đọc tài liệu Word:** `officecli get doc.docx /body`
- **Lưu và giải phóng file:** `officecli close doc.docx`
- **Tạo bảng tính Excel:** `officecli create data.xlsx`
- **Tạo slide PowerPoint:** `officecli create presentation.pptx`
