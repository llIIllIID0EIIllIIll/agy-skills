# Hướng dẫn Xử lý Văn bản Word (.docx) với OfficeCLI

Tài liệu hướng dẫn các thao tác tạo mới, đọc, phân tích và chỉnh sửa file Microsoft Word (`.docx`) bằng công cụ `officecli`.

---

## 1. Các lệnh thao tác cơ bản

```bash
# Tạo file trống
officecli create document.docx

# Lấy cấu trúc toàn bộ file (dưới dạng cây phân cấp)
officecli get document.docx /

# Lấy nội dung phần thân (body)
officecli get document.docx /body

# Thiết lập thuộc tính tài liệu (Khổ giấy, font, lề)
officecli set document.docx / --prop defaultFont="Times New Roman" --prop defaultFontSize="13pt"

# Đóng và lưu toàn bộ thay đổi ra đĩa
officecli close document.docx
```

---

## 2. Thêm đoạn văn (Paragraphs) & Định dạng

```bash
# Đoạn văn thường (căn đều 2 bên, thụt đầu dòng 1.25cm)
officecli add doc.docx /body --type p --prop text="Nội dung đoạn văn bản..." --prop alignment=both --prop lineSpacing="1.3x" --prop firstLineIndent="709" --prop spaceBefore="3pt" --prop spaceAfter="6pt"

# Tiêu đề mục (Heading / Bold)
officecli add doc.docx /body --type p --prop text="I. NỘI DUNG CHÍNH" --prop bold=true --prop size="13pt" --prop spaceBefore="10pt" --prop spaceAfter="4pt"

# Tiêu đề văn bản căn giữa (Title)
officecli add doc.docx /body --type p --prop text="BÁO CÁO KẾT QUẢ CÔNG TÁC" --prop alignment=center --prop bold=true --prop size="16pt" --prop spaceBefore="18pt" --prop spaceAfter="12pt"
```

---

## 3. Bảng biểu (Tables)

```bash
# Thêm bảng 3 cột, 4 hàng, viền chuẩn
officecli add doc.docx /body --type table --prop cols=3 --prop rows=4 --prop width="100%"

# Thêm bảng ẩn viền (dùng cho Header hoặc Footer chữ ký)
officecli add doc.docx /body --type table --prop cols=2 --prop rows=1 --prop border.all=none --prop width="100%"

# Gán nội dung vào ô cụ thể: tr[Hàng]/tc[Cột]
officecli add doc.docx "/body/tbl[1]/tr[1]/tc[1]" --type p --prop text="STT" --prop alignment=center --prop bold=true
officecli add doc.docx "/body/tbl[1]/tr[1]/tc[2]" --type p --prop text="Nội dung công việc" --prop alignment=center --prop bold=true
officecli add doc.docx "/body/tbl[1]/tr[1]/tc[3]" --type p --prop text="Kinh phí (VNĐ)" --prop alignment=center --prop bold=true
```

---

## 4. Mục lục tự động (TOC) & Header/Footer

```bash
# Thêm mục lục tự động
officecli add doc.docx / --type toc

# Thêm Header và Footer kèm số trang
officecli add doc.docx / --type footer
officecli add doc.docx /footer --type p --prop text="Trang {PAGE} / {NUMPAGES}" --prop alignment=center
```
