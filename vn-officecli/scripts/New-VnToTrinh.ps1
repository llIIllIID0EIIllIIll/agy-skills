<#
.SYNOPSIS
    Tạo văn bản Tờ trình (TTr) chuẩn thể thức Tập đoàn Hoàn Cầu (QyĐ 01/2025/QyĐ-PC) & Nghị định 30/2020/NĐ-CP.

.DESCRIPTION
    Tạo tài liệu Word (.docx) chuyên nghiệp tuân thủ nghiêm ngặt Mẫu số 06 (Phụ lục 05 - Trang 30):
    - Tự động chuẩn hóa nội dung: chống lặp "V/v:", "Số:", "Kính gửi:", tự động format ngày tháng, số lần trình.
    - Khổ giấy A4 (210 x 297mm), Lề: Trên 25mm, Dưới 20mm, Trái 30mm, Phải 15mm.
    - Font Times New Roman 13pt, giãn dòng 1.3x, thụt đầu dòng 1.25cm.
    - Header & Chân trang bố cục bằng bảng ẩn viền chống xô lệch trang.
    - Đầy đủ 3 phần chuẩn Tờ trình + Danh mục tài liệu gửi kèm + 2 khối chân trang đặc thù: Ý kiến thẩm định & Phê duyệt.
    - Thực thi siêu tốc qua cơ chế Batch Processing của OfficeCLI.
#>

[CmdletBinding()]
param(
    [string]$OutputPath = "ToTrinh_HoanCau_Chuan.docx",
    [string]$Title = "Phê duyệt đề xuất chủ trương triển khai hệ thống quản trị 2025/2026",
    [string]$SubmissionCount = "", # Ví dụ: "1", "Trình lần thứ nhất", "lần 2"
    [string]$Recipient = "Hội đồng Thành viên / Ban Tổng Giám đốc Tập đoàn Hoàn Cầu",
    [string]$Organization = "PHÒNG CÔNG NGHỆ VÀ VẬN HÀNH",
    [string]$ParentOrg = "CÔNG TY TNHH HOÀN CẦU",
    [string]$DocNumber = "01/2025/TTr-CNVH",
    [string]$Location = "TP. HCM",
    [object]$Date = $null,
    [string]$SignerPrefix = "", # "TM.", "KT.", "TL.", "TUQ." hoặc rỗng
    [string]$SignerTitle = "TRƯỞNG PHÒNG",
    [string]$SignerName = "Võ Thị Thanh Tuyền",
    [string[]]$Receivers = @("Như kính gửi;", "Ban Tổng Giám đốc;", "Khối Pháp chế & Kiểm soát;", "Phòng Kế toán - Tài chính;", "Lưu: VT, CNVH."),
    [string[]]$AttachDocs = @("Bản thuyết minh dự án chi tiết;", "Bảng dự toán kinh phí và hiệu quả đầu tư;", "Báo giá đối tác cung cấp."),
    [switch]$UseCorporateHeader # Header doanh nghiệp nội bộ (Tầng 9 Nam Á Bank, CMT8, TP.HCM)
)

$ErrorActionPreference = "Stop"

# --- CÁC HÀM CHUẨN HÓA DỮ LIỆU ĐẦU VÀO ---

function Normalize-VnSubject {
    param([string]$Text, [string]$Prefix = "V/v:")
    if ([string]::IsNullOrWhiteSpace($Text)) { return "$Prefix [Trích yếu nội dung tờ trình]" }
    $clean = $Text.Trim()
    # Bóc tách triệt để các dạng tiền tố: "v/v", "v/v:", "về việc", "về việc:", "ve viec", "trích yếu"
    $clean = $clean -replace '^(?i)(v/v|về việc|ve viec|trích yếu)\s*[:：\-.]?\s*', ''
    return "$Prefix $clean"
}

function Normalize-VnDocNumber {
    param([string]$Text, [string]$Default = "01/2025/TTr-CNVH")
    if ([string]::IsNullOrWhiteSpace($Text)) { return "Số: $Default" }
    $clean = $Text.Trim() -replace '^(?i)Số\s*[:：]?\s*', ''
    return "Số: $clean"
}

function Get-VnCleanRecipient {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return "Hội đồng Thành viên / Ban Tổng Giám đốc" }
    return $Text.Trim() -replace '^(?i)Kính gửi\s*[:：]?\s*', ''
}

function Normalize-VnSubmissionCount {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $clean = $Text.Trim() -replace '^\((.*)\)$', '$1'
    if ($clean -notmatch '^(?i)Trình') {
        if ($clean -match '^\d+$') {
            $clean = "Trình lần thứ $clean"
        } else {
            $clean = "Trình $clean"
        }
    }
    return "($clean)"
}

function Normalize-VnDate {
    param([string]$Loc = "TP. HCM", [object]$DInput = $null)
    if ($DInput -is [datetime]) { $d = $DInput }
    elseif ($DInput -is [string] -and $DInput -match '^\d{4}-\d{2}-\d{2}') { $d = [datetime]::Parse($DInput) }
    else { $d = Get-Date }

    $dayStr = if ($d.Day -lt 10) { "0" + $d.Day } else { [string]$d.Day }
    $monthStr = if ($d.Month -lt 10) { "0" + $d.Month } else { [string]$d.Month }
    return "$Loc, ngày $dayStr tháng $monthStr năm $($d.Year)"
}

# --- BƯỚC 1: XÁC ĐỊNH VÀ CHUẨN HÓA DỮ LIỆU ---

$normTitle = Normalize-VnSubject -Text $Title -Prefix "V/v:"
$normDocNumber = Normalize-VnDocNumber -Text $DocNumber
$cleanRecipient = Get-VnCleanRecipient -Text $Recipient
$recipientParagraph = "Kính gửi: $cleanRecipient"
$normSubmission = Normalize-VnSubmissionCount -Text $SubmissionCount
$normDateString = Normalize-VnDate -Loc $Location -DInput $Date

# --- BƯỚC 2: KIỂM TRA ENGINE OFFICECLI ---

$cli = Get-Command officecli -ErrorAction SilentlyContinue
if (-not $cli) {
    $candidates = @("$env:LOCALAPPDATA\agy\bin\officecli.exe", "$env:LOCALAPPDATA\OfficeCLI\officecli.exe")
    $cliPath = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $cliPath) {
        throw "OfficeCLI chưa được cài đặt. Vui lòng chạy lệnh: irm https://raw.githubusercontent.com/llIIllIID0EIIllIIll/agy-skills/main/install.ps1 | iex"
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

Write-Host "[*] Khởi tạo tài liệu Tờ trình chuẩn Hoàn Cầu: $fullPath" -ForegroundColor Cyan
& $cliPath create $fullPath

# --- BƯỚC 3: DỰNG DANH SÁCH LỆNH BATCH SIÊU TỐC ---

$batchCmds = [System.Collections.Generic.List[object]]::new()

# 1. Page Setup A4 & Margins
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

if ($UseCorporateHeader) {
    $batchCmds.Add(@{ command = "add"; parent = "/body/tbl[1]/tr[1]/tc[1]"; type = "p"; props = @{ text = "(LOGO)"; alignment = "center"; bold = "true"; size = "12pt" } })
    $batchCmds.Add(@{ command = "add"; parent = "/body/tbl[1]/tr[1]/tc[1]"; type = "p"; props = @{ text = $normDocNumber; alignment = "center"; size = "12pt" } })
    $batchCmds.Add(@{ command = "add"; parent = "/body/tbl[1]/tr[1]/tc[2]"; type = "p"; props = @{ text = "TẬP ĐOÀN HOÀN CẦU"; alignment = "center"; bold = "true"; size = "13pt" } })
    $batchCmds.Add(@{ command = "add"; parent = "/body/tbl[1]/tr[1]/tc[2]"; type = "p"; props = @{ text = "Tầng 9, Tòa nhà Nam Á, 201-203 CMT8, P. Bàn Cờ, TP.HCM"; alignment = "center"; size = "10pt" } })
    $batchCmds.Add(@{ command = "add"; parent = "/body/tbl[1]/tr[1]/tc[2]"; type = "p"; props = @{ text = "ĐT: 08.3832.9897 - Fax: 08.3832.9894"; alignment = "center"; size = "10pt" } })
    $batchCmds.Add(@{ command = "add"; parent = "/body/tbl[1]/tr[1]/tc[2]"; type = "p"; props = @{ text = $normDateString; alignment = "center"; italic = "true"; size = "12pt" } })
} else {
    if ($ParentOrg) {
        $batchCmds.Add(@{ command = "add"; parent = "/body/tbl[1]/tr[1]/tc[1]"; type = "p"; props = @{ text = $ParentOrg.ToUpper(); alignment = "center"; size = "11pt"; lineSpacing = "1.15x" } })
    }
    $batchCmds.Add(@{ command = "add"; parent = "/body/tbl[1]/tr[1]/tc[1]"; type = "p"; props = @{ text = $Organization.ToUpper(); alignment = "center"; bold = "true"; size = "12pt"; lineSpacing = "1.15x" } })
    $batchCmds.Add(@{ command = "add"; parent = "/body/tbl[1]/tr[1]/tc[1]"; type = "p"; props = @{ text = $normDocNumber; alignment = "center"; size = "12pt"; lineSpacing = "1.15x" } })

    $batchCmds.Add(@{ command = "add"; parent = "/body/tbl[1]/tr[1]/tc[2]"; type = "p"; props = @{ text = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"; alignment = "center"; bold = "true"; size = "12pt"; lineSpacing = "1.15x" } })
    $batchCmds.Add(@{ command = "add"; parent = "/body/tbl[1]/tr[1]/tc[2]"; type = "p"; props = @{ text = "Độc lập - Tự do - Hạnh phúc"; alignment = "center"; bold = "true"; size = "13pt"; lineSpacing = "1.15x" } })
    $batchCmds.Add(@{ command = "add"; parent = "/body/tbl[1]/tr[1]/tc[2]"; type = "p"; props = @{ text = $normDateString; alignment = "center"; italic = "true"; size = "13pt"; lineSpacing = "1.15x" } })
}

# 3. Tiêu đề TỜ TRÌNH & Trích yếu
$batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = @{ text = "TỜ TRÌNH"; alignment = "center"; bold = "true"; size = "16pt"; spaceBefore = "18pt"; spaceAfter = "4pt" } })

$titleSpaceAfter = if ($normSubmission) { "2pt" } else { "14pt" }
$batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = @{ text = $normTitle; alignment = "center"; bold = "true"; italic = "true"; size = "13pt"; spaceAfter = $titleSpaceAfter } })

if ($normSubmission) {
    $batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = @{ text = $normSubmission; alignment = "center"; italic = "true"; size = "12pt"; spaceAfter = "14pt" } })
}

# 4. Kính gửi
$batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = @{ text = $recipientParagraph; bold = "true"; size = "13pt"; spaceAfter = "12pt" } })

# 5. Nội dung Tờ trình (3 Phần chuẩn Hoàn Cầu)
$bodyPProps = @{ alignment = "both"; size = "13pt"; lineSpacing = "1.3x"; firstLineIndent = "709"; spaceBefore = "3pt"; spaceAfter = "6pt" }

$pIntro1 = $bodyPProps.Clone(); $pIntro1["text"] = "- Căn cứ Quy chế tổ chức và hoạt động của Tập đoàn Hoàn Cầu;"
$batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = $pIntro1 })

$pIntro2 = $bodyPProps.Clone(); $pIntro2["text"] = "- Căn cứ chức năng, nhiệm vụ và nhu cầu thực tế trong công tác quản trị, điều hành hoạt động của $Organization;"
$batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = $pIntro2 })

$pIntro3 = $bodyPProps.Clone(); $pIntro3["text"] = "$Organization kính trình $cleanRecipient xem xét và phê duyệt nội dung chi tiết như sau:"
$batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = $pIntro3 })

# Phần I
$batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = @{ text = "I. SỰ CẦN THIẾT VÀ CĂN CỨ THỰC TIỄN"; bold = "true"; size = "13pt"; spaceBefore = "10pt"; spaceAfter = "4pt" } })

$pI1 = $bodyPProps.Clone(); $pI1["text"] = "1. Bối cảnh và lý do đề xuất: Nêu rõ thực trạng công việc, các khó khăn bất cập tồn tại hoặc cơ hội phát triển mới đòi hỏi phải thực hiện đề xuất."
$batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = $pI1 })

$pI2 = $bodyPProps.Clone(); $pI2["text"] = "2. Mục tiêu hướng tới: Tối ưu hóa hiệu quả vận hành, tiết giảm chi phí, nâng cao năng suất lao động và đảm bảo tuân thủ nghiêm ngặt các quy định pháp luật."
$batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = $pI2 })

# Phần II
$batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = @{ text = "II. NỘI DUNG ĐỀ XUẤT VÀ PHƯƠNG ÁN TRIỂN KHAI"; bold = "true"; size = "13pt"; spaceBefore = "10pt"; spaceAfter = "4pt" } })

$pII1 = $bodyPProps.Clone(); $pII1["text"] = "1. Nội dung chi tiết phương án: Trình bày chi tiết các giải pháp kỹ thuật, phạm vi áp dụng, tiêu chuẩn lựa chọn và quy trình thực hiện."
$batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = $pII1 })

$pII2 = $bodyPProps.Clone(); $pII2["text"] = "2. Dự toán ngân sách và nguồn lực: Bảng tổng hợp khái toán kinh phí, nhân sự tham gia triển khai và kế hoạch phân kỳ giải ngân vốn."
$batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = $pII2 })

$pII3 = $bodyPProps.Clone(); $pII3["text"] = "3. Tiến độ thực hiện: Lộ trình triển khai cụ thể theo từng mốc thời gian từ khâu chuẩn bị, thẩm định đến hoàn thành bàn giao nghiệm thu."
$batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = $pII3 })

# Phần III
$batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = @{ text = "III. KIẾN NGHỊ VÀ ĐỀ XUẤT PHÊ DUYỆT"; bold = "true"; size = "13pt"; spaceBefore = "10pt"; spaceAfter = "4pt" } })

$pIIIIntro = $bodyPProps.Clone(); $pIIIIntro["text"] = "Để đảm bảo tiến độ triển khai thông suốt và đạt hiệu quả tối ưu, $Organization kính đề nghị $cleanRecipient xem xét và quyết định:"
$batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = $pIIIIntro })

$pIII1 = $bodyPProps.Clone(); $pIII1["text"] = "1. Phê duyệt chủ trương và phương án triển khai chi tiết theo nội dung nêu tại Phần II."
$batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = $pIII1 })

$pIII2 = $bodyPProps.Clone(); $pIII2["text"] = "2. Phê duyệt dự toán kinh phí thực hiện và giao các phòng ban, đơn vị liên quan phối hợp triển khai theo đúng thẩm quyền."
$batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = $pIII2 })

$closingPProps = $bodyPProps.Clone()
$closingPProps["text"] = "Kính trình Lãnh đạo xem xét, phê duyệt./."
$closingPProps["spaceAfter"] = "18pt"
$batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = $closingPProps })

# 6. Khối Chân trang 1: Bảng 2 cột (Nơi nhận & Chữ ký)
$batchCmds.Add(@{
    command = "add"; parent = "/body"; type = "table"
    props = @{ cols = 2; rows = 1; "border.all" = "none"; width = "100%" }
})

$batchCmds.Add(@{ command = "add"; parent = "/body/tbl[2]/tr[1]/tc[1]"; type = "p"; props = @{ text = "Nơi nhận:"; bold = "true"; italic = "true"; size = "10pt"; spaceAfter = "2pt" } })
foreach ($rcv in $Receivers) {
    $batchCmds.Add(@{ command = "add"; parent = "/body/tbl[2]/tr[1]/tc[1]"; type = "p"; props = @{ text = "- $rcv"; size = "10pt"; lineSpacing = "1.15x"; spaceAfter = "1pt" } })
}

if ($AttachDocs -and $AttachDocs.Count -gt 0) {
    $batchCmds.Add(@{ command = "add"; parent = "/body/tbl[2]/tr[1]/tc[1]"; type = "p"; props = @{ text = "Tài liệu gửi kèm:"; bold = "true"; size = "10pt"; spaceBefore = "6pt"; spaceAfter = "2pt" } })
    foreach ($doc in $AttachDocs) {
        $batchCmds.Add(@{ command = "add"; parent = "/body/tbl[2]/tr[1]/tc[1]"; type = "p"; props = @{ text = "- $doc"; italic = "true"; size = "10pt"; lineSpacing = "1.15x"; spaceAfter = "1pt" } })
    }
}

$signerTitleFull = if ($SignerPrefix) { "$SignerPrefix $SignerTitle" } else { $SignerTitle }
$batchCmds.Add(@{ command = "add"; parent = "/body/tbl[2]/tr[1]/tc[2]"; type = "p"; props = @{ text = $signerTitleFull; alignment = "center"; bold = "true"; size = "13pt"; spaceAfter = "45pt" } })
$batchCmds.Add(@{ command = "add"; parent = "/body/tbl[2]/tr[1]/tc[2]"; type = "p"; props = @{ text = $SignerName; alignment = "center"; bold = "true"; size = "13pt" } })

# 7. Khối Chân trang 2: Ý kiến thẩm định
$batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = @{ text = "Ý kiến thẩm định:"; bold = "true"; size = "13pt"; spaceBefore = "16pt"; spaceAfter = "4pt" } })
$batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = @{ text = "................................................................................................................................................................................"; size = "11pt"; spaceAfter = "4pt" } })
$batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = @{ text = "................................................................................................................................................................................"; size = "11pt"; spaceAfter = "4pt" } })
$batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = @{ text = "................................................................................................................................................................................"; size = "11pt"; spaceAfter = "10pt" } })

# 8. Khối Chân trang 3: Phê duyệt
$batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = @{ text = "Phê duyệt:"; bold = "true"; size = "13pt"; spaceBefore = "6pt"; spaceAfter = "4pt" } })
$batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = @{ text = "................................................................................................................................................................................"; size = "11pt"; spaceAfter = "4pt" } })
$batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = @{ text = "................................................................................................................................................................................"; size = "11pt"; spaceAfter = "4pt" } })
$batchCmds.Add(@{ command = "add"; parent = "/body"; type = "p"; props = @{ text = "................................................................................................................................................................................"; size = "11pt"; spaceAfter = "4pt" } })

# --- BƯỚC 4: THỰC THI BATCH QUA OFFICECLI ---

Write-Host "[+] Đang thực thi $($batchCmds.Count) thao tác định dạng văn bản bằng OfficeCLI Batch..." -ForegroundColor Yellow
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
Write-Host "  TẠO THÀNH CÔNG TỜ TRÌNH CHUẨN HOÀN CẦU 2025/2026!" -ForegroundColor Green
Write-Host "  Trích yếu : $normTitle" -ForegroundColor White
Write-Host "  Kính gửi  : $cleanRecipient" -ForegroundColor White
Write-Host "  Số hiệu   : $normDocNumber" -ForegroundColor White
Write-Host "  File xuất : $fullPath" -ForegroundColor White
Write-Host "===================================================" -ForegroundColor Cyan
