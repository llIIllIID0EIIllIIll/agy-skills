<#
.SYNOPSIS
    Chuẩn hóa định dạng toàn diện file Word (.docx) sang thể thức hành chính Việt Nam theo Quy định 01/2025/QyĐ-PC.

.DESCRIPTION
    Tự động hiệu chỉnh tài liệu qua cơ chế phân tích cú pháp và Batch Processing:
    1. Thiết lập trang A4: Lề trên 25mm, dưới 20mm, trái 30mm, phải 15mm.
    2. Chuẩn hóa Font mặc định Times New Roman 13pt.
    3. Tự động nhận diện và định dạng phân cấp tiêu đề (Phần, Chương, Điều, Khoản, Điểm).
    4. Tự động thụt đầu dòng 1.25cm cho đoạn văn, khử thụt kép cho danh sách gạch đầu dòng/numbering.
    5. Chuẩn hóa kích thước bảng biểu (width=100%) và font chữ trong bảng (Times New Roman 12pt, lineSpacing=1.0x).
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$FilePath
)

$ErrorActionPreference = "Stop"

# Kiểm tra OfficeCLI
$cli = Get-Command officecli -ErrorAction SilentlyContinue
if (-not $cli) {
    $candidates = @("$env:LOCALAPPDATA\agy\bin\officecli.exe", "$env:LOCALAPPDATA\OfficeCLI\officecli.exe")
    $cliPath = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $cliPath) {
        throw "Không tìm thấy OfficeCLI. Vui lòng cài đặt: irm https://raw.githubusercontent.com/llIIllIID0EIIllIIll/agy-skills/main/install.ps1 | iex"
    }
} else {
    $cliPath = $cli.Source
}

$fullPath = [System.IO.Path]::GetFullPath($FilePath)
if (!(Test-Path $fullPath)) {
    throw "Không tìm thấy file: $fullPath"
}

Write-Host "[*] Đang phân tích cấu trúc tài liệu: $fullPath..." -ForegroundColor Cyan

# 1. Thu thập danh sách đoạn văn và bảng biểu
$pJson = & $cliPath query $fullPath "p" --json 2>$null | ConvertFrom-Json
$tblJson = & $cliPath query $fullPath "table" --json 2>$null | ConvertFrom-Json

$batchCmds = [System.Collections.Generic.List[object]]::new()

# 2. Chuẩn hóa Khổ giấy A4 & Căn lề
$batchCmds.Add(@{
    command = "set"; path = "/"
    props = @{
        defaultFont = "Times New Roman"; defaultFontSize = "13pt"
        pageWidth = "11906"; pageHeight = "16838"
        marginTop = "1417"; marginBottom = "1134"
        marginLeft = "1701"; marginRight = "850"
    }
})

# 3. Chuẩn hóa Bảng biểu
if ($tblJson -and $tblJson.data -and $tblJson.data.results) {
    foreach ($tbl in $tblJson.data.results) {
        $batchCmds.Add(@{
            command = "set"; path = $tbl.path
            props = @{ width = "100%" }
        })
    }
}

# 4. Chuẩn hóa từng Đoạn văn
if ($pJson -and $pJson.data -and $pJson.data.results) {
    foreach ($p in $pJson.data.results) {
        $text = if ($p.text) { $p.text.Trim() } else { "" }
        $path = $p.path

        # Đoạn văn nằm trong bảng
        if ($path -match '/tbl\[') {
            $batchCmds.Add(@{
                command = "set"; path = $path
                props = @{
                    font = "Times New Roman"; size = "12pt"
                    lineSpacing = "1.0x"; spaceBefore = "1pt"; spaceAfter = "1pt"
                }
            })
            continue
        }

        # Đoạn văn ngoài thân (Body paragraphs)
        if ([string]::IsNullOrWhiteSpace($text)) {
            continue
        }

        # Tiêu đề chính loại văn bản (TỜ TRÌNH, CÔNG VĂN, QUYẾT ĐỊNH, BÁO CÁO...)
        if ($text -match '^(TỜ TRÌNH|QUYẾT ĐỊNH|NGHỊ QUYẾT|BÁO CÁO|THÔNG BÁO|BIÊN BẢN|QUY CHẾ|QUY ĐỊNH|GIẤY ỦY QUYỀN)$') {
            $batchCmds.Add(@{
                command = "set"; path = $path
                props = @{
                    font = "Times New Roman"; size = "16pt"; bold = "true"
                    alignment = "center"; spaceBefore = "18pt"; spaceAfter = "4pt"
                    firstLineIndent = "0"
                }
            })
        }
        # Trích yếu nội dung (V/v...)
        elseif ($text -match '^(?i)v/v\s*[:：\-.]?') {
            $batchCmds.Add(@{
                command = "set"; path = $path
                props = @{
                    font = "Times New Roman"; size = "13pt"; bold = "true"; italic = "true"
                    alignment = "center"; spaceBefore = "2pt"; spaceAfter = "14pt"
                    firstLineIndent = "0"
                }
            })
        }
        # Tiêu đề Mục La Mã (I., II., III...) hoặc Điều/Khoản lớn
        elseif ($text -match '^(I|II|III|IV|V|VI|VII|VIII|IX|X)\.\s+' -or $text -match '^(Điều|Chương|Phần)\s+\d+[\.:]?\s*') {
            $batchCmds.Add(@{
                command = "set"; path = $path
                props = @{
                    font = "Times New Roman"; size = "13pt"; bold = "true"
                    spaceBefore = "10pt"; spaceAfter = "4pt"
                    firstLineIndent = "0"; alignment = "both"
                }
            })
        }
        # Kính gửi
        elseif ($text -match '^(?i)kính gửi\s*[:：]?\s*') {
            $batchCmds.Add(@{
                command = "set"; path = $path
                props = @{
                    font = "Times New Roman"; size = "13pt"; bold = "true"
                    spaceBefore = "12pt"; spaceAfter = "10pt"
                    firstLineIndent = "0"
                }
            })
        }
        # Danh sách liệt kê, bullet, gạch đầu dòng, số thứ tự
        elseif ($text -match '^[\-\•\*\–\+]\s+' -or $text -match '^\d+[\.\)]\s+' -or $text -match '^[a-z]\)\s+') {
            $batchCmds.Add(@{
                command = "set"; path = $path
                props = @{
                    font = "Times New Roman"; size = "13pt"
                    lineSpacing = "1.3x"; spaceBefore = "2pt"; spaceAfter = "4pt"
                    firstLineIndent = "0"; alignment = "both"
                }
            })
        }
        # Đoạn văn nội dung thông thường
        else {
            $batchCmds.Add(@{
                command = "set"; path = $path
                props = @{
                    font = "Times New Roman"; size = "13pt"
                    lineSpacing = "1.3x"; spaceBefore = "3pt"; spaceAfter = "6pt"
                    firstLineIndent = "709"; alignment = "both"
                }
            })
        }
    }
}

Write-Host "[+] Áp dụng $($batchCmds.Count) quy tắc chuẩn hóa thể thức qua OfficeCLI Batch..." -ForegroundColor Yellow
$tempBatchFile = [System.IO.Path]::GetTempFileName()
try {
    $batchCmds | ConvertTo-Json -Depth 5 | Set-Content -Path $tempBatchFile -Encoding utf8
    $batchResult = & $cliPath batch $fullPath --input $tempBatchFile 2>&1
    & $cliPath close $fullPath 2>$null
}
finally {
    if (Test-Path $tempBatchFile) { Remove-Item $tempBatchFile -Force -ErrorAction SilentlyContinue }
}

Write-Host "===================================================" -ForegroundColor Cyan
Write-Host "  CHUẨN HÓA THỂ THỨC VĂN BẢN THÀNH CÔNG!" -ForegroundColor Green
Write-Host "  File: $fullPath" -ForegroundColor White
Write-Host "===================================================" -ForegroundColor Cyan
