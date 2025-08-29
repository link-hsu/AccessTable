<#
.SYNOPSIS
  依 reports.csv 自動複製檔案到多個目標位置，支援 {InputDate}, {ROCYM}, {ROCYMDate} 之替換與 wildcard 搜尋

.PARAMETER InputDate
  必要。西元年月日，格式 YYYYMMDD（例: 20250731）

.PARAMETER CsvPath
  CSV 檔案路徑（預設 .\reports.csv）。CSV 須含三欄（標題隨意），分別為：報表名稱, 原始檔案路徑, 目標檔案路徑

.PARAMETER BaseDir
  來源與目標相對路徑的基底資料夾（預設為目前工作目錄）

.PARAMETER DryRun
  若指定，僅列出會做的動作，但不實際複製檔案
#>

param(
    [Parameter(Mandatory=$true)][string]$InputDate,
    [string]$CsvPath = ".\reports_v1.csv",
    [string]$BaseDir = (Get-Location).Path,
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'

function Write-Log { param($msg) Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $msg" }

# 檢查 InputDate 格式
if ($InputDate -notmatch '^\d{8}$') {
    throw "InputDate 必須為 YYYYMMDD（8 位數）。輸入: $InputDate"
}
try {
    $dt = [datetime]::ParseExact($InputDate, 'yyyyMMdd', $null)
} catch {
    throw "無法解析 InputDate 為日期: $InputDate"
}

# 計算民國年 (Year - 1911)、ROCYM、ROCYMDate
[int]$rocYear = $dt.Year - 1911
$ROCYM = "{0}{1}" -f $rocYear.ToString(), $dt.ToString('MM')        # e.g. 11407
$ROCYMDate = "{0}{1}" -f $rocYear.ToString(), $dt.ToString('MMdd') # e.g. 1140731

Write-Log "InputDate = $InputDate ; ROCYM = $ROCYM ; ROCYMDate = $ROCYMDate"
Write-Log "BaseDir = $BaseDir ; CsvPath = $CsvPath ; DryRun = $DryRun"

# 需要事先建立的 rpt folders（會在 SAVE_PDF\{ROCYM} 下建立）
$rptFolders = @("AI821","CNY1","FB1","FB2","FB3","FB3A","FB5","FB5A","FM5","表2","FM13","FM11","表41","FM2","FM10","F1_F2","AI240")

# Helper: 找到 CSV 欄位的 key 名（因為中文、英文 header 都可能）
function Get-FieldValue($row, [string[]] $possibleNames) {
    foreach ($n in $possibleNames) {
        # if ($row.PSObject.Properties.Match($n)) { return $row.$n }
        if ($row.PSObject.Properties.Match($n).Count -gt 0) {
            return $row.$n
        }
    }
    # fallback: 取第二或第三欄 (index 1/2)
    $props = $row.PSObject.Properties | Select-Object -ExpandProperty Name
    if ($props.Count -ge 3) {
        return $row.${props[1]} , $row.${props[2]} | Select-Object -First 1
    } elseif ($props.Count -ge 2) {
        return $row.${props[1]}
    } else {
        return $null
    }
}

if (-not (Test-Path $CsvPath)) { throw "找不到 CSV 檔案: $CsvPath" }
# 嘗試以 utf8 讀入
$csv = Import-Csv -Path $CsvPath -Encoding UTF8

# 建立 SAVE_PDF\{ROCYM} 以及 rptFolders 子資料夾（相對於 BaseDir\外幣報表產製_final\SAVE_PDF\{ROCYM}）
$savePdfBase = Join-Path -Path $BaseDir -ChildPath "外幣報表產製_final\SAVE_PDF\$ROCYM"
if (-not $DryRun) {
    if (-not (Test-Path $savePdfBase)) {
        New-Item -Path $savePdfBase -ItemType Directory -Force | Out-Null
        Write-Log "建立資料夾: $savePdfBase"
    }
    foreach ($f in $rptFolders) {
        $p = Join-Path -Path $savePdfBase -ChildPath $f
        if (-not (Test-Path $p)) {
            New-Item -Path $p -ItemType Directory -Force | Out-Null
            Write-Log "建立子資料夾: $p"
        }
    }
} else {
    Write-Log "DryRun: 不會建立 SAVE_PDF 資料夾，僅模擬"
}

# 逐列處理 CSV
$rowIndex = 0
$notFoundList = @()
foreach ($row in $csv) {
    $rowIndex++
    # 嘗試常見欄位名稱
    $srcRaw = Get-FieldValue -row $row -possibleNames @('原始檔案路徑','原始','SourcePath','Source','原始檔案')
    $tgtRaw = Get-FieldValue -row $row -possibleNames @('目標檔案路徑','目標','TargetPath','Target','目標檔案')
    $reportName = Get-FieldValue -row $row -possibleNames @('報表名稱','ReportName','Name','報表')

    if (-not $srcRaw) {
        Write-Log "第 $rowIndex 列沒有辨識到原始檔案路徑，跳過。"
        continue
    }
    if (-not $tgtRaw) {
        Write-Log "第 $rowIndex 列沒有辨識到目標檔案路徑，跳過。"
        continue
    }

    # 先把常見的 placeholder 做替換（用 InputDate / ROCYM / ROCYMDate）
    $placeholders = @{
        '{InputDate}' = $InputDate
        '{ROCYM}'     = $ROCYM
        '{ROCYMDate}' = $ROCYMDate
    }

    # function: replace placeholders in a string
    function Expand-Placeholders($s) {
        $out = $s
        foreach ($k in $placeholders.Keys) { $out = $out -replace [regex]::Escape($k), $placeholders[$k] }
        return $out
    }

    $srcExpanded = Expand-Placeholders($srcRaw.Trim())
    $tgtExpandedRaw = Expand-Placeholders($tgtRaw.Trim())

    Write-Log "[$rowIndex] 處理: '$reportName'"
    Write-Log "  原始路徑範本: $srcRaw"
    Write-Log "  取代後(預設): $srcExpanded"

    # 產生要嘗試的 source pattern 列表（含可能的 ROC variant 與 wildcard）
    $patterns = New-Object System.Collections.Generic.List[string]

    # 1) 直接用替換後的當作一個可能性
    $patterns.Add($srcExpanded)

    # 2) 若替換後檔名仍包含西元 InputDate（$InputDate），再嘗試把它換成 ROCYMDate（處理原始檔名為民國年月日情境）
    if ($srcExpanded -match [regex]::Escape($InputDate)) {
        $alt = $srcExpanded -replace [regex]::Escape($InputDate), $ROCYMDate
        if (-not $patterns.Contains($alt)) { $patterns.Add($alt) }
        Write-Log "  加入 ROC 版本嘗試: $alt"
    }

    # 3) 嘗試把原始檔名中 numeric-prefix + '-' 或長數字前綴的情況轉成 wildcard（例如 1754300085627-cm2610.xls => *cm2610.xls）
    $leaf = [System.IO.Path]::GetFileName($srcExpanded)
    if ($leaf) {
        # 找 'cm' 或 'CM' 後接數字的 pattern（不區分大小寫）
        $m = [regex]::Match($leaf, '(?i)(cm\d{3,})')
        if ($m.Success) {
            $cmToken = $m.Groups[1].Value
            $wildLeaf = '*' + $cmToken + [System.IO.Path]::GetExtension($leaf)
            $parent = [System.IO.Path]::GetDirectoryName($srcExpanded)
            if (-not $parent) { $parent = $BaseDir }
            $wildFull = Join-Path -Path $parent -ChildPath $wildLeaf
            if (-not $patterns.Contains($wildFull)) {
                $patterns.Add($wildFull)
                Write-Log "  加入 wildcard 嘗試: $wildFull (based on cm pattern)"
            }
        }

        # 若檔名有長數字開頭再接 '-' 或 '_'，也把前面換成 *（例如 17543-xxx => *-xxx）
        if ($leaf -match '^\d{5,}[-_].+') {
            $wildLeaf2 = $leaf -replace '^\d{5,}', '*'
            $parent = [System.IO.Path]::GetDirectoryName($srcExpanded)
            if (-not $parent) { $parent = $BaseDir }
            $wildFull2 = Join-Path -Path $parent -ChildPath $wildLeaf2
            if (-not $patterns.Contains($wildFull2)) {
                $patterns.Add($wildFull2)
                Write-Log "  加入 wildcard 嘗試: $wildFull2 (based on numeric prefix)"
            }
        }
    }

    # 4) 若仍找不到，亦可嘗試在該目錄下做更寬鬆的搜尋 (同名開頭 + 任意尾)
    $parentDir = [System.IO.Path]::GetDirectoryName($srcExpanded)
    if (-not $parentDir) { $parentDir = $BaseDir }

    # 把 patterns 去重
    $patterns = $patterns | Select-Object -Unique

    $foundAnyForRow = $false
    foreach ($pat in $patterns) {
        # 把相對路徑轉成絕對（以 BaseDir 為根）
        $patNormalized = $pat
        if (-not [System.IO.Path]::IsPathRooted($patNormalized)) {
            $patNormalized = Join-Path -Path $BaseDir -ChildPath $patNormalized
        }

        # 分離資料夾與檔案 filter
        $searchDir = [System.IO.Path]::GetDirectoryName($patNormalized)
        if (-not $searchDir) { $searchDir = $BaseDir }
        $searchLeaf = [System.IO.Path]::GetFileName($patNormalized)
        if (-not $searchLeaf) { $searchLeaf = '*' }

        if (-not (Test-Path $searchDir)) {
            Write-Log "    搜尋目錄不存在: $searchDir （pattern: $patNormalized），跳過此 pattern。"
            continue
        }




        # 使用 Get-ChildItem 找到符合檔案（支援 wildcard）
        # $matches = Get-ChildItem -Path $searchDir -Filter $searchLeaf -File -ErrorAction SilentlyContinue

        Write-Log " 追蹤: searchDir='$searchDir'  searchLeaf='$searchLeaf'  patNormalized='$patNormalized'"

        $matches = @( Get-ChildItem -Path $searchDir -Filter $searchLeaf -File -ErrorAction SilentlyContinue )

        if ($null -eq $matches) {
            Write-Log " 追蹤: Get-ChildItem 回傳 $null"
        } else {
            # safe: array guaranteed
            Write-Log " 追蹤: matches 類型 = $($matches.GetType().FullName)  Count = $($matches.Count)"
        }

        if ($matches.Count -gt 0) {
            foreach ($f in $matches) {
                $foundAnyForRow = $true
                $sourceFile = $f.FullName
                Write-Log "    找到來源檔案: $sourceFile (pattern used: $patNormalized)"

                # 目標路徑支援多個（以 ; 分隔）
                $targets = $tgtExpandedRaw -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
                foreach ($t in $targets) {
                    $tReplaced = Expand-Placeholders($t)

                    # 合成絕對目標路徑（以 BaseDir 為根）
                    $targetFull = $tReplaced
                    if (-not [System.IO.Path]::IsPathRooted($targetFull)) {
                        $targetFull = Join-Path -Path $BaseDir -ChildPath $targetFull
                    }

                    # 如果目標路徑看起來是資料夾（沒有檔名），把檔名加回去（使用來源檔名）
                    $targetParent = [System.IO.Path]::GetDirectoryName($targetFull)
                    $targetLeaf = [System.IO.Path]::GetFileName($targetFull)
                    if (-not $targetLeaf) {
                        $targetFull = Join-Path -Path $targetFull -ChildPath $f.Name
                        $targetParent = [System.IO.Path]::GetDirectoryName($targetFull)
                    }

                    # 建資料夾（若需要）
                    if (-not (Test-Path $targetParent)) {
                        if (-not $DryRun) {
                            New-Item -Path $targetParent -ItemType Directory -Force | Out-Null
                            Write-Log "      建立目標資料夾: $targetParent"
                        } else {
                            Write-Log "      DryRun: 會建立目標資料夾: $targetParent"
                        }
                    }

                    # Copy
                    if (-not $DryRun) {
                        try {
                            Copy-Item -Path $sourceFile -Destination $targetFull -Force
                            Write-Log "      已複製: $sourceFile  -> $targetFull"
                        } catch {
                            Write-Log "      複製失敗: $sourceFile -> $targetFull ; 錯誤: $_"
                        }
                    } else {
                        Write-Log "      DryRun: 會複製: $sourceFile  -> $targetFull"
                    }
                } # end targets loop
            } # end matches foreach
        } else {
            Write-Log "    pattern 無符合檔案: $patNormalized"
        }
    } # end patterns foreach

    if (-not $foundAnyForRow) {
        $notFoundList += ,@{ Row=$rowIndex; Report=$reportName; SourcePattern=$srcExpanded }
        Write-Log "  尚未找到任何符合的來源檔案（第 $rowIndex 行）。"
    }
} # end csv foreach

if ($notFoundList.Count -gt 0) {
    Write-Log "=== 完成，但有 $($notFoundList.Count) 個項目找不到來源檔案 ==="
    foreach ($n in $notFoundList) {
        Write-Log "  Row $($n.Row)  Report: $($n.Report)  SourceTemplate: $($n.SourcePattern)"
    }
} else {
    Write-Log "=== 完成：所有列已處理（找不到的 0 筆） ==="
}