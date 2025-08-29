下面給你一個 完整的 PowerShell 腳本（CopyReports.ps1），會依你描述的需求做：
	•	接受參數 -InputDate（格式 YYYYMMDD，例：20250731）
	•	自動計算民國年月 ROCYM（例：11407）與民國年月日 ROCYMDate（例：1140731）
	•	讀取 reports.csv（相對路徑），解析每列的「原始檔案路徑」與「目標檔案路徑」
	•	支援目標路徑多個（以 ; 分隔）
	•	處理原始檔名兩種情況：直接包含 InputDate（西元）或包含民國格式（如 1140731）；若原始檔名有亂碼前綴但中間有可識別關鍵字（例：cm2610、CM2810），會用 wildcard (*cm2610.xls) 去尋找匹配檔案
	•	自動建立 外幣報表產製_final\SAVE_PDF\{ROCYM}，以及指定的 rptFolders 子資料夾
	•	支援 -BaseDir（可指定 CSV 路徑與相對來源/目標的基底資料夾，預設為目前工作目錄）與 -CsvPath（預設 .\reports.csv）與 -DryRun（模擬執行，不實際複製）
	•	詳細 log（哪些檔案被找到了、複製到哪裡、找不到的紀錄）

使用方式（範例）：

# 乾跑（模擬）
.\CopyReports.ps1 -InputDate 20250731 -CsvPath .\reports.csv -BaseDir "D:\MyProject" -DryRun
# 真正執行
.\CopyReports.ps1 -InputDate 20250731 -CsvPath .\reports.csv -BaseDir "D:\MyProject"



⸻

腳本（儲存為 CopyReports.ps1）

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
        if ($row.PSObject.Properties.Match($n)) { return $row.$n }
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
        $matches = Get-ChildItem -Path $searchDir -Filter $searchLeaf -File -ErrorAction SilentlyContinue
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


⸻

程式說明（逐步、逐段）

我把程式分成幾個主要動作並解釋每段在做什麼，方便你了解內部運作與今後維護。
	1.	參數與輸入驗證
	•	param(...) 讀入：-InputDate（必填）、-CsvPath、-BaseDir、-DryRun。
	•	檢查 InputDate 是否為 8 位數，並解析成 DateTime。
	•	計算民國年 rocYear = Year - 1911，再組成 ROCYM（年+月）與 ROCYMDate（年+月+日）。
	2.	建立 SAVE_PDF 與 rptFolders
	•	腳本會在 BaseDir\外幣報表產製_final\SAVE_PDF\{ROCYM} 建資料夾，並在其下建立 $rptFolders 列表中指定的多個資料夾（若已存在會跳過）。
	•	當 -DryRun 時不實際建立，只會列出會建立的項目。
	3.	讀 CSV
	•	使用 Import-Csv 讀 reports.csv。
	•	程式會用 Get-FieldValue 這個 helper 去容錯比對欄位名稱（支援中文與英文常見欄位名）。若 CSV header 用中文（你範例是中文），會自動抓到正確欄位。
	4.	處理每一列（每一個報表條目）
	•	先用 Expand-Placeholders 把 {InputDate}, {ROCYM}, {ROCYMDate} 取代到原始路徑與目標路徑範本字串。
	•	產生一組要嘗試的 source pattern：
	•	直接的替換後路徑（最常見）。
	•	若替換後仍包含西元 InputDate，再嘗試把那段換成 ROCYMDate（處理原始檔名以民國年為檔名的情形）。
	•	若檔名包含 cm\d+（如 cm2610 或 CM2810），則加入 *cm2610.xls 的 wildcard pattern（處理檔名前有亂碼長數字情況）。
	•	若檔名有大型數字前綴（^\d{5,}[-_]），也把前面換成 * 做 wildcard。
	•	對每個 pattern：
	•	將相對路徑轉絕對（以 BaseDir 為根），分出目錄與檔名 filter。
	•	在那目錄用 Get-ChildItem -Filter 找符合檔案（支援 wildcard）。
	•	若找到一或多個檔案，逐一把該來源檔複製到對應的 target（TargetPath 欄位以 ; 分隔多個目標），會先建立目標資料夾（若不存在）。
	•	若 DryRun，僅列出要建立/要複製的項目，不做實際動作。
	5.	錯誤/未找到處理
	•	若某列在所有 pattern 下都找不到任何符合來源檔，會把該列資訊記下，最後輸出未找到清單，方便你去人工確認或修正 CSV。

⸻

注意事項與建議（實務經驗）
	1.	BaseDir：我把所有相對路徑視為以 BaseDir（預設當前執行目錄）為根。你可以把 BaseDir 指定為你專案的根目錄（例如 D:\DavidHsu\ReportsRoot），這樣 CSV 中的 外幣月報\{ROCYM}\... 就會被正確解析成 D:\DavidHsu\ReportsRoot\外幣月報\11407\...。
	2.	CSV 編碼：請確保 reports.csv 是 UTF-8 編碼（若含中文）。若遇到亂碼，可改用 -Encoding Default 或先另存成 UTF-8。
	3.	Wildcard 規則：我加入了針對 cm\d+ 與長數字前綴的 wildcard 處理。若還有其他命名慣例（例如 xxx-<code>.pdf），可以再擴充對應的 regex 處理邏輯。
	4.	多個來源對應多個目標：在 CSV 中，若目標路徑以 ; 分隔多個目標，腳本會分別把來源檔複製到每個目標。
	5.	覆蓋行為：Copy-Item -Force 會覆蓋既有檔案。如需改成不覆蓋，把 -Force 移除並自行加判斷。
	6.	日誌/紀錄：目前輸出以 Write-Host 為主。如果你要把 log 存到檔案，建議改為把訊息 Append 到一個 log 檔（Out-File -Append 或 Start-Transcript）。
	7.	測試：先用 -DryRun 跑完整的 CSV，確認所有 pattern 與目標位置都正確，再移除 -DryRun 做正式複製。

⸻

如果你要我：
	•	把這支腳本調整成 回傳一個 CSV 報表（列出每一列的來源、找到的實際檔案、複製到的目標、錯誤訊息），我可以幫你加上輸出檔案功能；
	•	或者你要我把 rptFolders 改成讀自一個設定檔（例如 folders.txt），也可以改；
	•	或者你的 CSV 欄位名稱不一樣（或 CSV 存放路徑在遠端），把實際範例給我（CSV 片段）我可以直接微調腳本以完全匹配你的檔案結構。

要我直接產生一個已修改好能在你機器上執行的範例（包含測試 sample reports.csv），或需要我把 log 改成輸出到 run_log.txt 嗎？要哪種我直接幫你改好。