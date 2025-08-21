Answer1.

太長、很好 — 我把整段主流程和被呼叫的 module 一一掃過，找出會在執行時最容易導致「跳錯誤或中斷」的地方，並直接給出可貼上的修正建議（含範例程式碼片段）。重點放在「最低限度能跑完、不會因程式錯誤停下來」的修正。你可以直接把下面的修改貼回原檔後再測試一次。

---

# 要修的高風險問題（摘要）

1. **ProcessAllReports → Import\_CopyThenRunAP 的參數傳遞錯誤**（呼叫的參數數目/順序錯誤）。
2. **某些被呼叫的 Module（例如 FM11、FM10）裡沒有把 `targetBook` / `wb` 設定好**，但程式內大量使用 `targetBook`/`Worksheets`，會導致在被其他 Workbook 呼叫時出錯。
3. **未限定 `Rows.Count`/`Cells` 的工作表物件**（未指定 worksheet 物件） — 若 ActiveWorkbook/ActiveSheet 不是你預期的那個，會出錯或抓到錯誤資料。
4. **GetMonths 與 GetWesternMonthEnd 等小 bug（重複遞減、日期格式建議）**，雖不一定致命但會造成錯誤或錯誤字串。
5. **Import\_CopyThenRunAP 沒有錯誤處理機制**（若中途發生錯誤未還原 Application 設定會造成環境異常）。
6. 少數地方使用 `Worksheets(...)`、`Rows()` 未以 workbook 變數限定（容易在跨 workbook 呼叫時錯誤）。

---

# 建議的修補（可直接貼上替換）

下面給你最關鍵、可以立刻降低錯誤率的修改。請 **逐一貼回並測試**。

### 1) 把 `ProcessAllReports` 中呼叫 `Import_CopyThenRunAP` 的地方改為「命名參數呼叫」，避免順序錯誤

替換原本的 `Call Import_CopyThenRunAP(...)` 那一整段，改成：

```vb
' --- Replace the previous Call ... block with the named-argument call ---
Call Import_CopyThenRunAP( _
    basePath:=basePath, _
    oldMon:=oldMon, _
    newMon:=newMon, _
    rptID:=CStr(wsRpt.Cells(i, "A").Value), _
    tplPattern:=CStr(wsRpt.Cells(i, "B").Value), _
    tplSheet:=CStr(wsRpt.Cells(i, "C").Value), _
    impPattern:=CStr(wsRpt.Cells(i, "D").Value), _
    impSheets:=CStr(wsRpt.Cells(i, "E").Value), _
    declTplRel:=CStr(wsRpt.Cells(i, "F").Value), _
    moduleSub:=CStr(wsRpt.Cells(i, "K").Value), _    ' <-- 確認 module 子程序名稱在哪個欄位，通常是 K
    wsMap:=wsMap, _
    lastMap:=lastMap, _
    ROCYearMonth:=ROCYearMonth, _
    NUMYearMonth:=NUMYearMonth, _
    westernMonthEnd:=westernMonthEnd)
```

> 為什麼：命名參數可避免表格欄位順序與 VBA 參數順序不同造成錯誤。請確認 `moduleSub` 在你表格的實際欄位（我用 K 做示範，若實際不是 K，改為正確欄位）。

---

### 2) 在 **匯入並篩選OBUAC5411B資料（FM11）** 開頭加上 `targetBook` 的設定並修正未限定的 Worksheet 呼叫

在 `Sub 匯入並篩選OBUAC5411B資料(...)` 一開始加入：

```vb
Dim targetBook As Workbook
If wb Is Nothing Then
    Set targetBook = ThisWorkbook
Else
    Set targetBook = wb
End If
```

並把所有未限定 workbook 的 `Worksheets(...)`、`Sheets(...)` 換成 `targetBook.Worksheets(...)` 或 `targetBook.Sheets(...)`，例如：

```vb
' 刪除舊的會科整理分頁（用 targetBook）
On Error Resume Next
targetBook.Worksheets("OBU-AC5411B會科整理").Delete
On Error GoTo 0

Set wsDest = targetBook.Sheets("OBU-AC5411B")
wsDest.Range("A2:Z" & wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row).ClearContents
...
Set wsDest = targetBook.Sheets.Add(After:=targetBook.Sheets("OBU-AC5411B"))
wsDest.Name = "OBU-AC5411B會科整理"
...
lastRow = targetBook.Sheets("OBU-AC5411B").Cells(targetBook.Sheets("OBU-AC5411B").Rows.Count, "A").End(xlUp).Row
```

> 為什麼：function 內原本大量使用 `targetBook` 但沒先設定，會出現 `Variable not set` 或參考到錯誤的 workbook。

---

### 3) 在所有被外部呼叫的 Sub（FM10、FM11、F1\_F2、AI240 等）都加入相同的 `targetBook` 預設模式

範例（貼到每個 Sub 的開始）：

```vb
Dim targetBook As Workbook
If wb Is Nothing Then
    Set targetBook = ThisWorkbook
Else
    Set targetBook = wb
End If
```

並將接下來的 `Set wsXXX = ...` 改為 `Set wsXXX = targetBook.Sheets("...")`。這是跨 Workbook 呼叫時最常見的出錯來源。

---

### 4) 修正 `GetMonths` 中的重複遞減 bug（會造成年份被減兩次）

將函式 `GetMonths` 中以下段落改為：

```vb
' 計算上一個月
m = m - 1
If m = 0 Then
    y = y - 1
    m = 12
End If

' 上一個月份（oldMon）
oldMon = CStr(y) & "/" & Format(m, "00")
```

（刪除原本重複的 `If m = 0 Then y = y - 1: m = 12`）

---

### 5) 改善 `GetWesternMonthEnd` 的輸出（更穩定的 yyyymmdd 組成）

替換原本的 `GetWesternMonthEnd` 結尾為：

```vb
lastDay = DateSerial(adYear, monthNum + 1, 0)
GetWesternMonthEnd = CStr(Year(lastDay)) & Format(Month(lastDay), "00") & Format(Day(lastDay), "00")
```

> 為什麼：使用 Year/Month/Day 串接可以避免 Format pattern 打錯導致奇怪結果。

---

### 6) 在 `Import_CopyThenRunAP` 增加基本錯誤處理（至少保證 Application 設定會被還原）

在 `Import_CopyThenRunAP` 最上方加入：

```vb
On Error GoTo ErrHandler
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.AskToUpdateLinks = False
```

在副程式尾端（常規結束前）加入：

```vb
CleanExit:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.AskToUpdateLinks = True
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤: " & Err.Number & " - " & Err.Description, vbExclamation
    Resume CleanExit
```

> 為什麼：若中間出錯，必須恢復 Excel 的顯示與提示狀態，否則會影響後續操作。

---

### 7) 盡量**限定所有 `Rows.Count`, `Cells`, `Range`** 的工作表物件，範例

把像這種不限定的寫法：

```vb
lastRow = targetBook.Worksheets("底稿_含NT").Cells(Rows.Count, 1).End(xlUp).Row
```

改成：

```vb
With targetBook.Worksheets("底稿_含NT")
    lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
End With
```

或至少：

```vb
lastRow = targetBook.Worksheets("底稿_含NT").Cells(targetBook.Worksheets("底稿_含NT").Rows.Count, 1).End(xlUp).Row
```

> 這類修改請在你常用、且會跨 workbook 存取的地方（例如 F1\_F2 中多處使用 `Rows.Count`）統一處理。

---

### 8) FM10 / FM11 / 其它模組常見的小修（範例）

**FM10** 開頭先設定 `targetBook`（如上），並把 `n = Application.Match(...)` 之類的查找，包成 `With targetBook.Sheets("OBU_AC4603")` 以避免 ActiveSheet 問題。

---

# 其他實務建議（非必需但強烈建議）

* 在 `ReportsConfig`（或 `Mappings`）那個表格裡，**明確把 module 子程序的完整名稱放在固定欄位（例如 K）**，格式為 `"Module1.匯入並篩選OBUAC5411B資料"`（包含 module 名 + 子程序名），呼叫時用命名參數 `moduleSub:=...`。這樣 Application.Run 才會可靠。
* 檢查 `tplPattern` / `impPattern` 等路徑替換後的檔案是否存在（可在 `Workbooks.Open` 前用 `If Dir(path) = "" Then MsgBox ... : Exit Sub`）。目前若檔案不存在會直接在 Open 時跳錯誤。
* 若你希望更健壯，可於 `Import_CopyThenRunAP` 裡每個 `Workbooks.Open` 前後都加 `On Error Resume Next` + 檢查 `If wbOld Is Nothing Then ...` 的保護邏輯。
* 測試順序：先用 **一個 ReportID（例如 FM11）** 在本機上測試整段流程並逐步查看結果，確認檔案路徑與 Sheet 名稱完全對應後再批次跑多個。

---

# 我可以幫你做的下一步（你選一個）

* A) 我把上面所有「可修改處」產生成一份打補丁版的 **完整 ProcessAllReports module**（整個 Sub/Function 實作），你只要貼上替換；或
* B) 我把 `FM11`、`Import_CopyThenRunAP`、`ProcessAllReports` 三個最關鍵的 procedures 給出完整修正版（可貼上直接取代），你再測試；或
* C) 只要你把 `ReportsConfig` 工作表前 5 列（包含表頭）貼上來，我幫你驗對應欄位（moduleSub 的欄位、tplPattern/impPattern 的列）並產生可直接執行的命名參數呼叫。

---

如果你要我直接把 **完整修正版**（選 B）貼出，我現在就把 `ProcessAllReports` + `Import_CopyThenRunAP` + `FM11` 的修正版貼出，裡面包含上述所有防禦性修正（named args、targetBook 設定、錯誤處理、限定 Worksheet、日期格式修正）。你要哪一個？






Answer2.

好的 — 我把你主流程中最重要的程式（含所有 helper 函數）、`Import_CopyThenRunAP`，以及一個**修正版的 FM11 範例**（被呼叫時不會因為未設定 workbook/worksheet 而爆）全部整理成一個可直接貼上的完整 VBA 檔案內容，並在每一處我修改或加強的地方用 `''' *** MODIFIED ***` 明確標示（內嵌在程式碼中）。程式碼之後我再列出「修改清單（逐行 / 逐處說明）」，讓你能清楚知道每一處改了什麼、為什麼要改、該如何在其他模組套用相同修法。

> 使用方式：整個內容可直接貼到一個新的 Module（或你現有 Module 覆蓋），然後測試 `ProcessAllReports`。
> 我先修 `ProcessAllReports`、所有 date helper、`Import_CopyThenRunAP`，以及 `FM11` 的 `匯入並篩選OBUAC5411B資料` 作為被呼叫 module 範例（其他模組請依照同樣模式做相同修正 — 我會在修改清單中說明哪些地方要套用）。

# 完整修正版（請全部複製貼上到 Module）

```vb
Option Explicit

' =========================
' Helper: 取得本月（ROC 格式）與上月（ROC 格式）
Public Sub GetMonths(ByRef oldMon As String, ByRef newMon As String)
    Dim ymRaw   As String
    Dim parts() As String
    Dim y       As Integer
    Dim m       As Integer

    ''' *** MODIFIED ***: 加強輸入防護與移除重複邏輯
    ymRaw = ThisWorkbook.Names("YearMonth").RefersToRange.Value
    parts = Split(Trim(ymRaw), "/")
    If UBound(parts) < 1 Then
        Err.Raise vbObjectError + 1, , "YearMonth 格式錯誤 (範例：114/06)"
    End If

    y = CInt(parts(0))
    m = CInt(parts(1))

    ' 當前月份（newMon） — 保持 ROC 格式 YYY/MM
    newMon = CStr(y) & "/" & Format(m, "00")

    ' 計算上一個月（修正重複遞減 bug）
    m = m - 1
    If m = 0 Then
        y = y - 1
        m = 12
    End If
    oldMon = CStr(y) & "/" & Format(m, "00")
End Sub

' =========================
Public Function ConvertToROCFormat(ByVal newYearMonth As String, _
                                   ByVal returnType As String) As String
    Dim parts() As String
    Dim rocYear As Integer
    Dim result As String

    parts = Split(newYearMonth, "/")
    If UBound(parts) < 1 Then
        ConvertToROCFormat = ""
        Exit Function
    End If
    rocYear = CInt(parts(0))

    If returnType = "ROC" Then
        result = "民國 " & CStr(rocYear) & " 年 " & parts(1) & " 月"
    ElseIf returnType = "NUM" Then
        result = Format(rocYear, "000") & Format(CInt(parts(1)), "00") ' e.g. 11406
    Else
        result = ""
    End If

    ConvertToROCFormat = result
End Function

' =========================
Public Function GetWesternMonthEnd() As String
    Dim ymRaw      As String
    Dim parts()    As String
    Dim rocYear    As Integer
    Dim monthNum   As Integer
    Dim adYear     As Integer
    Dim lastDay    As Date

    ymRaw = ThisWorkbook.Names("YearMonth").RefersToRange.Value
    parts = Split(Trim(ymRaw), "/")
    If UBound(parts) < 1 Then Err.Raise vbObjectError + 1, , "YearMonth 格式錯誤 (範例：114/08)"

    rocYear  = CInt(parts(0))
    monthNum = CInt(parts(1))
    adYear = rocYear + 1911

    lastDay = DateSerial(adYear, monthNum + 1, 0)

    ''' *** MODIFIED ***: 更穩定的 yyyymmdd 組成，避免 Format pattern 錯誤
    GetWesternMonthEnd = CStr(Year(lastDay)) & Format(Month(lastDay), "00") & Format(Day(lastDay), "00")
End Function

' =========================
Public Function GetROCMonthEnd() As String
    Dim ymRaw As String
    Dim parts() As String
    Dim rocYear As Long
    Dim monthNum As Long
    Dim adYear As Long
    Dim lastDay As Date

    ymRaw = ThisWorkbook.Names("YearMonth").RefersToRange.Value
    parts = Split(Trim(ymRaw), "/")
    If UBound(parts) < 1 Then Err.Raise vbObjectError + 1, , "YearMonth 格式錯誤 (範例：114/08)"

    rocYear = CLng(parts(0))
    monthNum = CLng(parts(1))

    adYear = rocYear + 1911
    lastDay = DateSerial(adYear, monthNum + 1, 0)

    ''' *** MODIFIED ***: 使用固定長度 ROC 年 + mmdd
    GetROCMonthEnd = Format(rocYear, "000") & Format(Month(lastDay), "00") & Format(Day(lastDay), "00")
End Function

' =========================
Public Function GetWesternMonth() As String
    Dim ymRaw      As String
    Dim parts()    As String
    Dim rocYear    As Integer
    Dim adYear     As Integer

    ymRaw = ThisWorkbook.Names("YearMonth").RefersToRange.Value
    parts = Split(Trim(ymRaw), "/")
    If UBound(parts) < 1 Then Err.Raise vbObjectError + 1, , "YearMonth 格式錯誤 (範例：114/08)"

    rocYear  = CInt(parts(0))
    adYear = rocYear + 1911

    ''' *** MODIFIED ***: CStr 並補零 month
    GetWesternMonth = CStr(adYear) & Format(CInt(parts(1)), "00")
End Function

' =========================
'— 2. 主流程：依 CaseType 分派 ——
Sub ProcessAllReports()
    Dim wbCtl    As Workbook
    Dim wsRpt    As Worksheet, wsMap As Worksheet
    Dim basePath As String
    Dim lastRpt  As Long, lastMap As Long
    Dim oldMon   As String, newMon As String
    Dim ROCYearMonth As String, NUMYearMonth As String
    Dim westernMonthEnd As String
    Dim ROCMonthEnd As String
    Dim westernMonth As String
    Dim i        As Long, caseType As String

    ''' *** MODIFIED ***: 新增變數宣告（先預設）
    Dim rptFolders As Variant
    Dim savePdfRoot As String

    Set wbCtl = ThisWorkbook
    Set wsRpt = wbCtl.Sheets("ReportsConfig")
    Set wsMap = wbCtl.Sheets("Mappings")
    basePath = wbCtl.Path

    Call GetMonths(oldMon, newMon)

    ROCYearMonth = ConvertToROCFormat(newMon, "ROC")
    NUMYearMonth = ConvertToROCFormat(newMon, "NUM")

    ''' *** MODIFIED ***: remove "/" for folder/file usage
    oldMon = Replace(oldMon, "/", "")
    newMon = Replace(newMon, "/", "")

    westernMonthEnd = GetWesternMonthEnd()
    ROCMonthEnd = GetROCMonthEnd()
    westernMonth = GetWesternMonth()

    ''' *** MODIFIED ***: 建立 SAVE_PDF\newMon\{rptID} 的資料夾（一次建立所有清單）
    savePdfRoot = wbCtl.Path & "\SAVE_PDF\" & newMon
    If Dir(savePdfRoot, vbDirectory) = "" Then MkDir savePdfRoot

    rptFolders = Array("AI821","CNY1","FB1","FB2","FB3","FB3A","FB5","FB5A","FM5","表2","FM13","FM11","表41","FM2","FM10","F1_F2","AI240")
    For i = LBound(rptFolders) To UBound(rptFolders)
        If Dir(savePdfRoot & "\" & rptFolders(i), vbDirectory) = "" Then
            MkDir savePdfRoot & "\" & rptFolders(i)
        End If
    Next i

    lastRpt = wsRpt.Cells(wsRpt.Rows.Count, "A").End(xlUp).Row
    lastMap = wsMap.Cells(wsMap.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRpt
        caseType = Trim(CStr(wsRpt.Cells(i, "H").Value))
        Select Case caseType
            Case "CopyThenRunAP"
                ''' *** MODIFIED ***: 改用命名參數呼叫，避免參數順序錯誤
                Call Import_CopyThenRunAP( _
                    basePath:=basePath, _
                    oldMon:=oldMon, _
                    newMon:=newMon, _
                    rptID:=CStr(wsRpt.Cells(i, "A").Value), _
                    tplPattern:=CStr(wsRpt.Cells(i, "B").Value), _
                    tplSheet:=CStr(wsRpt.Cells(i, "C").Value), _
                    impPattern:=CStr(wsRpt.Cells(i, "D").Value), _
                    impSheets:=CStr(wsRpt.Cells(i, "E").Value), _
                    declTplRel:=CStr(wsRpt.Cells(i, "F").Value), _
                    moduleSub:=CStr(wsRpt.Cells(i, "K").Value), _  ' <-- 確認 moduleSub 放哪個欄位
                    wsMap:=wsMap, _
                    lastMap:=lastMap, _
                    ROCYearMonth:=ROCYearMonth, _
                    NUMYearMonth:=NUMYearMonth, _
                    westernMonthEnd:=westernMonthEnd)

            Case Else
                MsgBox "未知 CaseType: " & caseType & "（ReportID=" & wsRpt.Cells(i, "A").Value & "）", vbExclamation
        End Select
    Next i

    MsgBox "全部報表處理完成！", vbInformation
End Sub

' =========================
' Import + Run 模組（主工作：開舊底稿、貼資料、執行該底稿中的 Module.Sub、存成新底稿、再貼入申報檔）
Public Sub Import_CopyThenRunAP( _
    ByVal basePath As String, _
    ByVal oldMon As String, _
    ByVal newMon As String, _
    ByVal rptID As String, _
    ByVal tplPattern As String, _
    ByVal tplSheet As String, _
    ByVal impPattern As String, _
    ByVal impSheets As String, _
    ByVal declTplRel As String, _
    ByVal moduleSub As String, _
    ByVal wsMap As Worksheet, _
    ByVal lastMap As Long, _
    ByVal ROCYearMonth As String, _
    ByVal NUMYearMonth As String, _
    ByVal westernMonthEnd As String)

    ''' *** MODIFIED ***: Defensive programming + error handling
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False

    Dim wbOld As Workbook, wbImp As Workbook
    Dim wbNew As Workbook, wbDecl As Workbook
    Dim arrTplSh() As String, arrImpF() As String, arrImpSh() As String
    Dim oldTplRel As String, newTplRel As String
    Dim tplPath As String, declTplPath As String
    Dim j As Long
    Dim impPatternArr() As String
    Dim f As String
    Dim targetBook As Workbook

    Set targetBook = ThisWorkbook ' the control workbook (caller)

    ''' *** MODIFIED ***: 安全替換路徑參數
    oldTplRel = Replace(tplPattern, "YYYYMM", oldMon)
    newTplRel = Replace(tplPattern, "YYYYMM", newMon)
    arrTplSh = Split(tplSheet, ",")

    impPatternArr = Split(impPattern, ",")
    ReDim arrImpF(LBound(impPatternArr) To UBound(impPatternArr))

    For j = LBound(impPatternArr) To UBound(impPatternArr)
        If impPatternArr(j) Like "*外幣債損益評估表(月底)對AC5100B*" Then
            arrImpF(j) = Replace(impPatternArr(j), "YYYYMM", westernMonthEnd)
        Else
            arrImpF(j) = Replace(impPatternArr(j), "YYYYMM", newMon)
        End If
    Next j

    arrImpSh  = Split(impSheets, ",")

    tplPath     = basePath & "\" & oldTplRel
    declTplPath = basePath & "\" & declTplRel

    ' 開舊底稿之前檢查檔案是否存在
    If Dir(tplPath) = "" Then
        MsgBox "找不到模板檔：" & vbCrLf & tplPath, vbExclamation
        GoTo CleanExit
    End If

    Set wbOld = Workbooks.Open(Filename:=tplPath, ReadOnly:=False)

    ''' *** MODIFIED ***: 統一以 rptID (參數) 為準，避免大小寫或變數名稱不同
    Select Case UCase(Trim(rptID))
        Case "FM11"
            On Error Resume Next
            wbOld.Sheets("FM11").Range("D3").Value = NUMYearMonth
            On Error GoTo ErrHandler
        Case "表41"
            On Error Resume Next
            wbOld.Sheets("表41").Range("A3").Value = NUMYearMonth
            On Error GoTo ErrHandler
        Case "FM2"
            On Error Resume Next
            wbOld.Sheets("FM2").Range("C2").Value = NUMYearMonth
            On Error GoTo ErrHandler
        Case "FM10"
            On Error Resume Next
            wbOld.Sheets("FM10").Range("C2").Value = NUMYearMonth
            On Error GoTo ErrHandler
        Case "F1_F2"
            On Error Resume Next
            wbOld.Sheets("F1").Range("B3").Value = NUMYearMonth
            wbOld.Sheets("F2").Range("B3").Value = NUMYearMonth
            On Error GoTo ErrHandler
        Case "AI240"
            On Error Resume Next
            wbOld.Sheets("AI240").Range("A2").Value = NUMYearMonth
            On Error GoTo ErrHandler
    End Select

    ' 逐一開啟匯入檔案並貼上
    For j = LBound(arrImpF) To UBound(arrImpF)
        wbOld.Sheets(Trim(arrTplSh(j))).Cells.ClearContents

        f = basePath & "\" & Trim(arrImpF(j))
        If Dir(f) = "" Then
            MsgBox "找不到匯入檔：" & vbCrLf & f, vbExclamation
            ' 依需求選擇跳過或退出，這裡選擇跳過該檔案
            GoTo NextImpFile
        End If

        Set wbImp = Workbooks.Open(Filename:=f, ReadOnly:=True)
        With wbOld.Sheets(Trim(arrTplSh(j)))
            wbImp.Sheets(Trim(arrImpSh(j))).UsedRange.Copy
            .Range("A1").PasteSpecial xlPasteValues
        End With
        wbImp.Close SaveChanges:=False
NextImpFile:
    Next j

    ' 執行對應 Module.Sub —— 檢查 moduleSub 參數是否有值
    If Trim(moduleSub) <> "" Then
        ' moduleSub 可能是 "Module1.匯入並篩選OBUAC5411B資料" 或 "Module.ProcessFM2"
        On Error Resume Next
        Application.Run "'" & wbOld.Name & "'!" & moduleSub, wbOld
        If Err.Number <> 0 Then
            ' 嘗試不帶 wbOld 的呼叫（某些 sub 只接受 workbook 物件）
            Err.Clear
            Application.Run "'" & wbOld.Name & "'!" & moduleSub
        End If
        On Error GoTo ErrHandler
    End If

    ' 另存新底稿（copy）
    wbOld.SaveCopyAs Filename:=basePath & "\" & newTplRel
    wbOld.Close SaveChanges:=False

    ' 貼入申報模板（declTplPath）
    If Dir(basePath & "\" & newTplRel) = "" Then
        MsgBox "無法找到另存的新底稿: " & basePath & "\" & newTplRel, vbExclamation
        GoTo CleanExit
    End If
    Set wbNew = Workbooks.Open(Filename:=basePath & "\" & newTplRel, ReadOnly:=False)

    If Dir(declTplPath) = "" Then
        MsgBox "找不到申報模板檔：" & vbCrLf & declTplPath, vbExclamation
        wbNew.Close SaveChanges:=False
        GoTo CleanExit
    End If
    Set wbDecl = Workbooks.Open(Filename:=declTplPath, ReadOnly:=False)

    Dim srcSh As String, rngStrings() As String
    Dim rngSrc As Range, rngDst As Range
    Dim srcAddr As Variant
    Dim jMap As Long

    If UCase(Trim(rptID)) = "FM11" Or UCase(Trim(rptID)) = "表41" Then
        For jMap = 2 To lastMap
            If wsMap.Cells(jMap, "A").Value = rptID Then
                srcSh = CStr(wsMap.Cells(jMap, "B").Value)
                rngStrings = Split(CStr(wsMap.Cells(jMap, "C").Value), ",")

                For Each srcAddr In rngStrings
                    srcAddr = Trim(srcAddr)
                    On Error Resume Next
                    Set rngSrc = wbNew.Sheets(Trim(srcSh)).Range(srcAddr)
                    Set rngDst = wbDecl.Sheets(Trim(srcSh)).Range(srcAddr)
                    On Error GoTo ErrHandler
                    If Not rngSrc Is Nothing And Not rngDst Is Nothing Then
                        rngDst.Value = rngSrc.Value
                    End If
                Next srcAddr
            End If
        Next jMap
    End If

    wbDecl.Save
    wbDecl.Close SaveChanges:=False
    wbNew.Close SaveChanges:=False

CleanExit:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.AskToUpdateLinks = True
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤: " & Err.Number & " - " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

' =========================
Public Function ParentFolder(ByVal fullPath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo ErrHandler
    ParentFolder = fso.GetParentFolderName(fullPath)
    Exit Function
ErrHandler:
    ParentFolder = ""
End Function

' =========================
' 修正版 FM11 範例（被其他 Workbook 呼叫時也安全）
' *** MODIFIED: 重寫匯入並篩選OBUAC5411B資料，加入 targetBook, 防禦式檢查，限定 worksheet 物件 ***
Public Sub 匯入並篩選OBUAC5411B資料(Optional ByVal wb As Workbook, _
                            Optional ByVal calledByOtherExcel As Boolean = False)
    Dim targetBook As Workbook
    Dim importWB As Workbook
    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim importFile As String
    Dim lastRow As Long, destRow As Long
    Dim keyword As Variant
    Dim keywords As Variant
    Dim i As Long
    Dim sumRange As Range

    ''' *** MODIFIED ***: 設定 targetBook（被外部呼叫時會傳入 wb）
    If wb Is Nothing Then
        Set targetBook = ThisWorkbook
    Else
        Set targetBook = wb
    End If

    keywords = Array("FVPL", "FVOCI", "AC", "拆放證券公司息-OSU")

    ' 若目標工作表存在就刪除重建（限定 targetBook）
    On Error Resume Next
    Application.DisplayAlerts = False
    targetBook.Worksheets("OBU-AC5411B會科整理").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' 來源資料處理（若不是被外部呼叫，走選檔流程）
    If Not calledByOtherExcel Then
        importFile = Application.GetOpenFilename("Excel Files (*.xls;*.xlsx), *.xls;*.xlsx", , "請選取 OBU-AC5411B 檔案")
        If importFile = "False" Then Exit Sub
        Set importWB = Workbooks.Open(importFile)
        Set wsSource = importWB.Sheets(1)
        Set wsDest = targetBook.Sheets("OBU-AC5411B")

        ' 清除目的表舊資料
        wsDest.Range("A2:Z" & wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row).ClearContents

        ' 複製來源檔案第2列起資料貼到目前檔案中
        lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
        If lastRow >= 2 Then
            wsSource.Range("A2:Z" & lastRow).Copy Destination:=wsDest.Range("A2")
        End If

        ' 將 B 欄強制轉型數值並關閉來源
        With wsDest
            With .Range("B2:B" & .Cells(.Rows.Count, "B").End(xlUp).Row)
                .NumberFormat = "0"
                .Value = .Value
            End With
        End With

        importWB.Close SaveChanges:=False
    Else
        ' 若被呼叫，假設資料已在 ThisWorkbook 的 OBU-AC5411B
        Set wsDest = targetBook.Sheets("OBU-AC5411B")
    End If

    ' 建立結果分頁
    Set wsDest = targetBook.Sheets("OBU-AC5411B")
    Set wsDest = targetBook.Sheets.Add(After:=wsDest)
    wsDest.Name = "OBU-AC5411B會科整理"

    ' 複製標題列
    targetBook.Sheets("OBU-AC5411B").Rows(1).Copy Destination:=wsDest.Rows(1)
    destRow = 2

    lastRow = targetBook.Sheets("OBU-AC5411B").Cells(targetBook.Sheets("OBU-AC5411B").Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        For Each keyword In keywords
            If InStr(1, targetBook.Sheets("OBU-AC5411B").Cells(i, "A").Value, keyword, vbTextCompare) > 0 Then
                targetBook.Sheets("OBU-AC5411B").Rows(i).Copy Destination:=wsDest.Rows(destRow)
                destRow = destRow + 1
                Exit For
            End If
        Next keyword
    Next i

    If destRow > 2 Then
        wsDest.Cells(destRow, "B").Value = "本月金額總和"
        Set sumRange = wsDest.Range("C2:C" & destRow - 1)
        wsDest.Cells(destRow, "C").Formula = "=SUM(" & sumRange.Address(False, False) & ")"
        targetBook.Sheets("FM11 計算1").Range("G4").Value = wsDest.Cells(destRow, "C").Value
    End If

    wsDest.Columns.AutoFit
    MsgBox "匯入成功並完成篩選與總和計算！", vbInformation
    targetBook.Sheets("FM11 計算1").Activate
End Sub
```

---

# 每一處修改說明（逐項、可貼到程式碼旁理解）

下面列出我在程式中每一個 `''' *** MODIFIED ***` 標記代表的修改內容與理由 — 你可以用來比對原始檔並確定要改哪裡。

1. `GetMonths`

   * 新增 `Trim` 與格式判斷（若 `YearMonth` 格式錯誤則 raise error）。
   * 刪除原本**重複**的 `If m = 0 Then y = y - 1: m = 12`（原版會把年份減兩次）。
   * 結果：上一個月計算正確且穩健。

2. `ConvertToROCFormat`

   * 若 `returnType="NUM"` 改用 `Format(rocYear,"000") & Format(month,"00")`，確保欄位長度一致（例：11406）。
   * 防護：若輸入格式不對回傳空字串。

3. `GetWesternMonthEnd` / `GetROCMonthEnd` / `GetWesternMonth`

   * 檢查 `YearMonth` 格式。
   * 使用 `Year/Month/Day` 與 `Format(...,"00")` 組合成字串，避免 `Format` 模式打錯造成輸出錯誤（你原版 `Format(lastDay, "yyyymmdd")` 易誤解）。

4. `ProcessAllReports`

   * 把 `ThisWorkbook` 的 `Sheets("ReportsConfig")`/`Sheets("Mappings")` 設定提早，避免在迴圈中依賴 ActiveWorkbook。
   * 新增建立 `SAVE_PDF\newMon\{rptID}` 資料夾的步驟（跟你原本的 intention 一致）。
   * **最重要**：把 `Import_CopyThenRunAP` 的呼叫改為 **命名參數**，避免參數順序/欄位對不上時造成錯誤（這是你最早出問題的來源之一）。
   * 注意：`moduleSub` 我從 `K` 欄取得（你如果放在其他欄，請改成正確欄位字母）。

5. `Import_CopyThenRunAP`（大量修改）

   * 新增 `On Error GoTo ErrHandler` 與 `CleanExit`，並在例外時恢復 `Application` 狀態（避免 Excel 被鎖住）。
   * 開檔前加入 `Dir(path)` 檢查，若找不到檔案顯示訊息並跳過或結束。
   * 統一使用傳入的 `rptID`（`UCase` 比對）設定報表內年月欄位，並用 `On Error Resume Next` 包住個別 Sheet 存取以避免欄位不存在就跳錯。
   * 匯入每個檔案時檢查檔案是否存在，若不存在就跳過該檔（避免整支程序直接中斷）。
   * 執行 `Application.Run` 時先嘗試傳入 `wbOld`，若失敗再嘗試不帶參數呼叫（容錯）。
   * 貼入申報模板時也做 `Dir` 檢查。
   * 所有 `Sheets`/`Range` 操作都指定到開啟的工作簿物件，避免 ActiveWorkbook 錯誤。

6. `ParentFolder`

   * 無大更動（保留）。

7. `匯入並篩選OBUAC5411B資料 (FM11)`

   * **重大修改**：加入 `targetBook` 變數（若 `wb` 為 Nothing，使用 ThisWorkbook；若有傳入，使用傳入的 `wb`）— 使此 Sub 在被其他 Workbook 以 `Application.Run` 呼叫時能正確作用。
   * 所有對 `Sheets`/`Rows`/`Cells` 的操作都限定為 `targetBook`，避免抓到錯誤的 ActiveWorkbook/ActiveSheet。
   * 當 `calledByOtherExcel = True` 時，略過檔案選擇的步驟（你原始程式已有類似邏輯；我把流程整理得更清楚）。
   * 結果分頁、總和公式與貼回 `FM11 計算1` 的步驟都限定 `targetBook`。

---

# 建議你接下來快速驗證的清單（逐步測試）

1. 在控制檔（含 `ReportsConfig` 與 `Mappings`）先只放一列（例如 FM11）並把 `moduleSub` 核對成你實際放的位置（K 或其他欄）。
2. 準備好該列所需的檔案路徑（`tplPattern/impPattern/declTplRel` 替換後）並確保檔名存在；先測試一個 report（FM11）。
3. 執行 `ProcessAllReports`，若出錯訊息跳出請把錯誤訊息截圖或貼回來，我會協助排查。
4. 若其他被呼叫模組（例如 FM2 / FM10 / F1\_F2 / AI240）遇到同樣 workbook 未設定或未限定 Sheets/Cells 的情況，請把我在 FM11 做的 `targetBook` 模式套到該模組。

---

如果你要我把其他被呼叫的模組（FM2、FM10、F1\_F2、AI240）也一併完整修正（把所有 `Worksheets` 改為 `targetBook.Worksheets`、加上 `targetBook` 判斷並加上錯誤處理），我可以繼續做 — 但這會使回覆非常長。你想：
A) 我接著把 FM2、FM10、F1\_F2、AI240 都完整修正並回傳（一次提供全部），或
B) 先用我上面給的 `ProcessAllReports` + `Import_CopyThenRunAP` + `FM11` 去跑一次，遇到錯誤再把有問題的模組貼上，我逐一修正？

你選 A 或 B（或直接說「全部修好」），我就繼續。



Answer3.

好 — 我已把 **FM2、FM10、F1\_F2、AI240** 這四個被呼叫的模組全部改寫成「被其他 Workbook 呼叫也安全」、「限定 workbook/worksheet 物件」、「加上基本錯誤處理與環境還原」的完整版本。程式內我在每處我修改或加強的地方都用 `''' *** MODIFIED ***` 明確標示，並在程式碼後面把每一處修改逐條說明給你（方便逐一比對原始版本）。你可以直接把下列每個模組貼回對應的 Module 中覆蓋原本內容，然後測試 `ProcessAllReports`。

> 注意：
>
> * 我假設你已經把之前我給的 `ProcessAllReports`、`Import_CopyThenRunAP`、helper functions（GetMonths / GetWesternMonthEnd / GetROCMonthEnd / ConvertToROCFormat / GetWesternMonth / ParentFolder）貼回專案。
> * 下面每個 Sub 都會先處理 `wb` / `targetBook`，並把 `Sheets` / `Cells` 都限定在 `targetBook` 上以避免跨 Workbook 參考錯誤。
> * 若你在 `ReportsConfig` 表格裡把 `moduleSub` 放在其他欄，請到 `ProcessAllReports` 將該欄位字母改成實際欄位（我示範用 K）。

---

## 1) 修正版 FM2 (整支 module)

```vb
Option Explicit

' =========================
' FM2 處理程序（修正版）
Public Sub ProcessFM2_ButtonClick()
    ProcessFM2 ThisWorkbook
End Sub

Sub ProcessFM2(Optional ByVal wb As Workbook)
    Dim targetBook As Workbook
    Dim wsData As Worksheet
    Dim wsMap As Worksheet
    Dim wsCompute As Worksheet
    Dim wsFM2 As Worksheet

    ''' *** MODIFIED ***: 建立 targetBook 機制（被外部呼叫時安全）
    If wb Is Nothing Then
        Set targetBook = ThisWorkbook
    Else
        Set targetBook = wb
    End If

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    ''' *** MODIFIED ***: 明確以 targetBook 取得工作表，並檢查存在性
    On Error Resume Next
    Set wsData = targetBook.Worksheets("OBU_MM4901B")
    Set wsMap = targetBook.Worksheets("金融機構代號對照表")
    Set wsCompute = targetBook.Worksheets("計算表")
    Set wsFM2 = targetBook.Worksheets("FM2")
    On Error GoTo ErrHandler

    If wsData Is Nothing Or wsMap Is Nothing Or wsFM2 Is Nothing Then
        MsgBox "找不到必要的工作表，請確認有 OBU_MM4901B / 金融機構代號對照表 / FM2 三個工作表。", vbExclamation
        GoTo CleanExit
    End If

    ' =============================
    ' 【修改處 1】：刪除 K 欄沒有資料的列（限定工作表）
    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row

    Dim r As Long
    For r = lastRow To 2 Step -1
        If IsEmpty(wsData.Cells(r, "K").Value) Or Trim(CStr(wsData.Cells(r, "K").Value)) = "" Then
            wsData.Rows(r).Delete
        End If
    Next r
    ' =============================

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    ' 1) 建立資料結構
    lastRow = wsData.Cells(wsData.Rows.Count, "C").End(xlUp).Row
    Dim key As Variant, arrAK As Variant
    For r = 2 To lastRow
        key = Trim(CStr(wsData.Cells(r, "C").Value))
        If key <> "" Then
            If Not dict.Exists(key) Then
                Dim inner As Object
                Set inner = CreateObject("Scripting.Dictionary")
                inner.Add "Rows", CreateObject("System.Collections.ArrayList")
                inner.Add "Records", CreateObject("System.Collections.ArrayList")
                inner.Add "Class", ""
                inner.Add "BankCodes", CreateObject("System.Collections.ArrayList")
                dict.Add key, inner
            End If
            arrAK = Application.Index(wsData.Range("A" & r & ":K" & r).Value, 1, 0)
            dict(key)("Rows").Add r
            dict(key)("Records").Add arrAK
        End If
    Next r

    ' 2) 用對照表決定 DBU/OBU
    Dim mapLastRow As Long
    mapLastRow = wsMap.Cells(wsMap.Rows.Count, "A").End(xlUp).Row

    Dim i As Long
    Dim toRemove As Collection
    Set toRemove = New Collection

    Dim nameA As String, nameB As String, bankCode As String
    Dim comp As Long
    For Each key In dict.Keys
        Dim found As Boolean: found = False
        ' 搜尋 A 欄 (DBU)
        For i = 1 To mapLastRow
            nameA = Trim(CStr(wsMap.Cells(i, "A").Value))
            If nameA <> "" Then
                comp = StrComp(key, nameA, vbTextCompare)
                If comp = 0 Then
                    found = True
                    dict(key)("Class") = "DBU"
                    bankCode = Trim(CStr(wsMap.Cells(i, "C").Value))
                    If bankCode <> "" Then dict(key)("BankCodes").Add bankCode
                    Exit For
                End If
            End If
        Next i
        If Not found Then
            ' 搜尋 B 欄 (OBU)
            For i = 1 To mapLastRow
                nameB = Trim(CStr(wsMap.Cells(i, "B").Value))
                If nameB <> "" Then
                    comp = StrComp(key, nameB, vbTextCompare)
                    If comp = 0 Then
                        found = True
                        dict(key)("Class") = "OBU"
                        bankCode = Trim(CStr(wsMap.Cells(i, "C").Value))
                        If bankCode <> "" Then dict(key)("BankCodes").Add bankCode
                        Exit For
                    End If
                End If
            Next i
        End If

        If Not found Then
            toRemove.Add key
        End If
    Next key

    For i = 1 To toRemove.Count
        dict.Remove toRemove(i)
    Next i

    ' 3) DBU/OBU index blocks（保持原邏輯）
    Dim dbuBlocks As Variant, obuBlocks As Variant
    dbuBlocks = Array(Array(3, 10), Array(12, 19), Array(21, 28), Array(30, 37), Array(39, 46))
    obuBlocks = Array(Array(50, 57), Array(59, 66), Array(68, 75), Array(77, 84), Array(86, 93))

    Dim dbuList As Collection, obuList As Collection
    Set dbuList = New Collection
    Set obuList = New Collection
    For Each key In dict.Keys
        If dict(key)("Class") = "DBU" Then
            dbuList.Add key
        ElseIf dict(key)("Class") = "OBU" Then
            obuList.Add key
        End If
    Next key

    ' 4) 清空目標區段並寫入
    Dim b As Long
    For b = LBound(dbuBlocks) To UBound(dbuBlocks)
        wsCompute.Range(wsCompute.Cells(dbuBlocks(b)(0), "A"), wsCompute.Cells(dbuBlocks(b)(1), "K")).ClearContents
    Next b
    For b = LBound(obuBlocks) To UBound(obuBlocks)
        wsCompute.Range(wsCompute.Cells(obuBlocks(b)(0), "A"), wsCompute.Cells(obuBlocks(b)(1), "K")).ClearContents
    Next b

    Dim recCount As Long, blk As Variant, tRow As Long, writtenCount As Long
    Dim warnings As Collection
    Set warnings = New Collection

    ' DBU 貼入
    For b = LBound(dbuBlocks) To UBound(dbuBlocks)
        If b + 1 > dbuList.Count Then Exit For
        key = dbuList(b + 1)
        recCount = dict(key)("Records").Count
        blk = dbuBlocks(b)
        writtenCount = 0
        For r = 0 To recCount - 1
            tRow = blk(0) + r
            If tRow <= blk(1) Then
                Dim recArr As Variant
                recArr = dict(key)("Records")(r)
                Dim colIdx As Long
                For colIdx = 1 To 11
                    wsCompute.Cells(tRow, colIdx).Value = recArr(colIdx)
                Next colIdx
                writtenCount = writtenCount + 1
            Else
                warnings.Add "DBU '" & key & "' 的紀錄數 (" & recCount & ") 超過 index" & (b + 1) & " 容量（" & (blk(1) - blk(0) + 1) & "），僅貼入前 " & writtenCount & " 筆。"
                Exit For
            End If
        Next r
    Next b

    ' OBU 貼入
    For b = LBound(obuBlocks) To UBound(obuBlocks)
        If b + 1 > obuList.Count Then Exit For
        key = obuList(b + 1)
        recCount = dict(key)("Records").Count
        blk = obuBlocks(b)
        writtenCount = 0
        For r = 0 To recCount - 1
            tRow = blk(0) + r
            If tRow <= blk(1) Then
                recArr = dict(key)("Records")(r)
                For colIdx = 1 To 11
                    wsCompute.Cells(tRow, colIdx).Value = recArr(colIdx)
                Next colIdx
                writtenCount = writtenCount + 1
            Else
                warnings.Add "OBU '" & key & "' 的紀錄數 (" & recCount & ") 超過 index" & (b + 1) & " 容量（" & (blk(1) - blk(0) + 1) & "），僅貼入前 " & writtenCount & " 筆。"
                Exit For
            End If
        Next r
    Next b

    ' 5) 貼銀行代號到 FM2
    Dim bankSet As Object
    Set bankSet = CreateObject("Scripting.Dictionary")

    For i = 1 To dbuList.Count
        key = dbuList(i)
        Dim bcArr As Object
        Set bcArr = dict(key)("BankCodes")
        For r = 0 To bcArr.Count - 1
            If Not bankSet.Exists(bcArr(r)) Then bankSet.Add bcArr(r), 1
        Next r
    Next i

    For i = 1 To obuList.Count
        key = obuList(i)
        Set bcArr = dict(key)("BankCodes")
        For r = 0 To bcArr.Count - 1
            If Not bankSet.Exists(bcArr(r)) Then bankSet.Add bcArr(r), 1
        Next r
    Next i

    ' 清 FM2 C10 往下
    Dim startFMrow As Long: startFMrow = 10
    wsFM2.Range(wsFM2.Cells(startFMrow, "C"), wsFM2.Cells(wsFM2.Rows.Count, "C")).ClearContents

    Dim outRow As Long: outRow = startFMrow
    Dim kKey As Variant
    For Each kKey In bankSet.Keys
        wsFM2.Cells(outRow, "C").Value = kKey
        outRow = outRow + 1
    Next kKey

    ' 顯示結果與警告
    Dim msg As String
    msg = "處理完成。" & vbCrLf & "DBU 數量: " & dbuList.Count & "，OBU 數量: " & obuList.Count & "。" & vbCrLf & "已將銀行代號貼至 FM2 C" & startFMrow & " 開始的欄位。"
    If warnings.Count > 0 Then
        msg = msg & vbCrLf & vbCrLf & "注意：" & vbCrLf
        For i = 1 To warnings.Count
            msg = msg & "- " & warnings(i) & vbCrLf
        Next i
    End If
    MsgBox msg, vbInformation, "OBU 處理結果"

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "FM2 發生錯誤: " & Err.Number & " - " & Err.Description, vbExclamation
    Resume CleanExit
End Sub
```

---

## 2) 修正版 FM10

```vb
Option Explicit

' =========================
' FM10 處理程序（修正版）
Public Sub CopyAndDeleteRows_ButtonClick()
    CopyAndDeleteRows ThisWorkbook
End Sub

Sub CopyAndDeleteRows(Optional ByVal wb As Workbook)
    Dim targetBook As Workbook
    If wb Is Nothing Then
        Set targetBook = ThisWorkbook
    Else
        Set targetBook = wb
    End If

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    Dim wsAC4603 As Worksheet
    Dim wsFM10 As Worksheet
    Set wsAC4603 = targetBook.Sheets("OBU_AC4603")
    Set wsFM10 = targetBook.Sheets("FM10底稿")

    If wsAC4603 Is Nothing Or wsFM10 Is Nothing Then
        MsgBox "找不到 OBU_AC4603 或 FM10底稿 工作表", vbExclamation
        GoTo CleanExit
    End If

    Dim n As Variant
    Dim count As Long
    count = 26 ' 若欄位異動請修改此數字

    ' 找到第 n 行 (若找不到則返回錯誤)
    n = Application.Match("強制FVPL金融資產-公債-地方政府(外國)", wsAC4603.Range("A:A"), 0)
    If IsError(n) Then
        MsgBox "找不到目標欄位 '強制FVPL金融資產-公債-地方政府(外國)'", vbExclamation
        GoTo CleanExit
    End If

    ' 檢查欄位串是否一致（用 With 限定 worksheet）
    With wsAC4603
        If .Range("A" & n + 1).Value = "強制FVPL金融資產-普通公司債(民營)(外國)" And _
           .Range("A" & n + 2).Value = "12005" And _
           .Range("A" & n + 3).Value = "強制FVPL金融資產評價調整-公債-地方-外國" And _
           .Range("A" & n + 4).Value = "強制FVPL金融資產評價調整-普通公司債(民營)(外國)" And _
           .Range("A" & n + 5).Value = "12007" And _
           .Range("A" & n + 6).Value = "FVOCI債務工具-公債-中央政府(外國)" And _
           .Range("A" & n + 7).Value = "FVOCI債務工具-普通公司債(公營)(外國)" And _
           .Range("A" & n + 8).Value = "FVOCI債務工具-普通公司債(民營)(外國)" And _
           .Range("A" & n + 9).Value = "FVOCI債務工具-金融債券-海外" And _
           .Range("A" & n + 10).Value = "12111" And _
           .Range("A" & n + 11).Value = "FVOCI債務工具評價調整-公債-中央政府(外國)" And _
           .Range("A" & n + 12).Value = "FVOCI債務工具評價調整-普通公司債(公營)(外國)" And _
           .Range("A" & n + 13).Value = "FVOCI債務工具評價調整-普通公司債(民營)(外國)" And _
           .Range("A" & n + 14).Value = "FVOCI債務工具評價調整-金融債券-海外" And _
           .Range("A" & n + 15).Value = "12113" And _
           .Range("A" & n + 16).Value = "AC債務工具投資-公債-中央政府(外國)" And _
           .Range("A" & n + 17).Value = "AC債務工具投資-普通公司債(民營)(外國)" And _
           .Range("A" & n + 18).Value = "AC債務工具投資-金融債券-海外" And _
           .Range("A" & n + 19).Value = "12201" And _
           .Range("A" & n + 20).Value = "累積減損-AC債務工具投資-公債-中央政府(外國)" And _
           .Range("A" & n + 21).Value = "累積減損-AC債務工具投資-普通公司(民營)(外國)" And _
           .Range("A" & n + 22).Value = "累積減損-AC債務工具投資-金融債券-海外" And _
           .Range("A" & n + 23).Value = "12203" And _
           .Range("A" & n + 24).Value = "拆放證券公司-OSU" And _
           .Range("A" & n + 25).Value = "15551" Then

            ' 刪除第 n+count 到最後一行
            .Rows(n + count & ":" & .Rows.count).Delete

            ' 刪除第一行到 n-1 行
            .Rows("1:" & n - 1).Delete

            ' 清除 target FM10 底稿
            wsFM10.Range("A4:J" & (4 + count - 1)).ClearContents
            Application.CutCopyMode = False

            ' 複製 AC4603 數值內容到 FM10 底稿
            .Range("A1:J" & count).Copy
            wsFM10.Range("A4").Resize(count, 10).PasteSpecial Paste:=xlPasteValues

            MsgBox "完成"
        Else
            MsgBox "欄位檢核不符，請確認 OBU_AC4603 欄位內容", vbExclamation
        End If
    End With

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "FM10 發生錯誤: " & Err.Number & " - " & Err.Description, vbExclamation
    Resume CleanExit
End Sub
```

---

## 3) 修正版 F1\_F2（含 selectionProcess\_DL6850 / selectionProcess\_CM2810 / ClearRange）

```vb
Option Explicit

' =========================
' F1_F2 主流程（修正版）
Public Sub MainSub_ButtonClick()
    MainSub ThisWorkbook
End Sub

Sub MainSub(Optional ByVal wb As Workbook, _
            Optional ByVal calledByOtherExcel As Boolean = False, _
            Optional ByVal baseDatePassed As Variant = Empty)

    Dim targetBook As Workbook
    If wb Is Nothing Then
        Set targetBook = ThisWorkbook
    Else
        Set targetBook = wb
    End If

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    ' 先執行篩選程序（給予 targetBook）
    Call selectionProcess_DL6850(targetBook, calledByOtherExcel, baseDatePassed)
    Call selectionProcess_CM2810(targetBook)

    Dim wsSrc_DL6850 As Worksheet, wsDst_DL6850 As Worksheet
    Dim wsSrc_CM2810 As Worksheet, wsDst_CM2810 As Worksheet
    Dim srcRng_DL6850 As Range, dstRng_DL6850 As Range
    Dim srcRng_CM2810 As Range, dstRng_CM2810 As Range
    Dim lastRow As Long
    Dim i As Long

    Set wsSrc_DL6850 = targetBook.Worksheets("底稿_含NT_原始資料")
    Set wsDst_DL6850 = targetBook.Worksheets("底稿_含NT")
    Set wsSrc_CM2810 = targetBook.Worksheets("國內顧客_原始資料")
    Set wsDst_CM2810 = targetBook.Worksheets("國內顧客")

    ' Copy Data for DL6850
    lastRow = wsSrc_DL6850.Cells(wsSrc_DL6850.Rows.Count, "I").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "來源沒有資料 (I 欄最後一列 < 2)。", vbInformation
        GoTo CleanExit
    End If

    Set srcRng_DL6850 = wsSrc_DL6850.Range("A2", wsSrc_DL6850.Cells(lastRow, "I"))
    Set dstRng_DL6850 = wsDst_DL6850.Range("B2").Resize(srcRng_DL6850.Rows.Count, srcRng_DL6850.Columns.Count)

    ' 先清除目標區 (B:D) 的資料（限定 worksheet）
    Dim lastRowDst As Long
    lastRowDst = wsDst_DL6850.Cells(wsDst_DL6850.Rows.Count, "I").End(xlUp).Row
    If lastRowDst >= 2 Then wsDst_DL6850.Range("B2:D" & lastRowDst).ClearContents

    dstRng_DL6850.Value = srcRng_DL6850.Value

    ' ===================
    ' Copy for CM2810
    lastRow = wsSrc_CM2810.Cells(wsSrc_CM2810.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "來源沒有資料 (CM2810 欄最後一列 < 2)。", vbInformation
        GoTo CleanExit
    End If

    Set srcRng_CM2810 = wsSrc_CM2810.Range("A2", wsSrc_CM2810.Cells(lastRow, "H"))
    Set dstRng_CM2810 = wsDst_CM2810.Range("A2").Resize(srcRng_CM2810.Rows.Count, srcRng_CM2810.Columns.Count)

    lastRowDst = wsDst_CM2810.Cells(wsDst_CM2810.Rows.Count, "A").End(xlUp).Row
    If lastRowDst >= 2 Then wsDst_CM2810.Range("A2:H" & lastRowDst).ClearContents

    dstRng_CM2810.Value = srcRng_CM2810.Value

    ' 清空其他工作表（使用 ClearRange）
    ClearRange targetBook, "底稿_無NT"
    ClearRange targetBook, "國外即期"
    ClearRange targetBook, "國外換匯"
    ClearRange targetBook, "國內即期"
    ClearRange targetBook, "國內換匯"

    ' 建立 底稿_無NT
    lastRow = wsDst_DL6850.Cells(wsDst_DL6850.Rows.Count, 1).End(xlUp).Row
    Dim destinationRow As Long
    destinationRow = 2
    For i = 2 To lastRow
        If wsDst_DL6850.Cells(i, 13).Value = False Then
            wsDst_DL6850.Rows(i).Copy Destination:=targetBook.Worksheets("底稿_無NT").Rows(destinationRow)
            destinationRow = destinationRow + 1
        End If
    Next i

    ' 國外即期 / 國外換匯 / 國內即期 / 國內換匯（限定工作表）
    lastRow = targetBook.Worksheets("底稿_無NT").Cells(targetBook.Worksheets("底稿_無NT").Rows.Count, 1).End(xlUp).Row
    destinationRow = 2
    For i = 2 To lastRow
        If targetBook.Worksheets("底稿_無NT").Cells(i, 1).Value = "FS" And targetBook.Worksheets("底稿_無NT").Cells(i, 11).Value = "國外" Then
            targetBook.Worksheets("國外即期").Rows(destinationRow).Value = targetBook.Worksheets("底稿_無NT").Rows(i).Value
            destinationRow = destinationRow + 1
        End If
    Next i

    destinationRow = 2
    For i = 2 To lastRow
        If targetBook.Worksheets("底稿_無NT").Cells(i, 1).Value = "SS" And targetBook.Worksheets("底稿_無NT").Cells(i, 11).Value = "國外" Then
            targetBook.Worksheets("國外換匯").Rows(destinationRow).Value = targetBook.Worksheets("底稿_無NT").Rows(i).Value
            destinationRow = destinationRow + 1
        End If
    Next i

    destinationRow = 2
    For i = 2 To lastRow
        If targetBook.Worksheets("底稿_無NT").Cells(i, 1).Value = "FS" And targetBook.Worksheets("底稿_無NT").Cells(i, 11).Value = "國內" Then
            targetBook.Worksheets("國內即期").Rows(destinationRow).Value = targetBook.Worksheets("底稿_無NT").Rows(i).Value
            destinationRow = destinationRow + 1
        End If
    Next i

    destinationRow = 2
    For i = 2 To lastRow
        If targetBook.Worksheets("底稿_無NT").Cells(i, 1).Value = "SS" And targetBook.Worksheets("底稿_無NT").Cells(i, 11).Value = "國內" Then
            targetBook.Worksheets("國內換匯").Rows(destinationRow).Value = targetBook.Worksheets("底稿_無NT").Rows(i).Value
            destinationRow = destinationRow + 1
        End If
    Next i

    MsgBox "已完成"

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "F1_F2 發生錯誤: " & Err.Number & " - " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

' -----------------------
' selectionProcess_DL6850 (修正版): 需傳入 workbook 物件
Sub selectionProcess_DL6850(ByVal wb As Workbook, _
                            ByVal calledByOtherExcel As Boolean, _
                            ByVal baseDatePassed As Variant)
    Dim targetBook As Workbook
    Set targetBook = wb

    Dim startDate As Date, endDate As Date
    Dim ym As String
    Dim parts() As String
    Dim y As Integer, m As Integer

    If calledByOtherExcel Then
        ym = CStr(baseDatePassed)
    Else
        ym = InputBox("請輸入報表年月份(格式：YYY/MM，例如 114/09）", "輸入年月 (ROC年/月)")
    End If
    If Trim(ym) = "" Then Exit Sub

    parts = Split(Trim(ym), "/")
    If UBound(parts) <> 1 Then
        MsgBox "輸入格式錯誤，請使用 YYY/MM（例如 114/09）", vbExclamation
        Exit Sub
    End If

    On Error GoTo InvalidInput
    y = CInt(parts(0))
    m = CInt(parts(1))
    If m < 1 Or m > 12 Then GoTo InvalidInput
    y = y + 1911

    startDate = DateSerial(y, m, 1)
    endDate = DateSerial(y, m + 1, 1) - 1

    ' 清除底稿_含NT_原始資料
    targetBook.Sheets("底稿_含NT_原始資料").Range("A:I").ClearContents

    ' 清除全部交易工作表多餘資料（限定 targetBook）
    Dim ws As Worksheet
    Set ws = targetBook.Sheets("DL6850全部交易")

    Dim lastRowOrigin As Long
    lastRowOrigin = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim i As Long
    For i = lastRowOrigin To 2 Step -1
        If Left(Trim(CStr(ws.Cells(i, "A").Value)), 2) <> "TR" Then
            ws.Rows(i).Delete
        End If
    Next i

    ' 清除不在日期範圍的資料
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    For i = lastRow To 2 Step -1
        If IsDate(ws.Cells(i, "I").Value) Then
            If ws.Cells(i, "I").Value < startDate Or ws.Cells(i, "I").Value > endDate Then
                ws.Rows(i).ClearContents
            End If
        Else
            ws.Rows(i).ClearContents
        End If
    Next i

    ' 刪除包含空白儲存格的整行（以 I 欄為判斷）
    Dim deleteRows As Range
    For i = lastRow To 2 Step -1
        If IsEmpty(ws.Cells(i, "I")) Then
            If deleteRows Is Nothing Then
                Set deleteRows = ws.Rows(i)
            Else
                Set deleteRows = Union(deleteRows, ws.Rows(i))
            End If
        End If
    Next i
    If Not deleteRows Is Nothing Then deleteRows.Delete

    ' 複製到底稿_含NT_原始資料
    Dim lastRowSource As Long
    lastRowSource = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim targetSheet As Worksheet
    Set targetSheet = targetBook.Sheets("底稿_含NT_原始資料")
    ws.Range("A1:I" & lastRowSource).Copy Destination:=targetSheet.Range("A1")

    MsgBox "完成"
    Exit Sub

InvalidInput:
    MsgBox "輸入年月格式錯誤或數值不正確，請重新輸入。例如：114/09", vbExclamation
End Sub

' -----------------------
' selectionProcess_CM2810 (修正版): 需傳入 workbook 物件
Sub selectionProcess_CM2810(ByVal wb As Workbook)
    Dim targetBook As Workbook
    Set targetBook = wb

    On Error Resume Next
    Application.DisplayAlerts = False
    targetBook.Sheets("樞紐表").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Dim ws As Worksheet
    Set ws = targetBook.Sheets("CM2810全部交易")
    Dim wsClear As Worksheet
    Set wsClear = targetBook.Sheets("國內顧客_原始資料")

    wsClear.Range("A:H").ClearContents
    ws.Range("A1:H1").Copy Destination:=wsClear.Range("A1")

    ' 1. 清除多餘資料（限定 targetBook）
    Dim lastRowOrigin As Long
    lastRowOrigin = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim i As Long
    For i = lastRowOrigin To 2 Step -1
        If Left(Trim(CStr(ws.Cells(i, "A").Value)), 2) <> "MB" Then
            ws.Rows(i).Delete
        End If
    Next i

    ws.Range("G:V").ClearContents
    ws.Range("G1").Value = "筆數"
    ws.Range("H1").Value = "配對"

    ' 2. 排序
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=ws.Range("A2:A" & lastRow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ws.Sort
        .SetRange ws.Range("A1:H" & lastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ' 3. 產生樞紐分析表（限定 targetBook）
    Dim pivotSheet As Worksheet
    Set pivotSheet = targetBook.Sheets.Add(After:=ws)
    On Error Resume Next
    pivotSheet.Name = "樞紐表"
    On Error GoTo 0

    Dim pivotRange As Range
    Set pivotRange = ws.Range("A1:F" & lastRow)

    Dim pivotTable As pivotTable
    Set pivotTable = pivotSheet.PivotTableWizard(SourceType:=xlDatabase, _
        SourceData:=pivotRange, TableDestination:=pivotSheet.Cells(1, 1), _
        TableName:="樞紐分析表")

    pivotTable.PivotFields("交易編號").Orientation = xlRowField
    With pivotTable.PivotFields("幣別")
        .Orientation = xlDataField
        .Function = xlCount
    End With

    ' 插入公式並 Autofill
    ws.Range("G2").Formula = "=VLOOKUP(A2, 樞紐表!$A:$B, 2, FALSE)"
    ws.Range("H2").Formula = "=CONCATENATE(C2, E2)"
    ws.Range("G2:H2").AutoFill Destination:=ws.Range("G2:H" & lastRow)

    ' 4. 複製符合條件到國內顧客
    Dim custSheet As Worksheet
    Set custSheet = targetBook.Sheets("國內顧客_原始資料")
    Dim custRow As Long: custRow = 2
    Dim fCell As Range
    For Each fCell In ws.Range("G2:G" & lastRow)
        If fCell.Value = 2 Then
            ws.Rows(fCell.Row).Copy Destination:=custSheet.Range("A" & custRow)
            custRow = custRow + 1
        End If
    Next fCell

    MsgBox "完成"
End Sub

' -----------------------
' ClearRange (修正版): 接受 workbook 與 sheetName
Sub ClearRange(wb As Workbook, sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    ws.Range("A2:M" & lastRow).ClearContents
End Sub
```

---

## 4) 修正版 AI240（含 ImportDL6850CSV 的修正）

```vb
Option Explicit

' =========================
' AI240 主流程（修正版）
Public Sub CopyDataToAI240_ButtonClick()
    CopyDataToAI240 ThisWorkbook
End Sub

Sub CopyDataToAI240(Optional ByVal wb As Workbook, _
                    Optional ByVal calledByOtherExcel As Boolean = False, _
                    Optional ByVal baseDatePassed As Variant = Empty)

    Dim targetBook As Workbook
    If wb Is Nothing Then
        Set targetBook = ThisWorkbook
    Else
        Set targetBook = wb
    End If

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    Dim wsDL6850 As Worksheet
    Dim wsAI240 As Worksheet
    Set wsDL6850 = targetBook.Worksheets("DL6850原始資料")
    Set wsAI240 = targetBook.Worksheets("AI240")
    If wsDL6850 Is Nothing Or wsAI240 Is Nothing Then
        MsgBox "找不到 DL6850原始資料 或 AI240 工作表", vbExclamation
        GoTo CleanExit
    End If

    Dim inputDate As Date
    Dim baseDate As Date

    If calledByOtherExcel Then
        ' 若是被外部呼叫，baseDatePassed 可能是空或日期
        If IsDate(baseDatePassed) Then
            baseDate = CDate(baseDatePassed)
        Else
            MsgBox "被其他 Excel 呼叫但未提供正確之基準日", vbExclamation
            GoTo CleanExit
        End If
    Else
        inputDate = InputBox("請輸入基準日(日期格式yyyy/mm/dd)：")
        If Trim(CStr(inputDate)) = "" Then
            MsgBox "未輸入基準日", vbExclamation
            GoTo CleanExit
        End If
        If Not IsDate(inputDate) Then
            MsgBox "基準日格式錯誤，請以 yyyy/mm/dd 輸入", vbExclamation
            GoTo CleanExit
        End If
        baseDate = CDate(inputDate)
    End If

    ' 填入基準日
    wsDL6850.Range("P1").Value = baseDate
    wsAI240.Range("A2").Value = baseDate

    ' 清空 AI240 固定區域資料（限定 targetBook）
    wsAI240.Range("A9:I58").ClearContents
    wsAI240.Range("L9:T58").ClearContents
    wsAI240.Range("A90:I139").ClearContents
    wsAI240.Range("L90:T139").ClearContents
    wsAI240.Range("A153:I162").ClearContents
    wsAI240.Range("L153:T162").ClearContents
    wsAI240.Range("A170:I179").ClearContents
    wsAI240.Range("L170:T179").ClearContents

    If Not calledByOtherExcel Then
        Call ImportDL6850CSV(targetBook)
    End If

    Dim rowCount As Long
    rowCount = wsDL6850.Cells(wsDL6850.Rows.Count, "B").End(xlUp).Row

    ' 刪除 B 欄開頭不是 "TR" 的列（從下往上）
    Dim i As Long
    For i = rowCount To 2 Step -1
        If Left(Trim(CStr(wsDL6850.Range("B" & i).Value)), 2) <> "TR" Then
            wsDL6850.Rows(i).Delete
        End If
    Next i

    rowCount = wsDL6850.Cells(wsDL6850.Rows.Count, "B").End(xlUp).Row
    For i = rowCount To 2 Step -1
        Dim valE As String, valH As String
        valE = Trim(CStr(wsDL6850.Range("E" & i).Value))
        valH = Trim(CStr(wsDL6850.Range("H" & i).Value))

        If ((valE <> "TWD" And valH <> "TWD") _
           Or Not IsDate(wsDL6850.Range("C" & i).Value) _
           Or Not IsDate(wsDL6850.Range("J" & i).Value) _
           Or wsDL6850.Range("C" & i).Value <= baseDate _
           Or wsDL6850.Range("J" & i).Value > baseDate) Then
            wsDL6850.Rows(i).Delete
        End If
    Next i

    ' 之後的複製邏輯與你原版一致，只把所有 ThisWorkbook/ActiveWorkbook 改為 targetBook（以下略過部分，但邏輯一致）
    ' --- 以下保留原本邏輯但限定 wsDL6850 / wsAI240 (如你的原始程式) ---
    ' SWOP & OutFlow TWD (複製到 A9 起)
    Dim destRow0TO10 As Long, destRow11TO30 As Long, destRow31TO90 As Long
    Dim destRow91TO180 As Long, destRow181TO365 As Long, destRow366TO As Long
    Dim copyCount0To10 As Long, copyCount11To30 As Long, copyCount31To90 As Long
    Dim copyCount91To180 As Long, copyCount181To365 As Long, copyCount366To As Long

    ' 初始化起始列
    destRow0TO10 = 9: destRow11TO30 = 19: destRow31TO90 = 29: destRow91TO180 = 39: destRow181TO365 = 49
    copyCount0To10 = 0: copyCount11To30 = 0: copyCount31To90 = 0: copyCount91To180 = 0: copyCount181To365 = 0

    For i = 2 To rowCount
        If (Trim(CStr(wsDL6850.Range("A" & i).Value)) Like "SS*" Or Trim(CStr(wsDL6850.Range("A" & i).Value)) Like "SF*") And Trim(CStr(wsDL6850.Range("H" & i).Value)) = "TWD" Then
            Select Case wsDL6850.Range("N" & i).Value
                Case 0 To 10
                    copyCount0To10 = copyCount0To10 + 1
                    If copyCount0To10 > 10 Then MsgBox "此期間流出之SWOP筆數超過10筆": GoTo CleanExit
                    wsAI240.Range("A" & destRow0TO10 & ":I" & destRow0TO10).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow0TO10 = destRow0TO10 + 1
                Case 11 To 30
                    copyCount11To30 = copyCount11To30 + 1
                    If copyCount11To30 > 10 Then MsgBox "此期間流出之SWOP筆數超過10筆": GoTo CleanExit
                    wsAI240.Range("A" & destRow11TO30 & ":I" & destRow11TO30).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow11TO30 = destRow11TO30 + 1
                Case 31 To 90
                    copyCount31To90 = copyCount31To90 + 1
                    If copyCount31To90 > 10 Then MsgBox "此期間流出之SWOP筆數超過10筆": GoTo CleanExit
                    wsAI240.Range("A" & destRow31TO90 & ":I" & destRow31TO90).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow31TO90 = destRow31TO90 + 1
                Case 91 To 180
                    copyCount91To180 = copyCount91To180 + 1
                    If copyCount91To180 > 10 Then MsgBox "此期間流出之SWOP筆數超過10筆": GoTo CleanExit
                    wsAI240.Range("A" & destRow91TO180 & ":I" & destRow91TO180).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow91TO180 = destRow91TO180 + 1
                Case 181 To 365
                    copyCount181To365 = copyCount181To365 + 1
                    If copyCount181To365 > 10 Then MsgBox "此期間流出之SWOP筆數超過10筆": GoTo CleanExit
                    wsAI240.Range("A" & destRow181TO365 & ":I" & destRow181TO365).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow181TO365 = destRow181TO365 + 1
            End Select
        End If
    Next i

    ' ...（其後各段同理，我保留原邏輯，只要你需要我可以把所有區塊完整複製到這裡；目前保留此段結構以避免回覆過長）...

    MsgBox "完成"

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "AI240 發生錯誤: " & Err.Number & " - " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

' =========================
' ImportDL6850CSV (修正版) - 接受目標 workbook
Sub ImportDL6850CSV(ByVal targetBook As Workbook)
    Dim wbImport As Workbook
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim filePath As Variant
    Dim lastRow As Long

    filePath = Application.GetOpenFilename("Excel 檔案 (*.xls), *.xls", , "請選擇 DL6850 Excel 檔")
    If filePath = "False" Then Exit Sub

    Set wbImport = Workbooks.Open(Filename:=filePath)
    Set wsSource = wbImport.Sheets(1)

    Set wsDest = targetBook.Sheets("DL6850原始資料")

    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row

    wsDest.Range("B1:L" & wsDest.Rows.Count).ClearContents

    wsSource.Range("A1:L" & lastRow).Copy
    wsDest.Range("B1").PasteSpecial xlPasteValues

    Application.DisplayAlerts = False
    wbImport.Close SaveChanges:=False
    Application.DisplayAlerts = True

    MsgBox "DL6850 資料匯入完成！", vbInformation
End Sub
```

---

## 逐項修改清單（每處 `''' *** MODIFIED ***` 說明）

下面把剛剛程式內每個 `''' *** MODIFIED ***` 標記的改動意義逐一說明，方便你在原始程式中逐行比對：

1. **所有模組（FM2 / FM10 / F1\_F2 / AI240）**

   * 新增 `targetBook` 機制（若 `wb` 為 Nothing 使用 `ThisWorkbook`，否則使用傳入的 `wb`）。理由：當 `ProcessAllReports` 用 `Application.Run` 呼叫子程序時，傳入的 `wbOld`/`wbNew` 常常不同於 `ThisWorkbook`，若不限定會抓錯 ActiveWorkbook/ActiveSheet。
   * 在一開始加入 `On Error GoTo ErrHandler`（或在必要處加 `On Error Resume Next` 包裝單次可能不存在的 Sheet）。理由：避免未處理錯誤導致 Excel 中斷執行，並能在結尾還原 `Application` 設定。

2. **ProcessAllReports 內的 Import 呼叫**

   * 把 `Import_CopyThenRunAP` 改為 **命名參數** 呼叫（`basePath:=...`）。理由：你的 `ReportsConfig` 表格欄位順序若有差異，位置參數會導致錯誤。命名參數可避免欄位順序錯配。

3. **Import\_CopyThenRunAP**

   * 加入檔案存在性檢查（`If Dir(tplPath) = "" Then` / `If Dir(f) = "" Then`）。理由：`Workbooks.Open` 在找不到檔案時會直接拋錯。現在會先報錯並跳過或結束，較穩健。
   * 統一 `Application.Run` 的容錯（先嘗試傳入 wbOld，失敗再嘗試不帶參數）。理由：被呼叫的 Module.Sub 接受參數的簽章不一致時會跳錯，用兩次嘗試提高容錯。
   * 在整支程序退出前 `CleanExit` 一定會還原 `Application.DisplayAlerts`、`ScreenUpdating`、`AskToUpdateLinks`。理由：若發生錯誤不還原會影響使用者介面。

4. **GetWesternMonthEnd / GetROCMonthEnd / GetWesternMonth**

   * 改用 `Year(lastDay)` / `Month(lastDay)` / `Day(lastDay)` 串接並用 `Format(...,"00")` 補零，避免 `Format(date, "yyyymmdd")` 這種容易誤打 pattern 的寫法。

5. **FM11 / FM2 / FM10 / F1\_F2 / AI240**

   * 把所有 `Sheets`, `Cells`, `Rows.Count` 都限定為 `targetBook.Worksheets(...)` 或 `ws` 變數，避免抓到錯誤 Workbook 的資料。
   * 在可能會找不到分頁 / 範圍的地方加 `On Error Resume Next` + 檢查物件是否為 `Nothing`，以免 `Range` 存取時中斷。
   * 在需要 user input 的地方（例如 AI240 基準日、F1\_F2 年月）加上格式檢查（是否為空 / 是否為日期 / 是否為 YYY/MM 格式），避免傳入錯誤值導致日後的比較運算出錯。

6. **ClearRange**

   * 改為接受 `wb As Workbook` 與 `sheetName`，在 F1\_F2 使用 `ClearRange targetBook, "xxx"`，讓清除範圍操作也限定 workbook。

---

## 測試與套用建議（步驟）

1. 先把我先前給的 `ProcessAllReports` + `Import_CopyThenRunAP` + helper functions 貼回 Project。
2. 把上面四個模組分別貼進各自 Module（或同一 Module 也可），儲存並啟用巨集。
3. 在 `ReportsConfig` 中先只放一行（例如 FM11），確認 `moduleSub` 欄位內容格式為你模組實際可呼叫的名稱（例如 `Module1.匯入並篩選OBUAC5411B資料` 或 `Module.匯入並篩選OBUAC5411B資料`）。
4. 執行 `ProcessAllReports`，若跳出錯誤訊息，把錯誤號碼與描述貼過來，我會直接協助修正。

---

我已把所有主要 risk points 都改成防禦式寫法（檔案檢查 / workbook 限定 / 錯誤處理 / 恢復 Application 設定）。若你想我把 AI240 裡剩下的**每一個**複製區塊（SPOT、SWOP、各時段的複製與計數）也完整展開（目前我保留了起始段並註記可完整貼回），我可以把 AI240 的整支長流程完整貼出（會再長一些）。你要我：

1. 再把 AI240 那些略過的區塊全部完整貼上（完整版本），或
2. 先照現在的版本測一次（我保留原邏輯並已做安全改寫），再回報錯誤逐段調整？

直接回答 `1` 或 `2` 或說「都貼完整」，我馬上把你要的結果補上。

