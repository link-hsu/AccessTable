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
