下面一次把 版本 A（不新增右側 Report 編輯視窗） 與 版本 B（有新增右側 Report 編輯視窗） 的 完整可貼入 VBA 程式碼 分開給你 — 每個版本包含三個 Module（ConfigIO、Helpers、UIHandlers）與 UI_Main 的 Worksheet 事件程式。你可以直接把每個版本的三個 Module + Sheet code 各別貼入 VBA，即可啟用該版本的 UI 行為。

我已在程式中以 ********** MODIFIED ********** 明確標示所有我針對你要求（移除 lock、初始鎖定、右側 Report tab 等）所做的修改處。
注意：兩個版本都會使用相同的 CSV 格式與 config\ 資料夾結構（config\tblReports.csv、tblUpdateSheet.csv、tblExportPDF.csv、Mappings.csv）；若你已經有其他 modules（例如你原本的 ProcessAllReports_New），可保留不衝突，Process 前只要先呼叫 LoadAllCSVToSheets 即可讀入最新設定。

⸻

版本 A（不新增右側 tblReports 編輯視窗）——完整程式碼

Module: ConfigIO

貼入一個新 Module（名稱建議 ConfigIO）：

Option Explicit

' ----------------------------
' Config IO: load/save CSV and helpers
' ----------------------------

Public Const CONFIG_FOLDER As String = "config"
Public Const CONFIG_BACKUP_FOLDER As String = "config\backup"
Public Const CONFIG_LOCKS_FOLDER As String = "config\locks"

' Ensure folders exist
Public Sub EnsureConfigFolders()
    Dim base As String: base = ThisWorkbook.Path
    MkDirRecursive base & "\" & CONFIG_FOLDER
    MkDirRecursive base & "\" & CONFIG_BACKUP_FOLDER
    MkDirRecursive base & "\" & CONFIG_LOCKS_FOLDER
End Sub

' Load all CSV files into corresponding sheets (overwrite)
Public Sub LoadAllCSVToSheets()
    EnsureConfigFolders
    Dim base As String: base = ThisWorkbook.Path & "\" & CONFIG_FOLDER & "\"
    LoadCSVToSheet base & "tblReports.csv", "tblReports"
    LoadCSVToSheet base & "tblUpdateSheet.csv", "tblUpdateSheet"
    LoadCSVToSheet base & "tblExportPDF.csv", "tblExportPDF"
    LoadCSVToSheet base & "Mappings.csv", "Mappings"
End Sub

' Load single CSV to sheet (overwrite)
Public Sub LoadCSVToSheet(csvPath As String, targetSheetName As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(targetSheetName)
    ws.Cells.Clear ' 清空舊資料

    ' 建立 ADODB.Stream
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2 ' 文字
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile csvPath

    ' 讀整個 CSV 內容
    Dim fullText As String
    fullText = stm.ReadText
    stm.Close
    Set stm = Nothing

    ' 分行（支援 LF 或 CRLF）
    Dim lines() As String
    lines = Split(fullText, vbCrLf)
    If UBound(lines) = 0 Then lines = Split(fullText, vbLf)

    Dim r As Long, c As Long
    For r = LBound(lines) To UBound(lines)
        If Trim(lines(r)) <> "" Then
            Dim columns() As String
            columns = ParseCSVLine(lines(r)) ' 使用你原本的 CSV 解析函數或自己寫

            For c = LBound(columns) To UBound(columns)
                ws.Cells(r + 1, c + 1).Value = columns(c)
            Next c
        End If
    Next r
End Sub

' Public Sub LoadCSVToSheet(csvPath As String, targetSheetName As String)
'     On Error GoTo ErrHandler
'     Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
'     Dim ws As Worksheet
'     On Error Resume Next
'     Set ws = ThisWorkbook.Worksheets(targetSheetName)
'     On Error GoTo ErrHandler
'     If ws Is Nothing Then Exit Sub

'     If Not fso.FileExists(csvPath) Then
'         ' If file missing, just clear sheet
'         ws.Cells.Clear
'         Exit Sub
'     End If

'     Dim ts As Object: Set ts = fso.OpenTextFile(csvPath, 1, False, -1)
'     ws.Cells.Clear
'     Dim rowIndex As Long: rowIndex = 1
'     Do While Not ts.AtEndOfStream
'         Dim line As String: line = ts.ReadLine
'         Dim arr As Variant: arr = ParseCSVLine(line)
'         Dim c As Long
'         For c = LBound(arr) To UBound(arr)
'             ws.Cells(rowIndex, c + 1).Value = arr(c)
'         Next c
'         rowIndex = rowIndex + 1
'     Loop
'     ts.Close
'     Exit Sub
' ErrHandler:
'     LogError "LoadCSVToSheet error: " & Err.Number & " " & Err.Description & " file=" & csvPath
' End Sub    


' Save a sheet to CSV (with backup)
Public Sub SaveSheetToCSV(sheetName As String, csvPath As String)
    On Error GoTo ErrHandler
    EnsureConfigFolders
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(sheetName)
    If ws Is Nothing Then Exit Sub

    ' backup if exists
    If fso.FileExists(csvPath) Then
        Dim bkp As String
        bkp = ThisWorkbook.Path & "\" & CONFIG_BACKUP_FOLDER & "\" & Replace(Mid(csvPath, InStrRev(csvPath, "\") + 1), ".csv", "") & "_" & Format(Now, "yyyymmdd_HHnnss") & ".csv"
        FileCopy csvPath, bkp
    End If

    Dim ts As Object: Set ts = fso.CreateTextFile(csvPath, True, False)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow = 0 Then lastRow = 1
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol = 0 Then lastCol = 1

    Dim r As Long
    For r = 1 To lastRow
        Dim line As String: line = ""
        Dim c As Long
        For c = 1 To lastCol
            Dim v As String: v = CStr(Nz(ws.Cells(r, c).Value, ""))
            If InStr(v, ",") > 0 Or InStr(v, """") > 0 Then
                v = Replace(v, """", """""")
                v = """" & v & """"
            End If
            If c = 1 Then line = v Else line = line & "," & v
        Next c
        ts.WriteLine line
    Next r
    ts.Close
    Exit Sub
ErrHandler:
    LogError "SaveSheetToCSV error: " & Err.Number & " " & Err.Description & " sheet=" & sheetName
End Sub

' Helper: parse simple CSV line (handles double quote escaping)
Public Function ParseCSVLine(lineText As String) As Variant
    Dim pattern As String
    pattern = """[^""]*""|[^,]+"
    
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.Pattern = pattern

    Dim matches As Object
    Set matches = regEx.Execute(lineText)
    
    Dim result() As String
    ReDim result(0 To matches.Count - 1)

    Dim i As Long
    For i = 0 To matches.Count - 1
        Dim val As String
        val = matches(i).Value
        ' 移除雙引號
        If Left(val, 1) = """" And Right(val, 1) = """" Then
            val = Mid(val, 2, Len(val) - 2)
        End If
        result(i) = val
    Next i

    ParseCSVLine = result
End Function

' Public Function ParseCSVLine(line As String) As Variant
'     Dim result As Collection: Set result = New Collection
'     Dim i As Long: i = 1
'     Dim buf As String: buf = ""
'     Dim inQuotes As Boolean: inQuotes = False
'     Do While i <= Len(line)
'         Dim ch As String: ch = Mid(line, i, 1)
'         If ch = """" Then
'             If inQuotes And i < Len(line) And Mid(line, i + 1, 1) = """" Then
'                 buf = buf & """"
'                 i = i + 1
'             Else
'                 inQuotes = Not inQuotes
'             End If
'         ElseIf ch = "," And Not inQuotes Then
'             result.Add buf
'             buf = ""
'         Else
'             buf = buf & ch
'         End If
'         i = i + 1
'     Loop
'     result.Add buf
'     Dim arr() As String
'     ReDim arr(0 To result.Count - 1)
'     Dim idx As Long
'     For idx = 1 To result.Count
'         arr(idx - 1) = result(idx)
'     Next idx
'     ParseCSVLine = arr
' End Function

' Nz helper
Public Function Nz(v As Variant, opt As Variant) As Variant
    If IsError(v) Then Nz = opt: Exit Function
    If IsNull(v) Then Nz = opt: Exit Function
    If Trim(CStr(v & "")) = "" Then Nz = opt Else Nz = v
End Function


⸻

Module: Helpers

貼入一個 Module（名稱建議 Helpers）：

Option Explicit

' Logging + MkDirRecursive + small helpers

Public Sub LogInfo(msg As String)
    WriteLog "INFO", msg
End Sub
Public Sub LogWarn(msg As String)
    WriteLog "WARN", msg
End Sub
Public Sub LogError(msg As String)
    WriteLog "ERROR", msg
End Sub

Private Sub WriteLog(level As String, msg As String)
    On Error Resume Next
    Dim f As Integer: f = FreeFile
    Dim logFolder As String: logFolder = ThisWorkbook.Path & "\logs"
    If Dir(logFolder, vbDirectory) = "" Then MkDirRecursive logFolder
    Dim logPath As String: logPath = logFolder & "\RunLog_" & Format(Date, "yyyymmdd") & ".txt"
    Open logPath For Append As #f
    Print #f, Format(Now, "yyyy-mm-dd HH:nn:ss") & " | " & level & " | " & msg
    Close #f
End Sub

Public Sub MkDirRecursive(path As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(path) = 0 Then Exit Sub
    If Right(path, 1) = "\" Then path = Left(path, Len(path) - 1)
    Dim parts() As String: parts = Split(path, "\")
    Dim cur As String: cur = parts(0)
    Dim i As Long
    For i = 1 To UBound(parts)
        cur = cur & "\" & parts(i)
        If Not fso.FolderExists(cur) Then
            On Error Resume Next
            fso.CreateFolder cur
            On Error GoTo 0
        End If
    Next i
End Sub


⸻

Module: UIHandlers （版本 A）

把整個內容貼入一個 Module（名稱建議 UIHandlers） — 這是版本 A 的完整 UI 處理：

Option Explicit

' UIHandlers (Version A: without right-side tblReports editor)
' ********** MODIFIED **********: 去除 file-lock 實作，改為單人模式下的 in-memory editing flag
' ********** MODIFIED **********: 初始化時把相關 config sheets 與 UI sheet 設為受保護（不可編輯）
' ********** MODIFIED **********: EnterEditMode 只解鎖右側 staging 範圍以便編輯；Save/Cancel 會恢復鎖定

Public currentSelectedReportID As String
Public currentActiveTab As String  ' "UpdateSheet", "ExportPDF", "Mappings"
Public currentInEditMode As Boolean
Private stagingRightTableRangeAddress As String

' Initialize UI (call on Workbook_Open or manually)
Public Sub InitializeUI()
    LoadAllCSVToSheets
    EnsureConfigFolders

    currentActiveTab = "UpdateSheet"
    currentSelectedReportID = ""
    currentInEditMode = False

    RefreshLeftPanel
    SetInitialProtectionState
    RefreshRightPanel "", currentActiveTab

    MsgBox "UI 初始化完成。請選擇左側報表，按 Edit 進入編輯模式。", vbInformation
End Sub

' ********** MODIFIED **********
Public Sub SetInitialProtectionState()
    Dim shNames As Variant
    shNames = Array("tblReports", "tblUpdateSheet", "tblExportPDF", "Mappings", "UI_Main")
    Dim s As Variant
    For Each s In shNames
        On Error Resume Next
        Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(s)
        If Not ws Is Nothing Then
            ws.Cells.Locked = True
            ws.Protect Password:="", UserInterfaceOnly:=True
        End If
        On Error GoTo 0
    Next s
End Sub
' ********** MODIFIED END **********

Public Sub RefreshLeftPanel()
    Dim src As Worksheet: Set src = ThisWorkbook.Worksheets("tblReports")
    Dim dst As Worksheet: Set dst = ThisWorkbook.Worksheets("UI_Main")
    dst.Range("A1:C1000").Clear
    src.UsedRange.Copy Destination:=dst.Range("A1")
End Sub

Public Sub OnReportSelectedFromUI(reportID As String)
    currentSelectedReportID = reportID
    If Trim(currentSelectedReportID) = "" Then Exit Sub
    currentInEditMode = False
    RefreshRightPanel currentSelectedReportID, currentActiveTab
End Sub

Public Sub RefreshRightPanel(reportID As String, activeTab As String)
    Dim dst As Worksheet: Set dst = ThisWorkbook.Worksheets("UI_Main")
    Dim src As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim outTopLeft As Range: Set outTopLeft = dst.Range("E3")
    dst.Unprotect
    dst.Range("E3:Z1000").Clear

    Select Case activeTab
        Case "UpdateSheet": Set src = ThisWorkbook.Worksheets("tblUpdateSheet")
        Case "ExportPDF":  Set src = ThisWorkbook.Worksheets("tblExportPDF")
        Case "Mappings":   Set src = ThisWorkbook.Worksheets("Mappings")
        Case Else:         Set src = ThisWorkbook.Worksheets("tblUpdateSheet")
    End Select
    

    lastCol = src.Cells(1, src.Columns.Count).End(xlToLeft).Column
    src.Range(src.Cells(1, 1), src.Cells(1, lastCol)).Copy Destination:=outTopLeft

    If Trim(reportID) = "" Then
        stagingRightTableRangeAddress = outTopLeft.Resize(1, lastCol).Address
        dst.Protect UserInterfaceOnly:=True
        Exit Sub
    End If

    lastRow = src.Cells(src.Rows.Count, 1).End(xlUp).Row
    Dim r As Long, outRow As Long: outRow = 1
    For r = 2 To lastRow
        If Trim(CStr(src.Cells(r, 1).Value)) = reportID Then
            src.Range(src.Cells(r, 1), src.Cells(r, lastCol)).Copy Destination:=outTopLeft.Offset(outRow, 0)
            outRow = outRow + 1
        End If
    Next r

    If outRow = 1 Then
        stagingRightTableRangeAddress = outTopLeft.Resize(1, lastCol).Address
    Else
        stagingRightTableRangeAddress = outTopLeft.Resize(outRow, lastCol).Address
    End If

    dst.Range(stagingRightTableRangeAddress).Locked = True
    dst.Protect UserInterfaceOnly:=True
End Sub

' Navbar
Public Sub Nav_ShowUpdateSheet(): currentActiveTab = "UpdateSheet": If currentSelectedReportID <> "" Then RefreshRightPanel currentSelectedReportID, currentActiveTab: End Sub
Public Sub Nav_ShowExportPDF():  currentActiveTab = "ExportPDF":  If currentSelectedReportID <> "" Then RefreshRightPanel currentSelectedReportID, currentActiveTab: End Sub
Public Sub Nav_ShowMappings():   currentActiveTab = "Mappings":   If currentSelectedReportID <> "" Then RefreshRightPanel currentSelectedReportID, currentActiveTab: End Sub

' ********** MODIFIED **********
Public Sub EnterEditMode()
    If currentInEditMode Then
        MsgBox "已處於編輯模式。", vbInformation
        Exit Sub
    End If
    If Trim(currentSelectedReportID) = "" Then
        MsgBox "請先選擇一個 ReportID。", vbExclamation
        Exit Sub
    End If

    currentInEditMode = True
    Dim ui As Worksheet: Set ui = ThisWorkbook.Worksheets("UI_Main")
    ui.Unprotect
    On Error Resume Next
    ui.Range(stagingRightTableRangeAddress).Locked = False
    On Error GoTo 0
    ui.Activate
    On Error Resume Next
    ui.Range(stagingRightTableRangeAddress).Cells(2, 1).Select
    On Error GoTo 0

    MsgBox "已進入編輯模式。修改右側內容後請按 Save 或 Cancel。", vbInformation
End Sub
' ********** MODIFIED END **********

Public Sub CancelEdits()
    If Not currentInEditMode Then
        MsgBox "不在編輯模式，無需取消。", vbInformation
        Exit Sub
    End If
    If Trim(currentSelectedReportID) <> "" Then
        RefreshRightPanel currentSelectedReportID, currentActiveTab
    Else
        RefreshRightPanel "", currentActiveTab
    End If
    currentInEditMode = False
    MsgBox "已取消修改，回復為原始設定。", vbInformation
End Sub

' ********** MODIFIED **********
Public Sub SaveEdits()
    If Not currentInEditMode Then
        MsgBox "目前不在編輯模式。", vbExclamation
        Exit Sub
    End If
    If Trim(currentSelectedReportID) = "" Then
        MsgBox "沒有選取 ReportID，無法儲存。", vbExclamation
        Exit Sub
    End If

    Dim ui As Worksheet: Set ui = ThisWorkbook.Worksheets("UI_Main")
    Dim rng As Range: Set rng = ui.Range(stagingRightTableRangeAddress)

    Dim headerCols As Long: headerCols = rng.Columns.Count
    Dim r As Long
    For r = 2 To rng.Rows.Count
        Dim v As String: v = Trim(CStr(rng.Cells(r, 1).Value))
        If v = "" Then
            Dim allBlank As Boolean: allBlank = True
            Dim rr As Long
            For rr = r To rng.Rows.Count
                If Trim(CStr(rng.Cells(rr, 1).Value)) <> "" Then allBlank = False: Exit For
            Next rr
            If allBlank Then Exit For
        ElseIf v <> currentSelectedReportID Then
            MsgBox "第 " & r & " 列的 ReportID 不正確（應為 " & currentSelectedReportID & "）。請修正後再儲存。", vbExclamation
            Exit Sub
        End If
    Next r

    Dim targetSheetName As String
    Select Case currentActiveTab
        Case "UpdateSheet": targetSheetName = "tblUpdateSheet"
        Case "ExportPDF":  targetSheetName = "tblExportPDF"
        Case "Mappings":   targetSheetName = "Mappings"
        Case Else: targetSheetName = "tblUpdateSheet"
    End Select

    Dim wsTarget As Worksheet: Set wsTarget = ThisWorkbook.Worksheets(targetSheetName)
    Dim srcLastRow As Long: srcLastRow = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row
    Dim keepArr() As Variant
    Dim keepCount As Long: keepCount = 0
    Dim iRow As Long, j As Long

    If srcLastRow >= 2 Then
        For iRow = 2 To srcLastRow
            If Trim(CStr(wsTarget.Cells(iRow, 1).Value)) <> currentSelectedReportID Then
                keepCount = keepCount + 1
                ReDim Preserve keepArr(1 To keepCount, 1 To headerCols)
                For j = 1 To headerCols
                    keepArr(keepCount, j) = Nz(wsTarget.Cells(iRow, j).Value, "")
                Next j
            End If
        Next iRow
    End If

    wsTarget.Range(wsTarget.Cells(2, 1), wsTarget.Cells(wsTarget.Rows.Count, headerCols)).ClearContents
    Dim outR As Long: outR = 2
    For iRow = 1 To keepCount
        For j = 1 To headerCols
            wsTarget.Cells(outR, j).Value = keepArr(iRow, j)
        Next j
        outR = outR + 1
    Next iRow

    Dim dataRow As Long
    For dataRow = 2 To rng.Rows.Count
        Dim firstColVal As String: firstColVal = Trim(CStr(rng.Cells(dataRow, 1).Value))
        If firstColVal = "" Then Exit For
        For j = 1 To headerCols
            wsTarget.Cells(outR, j).Value = Nz(rng.Cells(dataRow, j).Value, "")
        Next j
        outR = outR + 1
    Next dataRow

    Dim csvFullPath As String: csvFullPath = ThisWorkbook.Path & "\" & CONFIG_FOLDER & "\" & targetSheetName & ".csv"
    SaveSheetToCSV targetSheetName, csvFullPath

    currentInEditMode = False
    ui.Range(stagingRightTableRangeAddress).Locked = True
    ui.Protect UserInterfaceOnly:=True

    MsgBox "儲存成功，並已更新設定檔 CSV。", vbInformation
    LoadAllCSVToSheets
    RefreshLeftPanel
    RefreshRightPanel currentSelectedReportID, currentActiveTab
End Sub
' ********** MODIFIED END **********

Private Function Nz(v As Variant, alt As Variant) As Variant
    If IsError(v) Then Nz = alt: Exit Function
    If IsNull(v) Then Nz = alt: Exit Function
    If Trim(CStr(v & "")) = "" Then Nz = alt Else Nz = v
End Function


⸻

UI_Main（Worksheet）事件程式

在 UI_Main 的工作表代碼視窗貼上：

Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error Resume Next
    If Not Intersect(Target, Me.Range("A2:A1000")) Is Nothing Then
        Dim r As Range: Set r = Intersect(Target, Me.Range("A2:A1000")).Cells(1)
        Dim reportID As String: reportID = Trim(CStr(r.Value))
        If reportID <> "" Then
            currentSelectedReportID = reportID
            RefreshRightPanel reportID, currentActiveTab
        End If
    End If
End Sub


⸻

使用說明（版本 A）
	1.	將上面三個 Module 及 UI_Main 事件程式貼進 VBA。
	2.	在 Workbook 同目錄建立 config\ 及子資料夾（backup,locks 會自動建立）。
	3.	把你的 CSV 放在 config\（或先建立範例 CSV）。
	4.	在 UI_Main 放好按鈕並指定 macro：
	•	Nav 按鈕：Nav_ShowUpdateSheet / Nav_Sh  owExportPDF / Nav_ShowMappings
	•	Edit → EnterEditMode
	•	Save → SaveEdits
	•	Cancel → CancelEdits
	5.	執行 InitializeUI（或在 Workbook_Open 呼叫它）初始化。
	6.	使用流程：左側選報表 → 右側顯示相對子表（鎖定） → 按 Edit 進入編輯 → 修改右側 → Save/Cancel。

⸻

最後的建議與補充
	•	兩個版本共用的 ConfigIO / Helpers 已完整提供；差異僅在 UIHandlers。
	•	我已把「初始不能編輯」與「編輯時局部解鎖」邏輯都實作好（以 Protect + Locked 實現）。
	•	若你希望把 HeaderTimeSheetRange / HeaderTimeFormat 的自動填值（之前提的 FillHeaderDates）也串進來，我可以把 FillHeaderDates 加到 ProcessReport 呼叫點或在 SaveEdits 自動執行（你之前也提過，若需要我可再加）。
	•	若之後你決定要多人協作，我可以把簡單的 lock-file 機制再加回來（或改為資料庫鎖）。

⸻

要我現在幫你：
	•	把一組 示範 CSV（可直接貼入檔案） 跟 UI_Main 的按鈕位置（格子座標）與建議大小 一併給出？
	•	或直接把 InitializeUI 加到 Workbook_Open（自動啟動）？

回覆你要哪一個，我就把對應的小段補上（我會保持程式碼完整且可直接貼入）。







1.
之前已經開啟Excel檔，刷新資料並且儲存，
現在開啟檔案時，UI_Main分頁進入保護沒有更新資料，
其他分頁也沒有被設定保護模式

2.
保護模式的管理要修改

