Option Explicit

' -------------------------
' Module: CopyThenRunAP_Rewrite
' Revised: support multi-row UpdateSheet, parse FilterFields/Values,
' centralize BuildTimeStringFormats usage, resolve tokens, wildcard support,
' import csv with filters, and single-save-after-all-updates flow.
' -------------------------

'------------------------------
' Logging utilities
'------------------------------
Private Const LOG_FOLDER As String = "\logs"

Private Function LogFilePath() As String
    Dim base As String: base = ThisWorkbook.Path
    Dim fld As String: fld = base & LOG_FOLDER
    If Dir(fld, vbDirectory) = "" Then MkDirRecursive fld
    LogFilePath = fld & "\RunLog_" & Format(Date, "yyyymmdd") & ".txt"
End Function

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
    Dim path As String: path = LogFilePath()
    Open path For Append As #f
    Print #f, Format(Now, "yyyy-mm-dd HH:nn:ss") & " | " & level & " | " & msg
    Close #f
End Sub

'------------------------------
' Utilities
'------------------------------
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

'------------------------------
' Loaders
'------------------------------
Public Function LoadAllReportConfigs() As Object
    Dim dictAll As Object: Set dictAll = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet
    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Worksheets("tblReports")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        Dim reportID As String: reportID = Trim(ws.Cells(r, "A").Value)
        If reportID = "" Then GoTo NextR
        Dim cfg As Object: Set cfg = CreateObject("Scripting.Dictionary")
        cfg("ReportID") = reportID
        cfg("ReportType") = Trim(ws.Cells(r, "B").Value) ' "月報"/"季報"/"半年報"        
        cfg("TplPathPattern") = Trim(ws.Cells(r, "C").Value)

        cfg("TplPathTimeFormat") = Trim(ws.Cells(r, "D").Value) ' e.g. "AD_YYYYMM" or "ROC_YYYMM"

        cfg("DeclPathPattern") = Trim(ws.Cells(r, "E").Value)
        cfg("DeclPathTimeFormat") = Trim(ws.Cells(r, "F").Value)
        Dim headers As String: headers = Trim(ws.Cells(r, "G").Value)
        If headers = "" Then
            cfg("HeaderTimeSheetRange") = Array()
        Else
            cfg("HeaderTimeSheetRange") = Split(Replace(headers, ";", "|"), "|")
        End If
        cfg("HeaderTimeFormat") = Trim(ws.Cells(r, "H").Value)
        cfg("IsToCreateDeclFile") = (UCase(Trim(ws.Cells(r, "I").Value)) = "TRUE")
        cfg("IsDeleteTplPattern") = (UCase(Trim(ws.Cells(r, "J").Value)) = "TRUE")
        cfg("ProcessingMacro") = Trim(ws.Cells(r, "K").Value)
        cfg("PDFParentFolder") = Trim(ws.Cells(r, "L").Value)

        ' Load UpdateRows aggregated by reportID
        Set cfg("UpdateSourceData") = LoadUpdateSourceData(reportID)
        ' Load per-report ExportPDF rows
        Set cfg("ExportPDFList") = LoadExportPDFForReport(reportID)
        Set dictAll(reportID) = cfg
NextR:
    Next r
    Set LoadAllReportConfigs = dictAll
    Exit Function
ErrHandler:
    LogError "LoadAllReportConfigs error: " & Err.Number & " " & Err.Description
    Set LoadAllReportConfigs = dictAll
End Function

' ---------- LoadUpdateSourceData: aggregate multiple rows per ReportID into dictionary of arrays ----------
Public Function LoadUpdateSourceData(reportID As String) As Object
    Dim result As Object: Set result = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet
    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Worksheets("tblUpdateSheet")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim tempUpdateSheet() As String
    Dim tempClearRange() As String
    Dim tempImportFilePath() As String
    Dim tempImportTimeFormat() As String
    Dim tempImportSheet() As String
    Dim tempImportProcessType() As String
    Dim tempPasteTopLeftRange() As String    
    Dim tempFilterField() As String

    Dim cnt As Long: cnt = 0
    Dim r As Long
    For r = 2 To lastRow
        If Trim(ws.Cells(r, "A").Value) = reportID Then
            ReDim Preserve tempUpdateSheet(0 To cnt)
            ReDim Preserve tempClearRange(0 To cnt)
            ReDim Preserve tempImportFilePath(0 To cnt)
            ReDim Preserve tempImportTimeFormat(0 To cnt)
            ReDim Preserve tempImportSheet(0 To cnt)
            ReDim Preserve tempImportProcessType(0 To cnt)            
            ReDim Preserve tempPasteTopLeftRange(0 To cnt)
            ReDim Preserve tempFilterField(0 To cnt)

            tempUpdateSheet(cnt) = Trim(ws.Cells(r, "B").Value)
            tempClearRange(cnt) = Trim(ws.Cells(r, "C").Value)
            tempImportFilePath(cnt) = Trim(ws.Cells(r, "D").Value)
            tempImportTimeFormat(cnt) = Trim(ws.Cells(r, "E").Value)
            tempImportSheet(cnt) = Trim(ws.Cells(r, "F").Value)
            tempImportProcessType(cnt) = Trim(ws.Cells(r, "G").Value)
            tempPasteTopLeftRange(cnt) = Trim(ws.Cells(r, "H").Value)
            tempFilterField(cnt) = Trim(ws.Cells(r, "I").Value)

            cnt = cnt + 1
        End If
    Next r

    If cnt = 0 Then
        Dim a0() As String
        result("Count") = 0
        result("UpdateSheetArr") = a0
        result("ClearRangeArr") = a0
        result("ImportPathPatternsArr") = a0
        result("ImportPathTimeFormatArr") = a0
        result("ImportSheetsArr") = a0
        result("ImportProcessTypesArr") = a0
        result("PasteTopLeftRangeArr") = a0
        result("FilterFieldsAndValuesRawArr") = a0
        Set LoadUpdateSourceData = result
        Exit Function
    End If

    result("Count") = cnt
    result("UpdateSheetArr") = tempUpdateSheet
    result("ClearRangeArr") = tempClearRange
    result("ImportPathPatternsArr") = tempImportFilePath
    result("ImportPathTimeFormatArr") = tempImportTimeFormat
    result("ImportSheetsArr") = tempImportSheet
    result("ImportProcessTypesArr") = tempImportProcessType
    result("PasteTopLeftRangeArr") = tempPasteTopLeftRange
    result("FilterFieldsAndValuesRawArr") = tempFilterField

    Set LoadUpdateSourceData = result
    Exit Function
ErrHandler:
    LogError "LoadUpdateSourceData error: " & Err.Number & " " & Err.Description
    Set LoadUpdateSourceData = result
End Function

Public Function LoadExportPDFForReport(reportID As String) As Object
    Dim result As Object: Set result = CreateObject("Scripting.Dictionary")
    On Error GoTo ErrHandler
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("tblExportPDF")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim tempPDFSheet() As String
    Dim tempParentFolder() As String
    Dim cnt As Long: cnt = 0
    Dim r As Long
    For r = 2 To lastRow
        If Trim(ws.Cells(r, "A").Value) = reportID Then
            ReDim Preserve tempPDFSheet(0 To cnt)
            ReDim Preserve tempParentFolder(0 To cnt)
            tempPDFSheet(cnt) = Trim(ws.Cells(r, "B").Value)
            tempParentFolder(cnt) = Trim(ws.Cells(r, "C").Value)
            cnt = cnt + 1
        End If
    Next r
    If cnt = 0 Then
        Dim a0() As String
        result("Count") = 0
        result("PDFSheetsArr") = a0
        result("ParentFolderArr") = a0
        Set LoadExportPDFForReport = result
        Exit Function
    End If
    result("Count") = cnt
    result("PDFSheetsArr") = tempPDFSheet
    result("ParentFolderArr") = tempParentFolder
    Set LoadExportPDFForReport = result
    Exit Function
ErrHandler:
    LogError "LoadExportPDFForReport error: " & Err.Number & " " & Err.Description
    Set LoadExportPDFForReport = result
End Function

'------------------------------
' ParseFilterFieldsAndValues:
' Input raw string like:
'   交易別:首購;續發;買斷|票類:CP1;CP2;TA
' Output: Array(filterFieldsArray, filterValuesArray) where filterValuesArray(i) is array of values (strings)
'------------------------------

Public Function ParseFilterFieldsAndValues(filterField As String) As Variant
    filterField = Trim(filterField & "")
    If filterField = "" Then
        Dim emptyFields(0 To -1) As String
        Dim emptyVals(0 To -1) As Variant
        ParseFilterFieldsAndValues = Array(emptyFields, emptyVals)
        Exit Function
    End If

    Dim parts() As String: parts = Split(filterField, "|")
    Dim fieldList() As String
    Dim valuesList() As Variant
    ReDim fieldList(0 To UBound(parts))
    ReDim valuesList(0 To UBound(parts))

    Dim i As Long
    For i = 0 To UBound(parts)
        Dim token As String: token = Trim(parts(i))
        If token = "" Then
            fieldList(i) = ""
            Dim a0() As String
            valuesList(i) = a0
        Else
            If InStr(token, ":") = 0 Then
                ' malformed - treat as field with no values
                fieldList(i) = token
                Dim a1(0 To -1) As String
                valuesList(i) = a1
            Else
                Dim kv() As String: kv = Split(token, ":", 2)
                fieldList(i) = Trim(kv(0))
                Dim vals() As String: vals = Split(kv(1), ";")
                Dim j As Long
                For j = 0 To UBound(vals)
                    vals(j) = Trim(vals(j))
                Next j
                valuesList(i) = vals
            End If
        End If
    Next i

    ParseFilterFieldsAndValues = Array(fieldList, valuesList)
End Function

'------------------------------
' ResolveImportPatternsWithMF
' - pattern: may contain tokens like YYYYMM, OLDYYYYMM, WESTERN_END, WEST_WORKDAY_END, ROC_WORKDAY_END
' - formatKey: e.g. "AD_YYYYMM", "ROC_YYYMM", used as suffix to NEW_/OLD_ keys in mf
' - mf: dictionary returned by BuildTimeStringFormats (expects keys like "NEW_AD_YYYYMM", "OLD_AD_YYYYMM", etc)
' - basePath: ThisWorkbook.Path used for wildcard search
' Returns array of relative paths (possibly empty strings)
'------------------------------
Public Function ResolveImportPatternsWithMF(pattern As String, _
                                            formatKey As String, _
                                            mf As Object, _
                                            basePath As String) As Variant
    Dim results() As String
    Dim outIdx As Long: outIdx = -1
    If Trim(pattern) = "" Then
        ReDim results(0 To 0): results(0) = "": ResolveImportPatternsWithMF = results: Exit Function
    End If

    ' >>> MODIFIED START <<< 
    ' 建立 FileSystemObject 用來轉換成絕對路徑
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 如果 basePath 是空的，自動預設為 ThisWorkbook.Path
    If Trim(basePath) = "" Then basePath = ThisWorkbook.Path
    ' >>> MODIFIED END <<<

    ' Allow multiple patterns separated by comma in the cell (user may input multiple)
    Dim patternParts() As String: patternParts = Split(pattern, ",")
    Dim pIdx As Long
    For pIdx = 0 To UBound(patternParts)
        Dim replaced As String: replaced = Trim(patternParts(pIdx))
        If replaced = "" Then
            outIdx = outIdx + 1: ReDim Preserve results(0 To outIdx): results(outIdx) = "": GoTo ContinueLoop2
        End If

        ' First replace tokens using formatKey if available
        On Error Resume Next
        If formatKey <> "" Then
            If mf.Exists(formatKey) Then replaced = Replace(replaced, "YYYYMM", mf(formatKey))
        End If
        On Error GoTo 0

        ' Handle wildcard: use Dir to find first matching file in basePath\replaced pattern
        If InStr(replaced, "*") > 0 Then
            Dim f As String
            ' 先取出資料夾部分（相對路徑的資料夾）
            Dim folderPart As String, filePart As String
            folderPart = Left(replaced, InStrRev(replaced, "\") - 1) ' "../批次"
            filePart = Mid(replaced, InStrRev(replaced, "\") + 1)    ' "*cm2610*"
            
            ' 用 Dir 搜尋符合萬用字元的檔案
            f = Dir(ThisWorkbook.Path & "\" & folderPart & "\" & filePart)
            
            If f <> "" Then
                outIdx = outIdx + 1: ReDim Preserve results(0 To outIdx)
                ' 保留資料夾，使用 GetAbsolutePathFromProject 拼完整路徑
                results(outIdx) = GetAbsolutePathFromProject(folderPart & "\" & f)
            Else
                LogWarn "Wildcard no match for pattern: " & replaced
                outIdx = outIdx + 1: ReDim Preserve results(0 To outIdx)
                results(outIdx) = "" ' keep slot
            End If
        Else
            ' >>> MODIFIED START <<< 
            ' 即使不是 wildcard，也要轉成絕對路徑
            Dim absPath As String
            absPath = GetAbsolutePathFromProject(replaced)
            
            outIdx = outIdx + 1: ReDim Preserve results(0 To outIdx)
            results(outIdx) = absPath
            ' >>> MODIFIED END <<< 
        End If
ContinueLoop2:
    Next pIdx

    ResolveImportPatternsWithMF = results
End Function

Function GetAbsolutePathFromProject(ByVal relPath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 如果傳入的是絕對路徑，直接回傳
    If fso.DriveExists(fso.GetDriveName(relPath)) Then
        GetAbsolutePathFromProject = fso.GetAbsolutePathName(relPath)
    Else
        ' 如果是相對路徑，就用 ThisWorkbook.Path 當基準
        GetAbsolutePathFromProject = fso.GetAbsolutePathName(ThisWorkbook.Path & "\" & relPath)
    End If
End Function



'------------------------------
' Top-level runner
'------------------------------
Public Sub ProcessAllReports()
    Dim allCfgs As Object: Set allCfgs = LoadAllReportConfigs()
    Dim key
    If allCfgs.Count = 0 Then
        LogWarn "No reports found in tblReports."
        Exit Sub
    End If

    ' Get YearMonth user-provided name
    Dim ymROC As String
    On Error GoTo ErrYM
    ymROC = ThisWorkbook.Names("YearMonth").RefersToRange.Value
    On Error GoTo 0

    Dim mf As Object: Set mf = BuildTimeStringFormats(ymROC)

    For Each key In allCfgs.Keys
        On Error GoTo ErrContinue
        Dim cfg As Object: Set cfg = allCfgs(key)
        LogInfo "Start report: " & cfg("ReportID")
        ProcessReport cfg, mf
        LogInfo "Finish report: " & cfg("ReportID")
ErrContinue:
        If Err.Number <> 0 Then
            LogError "ProcessAllReports_New encountered error for " & key & ": " & Err.Number & " " & Err.Description
            Err.Clear
        End If
    Next key

    MsgBox "全部報表處理完成！", vbInformation
    Exit Sub

ErrYM:
    MsgBox "找不到 Name 'YearMonth' 或其值格式錯誤。請先確認 ThisWorkbook.Names(""YearMonth"") 有正確的值 (例如: 114/06)。", vbExclamation
    Exit Sub
End Sub

'------------------------------
' Process single report (uses aggregated LoadUpdateSourceData arrays)
'------------------------------
Public Sub ProcessReport(cfg As Object, _
                         mf As Object)
    On Error GoTo ErrHandler
    Dim basePath As String: basePath = ThisWorkbook.Path

    ' Check report type (month/quarter/half) to decide whether to process
    Dim reportType As String: reportType = Trim(cfg("ReportType"))
    Dim newROC_MM As String
    If mf.Exists("NEW_ROC_MM") Then
      newROC_MM = mf("NEW_ROC_MM")
    Else
      ' ************LogWarn
      Exit Sub
    End If

    LogInfo "newROC_MM: " & newROC_MM

    Select Case reportType
        Case "季報"
            If Not (newROC_MM = "03" Or newROC_MM = "06" Or newROC_MM = "09" Or newROC_MM = "12") Then
                LogInfo "Skip seasonal report (季報) for month: " & newROC_MM & " ReportID=" & cfg("ReportID")
                Exit Sub
            End If
        Case "半年報"
            If Not (newROC_MM = "06" Or newROC_MM = "12") Then
                LogInfo "Skip 半年報 for month: " & newROC_MM & " ReportID=" & cfg("ReportID")
                Exit Sub
            End If
        Case "月報"
            ' always run
        Case "月報_COPY"
            ' ********** Logwarn
            oldTplPathRel = Replace(cfg("TplPathPattern"), "YYYYMM", mf("NEW_" & tplTimeFormatKey))

            newDeclPathRel = Replace(cfg("DeclPathPattern"), "YYYYMM", mf("NEW_" & declTimeFormatKey))

            ' copy oldTplPathRel to newDeclPathRel
            Exit Sub
        Case Else
            ' unknown -> run but warn
            LogWarn "未知報表類型 (will attempt to run): " & reportType & " ReportID=" & cfg("ReportID")
    End Select

    Dim updateData As Object: Set updateData = cfg("UpdateSourceData")
    Dim updateCount As Long: updateCount = 0
    If updateData.Exists("Count") Then updateCount = CLng(updateData("Count"))
    LogInfo "updateCount: " & updateCount

    ' Resolve old/new tpl and decl path using cfg("TplPathTimeFormat") and cfg("DeclPathTimeFormat")
    Dim tplTimeFormatKey As String: tplTimeFormatKey = Trim(cfg("TplPathTimeFormat"))
    Dim declTimeFormatKey As String: declTimeFormatKey = Trim(cfg("DeclPathTimeFormat"))

    Dim oldTplPathRel As String, newTplPathRel As String
    Dim oldDeclPathRel As String, newDeclPathRel As String

    On Error Resume Next
    If tplTimeFormatKey <> "" And mf.Exists("OLD_" & tplTimeFormatKey) Then
        oldTplPathRel = Replace(cfg("TplPathPattern"), "YYYYMM", mf("OLD_" & tplTimeFormatKey))
    Else
        oldTplPathRel = cfg("TplPathPattern")
        ' default: try OLD_AD_YYYYMM
        ' If mf.Exists("OLD_AD_YYYYMM") Then oldTplPathRel = Replace(cfg("TplPathPattern"), "YYYYMM", mf("OLD_AD_YYYYMM")) Else oldTplPathRel = Replace(cfg("TplPathPattern"), "YYYYMM", "")
        ' ***** 需要修改
        ' Log紀錄這邊有問題，直接Exit Sub
    End If
    LogInfo "oldTplPathRel: " & oldTplPathRel

    If tplTimeFormatKey <> "" And mf.Exists("NEW_" & tplTimeFormatKey) Then
        newTplPathRel = Replace(cfg("TplPathPattern"), "YYYYMM", mf("NEW_" & tplTimeFormatKey))
    Else
        newTplPathRel = cfg("TplPathPattern")
        ' If mf.Exists("NEW_AD_YYYYMM") Then newTplPathRel = Replace(cfg("TplPathPattern"), "YYYYMM", mf("NEW_AD_YYYYMM")) Else newTplPathRel = Replace(cfg("TplPathPattern"), "YYYYMM", "")
        ' ***** 需要修改
        ' Log紀錄這邊有問題，直接Exit Sub        
    End If
    LogInfo "newTplPathRel: " & newTplPathRel

    If declTimeFormatKey <> "" And mf.Exists("OLD_" & declTimeFormatKey) Then
        oldDeclPathRel = Replace(cfg("DeclPathPattern"), "YYYYMM", mf("OLD_" & declTimeFormatKey))
    Else
        oldDeclPathRel = cfg("DeclPathPattern")
        ' If mf.Exists("OLD_AD_YYYYMM") Then oldDeclPathRel = Replace(cfg("DeclPathPattern"), "YYYYMM", mf("OLD_AD_YYYYMM")) Else oldDeclPathRel = Replace(cfg("DeclPathPattern"), "YYYYMM", "")
        ' ***** 需要修改
        ' Log紀錄這邊有問題，直接Exit Sub        
    End If
    LogInfo "oldDeclPathRel: " & oldDeclPathRel

    If declTimeFormatKey <> "" And mf.Exists("NEW_" & declTimeFormatKey) Then
        newDeclPathRel = Replace(cfg("DeclPathPattern"), "YYYYMM", mf("NEW_" & declTimeFormatKey))
    Else
        newDeclPathRel = cfg("DeclPathPattern")
        ' If mf.Exists("NEW_AD_YYYYMM") Then newDeclPathRel = Replace(cfg("DeclPathPattern"), "YYYYMM", mf("NEW_AD_YYYYMM")) Else newDeclPathRel = Replace(cfg("DeclPathPattern"), "YYYYMM", "")
        ' ***** 需要修改
        ' Log紀錄這邊有問題，直接Exit Sub  
    End If
    LogInfo "newDeclPathRel: " & newDeclPathRel

    On Error GoTo ErrHandler

    ' Open old template once (read-only), process all update rows on it
    Dim wbOld As Workbook
    If Dir(basePath & "\" & oldTplPathRel) = "" Then
        LogWarn "Tpl not found: " & basePath & "\" & oldTplPathRel & " (ReportID=" & cfg("ReportID") & ")"
        Exit Sub
    End If

    Set wbOld = Workbooks.Open(basePath & "\" & oldTplPathRel, ReadOnly:=True)

    ' Fill header dates if any
    On Error Resume Next
    If cfg.Exists("HeaderTimeSheetRange") Then
        Dim headerFormat As String
        headerFormat = ""
        If cfg.Exists("HeaderTimeFormat") Then headerFormat = Trim(cfg("HeaderTimeFormat"))
        FillHeaderDates wbOld, cfg("HeaderTimeSheetRange"), mf, headerFormat
        LogInfo "headerFormat: " & headerFormat
    Else
        ' ***********LogWarn
        Exit Sub
    End If
    On Error GoTo ErrHandler


    ' ------------------ 替換區段：開始 ------------------
    ' Loop each update row
    Dim idx As Long
    For idx = 0 To updateCount - 1
        Dim targetSheet As String: targetSheet = updateData("UpdateSheetArr")(idx)
        Dim clearRange As String: clearRange = updateData("ClearRangeArr")(idx)
        Dim importPathRaw As String: importPathRaw = updateData("ImportPathPatternsArr")(idx)
        Dim importTimeFormatKey As String: importTimeFormatKey = updateData("ImportPathTimeFormatArr")(idx)
        Dim importSheetName As String: importSheetName = updateData("ImportSheetsArr")(idx) ' 單一名稱（不拆成多個）
        Dim importProcessType As String: importProcessType = updateData("ImportProcessTypesArr")(idx) ' 單一類型
        Dim PasteTopLeftRange As String: PasteTopLeftRange = updateData("PasteTopLeftRangeArr")(idx)
        Dim filterRaw As String: filterRaw = updateData("FilterFieldsAndValuesRawArr")(idx)

        ' Resolve import patterns -> array of resolved relative paths (handles wildcard)
        Dim resolvedImportPaths As Variant
        resolvedImportPaths = ResolveImportPatternsWithMF(importPathRaw, importTimeFormatKey, mf, basePath)

        ' Clear target range (if specified)
        On Error Resume Next
        If clearRange <> "" Then
            Select Case UCase(clearRange)
                Case "CELLS"
                    wbOld.Sheets(targetSheet).Cells.ClearContents
                Case Else
                    wbOld.Sheets(targetSheet).Range(clearRange).ClearContents
            End Select
        End If
        On Error GoTo ErrHandler
        ' CELLS	特殊關鍵字，表示整張清空	wbOld.Sheets(targetSheet).Cells.ClearContents
        ' A:AD	欄範圍	wbOld.Sheets(targetSheet).Range("A:AD").ClearContents
        ' A1:Z30	明確範圍	wbOld.Sheets(targetSheet).Range("A1:Z30").ClearContents
        ' A1	單一儲存格	wbOld.Sheets(targetSheet).Range("A1").ClearContents
        ' A1,B5,C10	多個離散範圍	wbOld.Sheets(targetSheet).Range("A1,B5,C10").ClearContents

        ' For each resolved import path, perform import
        Dim pIndex As Long
        For pIndex = LBound(resolvedImportPaths) To UBound(resolvedImportPaths)
            Dim impFullPath As String: impFullPath = resolvedImportPaths(pIndex)
            If Trim(impFullPath) = "" Then
                LogWarn "Resolved import path empty for pattern '" & importPathRaw & "' (ReportID=" & cfg("ReportID") & ")"
                GoTo NextResolved
            End If

            Dim importSheetName As String: importSheetName = ""
            Dim importType As String: importType = ""

            ' 使用同一列的 sheet 名稱與匯入類型（若為空則使用預設）
            If Trim(importSheetName) <> "" Then importSheetName = Trim(importSheetName)
            If Trim(importProcessType) <> "" Then importType = Trim(importProcessType)

            Dim ok As Boolean
            ok = ImportDataToTpl(impFullPath, wbOld, targetSheet, importSheetName, importType, PasteTopLeftRange, filterRaw)
            If Not ok Then
                LogWarn "Import failed for " & impFullPath & " into " & cfg("ReportID") & "!" & targetSheet
            Else
                LogInfo "Import success: " & impFullPath & " -> " & cfg("ReportID") & "!" & targetSheet
            End If

    NextResolved:
        Next pIndex
    Next idx
    ' ------------------ 替換區段：結束 ------------------

    ' After all updates completed on wbOld, run processing macro if configured
    ' If cfg.Exists("ProcessingMacro") Then
    '     If Trim(cfg("ProcessingMacro")) <> "" Then
    '         RunProcessingMacro cfg("ProcessingMacro"), wbOld
    '     End If
    ' End If

    If cfg.Exists("ProcessingMacro") Then
        If Trim(cfg("ProcessingMacro")) <> "" Then
            Dim macroName As String
            macroName = Trim(cfg("ProcessingMacro"))
            Dim reportID As String
            reportID = cfg("ReportID")

            Select Case reportID
                Case "FM11"
                    ' wbOld, True
                    RunProcessingMacroWithArgs macroName, wbOld, True

                Case "表41"
                    ' wbOld, True, CDate(<NEW_AD_YYYYMMDD_END>)
                    If mf.Exists("NEW_AD_YYYYMMDD_END") Then
                        Dim dtEnd As Date
                        dtEnd = CDate(DateSerial( _
                            CLng(Left(mf("NEW_AD_YYYYMMDD_END"), 4)), _
                            CLng(Mid(mf("NEW_AD_YYYYMMDD_END"), 5, 2)), _
                            CLng(Right(mf("NEW_AD_YYYYMMDD_END"), 2)) ))
                        RunProcessingMacroWithArgs macroName, wbOld, Array(True, dtEnd)
                    Else
                        ' fallback / warning
                        LogWarn "Missing NEW_AD_YYYYMMDD_END for ReportID=" & reportID
                        RunProcessingMacroWithArgs macroName, wbOld, True
                    End If

                Case "FM2"
                    ' wbOld (只有 wbOld)
                    RunProcessingMacroWithArgs macroName, wbOld

                Case "FM10"
                    ' wbOld
                    RunProcessingMacroWithArgs macroName, wbOld

                Case "F1_F2"
                    ' wbOld, True, mf("ROCYYYMM") (字串)
                    If mf.Exists("ROCYYYMM") Then
                        RunProcessingMacroWithArgs macroName, wbOld, Array(True, mf("ROCYYYMM"))
                    Else
                        LogWarn "Missing ROCYYYMM for ReportID=" & reportID
                        RunProcessingMacroWithArgs macroName, wbOld, True
                    End If

                Case "AI240"
                    ' wbOld, True, CDate(<NEW_AD_YYYYMMDD_WORKDAY_END>)
                    If mf.Exists("NEW_AD_YYYYMMDD_WORKDAY_END") Then
                        Dim dtWorkday As Date
                        dtWorkday = CDate(DateSerial( _
                            CLng(Left(mf("NEW_AD_YYYYMMDD_WORKDAY_END"), 4)), _
                            CLng(Mid(mf("NEW_AD_YYYYMMDD_WORKDAY_END"), 5, 2)), _
                            CLng(Right(mf("NEW_AD_YYYYMMDD_WORKDAY_END"), 2)) ))
                        RunProcessingMacroWithArgs macroName, wbOld, Array(True, dtWorkday)
                    Else
                        LogWarn "Missing NEW_AD_YYYYMMDD_WORKDAY_END for ReportID=" & reportID
                        RunProcessingMacroWithArgs macroName, wbOld, True
                    End If

                Case Else
                    ' 預設：只傳 wbOld（如果你要改成預設傳 wbOld, True 可改這裡）
                    RunProcessingMacroWithArgs macroName, wbOld
            End Select
        End If
    End If




    ' SaveCopyAs new template
    On Error Resume Next
    Dim newFullPath As String: newFullPath = basePath & "\" & newTplPathRel
    wbOld.SaveCopyAs newFullPath
    If Err.Number <> 0 Then
        LogError "SaveCopyAs failed for " & newFullPath & ": " & Err.Number & " " & Err.Description
        Err.Clear
    End If
    wbOld.Close SaveChanges:=False
    On Error GoTo ErrHandler

    ' Open new and apply .pings & export PDFs
    Dim wbNew As Workbook
    If Dir(newFullPath) = "" Then
        LogWarn "New template not found after SaveCopyAs: " & newFullPath
        Exit Sub
    End If
    Set wbNew = Workbooks.Open(newFullPath)

    If cfg.Exists("IsToCreateDeclFile") Then
        If cfg("IsToCreateDeclFile") Then
            If Dir(basePath & "\" & oldDeclPathRel) <> "" Then
                Dim wbDecl As Workbook: Set wbDecl = Workbooks.Open(basePath & "\" & oldDeclPathRel)
                ApplyMappings wbNew, wbDecl, cfg("ReportID")
                wbDecl.SaveCopyAs basePath & "\" & newDeclPathRel
                wbDecl.Close SaveChanges:=False
            Else
                LogWarn "Decl template not found: " & basePath & "\" & oldDeclPathRel
            End If            
        End If
    End If        

    ' Determine pdfSheets and parentFolder - consult ExportPDFList if present
    Dim savePdfSheets() As String
    Dim ParentFolder As String: ParentFolder = cfg("PDFParentFolder")
    If cfg.Exists("ExportPDFList") Then
        Dim exportList As Object: Set exportList = cfg("ExportPDFList")
        If exportList.Exists("Count") Then
            If exportList("Count") > 0 Then

                ' 這邊要寫一個 for 迴圈去處理，不能這樣處理

                ' Use first row's PDFSheets as default for now (you may adapt to multiple)
                Dim arrPDFs() As String
                arrPDFs = Split(exportList("PDFSheetsArr")(0), ",")
                savePdfSheets = arrPDFs
                If Trim(exportList("ParentFolderArr")(0)) <> "" Then ParentFolder = exportList("ParentFolderArr")(0)
            Else
                ' fallback to cfg("PDFParentFolder") and cfg("PDFSheets") if you have it
                ReDim savePdfSheets(0 To -1)
            End If
        End If
    Else
        ReDim savePdfSheets(0 To -1)
    End If

    ExportPDFs basePath, mf("NEW_AD_YYYYMM"), cfg("ReportID"), wbNew, savePdfSheets, ParentFolder

    wbNew.Close SaveChanges:=False

    ' Optionally delete original source files? cfg may have flag like IsDeleteTplPattern but that was for template deletion - be careful
    ' If you want to delete imported source files, implement here (cfg flag expected)

    Exit Sub

ErrHandler:
    LogError "ProcessReport error for " & cfg("ReportID") & ": " & Err.Number & " " & Err.Description
    On Error Resume Next
    If Not wbOld Is Nothing Then wbOld.Close SaveChanges:=False
    Resume Next
End Sub

'------------------------------
' ImportDataToTpl (unchanged)
'------------------------------
Public Function ImportDataToTpl(impFullPath As String, _
                                wbOld As Workbook, _
                                targetSheet As String, _
                                importSheet As String, _
                                importType As String, _
                                pasteRange As String, _
                                filterRaw As String) As Boolean
    On Error GoTo ErrHandler
    ImportDataToTpl = False
    Select Case Trim(importType)
        Case "Default_Copy"
            If Dir(impFullPath) = "" Then
                LogWarn "File not found for copy: " & impFullPath
                Exit Function
            End If
            Dim wbImp As Workbook: Set wbImp = Workbooks.Open(impFullPath, ReadOnly:=True)
            If importSheet = "" Then importSheet = wbImp.Sheets(1).Name
            On Error Resume Next
            wbImp.Sheets(importSheet).UsedRange.Copy
            wbOld.Sheets(targetSheet).Range(pasteRange).PasteSpecial xlPasteValues
            Application.CutCopyMode = False
            wbImp.Close SaveChanges:=False
            ImportDataToTpl = True
            Exit Function
                
        Case "CSV_Filter"
            If Dir(impFullPath) = "" Then
                LogWarn "CSV not found: " & impFullPath
                Exit Function
            End If

            ' parse filter
            Dim parseString As Variant
            parseString = ParseFilterFieldsAndValues(filterRaw)
            Dim filterFields() As String: filterFields = parseString(0)
            Dim filterValues As Variant: filterValues = parseString(1)            

            ImportCsvWithFilter impFullPath, wbOld.Sheets(targetSheet), wbOld.Sheets(targetSheet).Range(pasteRange), filterFields, filterValues
            ImportDataToTpl = True
            Exit Function

        Case "PNCDCAL"
            If Dir(impFullPath) = "" Then
                LogWarn "CloseRate file not found: " & impFullPath
                Exit Function
            End If

            Call PNCDCAL_FormatToCSV(impFullPath)
            Dim wbImp As Workbook: Set wbImp = Workbooks.Open(impFullPath, ReadOnly:=True)
            If importSheet = "" Then importSheet = wbImp.Sheets(1).Name
            On Error Resume Next
            wbImp.Sheets(importSheet).UsedRange.Copy
            wbOld.Sheets(targetSheet).Range(pasteRange).PasteSpecial xlPasteValues
            Application.CutCopyMode = False
            wbImp.Close SaveChanges:=False
            ImportDataToTpl = True
            Exit Function

        Case "CloseRate:USD"
            If Dir(impFullPath) = "" Then
                LogWarn "CloseRate file not found: " & impFullPath
                Exit Function
            End If

            Call ImportCloseRate(impFullPath, wbOld.Sheets(targetSheet), "USD", pasteRange)
            ImportDataToTpl = True
            Exit Function

        Case "CloseRate:TWD"
            If Dir(impFullPath) = "" Then
                LogWarn "CloseRate file not found: " & impFullPath
                Exit Function
            End If

            Call ImportCloseRate(impFullPath, wbOld.Sheets(targetSheet), "TWD", pasteRange)
            ImportDataToTpl = True
            Exit Function                

        Case "外幣債評估"
            If Dir(impFullPath) = "" Then
                LogWarn "CloseRate file not found: " & impFullPath
                Exit Function
            End If

            Call ImportFXDebtEvaluation(impFullPath)
            Dim wbImp As Workbook: Set wbImp = Workbooks.Open(impFullPath, ReadOnly:=True)
            If importSheet = "" Then importSheet = wbImp.Sheets(1).Name
            On Error Resume Next
            wbImp.Sheets(importSheet).UsedRange.Copy
            wbOld.Sheets(targetSheet).Range(pasteRange).PasteSpecial xlPasteValues
            Application.CutCopyMode = False
            wbImp.Close SaveChanges:=False
            ImportDataToTpl = True
            Exit Function                

        Case ""
            Exit Function
    End Select
    Exit Function
ErrHandler:
    LogError "ImportDataToTpl error: " & Err.Number & " " & Err.Description & " file=" & impFullPath
    ImportDataToTpl = False
End Function

'------------------------------
' RunProcessingMacro (unchanged)
'------------------------------

' *********** 修改到這邊 ***********

Public Sub RunProcessingMacro(macroSpec As String, _
                              wbOld As Workbook)
    On Error GoTo ErrHandler
    If Trim(macroSpec) = "" Then Exit Sub
    Dim parts() As String: parts = Split(macroSpec, "|")
    Dim macroName As String: macroName = parts(0)
    Dim params() As String
    If UBound(parts) >= 1 Then params = Split(parts(1), ",")
    Dim args() As Variant
    Dim argCount As Long: argCount = 0
    ReDim args(0 To 0)
    args(0) = wbOld
    If UBound(parts) >= 1 Then
        Dim i As Long
        For i = LBound(params) To UBound(params)
            argCount = argCount + 1
            ReDim Preserve args(0 To argCount)
            args(argCount) = params(i)
        Next i
    End If
    Select Case UBound(args)
        Case 0: Application.Run macroName, args(0)
        Case 1: Application.Run macroName, args(0), args(1)
        Case 2: Application.Run macroName, args(0), args(1), args(2)
        Case 3: Application.Run macroName, args(0), args(1), args(2), args(3)
        Case Else: Application.Run macroName, args(0)
    End Select
    Exit Sub
ErrHandler:
    LogError "RunProcessingMacro error: " & Err.Number & " " & Err.Description & " macro=" & macroSpec
End Sub


' --- Helper: 呼叫 macro 並支援可變參數
Public Sub RunProcessingMacroWithArgs(macroName As String, wbOld As Workbook, Optional extraArgs As Variant)
    On Error GoTo ErrHandler

    If Trim(macroName) = "" Then Exit Sub

    ' 確保以 workbook 為前綴，避免 Application.Run 找不到正確的 Macro
    Dim fullMacroName As String
    fullMacroName = "'" & wbOld.Name & "'!" & macroName

    ' 構建參數陣列：第一個參數一律是 wbOld
    Dim args() As Variant
    Dim argCount As Long
    argCount = 0
    ReDim args(0 To 0)
    args(0) = wbOld

    ' 如果有 extraArgs（可能是單一值，也可能是陣列），把它們加入 args
    If Not IsEmpty(extraArgs) Then
        Dim i As Long
        If IsArray(extraArgs) Then
            For i = LBound(extraArgs) To UBound(extraArgs)
                argCount = argCount + 1
                ReDim Preserve args(0 To argCount)
                args(argCount) = extraArgs(i)
            Next i
        Else
            argCount = argCount + 1
            ReDim Preserve args(0 To argCount)
            args(argCount) = extraArgs
        End If
    End If

    ' 依參數數量用 Application.Run 呼叫（避免傳入陣列直接當參數失敗）
    Select Case UBound(args)
        Case 0: Application.Run fullMacroName, args(0)
        Case 1: Application.Run fullMacroName, args(0), args(1)
        Case 2: Application.Run fullMacroName, args(0), args(1), args(2)
        Case 3: Application.Run fullMacroName, args(0), args(1), args(2), args(3)
        Case 4: Application.Run fullMacroName, args(0), args(1), args(2), args(3), args(4)
        Case Else
            ' 若有更多參數，可再擴充上面 Case 或改為 CallByName 等複雜方法
            Application.Run fullMacroName, args(0)
    End Select

    Exit Sub
ErrHandler:
    LogError "RunProcessingMacroWithArgs error: " & Err.Number & " " & Err.Description & " macro=" & macroName
End Sub



'------------------------------
' ApplyMappings (unchanged)
'------------------------------
Public Sub ApplyMappings(wbNew As Workbook, _
                         wbDecl As Workbook, _
                         reportID As String)
    On Error GoTo ErrHandler
    Dim wsMap As Worksheet: Set wsMap = ThisWorkbook.Worksheets("Mappings")
    Dim lastRow As Long: lastRow = wsMap.Cells(wsMap.Rows.Count, "A").End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        If Trim(wsMap.Cells(r, "A").Value) = reportID Then
            Dim srcSh As String: srcSh = Trim(wsMap.Cells(r, "B").Value)
            Dim rngStrings As String(): rngStrings = Split(wsMap.Cells(r, "C").Value, ",")

            For Each srcAddr In rngStrings
                srcAddr = Trim(srcAddr)
                Set rngSrc = wbNew.Sheets(srcSh).Range(srcAddr)
                Set rngDst = wbDecl.Sheets(srcSh).Range(srcAddr)

                ' 直接以值貼值，保留大小一致
                rngDst.Value = rngSrc.Value

            On Error GoTo 0
            Next srcAddr
        End If
    Next r
    Exit Sub
ErrHandler:
    LogError "ApplyMappings error: " & Err.Number & " " & Err.Description
End Sub

'------------------------------
' ExportPDFs (unchanged)
'------------------------------
Public Sub ExportPDFs(basePath As String, _
                      newMon As String, _
                      rptID As String, _
                      wb As Workbook, _
                      pdfSheets As Variant, _
                      Optional parentFolder As String = "")
    On Error GoTo ErrHandler
    Dim saveRoot As String
    If Trim(parentFolder) <> "" Then
        saveRoot = parentFolder & "\" & newMon & "\" & rptID & "\"
    Else
        saveRoot = basePath & "\SAVE_PDF\" & newMon & "\" & rptID & "\"
    End If
    If Dir(saveRoot, vbDirectory) = "" Then MkDirRecursive saveRoot
    Dim baseFile As String, fileNoExt As String
    baseFile = Mid(wb.Name, InStrRev(wb.Name, "\") + 1)
    If InStrRev(baseFile, ".") > 0 Then fileNoExt = Left(baseFile, InStrRev(baseFile, ".") - 1) Else fileNoExt = baseFile
    If IsEmpty(pdfSheets) Then Exit Sub
    Dim i As Long
    For i = LBound(pdfSheets) To UBound(pdfSheets)
        Dim shName As String: shName = Trim(pdfSheets(i))
        If shName <> "" Then
            On Error Resume Next
            Dim ws As Worksheet: Set ws = wb.Worksheets(shName)
            On Error GoTo 0
            If Not ws Is Nothing Then
                Dim pdfName As String: pdfName = saveRoot & fileNoExt & "_" & shName & ".pdf"
                ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfName, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
                LogInfo "Exported PDF: " & pdfName
            Else
                LogWarn "ExportPDFs: sheet not found: " & shName
            End If
        End If
    Next i
    Exit Sub
ErrHandler:
    LogError "ExportPDFs error: " & Err.Number & " " & Err.Description
End Sub


Public Function BuildTimeStringFormats(ByVal ymROC As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    On Error GoTo ErrHandler

    ymROC = Trim(ymROC & "")
    If ymROC = "" Then Err.Raise vbObjectError + 1, , "BuildTimeStringFormats: ymROC is empty."

    Dim parts() As String
    parts = Split(ymROC, "/")
    If UBound(parts) < 1 Then Err.Raise vbObjectError + 2, , "ymROC 格式錯誤，範例：114/06"

    Dim rocY As Long: rocY = CLng(parts(0))
    Dim m As Long: m = CLng(parts(1))
    If m < 1 Or m > 12 Then Err.Raise vbObjectError + 3, , "月份不正確"

    ' --- NEW (current month) ---
    Dim newROC_YYY As String: newROC_YYY = CStr(rocY)            ' e.g. "114"
    Dim newROC_MM As String: newROC_MM = Format(m, "00")     ' e.g. "06"
    Dim newROC_YYY_SLASH_MM As String: newROC_YYY_SLASH_MM = newROC_YYY & "/" & newROC_MM   ' "114/06"
    Dim newROC_YYYMM As String: newROC_YYYMM = newROC_YYY & newROC_MM   ' "11406"

    Dim newROC_CHINESE_YYYMM As String
    newROC_CHINESE_YYYMM = "民國 " & newROC_YYY & " 年 " & newROC_MM & " 月"   ' 例如 "民國 114 年 06 月"    

    Dim newAD_YYYY As String: newAD_YYYY = CStr(rocY + 1911)   ' "2025"
    Dim newAD_YYYYMM As String: newAD_YYYYMM = newAD_YYYY & newROC_MM   ' "202506" (NO SLASH)
    Dim newAD_YYYY_SLASH_MM As String: newAD_YYYY_SLASH_MM = newAD_YYYY & "/" & newROC_MM ' "2025/06" (WITH SLASH)
    Dim lastDay As Date: lastDay = DateSerial(CLng(newAD_YYYY), m + 1, 0)
    Dim newAD_YYYYMMDD_END As String: newAD_YYYYMMDD_END = Format(lastDay, "yyyymmdd") ' "20250630"

    ' workday-end using your helper (returns "YYYYMMDD" or "")
    Dim newAD_YYYYMMDD_WORKDAY_END As String
    newAD_YYYYMMDD_WORKDAY_END = GetWesternMonthWorkDayEnd(newAD_YYYYMMDD_END) ' e.g. "20250629"

    ' ROC workday end (as ROC string, no slash) and with slash variant
    Dim newROC_YYYMMDD_WORKDAY_END As String
    If newAD_YYYYMMDD_WORKDAY_END <> "" Then
        newROC_YYYMMDD_WORKDAY_END = CStr(CLng(Left(newAD_YYYYMMDD_WORKDAY_END, 4)) - 1911) & Mid(newAD_YYYYMMDD_WORKDAY_END, 5, 4) ' "1140629"
    Else
        newROC_YYYMMDD_WORKDAY_END = ""
    End If

    ' --- OLD (previous month) ---
    Dim old_m As Long: old_m = m - 1
    Dim old_rocY As Long: old_rocY = rocY
    If old_m = 0 Then old_m = 12: old_rocY = rocY - 1
    Dim oldROC_YYY As String: oldROC_YYY = CStr(old_rocY)
    Dim oldROC_MM As String: oldROC_MM = Format(old_m, "00")
    Dim oldROC_YYY_SLASH_MM As String: oldROC_YYY_SLASH_MM = oldROC_YYY & "/" & oldROC_MM
    Dim oldROC_YYYMM As String: oldROC_YYYMM = oldROC_YYY & oldROC_MM

    Dim oldROC_CHINESE_YYYMM As String
    oldROC_CHINESE_YYYMM = "民國 " & oldROC_YYY & " 年 " & oldROC_MM & " 月"   ' 例如 "民國 114 年 05 月"


    Dim oldAD_YYYY As String: oldAD_YYYY = CStr(old_rocY + 1911)
    Dim oldAD_YYYYMM As String: oldAD_YYYYMM = oldAD_YYYY & oldROC_MM
    Dim oldAD_YYYY_SLASH_MM As String: oldAD_YYYY_SLASH_MM = oldAD_YYYY & "/" & oldROC_MM
    Dim prevLastDay As Date: prevLastDay = DateSerial(CLng(oldAD_YYYY), old_m + 1, 0)
    Dim oldAD_YYYYMMDD_END As String: oldAD_YYYYMMDD_END = Format(prevLastDay, "yyyymmdd")
    Dim oldAD_YYYYMMDD_WORKDAY_END As String
    oldAD_YYYYMMDD_WORKDAY_END = GetWesternMonthWorkDayEnd(oldAD_YYYYMMDD_END)

    Dim oldROC_YYYMMDD_WORKDAY_END As String
    If oldAD_YYYYMMDD_WORKDAY_END <> "" Then
        oldROC_YYYMMDD_WORKDAY_END = CStr(CLng(Left(oldAD_YYYYMMDD_WORKDAY_END, 4)) - 1911) & Mid(oldAD_YYYYMMDD_WORKDAY_END, 5, 4)
    Else
        oldROC_YYYMMDD_WORKDAY_END = ""
    End If

    ' --- *** ADDED ***：季度（Season / 季報）相關計算（新增 key） ---
    ''' *** ADDED *** 計算當前的季序（1..4）
    Dim seasonNum As Long
    seasonNum = (m + 2) \ 3   ' m=1..3 ->1, 4..6->2, 7..9->3, 10..12->4

    ''' *** ADDED *** 當季結束月份（3,6,9,12）
    Dim seasonMonth As Long
    seasonMonth = seasonNum * 3

    ''' *** ADDED *** NEW：當季（ROC）季別表示（格式 "YYYss" ，如 "11402" 表示 114 年第 2 季）
    Dim newROC_SEASON As String: newROC_SEASON = newROC_YYY & Format(seasonNum, "00")  ' e.g. "11402"
    Dim newROC_SEASON_SLASH As String: newROC_SEASON_SLASH = newROC_YYY & "/" & Format(seasonNum, "00") ' "114/02"

    ''' *** ADDED *** NEW：當季結束的 ROC YYYMM（例如 11406）
    Dim newROC_SEASON_YYYMM As String: newROC_SEASON_YYYMM = newROC_YYY & Format(seasonMonth, "00") ' "11406"

    ''' *** ADDED *** OLD：計算上一季（若當季為第 1 季，上一季為前一 ROC 年第 4 季）
    Dim oldSeasonNum As Long
    Dim oldSeasonYearROC As Long
    If seasonNum > 1 Then
        oldSeasonNum = seasonNum - 1
        oldSeasonYearROC = rocY
    Else
        oldSeasonNum = 4
        oldSeasonYearROC = rocY - 1
    End If

    Dim oldROC_SEASON As String: oldROC_SEASON = CStr(oldSeasonYearROC) & Format(oldSeasonNum, "00")  ' e.g. "11401" -> previous "11304"
    Dim oldROC_SEASON_SLASH As String: oldROC_SEASON_SLASH = CStr(oldSeasonYearROC) & "/" & Format(oldSeasonNum, "00")

    ''' *** ADDED *** OLD：上一季結束的 ROC 年月（YYYMM），這個就是你說的「舊季度月份」：
    '''            例如傳入 114/06 -> oldSeasonEndYyyMm = "11403"
    '''            傳入 114/03 -> oldSeasonEndYyyMm = "11312"（跨年）
    Dim oldSeasonMonth As Long: oldSeasonMonth = oldSeasonNum * 3
    Dim oldSeasonEndYearROC As Long: oldSeasonEndYearROC = oldSeasonYearROC
    Dim oldROC_SEASON_YYYMM As String: oldROC_SEASON_YYYMM = CStr(oldSeasonEndYearROC) & Format(oldSeasonMonth, "00") ' e.g. "11403" or "11312"

    ' --- 半年度 (NEW keys) ---
    ' *** MODIFIED START ***
    ' 說明（修改重點）：
    '  - halfIndex: 代表「半年度序號」(1=上半, 2=下半)
    '  - halfNum: 實際要放入 key 的數值，依你要求為 2 或 4（上半 -> 02, 下半 -> 04）
    Dim halfIndex As Long
    Dim halfNum As Long
    Dim halfMonth As Long

    If m <= 6 Then
        halfIndex = 1
    Else
        halfIndex = 2
    End If

    halfNum = halfIndex * 2        ' 1 -> 2, 2 -> 4   (符合你要求)
    If halfIndex = 1 Then
        halfMonth = 6
    Else
        halfMonth = 12
    End If

    Dim newROC_HALF As String, newROC_HALF_SLASH As String, newROC_HALF_YYYMM As String
    newROC_HALF = newROC_YYY & Format(halfNum, "00")
    newROC_HALF_SLASH = newROC_YYY & "/" & Format(halfNum, "00")
    newROC_HALF_YYYMM = newROC_YYY & Format(halfMonth, "00")
    ' *** MODIFIED END ***

    ' *** MODIFIED START for OLD half ***
    Dim oldHalfIndex As Long, oldHalfYearROC As Long, oldHalfNum As Long, oldHalfMonth As Long
    If halfIndex > 1 Then
        oldHalfIndex = halfIndex - 1
        oldHalfYearROC = rocY
    Else
        oldHalfIndex = 2
        oldHalfYearROC = rocY - 1
    End If
    oldHalfNum = oldHalfIndex * 2
    If oldHalfIndex = 1 Then
        oldHalfMonth = 6
    Else
        oldHalfMonth = 12
    End If

    Dim oldROC_HALF As String, oldROC_HALF_SLASH As String, oldROC_HALF_YYYMM As String
    oldROC_HALF = CStr(oldHalfYearROC) & Format(oldHalfNum, "00")
    oldROC_HALF_SLASH = CStr(oldHalfYearROC) & "/" & Format(oldHalfNum, "00")
    oldROC_HALF_YYYMM = CStr(oldHalfYearROC) & Format(oldHalfMonth, "00")
    ' *** MODIFIED END *** 

    ' --- pack into dictionary with very explicit keys ---
    With d
        .RemoveAll
        ' NEW (no-slash / slash variants where applicable)
        .Add "NEW_ROC_YYY", newROC_YYY                     ' "114"
        .Add "NEW_ROC_MM", newROC_MM                   ' "06"
        .Add "NEW_ROC_YYY_SLASH_MM", newROC_YYY_SLASH_MM ' "114/06"
        .Add "NEW_ROC_YYYMM", newROC_YYYMM                 ' "11406" (no slash)

        .Add "NEW_AD_YYYY", newAD_YYYY                 ' "2025"
        .Add "NEW_AD_YYYYMM", newAD_YYYYMM             ' "202506" (no slash)  <-- **主要 AD(無slash) key**
        .Add "NEW_AD_YYYY_SLASH_MM", newAD_YYYY_SLASH_MM ' "2025/06" (有 slash)
        .Add "NEW_AD_YYYYMMDD_END", newAD_YYYYMMDD_END ' "20250630"
        .Add "NEW_AD_YYYYMMDD_WORKDAY_END", newAD_YYYYMMDD_WORKDAY_END ' "20250629" or ""
        .Add "NEW_ROC_YYYMMDD_WORKDAY_END", newROC_YYYMMDD_WORKDAY_END ' "1140629" (no slash)

        .Add "NEW_ROC_CHINESE_YYYMM", newROC_CHINESE_YYYMM         ' "民國 114 年 06 月"        

        ' --- *** ADDED *** 季度 (NEW keys) ---
        .Add "NEW_ROC_SEASON", newROC_SEASON             ' e.g. "11402" <- 6 月時為第 2 季 (*** NEW ***)
        .Add "NEW_ROC_SEASON_SLASH", newROC_SEASON_SLASH ' e.g. "114/02"
        .Add "NEW_ROC_SEASON_YYYMM", newROC_SEASON_YYYMM ' e.g. "11406" (當季結束月)
        .Add "NEW_SEASON_NUM", seasonNum                 ' 1..4
        .Add "NEW_SEASON_MONTH", seasonMonth      ' 3,6,9,12

        ' 半年度 (NEW)
        .Add "NEW_ROC_HALF", newROC_HALF
        .Add "NEW_ROC_HALF_SLASH", newROC_HALF_SLASH
        .Add "NEW_ROC_HALF_YYYMM", newROC_HALF_YYYMM
        .Add "NEW_HALF_NUM", halfNum
        .Add "NEW_HALF_MONTH", halfMonth        

        ' OLD (previous month)
        .Add "OLD_ROC_YYY", oldROC_YYY
        .Add "OLD_ROC_MM", oldROC_MM
        .Add "OLD_ROC_YYY_SLASH_MM", oldROC_YYY_SLASH_MM
        .Add "OLD_ROC_YYYMM", oldROC_YYYMM

        .Add "OLD_AD_YYYY", oldAD_YYYY
        .Add "OLD_AD_YYYYMM", oldAD_YYYYMM
        .Add "OLD_AD_YYYY_SLASH_MM", oldAD_YYYY_SLASH_MM
        .Add "OLD_AD_YYYYMMDD_END", oldAD_YYYYMMDD_END
        .Add "OLD_AD_YYYYMMDD_WORKDAY_END", oldAD_YYYYMMDD_WORKDAY_END
        .Add "OLD_ROC_YYYMMDD_WORKDAY_END", oldROC_YYYMMDD_WORKDAY_END

        .Add "OLD_ROC_CHINESE_YYYMM", oldROC_CHINESE_YYYMM         ' "民國 114 年 05 月"

        ' --- *** ADDED *** 季度 (OLD keys) ---
        .Add "OLD_ROC_SEASON", oldROC_SEASON                 ' e.g. "11401" or "11304"
        .Add "OLD_ROC_SEASON_SLASH", oldROC_SEASON_SLASH
        .Add "OLD_ROC_SEASON_YYYMM", oldROC_SEASON_YYYMM  ' 這就是你要的「舊季度月份」，例如 "11403" 或 "11312"        
        .Add "OLD_SEASON_NUM", oldSeasonNum
        .Add "OLD_SEASON_MONTH", oldSeasonMonth

        ' 半年度 (OLD)
        .Add "OLD_ROC_HALF", oldROC_HALF
        .Add "OLD_ROC_HALF_SLASH", oldROC_HALF_SLASH
        .Add "OLD_ROC_HALF_YYYMM", oldROC_HALF_YYYMM
        .Add "OLD_HALF_NUM", oldHalfNum
        .Add "OLD_HALF_MONTH", oldHalfMonth        

        .Add "NONE", ""
    End With

    Set BuildTimeStringFormats = d
    Exit Function

ErrHandler:
    LogError "BuildTimeStringFormats error: " & Err.Number & " " & Err.Description
    Set BuildTimeStringFormats = d
End Function


Public Sub FillHeaderDates(wb As Workbook, _
                           headerRefs As Variant, _
                           mf As Object, _
                           Optional headerTimeFormat As String = "")
    On Error Resume Next
    Dim valueToWrite As String
    valueToWrite = ""

    ' 1) 決定要寫入的值（優先使用 mf("NEW_" & headerTimeFormat)）
    If Trim(headerTimeFormat & "") <> "" Then
        If mf.Exists(headerTimeFormat) Then
            valueToWrite = mf(headerTimeFormat)
        Else
          ' ************Logwarn
          Exit Function
        End If
    Else
        ' *************LogWarn
        Exit Function
    End If

    ' 2) 支援 headerRefs 為單一字串或陣列
    Dim refs As Variant
    If IsArray(headerRefs) Then
        refs = headerRefs
    Else
        ' *********LogWarn
        Exit Function
    End If

    Dim i As Long
    For i = LBound(refs) To UBound(refs)
        Dim item As String = Trim(CStr(refs(i)))
        If item = "" Then GoTo NextRef
        ' 若使用者在同一儲存格還用 '|' 串多個 target，這裡再拆一次（容錯）
        Dim subParts() As String
        If InStr(item, "|") > 0 Then
            subParts = Split(item, "|")
        Else
            ReDim subParts(0 To 0)
            subParts(0) = item
        End If

        Dim s As Long
        For s = LBound(subParts) To UBound(subParts)
            Dim oneRef As String: oneRef = Trim(subParts(s))
            If oneRef = "" Then GoTo NextSub
            Dim ex As Long: ex = InStr(oneRef, "!")
            If ex = 0 Then
                ' 非 "Sheet!Addr" 格式，跳過（或可擴充支援 "Sheet Addr"）
                LogWarn "FillHeaderDates: 無效的 Header 參考格式: " & oneRef
                GoTo NextSub
            End If
            Dim shName As String: shName = Left(oneRef, ex - 1)
            Dim addr As String: addr = Mid(oneRef, ex + 1)
            On Error Resume Next
            If wb.Worksheets(shName) Is Nothing Then
                LogWarn "FillHeaderDates: 找不到工作表: " & shName
            Else
                ' 嘗試寫入（若 addr 非法會被忽略）
                wb.Worksheets(shName).Range(addr).Value = valueToWrite
            End If
            On Error GoTo 0
NextSub:
        Next s
NextRef:
    Next i
End Sub



' 1. 專案沒有處理路徑WildCard情況，可能用擷取Excel儲存格內的值，當作變數，丟進去一個搜尋的wildCard，讓他會隨著Input進去的值，
'    抓到那個檔名和檔案路徑，寫一個流程讓他可以去抓到檔案，有點算是動態抓取資料感覺。

' Ans: 應該有解決

' 2. 台幣有一張表要自己填入那張還沒處理

' 3. 支援可以不用一次處理所有的表

' 4. 如果已經處理過的話，留下紀錄，避免重複跑資料

' 5. updateSheet 欄位新增 "貼入起始位置PasteRange(TopLeft)"

' Answer: 解決

' 6. 新增路徑辨識 ..\ 往前一層folder功能

' Answer: 應該有解決

' 7. 部分報表不需要產生申報檔，新增 "IsToCreateDeclFile" 欄位

' Answer: 解決

' 8. 直接在Build裡面加入季節判斷，到時候直接判斷，
' 季報，原始檔案路徑要抓三個月前的，Build裡面要建立一個季報的OldYearMonth可以抓，

' 表頭也要特別Create新的去處理

' Answer: 解決

' 9. headerFormat 那邊要處理為字串




' ============================
' TypeHandling

Public Sub PNCDCAL_FormatToCSV(ByVal fullFilePath As String)
    Dim fso As Object
    Dim startHandle As Boolean
    Dim inputFile As Object
    Dim outputFile As Object
    Dim regEx As Object

    Dim i As Integer
    Dim line As String
    Dim outputLine As String
    
    Dim fields As Variant

    Dim outputFilePath As String
    Dim fileExtension As String
    
    ' 檔案檢查
    If Dir(fullFilePath) = "" Then
        MsgBox "File not found: " & fullFilePath
        Exit Sub
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")

    fileExtension = LCase(fso.GetExtensionName(fullFilePath))
    
    If fileExtension = "txt" Then
        outputFilePath = Left(fullFilePath, Len(fullFilePath) - Len(fileExtension)) & "csv"
    Else
        MsgBox "Error for FileExtension[FullFilePath: " & fullFilePath & "]"
        Exit Sub
    End If

    Set inputFile = fso.OpenTextFile(fullFilePath, 1, False)
    Set outputFile = fso.CreateTextFile(outputFilePath, True)
    Set regEx = CreateObject("VBScript.RegExp")
    
    With regEx
        .Global = False
        .Pattern = "^\d+-\d+$"    ' 檢查是否是 "1-30" 這種格式
    End With
    
    ' 寫入標題及初始化處理Row
    outputFile.WriteLine "期間,代號,上月餘額,本月利率,本月金額,本月償還,本月餘額(仟),佔比(%)"
    startHandle = False
    Do Until inputFile.AtEndOfStream
        line = Trim(inputFile.ReadLine)
        
        If InStr(line, "----") > 0 Then
            startHandle = True
        ElseIf startHandle Then
            If InStr(line, "===") > 0 Or Len(line) = 0 Then Exit Do
            If InStr(line, "合計") = 0 Then
                Do While InStr(line, "  ") > 0
                    line = Replace(line, "  ", " ")
                Loop
                
                line = Replace(line, ChrW(12288), " ") ' 去掉全形空格
                fields = Split(WorksheetFunction.Trim(line), " ")
                
                ' --- 新增邏輯：檢查前兩欄是否為 "1-30 + 天" 型態 ---
                If UBound(fields) >= 1 Then
                    If regEx.Test(fields(0)) And fields(1) = "天" Then
                        ' 合併 "1-30" + "天"
                        fields(0) = fields(0) & "天"
                        ' 將第二欄（"天"）移除
                        For i = 1 To UBound(fields) - 1
                            fields(i) = fields(i + 1)
                        Next i
                        ReDim Preserve fields(UBound(fields) - 1)
                    End If
                End If
                
                ' 去掉數字逗號
                For i = LBound(fields) To UBound(fields)
                    fields(i) = Replace(fields(i), ",", "")
                Next i
                
                outputLine = Join(fields, ",")
                outputFile.WriteLine outputLine
                
            End If
        End If
    Loop
    inputFile.Close
    outputFile.Close
End Sub



Public Sub ImportCsvWithFilter( _
    ByVal csvPath As String, _
    ByVal targetWS As Worksheet, _
    ByVal targetCell As Range, _
    ByVal filterFields As Variant, _
    ByVal filterValues As Variant)

    Dim wbCsv As Workbook
    Dim shCsv As Worksheet
    Dim fRange As Range
    Dim dataRange As Range
    Dim cell As Range
    Dim i As Long
    Dim colIndex As Long

    ' 1. 開啟 CSV 檔案
    Set wbCsv = Workbooks.Open(Filename:=csvPath)
    Set shCsv = wbCsv.Sheets(1)

    ' 2. 取得標題列範圍
    Set fRange = shCsv.Range(shCsv.Cells(1, 1), shCsv.Cells(1, shCsv.UsedRange.Columns.Count))
    fRange.AutoFilter

    ' 🔧 修正欄位資料：移除不可見空白字元
    Set dataRange = shCsv.UsedRange
    For Each cell In dataRange
        If VarType(cell.Value) = vbString Then
            cell.Value = Trim(Replace(cell.Value, Chr(160), ""))
        End If
    Next cell

    ' 3. 套用篩選條件
    For i = LBound(filterFields) To UBound(filterFields)
        colIndex = Application.Match(filterFields(i), fRange, 0)

        If Not IsError(colIndex) Then
            ' 🔧 修正：若該欄條件是多個值

            If IsArray(filterValues(i)) Then
                '──【改動】──
                ' 如果第一個元素以 "<>" 開頭，代表要做排除條件
                If Left(filterValues(i)(0), 2) = "<>" Then
                    Select Case UBound(filterValues(i))
                        Case 0
                            ' 一個排除條件
                            shCsv.UsedRange.AutoFilter _
                                Field:=colIndex, _
                                Criteria1:=filterValues(i)(0)
                        Case 1
                            ' 兩個排除條件
                            shCsv.UsedRange.AutoFilter _
                                Field:=colIndex, _
                                Criteria1:=filterValues(i)(0), _
                                Operator:=xlAnd, _
                                Criteria2:=filterValues(i)(1)
                        Case Else
                            MsgBox "排除條件最多只能兩個: " & filterFields(i), vbExclamation
                    End Select                

                Else
                    shCsv.UsedRange.AutoFilter Field:=colIndex, Criteria1:=filterValues(i), Operator:=xlFilterValues
                End If
            Else
                shCsv.UsedRange.AutoFilter Field:=colIndex, Criteria1:=filterValues(i)
            End If
        Else
            MsgBox "ImportCsvWithFilter 找不到欄位: " & filterFields(i), vbExclamation
        End If
    Next i

    ' 4. 複製可見列（含標題）
    On Error Resume Next
    shCsv.UsedRange.SpecialCells(xlCellTypeVisible).Copy
    On Error GoTo 0

    targetWS.Paste targetCell

    ' 5. 關閉 CSV
    wbCsv.Close SaveChanges:=False
End Sub


Public Sub ImportCloseRate(ByVal csvPath As String, _
                           ByVal targetWS As Worksheet, _
                           ByVal TargetCurrency As String, _
                           ByVal pasteRange As String)
    Dim wbCsv As Workbook
    Dim shCsv As Worksheet
    Dim i As Long, lastRow As Long
    Dim isRowDelete As Boolean
    Dim BaseCurrency As String

    ' 1. 開啟 CSV 檔案
    Set wbCsv = Workbooks.Open(Filename:=csvPath)
    Set shCsv = wbCsv.Sheets(1)
    
    shCsv.Columns("E").Delete
    shCsv.Columns("C").Delete
    lastRow = shCsv.Cells(shCsv.Rows.count, "A").End(xlUp).Row

    For i = 2 To lastRow
        If shCsv.Cells(i, "A").Value <> "" Then
            BaseCurrency = shCsv.Cells(i, "A").Value
        Else
            shCsv.Cells(i, "A").Value = BaseCurrency
        End If
    Next i

    If TargetCurrency = "USD" Then
        TargetCurrency = "TWD"
    ElseIf TargetCurrency = "TWD" Then
        TargetCurrency = "USD"
    Else
        MsgBox "未填入或Currency錯誤"
        Exit Sub
    End If

    ' 反向遍歷刪除列
    For i = lastRow To 2 Step -1
        isRowDelete = False
        If IsEmpty(shCsv.Cells(i, "A").Value) Or IsEmpty(shCsv.Cells(i, "B").Value) Or IsEmpty(shCsv.Cells(i, "C").Value) Or _
           Left(shCsv.Cells(i, "A").Value, 4) = "經副襄理" Or shCsv.Cells(i, "A").Value = TargetCurrency Then
            isRowDelete = True
        End If

        ' Delete Row
        If isRowDelete Then shCsv.Rows(i).Delete
    Next i

    On Error Resume Next
    shCsv.UsedRange.Copy
    On Error GoTo 0
    targetWS.Range(pasteRange).PasteSpecial xlPasteValues
    ' 5. 關閉 CSV
    wbCsv.Close SaveChanges:=False
End Sub


Public Sub ImportFXDebtEvaluation(ByVal csvPath As String)
    Dim ws As Worksheet
    Dim xlbk As WorkBook
    Dim xlsht As Worksheet
    
    Dim copyRg As Range
    Dim Rngs As Range
    Dim oneRng As Range
    
    Dim outputArr() As Variant
    Dim fvArray As Variant
    Dim mapGroupMeasurement As Object
    Dim groupMeasurement As Variant

    Dim i As Integer, j As Integer, k As Integer
    Dim lastRow As Integer
    
    Dim securityRows As Collection
    Dim category As Variant
    Dim tableColumns As Variant

    Set xlbk = Workbooks.Open(Filename:=csvPath)
    Set ws = xlbk.Worksheets("評估表")
    ws.Copy After:=xlbk.Sheets(xlbk.Sheets.count)
    ActiveSheet.Name = "評估表cp"
    Set xlsht = ActiveSheet
    
    Set copyRg = xlsht.UsedRange
    copyRg.Value = copyRg.Value
        
    Set Rngs = copyRg.Range("A5:T5")
    
    Dim columnsArray() As Variant
    Dim tempSplit As Variant
    Dim tempSave() As Variant
    Dim count As Long
    Dim splitCount As Long

    count = 0
    splitCount = 0
    
    For Each oneRng In Rngs
        tempSplit = Split(oneRng, vbLf)
        ReDim Preserve columnsArray(count)
        columnsArray(count) = Trim(tempSplit(0))
        count = count + 1
        
        If UBound(tempSplit) >= 1 Then
            ReDim Preserve tempSave(splitCount)
            tempSave(splitCount) = Trim(tempSplit(1))
            splitCount = splitCount + 1
        End If
    Next oneRng
    
    ReDim Preserve tempSave(splitCount)
    tempSave(splitCount) = "評價資產類別"
    
    For i = LBound(tempSave) To UBound(tempSave)
        ReDim Preserve columnsArray(count)
        columnsArray(count) = tempSave(i)
        count = count + 1
    Next i
    
    fvArray = Array("FVPL-公債", _
                    "FVPL-公司債(公營)", _
                    "FVPL-公司債(民營)", _
                    "FVPL-金融債", _
                    "FVOCI-公債", _
                    "FVOCI-公司債(公營)", _
                    "FVOCI-公司債(民營)", _
                    "FVOCI-金融債", _
                    "AC-公債", _
                    "AC-公司債(公營)", _
                    "AC-公司債(民營)", _
                    "AC-金融債")

    groupMeasurement = Array("FVPL_GovBond_Foreign", _
                             "FVPL_CompanyBond_Foreign", _
                             "FVPL_CompanyBond_Foreign", _
                             "FVPL_FinancialBond_Foreign", _
                             "FVOCI_GovBond_Foreign", _
                             "FVOCI_CompanyBond_Foreign", _
                             "FVOCI_CompanyBond_Foreign", _
                             "FVOCI_FinancialBond_Foreign", _
                             "AC_GovBond_Foreign", _
                             "AC_CompanyBond_Foreign", _
                             "AC_CompanyBond_Foreign", _
                             "AC_FinancialBond_Foreign")
    
    Set mapGroupMeasurement = CreateObject("Scripting.Dictionary")
    For i = LBound(fvArray) To UBound(fvArray)
        mapGroupMeasurement.Add fvArray(i), groupMeasurement(i)
    Next i

    lastRow = xlsht.Cells(xlsht.Rows.count, 1).End(xlUp).Row
    
    For i = lastRow To 1 Step -1
        If IsEmpty(xlsht.Cells(i, 1).Value) Or _
            xlsht.Cells(i, 1).Value = "Security_Id" Then
                xlsht.Rows(i).Delete
        End If
        
        If Left(Trim(xlsht.Cells(i, 1).Value), 2) = "標註" Then
            xlsht.Rows(i & ":" & lastRow).Delete
        End If
    Next i

    lastRow = xlsht.Cells(xlsht.Rows.count, 1).End(xlUp).Row
    
    Set securityRows = New Collection

    For i = 1 To lastRow
        For Each category In fvArray
            If xlsht.Cells(i, 1).Value = category Then
                securityRows.Add i
            End If
        Next category
    Next i
    
    Dim startRow As Integer
    Dim endRow As Integer
    Dim numRows As Integer
    Dim numCols As Integer

    For i = 1 To securityRows.count
        If i = 1 Then
            ReDim outputArr(1 To lastRow, 1 To 32)
        End If

        If i + 1 <= securityRows.count Then
            If securityRows(i) + 1 = securityRows(i + 1) Then
                GoTo ContinueLoop
            Else
                startRow = securityRows(i) + 1
                endRow = securityRows(i + 1) - 1
            End If
        Else
            startRow = securityRows(i) + 1
            endRow = lastRow
        End If
        
        category = xlsht.Cells(startRow - 1, 1).Value

        For j = startRow To endRow Step 2
            For k = 1 To 40
                If k >= 1 And k <= 20 Then
                    outputArr(j, k) = xlsht.Cells(j, k).Value
                    If category = "AC-公債" Or category = "AC-公司債(公營)" Or _
                       category = "AC-公司債(民營)" Or category = "AC-金融債" Then
                       outputArr(j, 20) = xlsht.Cells(j, 17).Value
                    End If
                    outputArr(j, 17) = ""
                ElseIf k = 22 Then 'Issuer
                    outputArr(j, 21) = xlsht.Cells(j + 1, k - 20).Value
                ElseIf k = 28 Then 'Avg_Txnt_Rate
                    outputArr(j, 22) = xlsht.Cells(j + 1, k - 20).Value
                ElseIf k = 29 Then 'Avg_Buy_Price
                    outputArr(j, 23) = xlsht.Cells(j + 1, k - 20).Value
                ElseIf k = 30 Then 'Tot_Nominal_Amt_USD
                    outputArr(j, 24) = xlsht.Cells(j + 1, k - 20).Value
                ElseIf k = 31 Then 'Book_Value
                    outputArr(j, 25) = xlsht.Cells(j + 1, k - 20).Value
                ElseIf k = 32 Then 'PL_Amt_USD
                    outputArr(j, 26) = xlsht.Cells(j + 1, k - 20).Value
                ElseIf k = 33 Then 'Amortize_Amt
                    outputArr(j, 27) = xlsht.Cells(j + 1, k - 20).Value
                ElseIf k = 34 Then 'DVO1_USD
                    outputArr(j, 28) = xlsht.Cells(j + 1, k - 20).Value
                ElseIf k = 35 Then 'Interest_receivable_USD
                    outputArr(j, 29) = xlsht.Cells(j + 1, k - 20).Value
                ElseIf k = 36 Then '當日評等
                    outputArr(j, 30) = xlsht.Cells(j + 1, k - 20).Value
                ElseIf k = 37 Then '評價類別
                    outputArr(j, 31) = category
                End If
            Next k

            ' Add col 32 for groupMeasurement
            If mapGroupMeasurement.Exists(category) Then
                outputArr(j, 32) = mapGroupMeasurement(category)
            Else
                outputArr(j, 32) = ""
            End If
        Next j
ContinueLoop:
    Next i
    
    xlbk.Sheets.Add After:=xlbk.Sheets(xlbk.Sheets.count)
    ActiveSheet.Name = "OutputData"

    numRows = UBound(outputArr, 1)
    numCols = UBound(outputArr, 2)
    ActiveSheet.Range(ActiveSheet.Cells(2, 1), ActiveSheet.Cells(numRows + 1, numCols)).Value = outputArr
    
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.count, 1).End(xlUp).Row
    
    ' For i = lastRow To 2 Step -1
    '     If ActiveSheet.Cells(i, 1).Value = "" Then
    '     ActiveSheet.Rows(i).Delete
    '     End If
    ' Next i

    For i = lastRow To 2 Step -1
        ' 如果 A 欄是空白，或 AE 欄不符合 "AC*" 就刪除整列
        If ActiveSheet.Cells(i, 1).Value = "" _
           Or Not (ActiveSheet.Cells(i, 31).Value Like "AC*") Then
            ActiveSheet.Rows(i).Delete
        End If
    Next i

    For i = 1 To (UBound(columnsArray) - LBound(columnsArray) + 1)
        ActiveSheet.Range("A1").Resize(1, (UBound(columnsArray) - LBound(columnsArray) + 1)).Value = columnsArray
    Next i

    If ActiveSheet.Range("AH1").Value = "評價資產類別" Then
        ActiveSheet.Range("AH1").Value = ""
    End If
    
    For Each ws In xlbk.Sheets
        If Not ws Is ActiveSheet Then
            ws.Delete
        End If
    Next ws

    Set securityRows = Nothing

    xlbk.Save
    xlbk.Close False
    Set xlsht = Nothing
    Set xlbk = Nothing
End Sub

' ============================


' 將 Holidays 工作表 C 欄的值標準化為 "YYYYMMDD"（無斜線）
Private Function NormalizeDate(ByVal v As Variant) As String
    On Error GoTo ErrHandler
    NormalizeDate = ""
    If IsNull(v) Then Exit Function
    If Trim(CStr(v)) = "" Then Exit Function
    
    ' 若 Excel 能辨識為日期，直接用 CDate 轉
    If IsDate(v) Then
        NormalizeDate = Format(CDate(v), "yyyymmdd")
        Exit Function
    End If
    
    ' 否則嘗試把字串中的 / 或 - 移除，再檢查是否為 8 位數 YYYYMMDD
    Dim s As String
    s = Trim(CStr(v))
    s = Replace(s, "/", "")
    s = Replace(s, "-", "")
    s = Replace(s, " ", "")
    
    If Len(s) = 8 And IsNumeric(s) Then
        ' 再檢查是否為合法日期（避免像 20250230 之類不合法）
        Dim y As Integer, m As Integer, d As Integer
        y = CInt(Left(s, 4))
        m = CInt(Mid(s, 5, 2))
        d = CInt(Right(s, 2))
        If y >= 1900 And m >= 1 And m <= 12 Then
            ' 用 DateSerial 嘗試建構日期（若非法會錯誤）
            Dim dt As Date
            dt = DateSerial(y, m, d)
            NormalizeDate = Format(dt, "yyyymmdd")
        End If
    End If
    Exit Function
ErrHandler:
    NormalizeDate = ""
End Function

' 主函式：輸入 "YYYYMMDD"（無斜線） -> 回傳調整後之 "YYYYMMDD"（無斜線）
' holidaysSheetName 預設為 "Holidays"，會讀取該工作表之 C 欄作為「不用上班」清單
Public Function GetWesternMonthWorkDayEnd(ByVal inputYMD As String, Optional ByVal holidaysSheetName As String = "Holidays") As String
    On Error GoTo ErrHandler
    GetWesternMonthWorkDayEnd = ""
    Dim s As String: s = Trim(inputYMD)
    If s = "" Then Exit Function
    ' 必須為 8 位數
    If Len(s) <> 8 Or Not IsNumeric(s) Then Exit Function
    
    Dim y As Long, m As Long, d As Long
    y = CLng(Left(s, 4))
    m = CLng(Mid(s, 5, 2))
    d = CLng(Right(s, 2))
    
    ' 嘗試建立 Date（若輸入非法會跳到 ErrHandler）
    Dim cur As Date
    cur = DateSerial(y, m, d)
    
    ' 讀取 Holidays!C 欄，建立字典（key = "YYYYMMDD"）
    Dim holidays As Object: Set holidays = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(holidaysSheetName)
    On Error GoTo ErrHandler
    
    If Not ws Is Nothing Then
        Dim lastRow As Long, r As Long
        lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row ' 第3欄 C
        For r = 1 To lastRow
            Dim rawVal As Variant
            rawVal = ws.Cells(r, 3).Value
            Dim key As String
            key = NormalizeDate(rawVal)
            If key <> "" Then
                If Not holidays.Exists(key) Then holidays.Add key, True
            End If
        Next r
    End If
    
    ' 往前搜尋：若該日 NOT 在 holidays 中，則為工作日 -> 回傳 "YYYYMMDD"
    Dim safety As Long: safety = 0
    Do
        Dim curKey As String: curKey = Format(cur, "yyyymmdd")
        If Not holidays.Exists(curKey) Then
            GetWesternMonthWorkDayEnd = curKey
            Exit Function
        End If
        cur = DateAdd("d", -1, cur)
        safety = safety + 1
        If safety > 10000 Then Exit Do ' safety 防止極端無窮迴圈
    Loop
    
    ' 若失敗或超過安全次數則回傳空字串
    GetWesternMonthWorkDayEnd = ""
    Exit Function
ErrHandler:
    ' 如欲丟錯可改為 Err.Raise；目前設計為回傳空字串表示錯誤/格式不合法
    GetWesternMonthWorkDayEnd = ""
End Function
