' Module.bas

Option Explicit

Private Sub GetMonths(ByRef oldMon As String, ByRef newMon As String)
    Dim ymRaw   As String
    Dim parts() As String
    Dim y       As Integer
    Dim m       As Integer

    ' 假設儲存在 Name "YearMonth" 的儲存格是類似 "114/06" 的字串
    ymRaw = ThisWorkbook.Names("YearMonth").RefersToRange.Value
    ' 拆出年與月
    parts = Split(ymRaw, "/")
    y = CInt(parts(0))
    m = CInt(parts(1))
    ' 當前月份（newMon）
    newMon = CStr(y) & "/" & Format(m, "00")
    ' 計算上一個月
    m = m - 1
    If m = 0 Then
        y = y - 1
        m = 12
    End If

    If m = 0 Then y = y - 1: m = 12
    ' 上一個月份（oldMon）
    oldMon = CStr(y) & "/" & Format(m, "00")
End Sub

Public Function ConvertToROCFormat(ByVal newYearMonth As String, _
                                   ByVal returnType As String) As String
    Dim parts() As String
    Dim rocYear As Integer
    Dim result As String

    parts = Split(newYearMonth, "/")
    rocYear = CInt(parts(0))

    If returnType = "ROC" Then
        result = " 民國 " & CStr(rocYear) & " 年 " & parts(1) & " 月"
    ElseIf returnType = "NUM" Then
        result = CStr(rocYear) & parts(1)
    End If
    
    ConvertToROCFormat = result
End Function


' ###修改開始###

Public Function GetWesternMonthEnd() As String
    Dim ymRaw      As String
    Dim parts()    As String
    Dim rocYear    As Integer
    Dim monthNum   As Integer
    Dim adYear     As Integer
    Dim lastDay    As Date

    ymRaw = ThisWorkbook.Names("YearMonth").RefersToRange.Value
    parts = Split(ymRaw, "/")
    rocYear  = CInt(parts(0))
    monthNum = CInt(parts(1))
    
    ' ROC 年 + 1911 = 西元年
    adYear = rocYear + 1911
    ' DateSerial(year, month+1, 0) 會得到該月份的最後一天
    lastDay = DateSerial(adYear, monthNum + 1, 0)
    
    GetWesternMonthEnd = Format(lastDay, "yyyymmdd")
End Function

' ###修改結束###


'— 2. 主流程：依 CaseType 分派 ——
Sub ProcessAllReports()
    Dim wbCtl    As Workbook
    Dim wsRpt    As Worksheet, wsMap As Worksheet
    Dim basePath As String
    Dim lastRpt  As Long, lastMap As Long
    Dim oldMon   As String, newMon As String
    Dim ROCYearMonth As String, NUMYearMonth As String
    Dim westernMonthEnd As String
    Dim i        As Long, caseType As String
    ' ###修改開始###    
    Dim rptType As String
    Dim monString  As String
    Dim sendNUM As String
    Dim sendOldMon As String
    ' ###修改結束###    

    Call GetMonths(oldMon, newMon)

    ROCYearMonth = ConvertToROCFormat(newMon, "ROC")
    NUMYearMonth = ConvertToROCFormat(newMon, "NUM")
    oldMon = Replace(oldMon, "/", "")
    newMon = Replace(newMon, "/", "")
    westernMonthEnd = GetWesternMonthEnd()

    ' ###修改開始###    
    monString = Right(newMon, 2)
    ' ###修改結束###    

    Set wbCtl    = ThisWorkbook
    Set wsRpt    = wbCtl.Sheets("ReportsConfig")
    Set wsMap    = wbCtl.Sheets("Mappings")
    basePath     = wbCtl.Path
    lastRpt      = wsRpt.Cells(wsRpt.Rows.Count, "A").End(xlUp).Row
    lastMap      = wsMap.Cells(wsMap.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRpt
        ' ###修改開始###
        rptType = wsRpt.Cells(i, "G").Value
        Select Case rptType
            Case "季報"
                If Not (monString = "03" Or monString = "06" Or monString = "09" Or monString = "12") Then GoTo NextRpt
            Case "半年報"
                If Not (monString = "06" Or monString = "12") Then GoTo NextRpt
            Case "月報"
                ' 都執行
            Case Else
                Debug.Print "未知報表類型：" & rptType
                GoTo NextRpt
        End Select
        ' ###修改結束###

        ' ###修改開始###
        ' 預設送出原始 NUMYearMonth
        sendNUM = NUMYearMonth
        sendOldMon = oldMon
        If rptType = "季報" Then
            Select Case monString
                Case "03"
                    sendNUM = Left(NUMYearMonth, Len(NUMYearMonth) - 2) & "01"
                    sendOldMon = CStr(CInt(Left(oldMon, Len(oldMon) - 2)) - 1) & "12"
                Case "06"
                    sendNUM = Left(NUMYearMonth, Len(NUMYearMonth) - 2) & "02"
                    sendOldMon = Left(oldMon, Len(oldMon) - 2) & "03"
                Case "09"
                    sendNUM = Left(NUMYearMonth, Len(NUMYearMonth) - 2) & "03"
                    sendOldMon = Left(oldMon, Len(oldMon) - 2) & "06"
                Case "12"
                    sendNUM = Left(NUMYearMonth, Len(NUMYearMonth) - 2) & "04"
                    sendOldMon = Left(oldMon, Len(oldMon) - 2) & "09"
            End Select
        ElseIf rptType = "半年報" Then
            Select Case monString
                Case "06"
                    sendNUM = Left(NUMYearMonth, Len(NUMYearMonth) - 2) & "02"
                    sendOldMon = CStr(CInt(Left(oldMon, Len(oldMon) - 2)) - 1) & "12"
                Case "12"
                    sendNUM = Left(NUMYearMonth, Len(NUMYearMonth) - 2) & "04"
                    sendOldMon = Left(oldMon, Len(oldMon) - 2) & "06"
            End Select            
        End If
        ' ###修改結束###

        caseType = wsRpt.Cells(i, "H").Value
        Select Case caseType
            Case "會計資料庫"
                Call Import_會計資料庫( _
                    basePath, sendOldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, _
                    lastMap, _
                    ROCYearMonth, _
                    sendNUM)
            Case "PNCDCAL"
                Call Import_PNCDCAL(basePath, _
                                    sendOldMon, _
                                    newMon, _
                                    wsRpt.Cells(i, "A").Value, _
                                    wsRpt.Cells(i, "B").Value, _
                                    wsRpt.Cells(i, "C").Value, _
                                    wsRpt.Cells(i, "D").Value, _
                                    wsRpt.Cells(i, "E").Value, _
                                    wsRpt.Cells(i, "F").Value, _
                                    wsMap, _
                                    lastMap, _
                                    ROCYearMonth, _
                                    sendNUM)
            Case "票券交易明細表_交易日"
                Call Import_票券交易明細表_交易日( _
                    basePath, sendOldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, _
                    lastMap, _
                    ROCYearMonth, _
                    sendNUM)
            Case "會計資料庫及票券庫存明細表日結"
                Call Import_會計資料庫及票券庫存明細表日結( _
                    basePath, sendOldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, _
                    lastMap, _
                    ROCYearMonth, _
                    sendNUM)
            Case "會計資料庫及債券風險部位餘額表和票券風險部位餘額表"
                Call Import_會計資料庫及債券風險部位餘額表和票券風險部位餘額表( _
                    basePath, sendOldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, _
                    lastMap, _
                    ROCYearMonth, _
                    sendNUM)
            Case "會計資料庫及關帳匯率"
                Call Import_會計資料庫及關帳匯率( _
                    basePath, sendOldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, _
                    lastMap, _
                    ROCYearMonth, _
                    sendNUM)
            Case "會計資料庫及債券交易明細表及匯率表"
                Call Import_會計資料庫及債券交易明細表及匯率表( _
                    basePath, sendOldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, _
                    lastMap, _
                    ROCYearMonth, _
                    sendNUM)
            Case "票券交易明細表_交割日"
                Call Import_票券交易明細表_交割日( _
                    basePath, sendOldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, _
                    lastMap, _
                    ROCYearMonth, _
                    sendNUM)

            Case "會計資料庫及外幣債評估表及債券評價表及票券評價表及匯率表"
                Call Import_會計資料庫及外幣債評估表及債券評價表及票券評價表及匯率表( _
                    basePath, sendOldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, _
                    lastMap, _
                    ROCYearMonth, _
                    sendNUM, _
                    westernMonthEnd)

            Case "會計資料庫及匯率及債券評價及外幣債評估表"
                Call Import_會計資料庫及匯率及債券評價及外幣債評估表( _
                    basePath, sendOldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, _
                    lastMap, _
                    ROCYearMonth, _
                    sendNUM, _
                    westernMonthEnd)

            Case "無"
                Call Import_無資料( _
                    basePath, sendOldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, _
                    lastMap, _
                    ROCYearMonth, _
                    sendNUM)                    
            Case Else
                MsgBox "未知 CaseType: " & caseType & "（ReportID=" & wsRpt.Cells(i, "A").Value & "）", vbExclamation
        End Select
NextRpt:
    Next i

    MsgBox "全部報表處理完成！", vbInformation
End Sub

Private Sub Import_會計資料庫(ByVal basePath   As String, _
                             ByVal oldMon     As String, _
                             ByVal newMon     As String, _
                             ByVal rptID      As String, _
                             ByVal tplPattern As String, _
                             ByVal tplSheet   As String, _
                             ByVal impPattern As String, _
                             ByVal impSheets  As String, _
                             ByVal declTplRel As String, _
                             ByVal wsMap      As Worksheet, _
                             ByVal lastMap    As Long, _
                             ByVal ROCYearMonth As String, _
                             ByVal NUMYearMonth As String)

    Dim wbOld As Workbook, wbImp As Workbook
    Dim wbNew As Workbook, wbDecl As Workbook
    Dim arrImpF() As String, arrImpSh() As String
    Dim oldTplRel As String, newTplRel As String
    Dim tplPath As String, declTplPath As String
    Dim j As Long

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False

    '— 路徑置換 ——
    oldTplRel = Replace(tplPattern, "YYYYMM", oldMon)
    newTplRel = Replace(tplPattern, "YYYYMM", newMon)
    arrImpF   = Split(Replace(impPattern, "YYYYMM", newMon), ",")
    arrImpSh  = Split(impSheets, ",")

    tplPath     = basePath & "\" & oldTplRel
    declTplPath = basePath & "\" & declTplRel

    '— 開舊底稿 + 匯入資料 ——
    Set wbOld = Workbooks.Open(tplPath, ReadOnly:=True)

    Select Case rptId
        Case "表10"
            wbOld.Sheets("表10").Range("A2").Value = ROCYearMonth
        Case "表20"
            wbOld.Sheets("表20").Range("G3").Value = ROCYearMonth
        Case "AI430"
            wbOld.Sheets("Table1").Range("B3").Value = NUMYearMonth
    End Select    


    For j = LBound(arrImpF) To UBound(arrImpF)
        Set wbImp = Workbooks.Open(basePath & "\" & arrImpF(j), ReadOnly:=True)
        With wbOld.Sheets(tplSheet)
            .Cells.ClearContents

            wbImp.Sheets(arrImpSh(j)).UsedRange.Copy
            .Range("A1").PasteSpecial xlPasteValues
        End With
        wbImp.Close False
    Next j

    '— 另存新底稿 ——
    wbOld.SaveCopyAs basePath & "\" & newTplRel
    wbOld.Close False

    '— 刪除舊底稿檔案 ——
    Kill tplPath

    '— 貼入申報模板 ——
    Set wbNew  = Workbooks.Open(basePath & "\" & newTplRel)
    Set wbDecl = Workbooks.Open(declTplPath)

    Dim srcSh As String, rngStrings() As String
    Dim rngSrc As Range, rngDst As Range
    Dim srcAddr As Variant

    For j = 2 To lastMap
        If wsMap.Cells(j, "A").Value = rptID Then

            srcSh = wsMap.Cells(j, "B").Value
            rngStrings = Split(wsMap.Cells(j, "C").Value, ",")

            For Each srcAddr In rngStrings
                srcAddr = Trim(srcAddr)
                Set rngSrc = wbNew.Sheets(srcSh).Range(srcAddr)
                Set rngDst = wbDecl.Sheets(srcSh).Range(srcAddr)

                ' 直接以值貼值，保留大小一致
                rngDst.Value = rngSrc.Value
            Next srcAddr            
        End If
    Next j

    wbDecl.Save: wbDecl.Close False
    wbNew.Close False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.AskToUpdateLinks = True
End Sub

Private Sub Import_PNCDCAL(ByVal basePath   As String, _
                           ByVal oldMon     As String, _
                           ByVal newMon     As String, _
                           ByVal rptID      As String, _
                           ByVal tplPattern As String, _
                           ByVal tplSheet   As String, _
                           ByVal impPattern As String, _
                           ByVal impSheets  As String, _
                           ByVal declTplRel As String, _
                           ByVal wsMap      As Worksheet, _
                           ByVal lastMap    As Long, _
                           ByVal ROCYearMonth As String, _
                           ByVal NUMYearMonth As String)

    Dim wbOld As Workbook, wbImp As Workbook
    Dim wbNew As Workbook, wbDecl As Workbook
    Dim arrImpF() As String, arrImpSh() As String
    Dim oldTplRel As String, newTplRel As String
    Dim tplPath   As String, declTplPath As String
    Dim j         As Long

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False

    '— 路徑置換 ——
    oldTplRel = Replace(tplPattern, "YYYYMM", oldMon)
    newTplRel = Replace(tplPattern, "YYYYMM", newMon)
    arrImpF   = Split(Replace(impPattern, "YYYYMM", newMon), ",")
    arrImpSh  = Split(impSheets, ",")

    tplPath     = basePath & "\" & oldTplRel
    declTplPath = basePath & "\" & declTplRel

    '— 開舊底稿 + 匯入資料 ——
    Set wbOld = Workbooks.Open(tplPath, ReadOnly:=True)

    Select Case rptId
        Case "表15A"
            wbOld.Sheets("新台幣可轉讓定期存單發行、償還及餘額統計表").Range("A2").Value = ROCYearMonth
    End Select

    For j = LBound(arrImpF) To UBound(arrImpF)

        Call PNCDCAL_FormatToCSV(basePath & "\" & arrImpF(j))
        Set wbImp = Workbooks.Open(Replace(basePath & "\" & arrImpF(j), "txt", "csv"), ReadOnly:=True)
        
        With wbOld.Sheets(Trim(tplSheet))
            ' .Cells.Clear
            .Cells.ClearContents

            wbImp.Sheets(1).UsedRange.Copy
            .Range("A1").PasteSpecial xlPasteValues
        End With
        wbImp.Close False
    Next j

    '— 另存新底稿 ——
    wbOld.SaveCopyAs basePath & "\" & newTplRel
    wbOld.Close False

    '— 刪除舊底稿檔案 ——
    Kill tplPath

    '— 貼入申報模板 ——
    Set wbNew  = Workbooks.Open(basePath & "\" & newTplRel)
    Set wbDecl = Workbooks.Open(declTplPath)

    Dim srcSh As String, rngStrings() As String
    Dim rngSrc As Range, rngDst As Range
    Dim srcAddr As Variant

    For j = 2 To lastMap
        If wsMap.Cells(j, "A").Value = rptID Then

            srcSh = wsMap.Cells(j, "B").Value
            rngStrings = Split(wsMap.Cells(j, "C").Value, ",")

            For Each srcAddr In rngStrings
                srcAddr = Trim(srcAddr)
                Set rngSrc = wbNew.Sheets(srcSh).Range(srcAddr)
                Set rngDst = wbDecl.Sheets(srcSh).Range(srcAddr)

                ' 直接以值貼值，保留大小一致
                rngDst.Value = rngSrc.Value
            Next srcAddr            
        End If
    Next j

    wbDecl.Save: wbDecl.Close False
    wbNew.Close False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.AskToUpdateLinks = True
End Sub

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

Private Sub Import_票券交易明細表_交易日(ByVal basePath   As String, _
                                       ByVal oldMon     As String, _
                                       ByVal newMon     As String, _
                                       ByVal rptID      As String, _
                                       ByVal tplPattern As String, _
                                       ByVal tplSheet   As String, _
                                       ByVal impPattern As String, _
                                       ByVal impSheets  As String, _
                                       ByVal declTplRel As String, _
                                       ByVal wsMap      As Worksheet, _
                                       ByVal lastMap    As Long, _
                                       ByVal ROCYearMonth As String, _
                                       ByVal NUMYearMonth As String)
    Dim wbOld    As Workbook, wbImp As Workbook
    Dim wbNew    As Workbook, wbDecl As Workbook
    Dim arrImpF() As String, arrImpSh() As String
    Dim oldTplRel As String, newTplRel As String
    Dim tplPath   As String, declTplPath As String
    Dim j         As Long

    Dim filterFields As Variant
    Dim filterValues As Variant

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False

    '— 路徑置換 ——
    oldTplRel = Replace(tplPattern, "YYYYMM", oldMon)
    newTplRel = Replace(tplPattern, "YYYYMM", newMon)
    arrImpF   = Split(Replace(impPattern, "YYYYMM", newMon), ",")
    arrImpSh  = Split(impSheets, ",")

    tplPath     = basePath & "\" & oldTplRel
    declTplPath = basePath & "\" & declTplRel

    '— 開舊底稿 + 匯入資料 ——
    Set wbOld = Workbooks.Open(tplPath, ReadOnly:=True)

    Select Case rptId
        Case "表22"
            ' 設定匯入資料篩選條件
            filterFields = Array("交易別", "票類")
            filterValues = Array(Array("首購", "續發", "買斷", "附買回", "附賣回", "承銷發行", "代銷買入", "賣斷", "代銷發行", "代銷賣出", "附買回續做", "附賣回續作"), Array("CP1", "CP2", "TA", "同業NCD", "ABCP", "BA", "TB1", "TB2", "MN"))

            ' 設定製作報表年月份
            wbOld.Sheets("表22").Range("A2").Value = ROCYearMonth
        Case "表23"
            ' 設定匯入資料篩選條件
            filterFields = Array("交易別")
            filterValues = Array(Array("承銷發行"))

            ' 設定製作報表年月份
            wbOld.Sheets("票券利率統計表").Range("E2").Value = ROCYearMonth
    End Select

    ' 表23底稿右邊有公式，不要拿掉



    wbOld.Sheets(tplSheet).Range("A:AD").ClearContents
    For j = LBound(arrImpF) To UBound(arrImpF)
        Call ImportCsvWithFilter(basePath & "\" & arrImpF(j), wbOld.Sheets(tplSheet), wbOld.Sheets(tplSheet).Range("A1"), filterFields, filterValues)
    Next j

    '— 另存新底稿 ——
    wbOld.SaveCopyAs basePath & "\" & newTplRel
    wbOld.Close False

    '— 刪除舊底稿檔案 ——
    Kill tplPath

    '— 貼入申報模板 ——
    Set wbNew  = Workbooks.Open(basePath & "\" & newTplRel)
    Set wbDecl = Workbooks.Open(declTplPath)

    Dim srcSh As String, rngStrings() As String
    Dim rngSrc As Range, rngDst As Range
    Dim srcAddr As Variant

    For j = 2 To lastMap
        If wsMap.Cells(j, "A").Value = rptID Then

            srcSh = wsMap.Cells(j, "B").Value
            rngStrings = Split(wsMap.Cells(j, "C").Value, ",")

            For Each srcAddr In rngStrings
                srcAddr = Trim(srcAddr)
                Set rngSrc = wbNew.Sheets(srcSh).Range(srcAddr)
                Set rngDst = wbDecl.Sheets(srcSh).Range(srcAddr)

                ' 直接以值貼值，保留大小一致
                rngDst.Value = rngSrc.Value
            Next srcAddr            
        End If
    Next j

    wbDecl.Save: wbDecl.Close False
    wbNew.Close False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.AskToUpdateLinks = True
End Sub

Private Sub Import_票券交易明細表_交割日(ByVal basePath   As String, _
                                       ByVal oldMon     As String, _
                                       ByVal newMon     As String, _
                                       ByVal rptID      As String, _
                                       ByVal tplPattern As String, _
                                       ByVal tplSheet   As String, _
                                       ByVal impPattern As String, _
                                       ByVal impSheets  As String, _
                                       ByVal declTplRel As String, _
                                       ByVal wsMap      As Worksheet, _
                                       ByVal lastMap    As Long, _
                                       ByVal ROCYearMonth As String, _
                                       ByVal NUMYearMonth As String)
    Dim wbOld    As Workbook, wbImp As Workbook
    Dim wbNew    As Workbook, wbDecl As Workbook
    Dim arrImpF() As String, arrImpSh() As String
    Dim oldTplRel As String, newTplRel As String
    Dim tplPath   As String, declTplPath As String
    Dim j         As Long

    Dim filterFields As Variant
    Dim filterValues As Variant

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False

    '— 路徑置換 ——
    oldTplRel = Replace(tplPattern, "YYYYMM", oldMon)
    newTplRel = Replace(tplPattern, "YYYYMM", newMon)
    arrImpF   = Split(Replace(impPattern, "YYYYMM", newMon), ",")
    arrImpSh  = Split(impSheets, ",")

    tplPath     = basePath & "\" & oldTplRel
    declTplPath = basePath & "\" & declTplRel

    '— 開舊底稿 + 匯入資料 ——
    Set wbOld = Workbooks.Open(tplPath, ReadOnly:=True)

    Select Case rptId
        Case "表22"
            ' 設定匯入資料篩選條件
            filterFields = Array("交易別", "票類")
            filterValues = Array(Array("首購", "續發", "買斷", "附買回", "附賣回", "承銷發行", "代銷買入", "賣斷", "代銷發行", "代銷賣出", "附買回續做", "附賣回續作"), Array("CP1", "CP2", "TA", "同業NCD", "ABCP", "BA", "TB1", "TB2", "MN"))

            ' 設定製作報表年月份
            wbOld.Sheets("表22").Range("A2").Value = ROCYearMonth
            
        Case "表23"
            ' 設定匯入資料篩選條件
            filterFields = Array("交易別")
            filterValues = Array(Array("承銷發行"))

            ' 設定製作報表年月份
            wbOld.Sheets("票券利率統計表").Range("E2").Value = ROCYearMonth

        Case "AI410"
            ' 設定匯入資料篩選條件
            filterFields = Array("交易別")
            filterValues = Array(Array("首購", "承銷發行"))

            ' 設定製作報表年月份
            wbOld.Sheets("Table1").Range("B3").Value = NUMYearMonth
        Case "AI415"
            ' 設定匯入資料篩選條件
            filterFields = Array("交易別")
            filterValues = Array(Array("首購", "買斷", "附買回", "附賣回", "承銷發行", "附買回續做", "附賣回續作"))

            ' 設定製作報表年月份
            wbOld.Sheets("Table1").Range("B3").Value = NUMYearMonth
    End Select

    ' 表23底稿右邊有公式，不要拿掉

    wbOld.Sheets(tplSheet).Range("A:AD").ClearContents
    For j = LBound(arrImpF) To UBound(arrImpF)
        Call ImportCsvWithFilter(basePath & "\" & arrImpF(j), wbOld.Sheets(tplSheet), wbOld.Sheets(tplSheet).Range("A1"), filterFields, filterValues)
    Next j

    '— 另存新底稿 ——
    wbOld.SaveCopyAs basePath & "\" & newTplRel
    wbOld.Close False

    '— 刪除舊底稿檔案 ——
    Kill tplPath

    '— 貼入申報模板 ——
    Set wbNew  = Workbooks.Open(basePath & "\" & newTplRel)
    Set wbDecl = Workbooks.Open(declTplPath)

    Dim srcSh As String, rngStrings() As String
    Dim rngSrc As Range, rngDst As Range
    Dim srcAddr As Variant

    For j = 2 To lastMap
        If wsMap.Cells(j, "A").Value = rptID Then

            srcSh = wsMap.Cells(j, "B").Value
            rngStrings = Split(wsMap.Cells(j, "C").Value, ",")

            For Each srcAddr In rngStrings
                srcAddr = Trim(srcAddr)
                Set rngSrc = wbNew.Sheets(srcSh).Range(srcAddr)
                Set rngDst = wbDecl.Sheets(srcSh).Range(srcAddr)

                ' 直接以值貼值，保留大小一致
                rngDst.Value = rngSrc.Value
            Next srcAddr            
        End If
    Next j

    wbDecl.Save: wbDecl.Close False
    wbNew.Close False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.AskToUpdateLinks = True
End Sub

Private Sub Import_會計資料庫及票券庫存明細表日結(ByVal basePath   As String, _
                                               ByVal oldMon     As String, _
                                               ByVal newMon     As String, _
                                               ByVal rptID      As String, _
                                               ByVal tplPattern As String, _
                                               ByVal tplSheet   As String, _
                                               ByVal impPattern As String, _
                                               ByVal impSheets  As String, _
                                               ByVal declTplRel As String, _
                                               ByVal wsMap      As Worksheet, _
                                               ByVal lastMap    As Long, _
                                               ByVal ROCYearMonth As String, _
                                               ByVal NUMYearMonth As String)

    Dim wbOld    As Workbook, wbImp As Workbook
    Dim wbNew    As Workbook, wbDecl As Workbook
    Dim arrTplSh() As String, arrImpF() As String, arrImpSh() As String
    Dim oldTplRel As String, newTplRel As String
    Dim tplPath   As String, declTplPath As String
    Dim j         As Long

    Dim filterFields As Variant
    Dim filterValues As Variant

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False

    '— 路徑置換 ——
    oldTplRel = Replace(tplPattern, "YYYYMM", oldMon)
    newTplRel = Replace(tplPattern, "YYYYMM", newMon)
    arrTplSh = Split(tplSheet, ",")
    arrImpF   = Split(Replace(impPattern, "YYYYMM", newMon), ",")
    arrImpSh  = Split(impSheets, ",")

    tplPath     = basePath & "\" & oldTplRel
    declTplPath = basePath & "\" & declTplRel

    '— 開舊底稿 + 匯入資料 ——
    Set wbOld = Workbooks.Open(tplPath, ReadOnly:=True)

    Select Case rptId
        Case "表36", "表6"
            ' 設定匯入資料篩選條件
            filterFields = Array("票類")
            filterValues = Array(Array("CP1", "CP2", "TA", "ABCP", "BA", "TB1", "TB2", "MN"))

            ' 設定製作報表年月份
            If rptId = "表36" Then
                wbOld.Sheets("表36").Range("E3").Value = ROCYearMonth
            ElseIf rptId = "表6" Then
                wbOld.Sheets("FOA").Range("D4").Value = ROCYearMonth
            End If
        
        Case "表7"
            ' 設定匯入資料篩選條件
            filterFields = Array("發票人", "票類")
            filterValues = Array(Array("<>中租迪和", "<>中華航空"), Array("CP1", "CP2", "TA", "同業NCD", "ABCP", "BA", "TB1", "TB2", "MN"))

            ' 設定製作報表年月份
            wbOld.Sheets("FOA").Range("J4").Value = ROCYearMonth
    End Select

    ' 表23底稿右邊有公式，不要拿掉
    
    For j = LBound(arrImpF) To UBound(arrImpF)
        Select Case True
            Case Mid(arrImpF(j), InStrRev(arrImpF(j), "\") + 1) Like "票券庫存明細表(日結)*"
                wbOld.Sheets(arrTplSh(j)).Range("A:AE").ClearContents
                Call ImportCsvWithFilter(basePath & "\" & Trim(arrImpF(j)), wbOld.Sheets(arrTplSh(j)), wbOld.Sheets(arrTplSh(j)).Range("A1"), filterFields, filterValues)

            Case Mid(arrImpF(j), InStrRev(arrImpF(j), "\") + 1) Like "會計資料庫*"
                wbOld.Sheets(arrTplSh(j)).Cells.ClearContents
                Set wbImp = Workbooks.Open(basePath & "\" & Trim(arrImpF(j)), ReadOnly:=True)

                With wbOld.Sheets(arrTplSh(j))
                    wbImp.Sheets(Trim(arrImpSh(j))).UsedRange.Copy
                    .Range("A1").PasteSpecial xlPasteValues
                End With
                wbImp.Close False
        End Select        
    Next j

    '— 另存新底稿 ——
    wbOld.SaveCopyAs basePath & "\" & newTplRel
    wbOld.Close False

    '— 刪除舊底稿檔案 ——
    Kill tplPath

    If rptId <> "表6" And rptId <> "表7" Then
        '— 貼入申報模板 ——
        Set wbNew  = Workbooks.Open(basePath & "\" & newTplRel)
        Set wbDecl = Workbooks.Open(declTplPath)

        Dim srcSh As String, rngStrings() As String
        Dim rngSrc As Range, rngDst As Range
        Dim srcAddr As Variant

        For j = 2 To lastMap
            If wsMap.Cells(j, "A").Value = rptID Then

                srcSh = wsMap.Cells(j, "B").Value
                rngStrings = Split(wsMap.Cells(j, "C").Value, ",")

                For Each srcAddr In rngStrings
                    srcAddr = Trim(srcAddr)
                    Set rngSrc = wbNew.Sheets(srcSh).Range(srcAddr)
                    Set rngDst = wbDecl.Sheets(srcSh).Range(srcAddr)

                    ' 直接以值貼值，保留大小一致
                    rngDst.Value = rngSrc.Value
                Next srcAddr            
            End If
        Next j

        wbDecl.Save: wbDecl.Close False
        wbNew.Close False
    End If

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.AskToUpdateLinks = True
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

Private Sub Import_會計資料庫及債券交易明細表及匯率表(ByVal basePath   As String, _
                                                   ByVal oldMon     As String, _
                                                   ByVal newMon     As String, _
                                                   ByVal rptID      As String, _
                                                   ByVal tplPattern As String, _
                                                   ByVal tplSheet   As String, _
                                                   ByVal impPattern As String, _
                                                   ByVal impSheets  As String, _
                                                   ByVal declTplRel As String, _
                                                   ByVal wsMap      As Worksheet, _
                                                   ByVal lastMap    As Long, _
                                                   ByVal ROCYearMonth As String, _
                                                   ByVal NUMYearMonth As String)

    Dim wbOld    As Workbook, wbImp As Workbook
    Dim wbNew    As Workbook, wbDecl As Workbook
    Dim arrTplSh() As String, arrImpF() As String, arrImpSh() As String
    Dim oldTplRel As String, newTplRel As String
    Dim tplPath   As String, declTplPath As String
    Dim j         As Long

    Dim filterFields As Variant
    Dim filterValues As Variant

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False

    '— 路徑置換 ——
    oldTplRel = Replace(tplPattern, "YYYYMM", oldMon)
    newTplRel = Replace(tplPattern, "YYYYMM", newMon)
    arrTplSh = Split(tplSheet, ",")

    ' ###修改開始###
    Dim impPatternArr() As String
    Dim f          As String
    Dim searchPath As String

    searchPath = basePath & "\批次報表\*cm2610.xls"

    impPatternArr = Split(impPattern, ",")
    ReDim arrImpF(LBound(impPatternArr) To UBound(impPatternArr))

    For j = LBound(impPatternArr) To UBound(impPatternArr)
        If impPatternArr(j) Like "*CloseRate*" Then
            f = Dir(searchPath)
            If f <> "" Then
                arrImpF(j) = "批次報表\" & f
            Else
                Debug.Print "找不到 CloseRate 檔案：" & searchPath
            End If
        Else
            arrImpF(j) = Replace(impPatternArr(j), "YYYYMM", newMon)
        End If
    Next j

    ' ###修改結束###

    arrImpSh  = Split(impSheets, ",")

    tplPath     = basePath & "\" & oldTplRel
    declTplPath = basePath & "\" & declTplRel

    '— 開舊底稿 + 匯入資料 ——
    Set wbOld = Workbooks.Open(tplPath, ReadOnly:=True)

    Select Case rptId
        
        Case "AI405"
            ' 設定匯入資料篩選條件
            filterFields = Array("交易類別")
            filterValues = Array(Array("發行前買斷", "買斷", "附買回", "附賣回", "發行前賣斷", "賣斷", "首購", "附買回續作", "附賣回續作"))

            ' 設定製作報表年月份
            wbOld.Sheets("Table1").Range("B3").Value = NUMYearMonth
    End Select    

    For j = LBound(arrImpF) To UBound(arrImpF)
        Select Case True
            Case Mid(arrImpF(j), InStrRev(arrImpF(j), "\") + 1) Like "債券交易明細表*"
                wbOld.Sheets(arrTplSh(j)).Range("A:AC").ClearContents
                Call ImportCsvWithFilter(basePath & "\" & Trim(arrImpF(j)), wbOld.Sheets(arrTplSh(j)), wbOld.Sheets(arrTplSh(j)).Range("A1"), filterFields, filterValues)

            Case Mid(arrImpF(j), InStrRev(arrImpF(j), "\") + 1) Like "會計資料庫*"
                wbOld.Sheets(arrTplSh(j)).Cells.ClearContents
                Set wbImp = Workbooks.Open(basePath & "\" & Trim(arrImpF(j)), ReadOnly:=True)

                With wbOld.Sheets(arrTplSh(j))
                    wbImp.Sheets(Trim(arrImpSh(j))).UsedRange.Copy
                    .Range("A1").PasteSpecial xlPasteValues
                End With
                wbImp.Close False
                
            Case Mid(arrImpF(j), InStrRev(arrImpF(j), "\") + 1) Like "*cm2610*"
                wbOld.Sheets(arrTplSh(j)).Range("A1:Z30").ClearContents
                Call ImportCloseRate(basePath & "\" & Trim(arrImpF(j)), wbOld.Sheets(arrTplSh(j)))                
        End Select
    Next j

    '— 另存新底稿 ——
    wbOld.SaveCopyAs basePath & "\" & newTplRel
    wbOld.Close False

    '— 刪除舊底稿檔案 ——
    Kill tplPath

    '— 貼入申報模板 ——
    Set wbNew  = Workbooks.Open(basePath & "\" & newTplRel)
    Set wbDecl = Workbooks.Open(declTplPath)

    Dim srcSh As String, rngStrings() As String
    Dim rngSrc As Range, rngDst As Range
    Dim srcAddr As Variant

    For j = 2 To lastMap
        If wsMap.Cells(j, "A").Value = rptID Then

            srcSh = wsMap.Cells(j, "B").Value
            rngStrings = Split(wsMap.Cells(j, "C").Value, ",")

            For Each srcAddr In rngStrings
                srcAddr = Trim(srcAddr)
                Set rngSrc = wbNew.Sheets(srcSh).Range(srcAddr)
                Set rngDst = wbDecl.Sheets(srcSh).Range(srcAddr)

                ' 直接以值貼值，保留大小一致
                rngDst.Value = rngSrc.Value
            Next srcAddr            
        End If
    Next j

    wbDecl.Save: wbDecl.Close False
    wbNew.Close False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.AskToUpdateLinks = True
End Sub

Private Sub Import_會計資料庫及債券風險部位餘額表和票券風險部位餘額表(ByVal basePath   As String, _
                                                                 ByVal oldMon     As String, _
                                                                 ByVal newMon     As String, _
                                                                 ByVal rptID      As String, _
                                                                 ByVal tplPattern As String, _
                                                                 ByVal tplSheet   As String, _
                                                                 ByVal impPattern As String, _
                                                                 ByVal impSheets  As String, _
                                                                 ByVal declTplRel As String, _
                                                                 ByVal wsMap      As Worksheet, _
                                                                 ByVal lastMap    As Long, _
                                                                 ByVal ROCYearMonth As String, _
                                                                 ByVal NUMYearMonth As String)

    Dim wbOld    As Workbook, wbImp As Workbook
    Dim wbNew    As Workbook, wbDecl As Workbook
    Dim arrTplSh() As String, arrImpF() As String, arrImpSh() As String
    Dim oldTplRel As String, newTplRel As String
    Dim tplPath   As String, declTplPath As String
    Dim j         As Long

    Dim filterFields As Variant
    Dim filterValues As Variant

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False

    '— 路徑置換 ——
    oldTplRel = Replace(tplPattern, "YYYYMM", oldMon)
    newTplRel = Replace(tplPattern, "YYYYMM", newMon)
    arrTplSh = Split(tplSheet, ",")
    arrImpF   = Split(Replace(impPattern, "YYYYMM", newMon), ",")
    arrImpSh  = Split(impSheets, ",")

    tplPath     = basePath & "\" & oldTplRel
    declTplPath = basePath & "\" & declTplRel 

    '— 開舊底稿 + 匯入資料 ——
    Set wbOld = Workbooks.Open(tplPath, ReadOnly:=True)

    Select Case rptId
        Case "AI233"
            ' 設定製作報表年月份
            wbOld.Sheets("Table1").Range("B3").Value = NUMYearMonth
    End Select


    For j = LBound(arrImpF) To UBound(arrImpF)
        wbOld.Sheets(arrTplSh(j)).Cells.ClearContents
        Set wbImp = Workbooks.Open(basePath & "\" & Trim(arrImpF(j)), ReadOnly:=True)

        With wbOld.Sheets(arrTplSh(j))
            wbImp.Sheets(Trim(arrImpSh(j))).UsedRange.Copy
            .Range("A1").PasteSpecial xlPasteValues
        End With
        wbImp.Close False
    Next j

    '— 另存新底稿 ——
    wbOld.SaveCopyAs basePath & "\" & newTplRel
    wbOld.Close False

    '— 刪除舊底稿檔案 ——
    Kill tplPath

    '— 貼入申報模板 ——
    Set wbNew  = Workbooks.Open(basePath & "\" & newTplRel)
    Set wbDecl = Workbooks.Open(declTplPath)

    Dim srcSh As String, rngStrings() As String
    Dim rngSrc As Range, rngDst As Range
    Dim srcAddr As Variant

    For j = 2 To lastMap
        If wsMap.Cells(j, "A").Value = rptID Then

            srcSh = wsMap.Cells(j, "B").Value
            rngStrings = Split(wsMap.Cells(j, "C").Value, ",")

            For Each srcAddr In rngStrings
                srcAddr = Trim(srcAddr)
                Set rngSrc = wbNew.Sheets(srcSh).Range(srcAddr)
                Set rngDst = wbDecl.Sheets(srcSh).Range(srcAddr)

                ' 直接以值貼值，保留大小一致
                rngDst.Value = rngSrc.Value
            Next srcAddr            
        End If
    Next j

    wbDecl.Save: wbDecl.Close False
    wbNew.Close False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.AskToUpdateLinks = True
End Sub

Private Sub Import_會計資料庫及關帳匯率(ByVal basePath   As String, _
                                      ByVal oldMon     As String, _
                                      ByVal newMon     As String, _
                                      ByVal rptID      As String, _
                                      ByVal tplPattern As String, _
                                      ByVal tplSheet   As String, _
                                      ByVal impPattern As String, _
                                      ByVal impSheets  As String, _
                                      ByVal declTplRel As String, _
                                      ByVal wsMap      As Worksheet, _
                                      ByVal lastMap    As Long, _
                                      ByVal ROCYearMonth As String, _
                                      ByVal NUMYearMonth As String)

    Dim wbOld    As Workbook, wbImp As Workbook
    Dim wbNew    As Workbook, wbDecl As Workbook
    Dim arrTplSh() As String, arrImpF() As String, arrImpSh() As String
    Dim oldTplRel As String, newTplRel As String
    Dim tplPath   As String, declTplPath As String
    Dim j         As Long

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False

    '— 路徑置換 ——
    oldTplRel = Replace(tplPattern, "YYYYMM", oldMon)
    newTplRel = Replace(tplPattern, "YYYYMM", newMon)
    arrTplSh = Split(tplSheet, ",")

    ' ###修改開始###
    Dim impPatternArr() As String
    Dim f          As String
    Dim searchPath As String

    searchPath = basePath & "\批次報表\*cm2610.xls"

    impPatternArr = Split(impPattern, ",")
    ReDim arrImpF(LBound(impPatternArr) To UBound(impPatternArr))

    For j = LBound(impPatternArr) To UBound(impPatternArr)
        If impPatternArr(j) Like "*CloseRate*" Then
            f = Dir(searchPath)
            If f <> "" Then
                arrImpF(j) = "批次報表\" & f
            Else
                Debug.Print "找不到 CloseRate 檔案：" & searchPath
            End If
        Else
            arrImpF(j) = Replace(impPatternArr(j), "YYYYMM", newMon)
        End If
    Next j

    ' ###修改結束###

    arrImpSh  = Split(impSheets, ",")

    tplPath     = basePath & "\" & oldTplRel
    declTplPath = basePath & "\" & declTplRel

    '— 開舊底稿 + 匯入資料 ——
    Set wbOld = Workbooks.Open(tplPath, ReadOnly:=True)

    Select Case rptId
        Case "AI345"
            wbOld.Sheets("AI345_NEW").Range("A2").Value = NUMYearMonth
    End Select
    
    For j = LBound(arrImpF) To UBound(arrImpF)

        Select Case True
            Case Mid(arrImpF(j), InStrRev(arrImpF(j), "\") + 1) Like "*cm2610*"
                wbOld.Sheets(arrTplSh(j)).Range("A1:Z30").ClearContents
                Call ImportCloseRate(basePath & "\" & Trim(arrImpF(j)), wbOld.Sheets(arrTplSh(j)))

            Case Mid(arrImpF(j), InStrRev(arrImpF(j), "\") + 1) Like "會計資料庫*"
                wbOld.Sheets(arrTplSh(j)).Cells.ClearContents
                Set wbImp = Workbooks.Open(basePath & "\" & Trim(arrImpF(j)), ReadOnly:=True)

                With wbOld.Sheets(arrTplSh(j))
                    wbImp.Sheets(Trim(arrImpSh(j))).UsedRange.Copy
                    .Range("A1").PasteSpecial xlPasteValues
                End With
                wbImp.Close False                
        End Select        
    Next j

    '— 另存新底稿 ——
    wbOld.SaveCopyAs basePath & "\" & newTplRel
    wbOld.Close False

    '— 刪除舊底稿檔案 ——
    Kill tplPath

    '— 貼入申報模板 ——
    ' Set wbNew  = Workbooks.Open(basePath & "\" & newTplRel)
    ' Set wbDecl = Workbooks.Open(declTplPath)

    ' Dim srcSh As String, rngStrings() As String
    ' Dim rngSrc As Range, rngDst As Range
    ' Dim srcAddr As Variant

    ' For j = 2 To lastMap
    '     If wsMap.Cells(j, "A").Value = rptID Then

    '         srcSh = wsMap.Cells(j, "B").Value
    '         rngStrings = Split(wsMap.Cells(j, "C").Value, ",")

    '         For Each srcAddr In rngStrings
    '             srcAddr = Trim(srcAddr)
    '             Set rngSrc = wbNew.Sheets(srcSh).Range(srcAddr)
    '             Set rngDst = wbDecl.Sheets(srcSh).Range(srcAddr)

    '             ' 直接以值貼值，保留大小一致
    '             rngDst.Value = rngSrc.Value
    '         Next srcAddr            
    '     End If
    ' Next j

    ' wbDecl.Save: wbDecl.Close False
    ' wbNew.Close False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.AskToUpdateLinks = True
End Sub


Public Sub ImportCloseRate(ByVal csvPath As String, _
                           ByVal targetWS As Worksheet)
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

    ' 反向遍歷刪除列
    For i = lastRow To 2 Step -1
        isRowDelete = False
        If IsEmpty(shCsv.Cells(i, "A").Value) Or IsEmpty(shCsv.Cells(i, "B").Value) Or IsEmpty(shCsv.Cells(i, "C").Value) Or _
           Left(shCsv.Cells(i, "A").Value, 4) = "經副襄理" Or shCsv.Cells(i, "A").Value = "USD" Then
            isRowDelete = True
        End If

        ' Delete Row
        If isRowDelete Then shCsv.Rows(i).Delete
    Next i

    On Error Resume Next
    shCsv.UsedRange.Copy
    On Error GoTo 0
    targetWS.Range("A1").PasteSpecial xlPasteValues
    ' 5. 關閉 CSV
    wbCsv.Close SaveChanges:=False
End Sub


Private Sub Import_會計資料庫及外幣債評估表及債券評價表及票券評價表及匯率表(ByVal basePath   As String, _
                                                                       ByVal oldMon     As String, _
                                                                       ByVal newMon     As String, _
                                                                       ByVal rptID      As String, _
                                                                       ByVal tplPattern As String, _
                                                                       ByVal tplSheet   As String, _
                                                                       ByVal impPattern As String, _
                                                                       ByVal impSheets  As String, _
                                                                       ByVal declTplRel As String, _
                                                                       ByVal wsMap      As Worksheet, _
                                                                       ByVal lastMap    As Long, _
                                                                       ByVal ROCYearMonth As String, _
                                                                       ByVal NUMYearMonth As String, _
                                                                       ByVal westernMonthEnd As String)

    Dim wbOld    As Workbook, wbImp As Workbook
    Dim wbNew    As Workbook, wbDecl As Workbook
    Dim arrTplSh() As String, arrImpF() As String, arrImpSh() As String
    Dim oldTplRel As String, newTplRel As String
    Dim tplPath   As String, declTplPath As String
    Dim j         As Long

    Dim filterFields As Variant
    Dim filterValues As Variant

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False

    '— 路徑置換 ——
    oldTplRel = Replace(tplPattern, "YYYYMM", oldMon)
    newTplRel = Replace(tplPattern, "YYYYMM", newMon)
    arrTplSh = Split(tplSheet, ",")

    ' ###修改開始###
    Dim impPatternArr() As String
    Dim f          As String
    Dim searchPath As String

    searchPath = basePath & "\批次報表\*cm2610.xls"

    impPatternArr = Split(impPattern, ",")
    ReDim arrImpF(LBound(impPatternArr) To UBound(impPatternArr))

    For j = LBound(impPatternArr) To UBound(impPatternArr)
        If impPatternArr(j) Like "*外幣債損益評估表(月底)對AC5100B*" Then
            arrImpF(j) = Replace(impPatternArr(j), "YYYYMM", westernMonthEnd)
        ElseIf impPatternArr(j) Like "*CloseRate*" Then
            f = Dir(searchPath)
            If f <> "" Then
                arrImpF(j) = "批次報表\" & f
            Else
                Debug.Print "找不到 CloseRate 檔案：" & searchPath
            End If
        Else
            arrImpF(j) = Replace(impPatternArr(j), "YYYYMM", newMon)
        End If
    Next j

    ' ###修改結束###

    arrImpSh  = Split(impSheets, ",")

    tplPath     = basePath & "\" & oldTplRel
    declTplPath = basePath & "\" & declTplRel

    '— 開舊底稿 + 匯入資料 ——
    Set wbOld = Workbooks.Open(tplPath, ReadOnly:=True)

    
    Select Case rptId
        Case "AI601"
            ' 設定匯入資料篩選條件
            filterFields = Array("帳務目的")
            filterValues = Array(Array("攤銷後成本衡量"))            

            ' 設定製作報表年月份
            wbOld.Sheets("Table1").Range("B3").Value = NUMYearMonth
    End Select
    
    For j = LBound(arrImpF) To UBound(arrImpF)

        Select Case True
            Case Mid(arrImpF(j), InStrRev(arrImpF(j), "\") + 1) Like "*cm2610*"
                wbOld.Sheets(arrTplSh(j)).Range("A1:Z30").ClearContents
                Call ImportCloseRate(basePath & "\" & Trim(arrImpF(j)), wbOld.Sheets(arrTplSh(j)))

            Case Mid(arrImpF(j), InStrRev(arrImpF(j), "\") + 1) Like "會計資料庫*"
                wbOld.Sheets(arrTplSh(j)).Cells.ClearContents
                Set wbImp = Workbooks.Open(basePath & "\" & Trim(arrImpF(j)), ReadOnly:=True)

                With wbOld.Sheets(arrTplSh(j))
                    wbImp.Sheets(Trim(arrImpSh(j))).UsedRange.Copy
                    .Range("A1").PasteSpecial xlPasteValues
                End With
                wbImp.Close False

            Case Mid(arrImpF(j), InStrRev(arrImpF(j), "\") + 1) Like "債券評價表*"
                wbOld.Sheets(arrTplSh(j)).Range("A1:V50").ClearContents
                Call ImportCsvWithFilter(basePath & "\" & Trim(arrImpF(j)), wbOld.Sheets(arrTplSh(j)), wbOld.Sheets(arrTplSh(j)).Range("A1"), filterFields, filterValues)

            Case Mid(arrImpF(j), InStrRev(arrImpF(j), "\") + 1) Like "票券評價表*"
                wbOld.Sheets(arrTplSh(j)).Range("A1:V50").ClearContents
                Call ImportCsvWithFilter(basePath & "\" & Trim(arrImpF(j)), wbOld.Sheets(arrTplSh(j)), wbOld.Sheets(arrTplSh(j)).Range("A1"), filterFields, filterValues)
                
            Case Mid(arrImpF(j), InStrRev(arrImpF(j), "\") + 1) Like "*外幣債損益評估表(月底)*"
                wbOld.Sheets(arrTplSh(j)).Cells.ClearContents
                Set wbImp = Workbooks.Open(basePath & "\" & Trim(arrImpF(j)), ReadOnly:=True)

                With wbOld.Sheets(arrTplSh(j))
                    wbImp.Sheets(Trim(arrImpSh(j))).UsedRange.Copy
                    .Range("A1").PasteSpecial xlPasteValues
                End With
                wbImp.Close False
        End Select        
    Next j

    '— 另存新底稿 ——
    wbOld.SaveCopyAs basePath & "\" & newTplRel
    wbOld.Close False

    '— 刪除舊底稿檔案 ——
    Kill tplPath

    '— 貼入申報模板 ——
    Set wbNew  = Workbooks.Open(basePath & "\" & newTplRel)
    Set wbDecl = Workbooks.Open(declTplPath)

    Dim srcSh As String, rngStrings() As String
    Dim rngSrc As Range, rngDst As Range
    Dim srcAddr As Variant
    
    For j = 2 To lastMap
        If wsMap.Cells(j, "A").Value = rptID Then

            srcSh = wsMap.Cells(j, "B").Value
            rngStrings = Split(wsMap.Cells(j, "C").Value, ",")

            For Each srcAddr In rngStrings
                srcAddr = Trim(srcAddr)

                Set rngSrc = wbNew.Sheets(Trim(srcSh)).Range(srcAddr)
                Set rngDst = wbDecl.Sheets(Trim(srcSh)).Range(srcAddr)

                ' 直接以值貼值，保留大小一致
                rngDst.Value = rngSrc.Value
            Next srcAddr
        End If
    Next j

    wbDecl.Save: wbDecl.Close False
    wbNew.Close False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.AskToUpdateLinks = True
End Sub

Private Sub Import_會計資料庫及匯率及債券評價及外幣債評估表(ByVal basePath   As String, _
                                                        ByVal oldMon     As String, _
                                                        ByVal newMon     As String, _
                                                        ByVal rptID      As String, _
                                                        ByVal tplPattern As String, _
                                                        ByVal tplSheet   As String, _
                                                        ByVal impPattern As String, _
                                                        ByVal impSheets  As String, _
                                                        ByVal declTplRel As String, _
                                                        ByVal wsMap      As Worksheet, _
                                                        ByVal lastMap    As Long, _
                                                        ByVal ROCYearMonth As String, _
                                                        ByVal NUMYearMonth As String, _
                                                        ByVal westernMonthEnd As String)

    Dim wbOld    As Workbook, wbImp As Workbook
    Dim wbNew    As Workbook, wbDecl As Workbook
    Dim arrTplSh() As String, arrImpF() As String, arrImpSh() As String
    Dim oldTplRel As String, newTplRel As String
    Dim tplPath   As String, declTplPath As String
    Dim j         As Long

    Dim filterFields As Variant
    Dim filterValues As Variant    

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False

    '— 路徑置換 ——
    oldTplRel = Replace(tplPattern, "YYYYMM", oldMon)
    newTplRel = Replace(tplPattern, "YYYYMM", newMon)
    arrTplSh = Split(tplSheet, ",")

    ' ###修改開始###
    Dim impPatternArr() As String
    Dim f          As String
    Dim searchPath As String

    searchPath = basePath & "\批次報表\*cm2610.xls"    

    impPatternArr = Split(impPattern, ",")
    ReDim arrImpF(LBound(impPatternArr) To UBound(impPatternArr))

    For j = LBound(impPatternArr) To UBound(impPatternArr)
        If impPatternArr(j) Like "*外幣債損益評估表(月底)對AC5100B*" Then
            arrImpF(j) = Replace(impPatternArr(j), "YYYYMM", westernMonthEnd)
        ElseIf impPatternArr(j) Like "*CloseRate*" Then
            f = Dir(searchPath)
            If f <> "" Then
                arrImpF(j) = "批次報表\" & f
            Else
                Debug.Print "找不到 CloseRate 檔案：" & searchPath
            End If            
        Else
            arrImpF(j) = Replace(impPatternArr(j), "YYYYMM", newMon)
        End If
    Next j

    ' ###修改結束###

    arrImpSh  = Split(impSheets, ",")

    tplPath     = basePath & "\" & oldTplRel
    declTplPath = basePath & "\" & declTplRel

    '— 開舊底稿 + 匯入資料 ——
    Set wbOld = Workbooks.Open(tplPath, ReadOnly:=True)

    Select Case rptId
        Case "AI605"
            ' 設定匯入資料篩選條件
            filterFields = Array("帳務目的")
            filterValues = Array(Array("攤銷後成本衡量"))

            ' 設定製作報表年月份
            wbOld.Sheets("Table1").Range("B3").Value = NUMYearMonth
    End Select
    
    For j = LBound(arrImpF) To UBound(arrImpF)

        Select Case True
            Case Mid(arrImpF(j), InStrRev(arrImpF(j), "\") + 1) Like "*cm2610*"
                wbOld.Sheets(arrTplSh(j)).Range("A1:Z30").ClearContents
                Call ImportCloseRate(basePath & "\" & Trim(arrImpF(j)), wbOld.Sheets(arrTplSh(j)))

            Case Mid(arrImpF(j), InStrRev(arrImpF(j), "\") + 1) Like "會計資料庫*"
                wbOld.Sheets(arrTplSh(j)).Cells.ClearContents
                Set wbImp = Workbooks.Open(basePath & "\" & Trim(arrImpF(j)), ReadOnly:=True)

                With wbOld.Sheets(arrTplSh(j))
                    wbImp.Sheets(Trim(arrImpSh(j))).UsedRange.Copy
                    .Range("A1").PasteSpecial xlPasteValues
                End With
                wbImp.Close False

            Case Mid(arrImpF(j), InStrRev(arrImpF(j), "\") + 1) Like "債券評價表*"
                wbOld.Sheets(arrTplSh(j)).Range("A:V").ClearContents

                Call ImportCsvWithFilter(basePath & "\" & Trim(arrImpF(j)), wbOld.Sheets(arrTplSh(j)), wbOld.Sheets(arrTplSh(j)).Range("A1"), filterFields, filterValues)                

            Case Mid(arrImpF(j), InStrRev(arrImpF(j), "\") + 1) Like "票券評價表*"
                wbOld.Sheets(arrTplSh(j)).Range("A1:V50").ClearContents
                Set wbImp = Workbooks.Open(basePath & "\" & Trim(arrImpF(j)), ReadOnly:=True)

                With wbOld.Sheets(arrTplSh(j))
                    wbImp.Sheets(Trim(arrImpSh(j))).UsedRange.Copy
                    .Range("A1").PasteSpecial xlPasteValues
                End With
                wbImp.Close False
                
            Case Mid(arrImpF(j), InStrRev(arrImpF(j), "\") + 1) Like "*外幣債損益評估表(月底)*"
                Call ImportFXDebtEvaluation(basePath & "\" & Trim(arrImpF(j)))
                wbOld.Sheets(arrTplSh(j)).Range("A:AF").ClearContents
                Set wbImp = Workbooks.Open(basePath & "\" & Trim(arrImpF(j)), ReadOnly:=True)

                With wbOld.Sheets(arrTplSh(j))
                    wbImp.Sheets(Trim(arrImpSh(j))).UsedRange.Copy
                    .Range("A1").PasteSpecial xlPasteValues
                End With
                wbImp.Close False
        End Select        
    Next j

    '— 另存新底稿 ——
    wbOld.SaveCopyAs basePath & "\" & newTplRel
    wbOld.Close False

    '— 刪除舊底稿檔案 ——
    Kill tplPath

    '— 貼入申報模板 ——
    Set wbNew  = Workbooks.Open(basePath & "\" & newTplRel)
    Set wbDecl = Workbooks.Open(declTplPath)

    Dim srcSh As String, rngStrings() As String
    Dim rngSrc As Range, rngDst As Range
    Dim srcAddr As Variant
    
    For j = 2 To lastMap
        If wsMap.Cells(j, "A").Value = rptID Then

            srcSh = wsMap.Cells(j, "B").Value
            rngStrings = Split(wsMap.Cells(j, "C").Value, ",")

            For Each srcAddr In rngStrings
                srcAddr = Trim(srcAddr)

                Set rngSrc = wbNew.Sheets(Trim(srcSh)).Range(srcAddr)
                Set rngDst = wbDecl.Sheets(Trim(srcSh)).Range(srcAddr)

                ' 直接以值貼值，保留大小一致
                rngDst.Value = rngSrc.Value
            Next srcAddr
        End If
    Next j

    wbDecl.Save: wbDecl.Close False
    wbNew.Close False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.AskToUpdateLinks = True
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

Private Sub Import_無資料(ByVal basePath   As String, _
                         ByVal oldMon     As String, _
                         ByVal newMon     As String, _
                         ByVal rptID      As String, _
                         ByVal tplPattern As String, _
                         ByVal tplSheet   As String, _
                         ByVal impPattern As String, _
                         ByVal impSheets  As String, _
                         ByVal declTplRel As String, _
                         ByVal wsMap      As Worksheet, _
                         ByVal lastMap    As Long, _
                         ByVal ROCYearMonth As String, _
                         ByVal NUMYearMonth As String)

    Dim wbOld    As Workbook, wbImp As Workbook
    Dim wbNew    As Workbook, wbDecl As Workbook
    Dim arrTplSh() As String, arrImpF() As String, arrImpSh() As String
    Dim oldTplRel As String, newTplRel As String
    Dim tplPath   As String, declTplPath As String
    Dim j         As Long

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False

    '— 路徑置換 ——
    oldTplRel = Replace(tplPattern, "YYYYMM", oldMon)
    newTplRel = Replace(tplPattern, "YYYYMM", newMon)
    arrTplSh = Split(tplSheet, ",")
    arrImpF   = Split(Replace(impPattern, "YYYYMM", newMon), ",")
    arrImpSh  = Split(impSheets, ",")

    tplPath     = basePath & "\" & oldTplRel
    declTplPath = basePath & "\" & declTplRel

    '— 開舊底稿 + 匯入資料 ——
    Set wbOld = Workbooks.Open(tplPath, ReadOnly:=True)    

    Select Case rptId
        Case "表15B"
            ' 設定製作報表年月份
            wbOld.Sheets("外幣可轉讓定期存單發行、償還及餘額統計表").Range("A2").Value = ROCYearMonth

        Case "表16"
            ' 設定製作報表年月份
            wbOld.Sheets("可轉讓定期存單每日持有額").Range("B2").Value = ROCYearMonth

        Case "表24"
            ' 設定製作報表年月份
            wbOld.Sheets("表24").Range("E2").Value = ROCYearMonth

        Case "表27"
            ' 設定製作報表年月份
            wbOld.Sheets("表27").Range("D3").Value = ROCYearMonth

        Case "AI816"
            ' 設定製作報表年月份
            wbOld.Sheets("Table1").Range("B3").Value = NUMYearMonth

        Case "AI271"
            ' 設定製作報表年月份
            wbOld.Sheets("Table1").Range("B3").Value = NUMYearMonth

        Case "AI272"
            ' 設定製作報表年月份
            wbOld.Sheets("Table1").Range("B3").Value = NUMYearMonth

        Case "AI273"
            ' 設定製作報表年月份
            wbOld.Sheets("Table1").Range("B3").Value = NUMYearMonth

        Case "AI281"
            ' 設定製作報表年月份
            wbOld.Sheets("Table1").Range("B3").Value = NUMYearMonth

        Case "AI282"
            ' 設定製作報表年月份
            wbOld.Sheets("Table1").Range("B3").Value = NUMYearMonth
    End Select


    '— 另存新底稿 ——
    wbOld.SaveCopyAs basePath & "\" & newTplRel
    wbOld.Close False

    '— 刪除舊底稿檔案 ——
    Kill tplPath

    '— 貼入申報模板 ——
    Set wbNew  = Workbooks.Open(basePath & "\" & newTplRel)
    Set wbDecl = Workbooks.Open(declTplPath)

    Dim srcSh As String, rngStrings() As String
    Dim rngSrc As Range, rngDst As Range
    Dim srcAddr As Variant

    For j = 2 To lastMap
        If wsMap.Cells(j, "A").Value = rptID Then

            srcSh = wsMap.Cells(j, "B").Value
            rngStrings = Split(wsMap.Cells(j, "C").Value, ",")

            For Each srcAddr In rngStrings
                srcAddr = Trim(srcAddr)
                Set rngSrc = wbNew.Sheets(srcSh).Range(srcAddr)
                Set rngDst = wbDecl.Sheets(srcSh).Range(srcAddr)

                ' 直接以值貼值，保留大小一致
                rngDst.Value = rngSrc.Value
            Next srcAddr            
        End If
    Next j

    wbDecl.Save: wbDecl.Close False
    wbNew.Close False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.AskToUpdateLinks = True   
End Sub
