

以下是在既有專案基礎上，加入 CaseType 分派機制的完整專案說明與程式碼。

⸻

一、主要變更概覽
	1.	在 ReportsConfig 增加一欄 CaseType，用來決定這張報表要呼叫哪一支 Sub
	2.	在主流程 ProcessAllReports，讀入 CaseType，以 Select Case 分派呼叫不同的子程序
	3.	保留原來 ProcessStandardReport（即先前的 ProcessSingleReport），並示範新增一支 ProcessSpecialReport：在匯入前做特殊處理

⸻

二、更新後的 ReportsConfig 欄位

欄位	說明
A:ReportID	報表識別碼
B:TplPath	底稿路徑範本（含 YYYYMM）
C:TplSheet	底稿要更新的分頁名稱
D:ImpPath	匯入檔案路徑範本（含 YYYYMM，多筆逗號分隔）
E:ImpSheets	匯入檔案對應分頁名稱（逗號分隔）
F:DeclTpl	申報模板路徑（不含年月）
G:Freq	月報／季報
H:CaseType	處理類型（如 Standard、Special1）

範例一行：
表10 |
央行\YYYYMM 表10…新版.xls |
E |
批次…-YYYYMM.xlsx,… |
餘額E (2),E |
央行\申報模板.xlsx |
月報 |
Special1

⸻

三、完整 VBA 程式碼

Option Explicit

'— 1. 讀取 YearMonth，計算 oldMon/newMon ——
' Private Sub GetMonths(ByRef oldMon As String, ByRef newMon As String)
'     Dim ym As String, ry As Integer, m As Integer
'     ym = ThisWorkbook.Names("YearMonth").RefersToRange.Value
'     newMon = ym
'     ry = CInt(Left(ym, Len(ym) - 2))
'     m  = CInt(Right(ym, 2)) - 1
'     If m = 0 Then ry = ry - 1: m = 12
'     oldMon = CStr(ry) & Format(m, "00")
' End Sub

' '—— 公用：開檔 —— 
' Function OpenWB( _
'     ByVal fullPath As String, _
'     Optional ByVal readOnly As Boolean = True, _
'     Optional ByVal updateLinks As XlUpdateLinks = xlUpdateLinksNever _
' ) As Workbook
'     ' 可統一設定 ReadOnly、連結更新行為、IgnoreReadOnlyRecommended 等
'     Set OpenWB = Workbooks.Open( _
'         Filename:=fullPath, _
'         ReadOnly:=readOnly, _
'         UpdateLinks:=updateLinks, _
'         IgnoreReadOnlyRecommended:=True _
'     )
' End Function

' '—— 公用：存檔並關閉 —— 
' Sub SaveCloseWB( _
'     ByVal wb As Workbook, _
'     Optional ByVal saveCopyPath As String = "", _
'     Optional ByVal saveOriginal As Boolean = False _
' )
'     If saveCopyPath <> "" Then
'         wb.SaveCopyAs Filename:=saveCopyPath
'     ElseIf saveOriginal Then
'         wb.Save
'     End If
'     wb.Close SaveChanges:=False
' End Sub

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

'— 2. 主流程：依 CaseType 分派 ——
Sub ProcessAllReports()
    Dim wbCtl    As Workbook
    Dim wsRpt    As Worksheet, wsMap As Worksheet
    Dim basePath As String
    Dim lastRpt  As Long, lastMap As Long
    Dim oldMon   As String, newMon As String
    Dim ROCYearMonth As String, NUMYearMonth As String, 
    Dim i        As Long, caseType As String

    Call GetMonths(oldMon, newMon)

    ROCYearMonth = ConvertToROCFormat(newMon, "ROC")
    NUMYearMonth = ConvertToROCFormat(newMon, "NUM")

    Set wbCtl    = ThisWorkbook
    Set wsRpt    = wbCtl.Sheets("ReportsConfig")
    Set wsMap    = wbCtl.Sheets("Mappings")
    basePath     = wbCtl.Path
    lastRpt      = wsRpt.Cells(wsRpt.Rows.Count, "A").End(xlUp).Row
    lastMap      = wsMap.Cells(wsMap.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRpt
        caseType = wsRpt.Cells(i, "H").Value
        Select Case LCase(caseType)
            Case "會計資料庫"
                Call Import_會計資料庫( _
                    basePath, oldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, _
                    lastMap, _
                    ROCYearMonth, _
                    NUMYearMonth)
            Case "PNCDCAL"
                Call Import_PNCDCAL(basePath, _
                                    oldMon, _
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
                                    NUMYearMonth)
            Case "票券交易明細表_交易日"
                Call Import_票券交易明細表_交易日( _
                    basePath, oldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, _
                    lastMap, _
                    ROCYearMonth, _
                    NUMYearMonth)
            Case "會計資料庫及票券庫存明細表日結"
                Call Import_會計資料庫及票券庫存明細表日結( _
                    basePath, oldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, _
                    lastMap, _
                    ROCYearMonth, _
                    NUMYearMonth)
            Case "會計資料庫及債券風險部位餘額表和票券風險部位餘額表"
                Call Import_會計資料庫及債券風險部位餘額表和票券風險部位餘額表( _
                    basePath, oldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, _
                    lastMap, _
                    ROCYearMonth, _
                    NUMYearMonth)
            Case "會計資料庫及關帳匯率"
                Call Import_會計資料庫及關帳匯率( _
                    basePath, oldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, _
                    lastMap, _
                    ROCYearMonth, _
                    NUMYearMonth)
            Case "會計資料庫及債券交易明細表"
                Call Import_會計資料庫及債券交易明細表( _
                    basePath, oldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, _
                    lastMap, _
                    ROCYearMonth, _
                    NUMYearMonth)
            Case "票券交易明細表_交割日"
                Call Import_票券交易明細表_交割日( _
                    basePath, oldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, _
                    lastMap, _
                    ROCYearMonth, _
                    NUMYearMonth)

            Case "會計資料庫及外幣債評估表及債券評價表及票券評價表及匯率表"
                Call Import_會計資料庫及外幣債評估表及債券評價表及票券評價表及匯率表( _
                    basePath, oldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, _
                    lastMap, _
                    ROCYearMonth, _
                    NUMYearMonth)

            Case "無"
                Call Import_無資料( _
                    basePath, oldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, _
                    lastMap, _
                    ROCYearMonth, _
                    NUMYearMonth)                    
            Case Else
                MsgBox "未知 CaseType: " & caseType & "（ReportID=" & wsRpt.Cells(i, "A").Value & "）", vbExclamation
        End Select
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
                             ByVal ROCYearMonth As Long, _
                             ByVal NUMYearMonth As Long)

    Dim wbOld    As Workbook, wbImp As Workbook
    Dim wbNew    As Workbook, wbDecl As Workbook
    Dim arrOldSh() As String, arrImpF() As String, arrImpSh() As String
    Dim oldTplRel As String, newTplRel As String
    Dim tplPath   As String, declTplPath As String
    Dim j         As Long

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

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
            wbOld.Sheets("AI430").Range("C2").Value = NUMYearMonth
    End Select    


    For j = LBound(arrImpF) To UBound(arrImpF)
        Set wbImp = Workbooks.Open(basePath & "\" & Trim(arrImpF(j)), ReadOnly:=True)
        With wbOld.Sheets(Trim(tplSheet))
            ' .Cells.Clear
            .Cells.ClearContents

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
                           ByVal ROCYearMonth As Long, _
                           ByVal NUMYearMonth As Long)

    Dim wbOld    As Workbook, wbImp As Workbook
    Dim wbNew    As Workbook, wbDecl As Workbook
    Dim arrOldSh() As String, arrImpF() As String, arrImpSh() As String
    Dim oldTplRel As String, newTplRel As String
    Dim tplPath   As String, declTplPath As String
    Dim j         As Long

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

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

        Call PNCDCAL_FormatToCSV(basePath & "\" & Trim(arrImpF(j)))
        Set wbImp = Workbooks.Open(Replace(basePath & "\" & Trim(arrImpF(j)), "txt", "csv"), ReadOnly:=True)
        
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
                                       ByVal ROCYearMonth As Long, _
                                       ByVal NUMYearMonth As Long)
    Dim wbOld    As Workbook, wbImp As Workbook
    Dim wbNew    As Workbook, wbDecl As Workbook
    Dim arrOldSh() As String, arrImpF() As String, arrImpSh() As String
    Dim oldTplRel As String, newTplRel As String
    Dim tplPath   As String, declTplPath As String
    Dim j         As Long

    Dim filterFields As Variant
    Dim filterValues As Variant

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    '— 路徑置換 ——
    oldTplRel = Replace(tplPattern, "YYYYMM", oldMon)
    newTplRel = Replace(tplPattern, "YYYYMM", newMon)
    arrImpF   = Split(Replace(impPattern, "YYYYMM", newMon), ",")
    arrImpSh  = Split(impSheets, ",")

    tplPath     = basePath & "\" & oldTplRel
    declTplPath = basePath & "\" & declTplRel

    Select Case rptId
        Case "表22"
            filterFields = Array("交易別", "票類")
            filterValues = Array(Array("首購", "續發", "買斷", "附買回", "附賣回", "承銷發行", "代銷買入", "賣斷", "代銷發行", "代銷賣出", "附買回續做", "附賣回續作"), Array("CP1", "CP2", "TA", "同業NCD", "ABCP", "BA", "TB1", "TB2", "MN"))
        Case "表23"
            filterFields = Array("交易別", "票類")
            filterValues = Array(Array("首購", "續發", "買斷", "附買回", "附賣回", "承銷發行", "代銷買入", "賣斷", "代銷發行", "代銷賣出", "攤提", "附買回續做", "附賣回續作"), Array("CP1", "CP2", "TA", "同業NCD", "ABCP", "BA", "TB1", "TB2", "MN"))                  
    End Select

    ' 表23底稿右邊有公式，不要拿掉

    '— 開舊底稿 + 匯入資料 ——
    Set wbOld = Workbooks.Open(tplPath, ReadOnly:=True)

    Select Case rptId
        Case "表22"
            wbOld.Sheets("表22").Range("A2").Value = ROCYearMonth
        Case "表23"
            wbOld.Sheets("票券利率統計表").Range("E2").Value = ROCYearMonth
    End Select

    wbOld.Sheets(Trim(tplSheet)).Range("A:AD").ClearContents
    For j = LBound(arrImpF) To UBound(arrImpF)
        Call ImportCsvWithFilter(basePath & "\" & Trim(arrImpF(j)), wbOld.Sheets(Trim(tplSheet)), wbOld.Sheets(Trim(tplSheet)).Range("A1"), filterFields, filterValues)
    Next j

    '— 另存新底稿 ——
    wbOld.SaveCopyAs basePath & "\" & newTplRel
    wbOld.Close False

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
                                       ByVal ROCYearMonth As Long, _
                                       ByVal NUMYearMonth As Long)
    Dim wbOld    As Workbook, wbImp As Workbook
    Dim wbNew    As Workbook, wbDecl As Workbook
    Dim arrOldSh() As String, arrImpF() As String, arrImpSh() As String
    Dim oldTplRel As String, newTplRel As String
    Dim tplPath   As String, declTplPath As String
    Dim j         As Long

    Dim filterFields As Variant
    Dim filterValues As Variant

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    '— 路徑置換 ——
    oldTplRel = Replace(tplPattern, "YYYYMM", oldMon)
    newTplRel = Replace(tplPattern, "YYYYMM", newMon)
    arrImpF   = Split(Replace(impPattern, "YYYYMM", newMon), ",")
    arrImpSh  = Split(impSheets, ",")

    tplPath     = basePath & "\" & oldTplRel
    declTplPath = basePath & "\" & declTplRel

    Select Case rptId
        Case "AI410"
            filterFields = Array("交易別")
            filterValues = Array(Array("首購", "承銷發行"))
        Case "AI415"
            filterFields = Array("交易別")
            filterValues = Array(Array("首購", "買斷", "附買回", "附賣回", "承銷發行", "附買回續做", "附賣回續作"))
    End Select

    ' 表23底稿右邊有公式，不要拿掉

    '— 開舊底稿 + 匯入資料 ——
    Set wbOld = Workbooks.Open(tplPath, ReadOnly:=True)

    Select Case rptId
        Case "AI410"
            wbOld.Sheets("AI410").Range("C2").Value = NUMYearMonth
        Case "AI415"
            wbOld.Sheets("AI415").Range("C2").Value = NUMYearMonth
    End Select


    wbOld.Sheets(Trim(tplSheet)).Range("A:AD").ClearContents
    For j = LBound(arrImpF) To UBound(arrImpF)
        Call ImportCsvWithFilter(basePath & "\" & Trim(arrImpF(j)), wbOld.Sheets(Trim(tplSheet)), wbOld.Sheets(Trim(tplSheet)).Range("A1"), filterFields, filterValues)
    Next j

    '— 另存新底稿 ——
    wbOld.SaveCopyAs basePath & "\" & newTplRel
    wbOld.Close False

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
                                               ByVal ROCYearMonth As Long, _
                                               ByVal NUMYearMonth As Long)

    Dim wbOld    As Workbook, wbImp As Workbook
    Dim wbNew    As Workbook, wbDecl As Workbook
    Dim arrOldSh() As String, arrTplSh() As String, arrImpF() As String, arrImpSh() As String
    Dim oldTplRel As String, newTplRel As String
    Dim tplPath   As String, declTplPath As String
    Dim j         As Long

    Dim filterFields As Variant
    Dim filterValues As Variant

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    '— 路徑置換 ——
    oldTplRel = Replace(tplPattern, "YYYYMM", oldMon)
    newTplRel = Replace(tplPattern, "YYYYMM", newMon)
    arrTplSh = Split(tplSheet, ",")
    arrImpF   = Split(Replace(impPattern, "YYYYMM", newMon), ",")
    arrImpSh  = Split(impSheets, ",")

    tplPath     = basePath & "\" & oldTplRel
    declTplPath = basePath & "\" & declTplRel

    Select Case rptId
        Case "表36", "表6"
            filterFields = Array("票類")
            filterValues = Array(Array("CP1", "CP2", "TA", "ABCP", "BA", "TB1", "TB2", "MN"))
        Case "表7"
            filterFields = Array("發票人", "票類")
            filterValues = Array(Array("<>中租迪和", "<>中華航空"), Array("CP1", "CP2", "TA", "同業NCD", "ABCP", "BA", "TB1", "TB2", "MN"))
    End Select

    ' 表23底稿右邊有公式，不要拿掉

    '— 開舊底稿 + 匯入資料 ——
    Set wbOld = Workbooks.Open(tplPath, ReadOnly:=True)

    Select Case rptId
        Case "表36"
            wbOld.Sheets("表36").Range("E3").Value = ROCYearMonth
    End Select
    
    For j = LBound(arrImpF) To UBound(arrImpF)

        ' If檔名startwith 會計資料庫 then go through old path
        ' If檔名Startwith 票券庫存明細表日結 then go through new path

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
    Dim i As Long, colIndex As Long

    ' 1. 開啟 CSV
    Set wbCsv = Workbooks.Open(Filename:=csvPath)
    Set shCsv = wbCsv.Sheets(1)

    ' 2. 標題列範圍
    Set fRange = shCsv.Range( _
        shCsv.Cells(1, 1), _
        shCsv.Cells(1, shCsv.UsedRange.Columns.Count) _
    )
    fRange.AutoFilter

    ' 🔧 修正欄位資料：移除隱藏空白
    Set dataRange = shCsv.UsedRange
    For Each cell In dataRange
        If VarType(cell.Value) = vbString Then
            cell.Value = Trim(Replace(cell.Value, Chr(160), ""))
        End If
    Next cell

    ' 3. 套用篩選條件
    For i = LBound(filterFields) To UBound(filterFields)
        ' 找出欄位在標題列中的索引
        colIndex = Application.Match(filterFields(i), fRange, 0)
        If Not IsError(colIndex) Then

            ' 🔧【新增】若條件為陣列，且每個字串都以 "<>" 開頭 → 使用 And 排除兩個值
            If IsArray(filterValues(i)) _
               And UBound(filterValues(i)) = 1 _
               And Left(filterValues(i)(0), 2) = "<>" _
               And Left(filterValues(i)(1), 2) = "<>" Then

                shCsv.UsedRange.AutoFilter _
                    Field:=colIndex, _
                    Criteria1:=filterValues(i)(0), _
                    Operator:=xlAnd, _
                    Criteria2:=filterValues(i)(1)

            ' 🔧【保留】原有：若是多重包含式篩選（多選一）
            ElseIf IsArray(filterValues(i)) Then

                shCsv.UsedRange.AutoFilter _
                    Field:=colIndex, _
                    Criteria1:=filterValues(i), _
                    Operator:=xlFilterValues

            ' 單一值篩選
            Else

                shCsv.UsedRange.AutoFilter _
                    Field:=colIndex, _
                    Criteria1:=filterValues(i)

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




Private Sub Import_會計資料庫及債券交易明細表(ByVal basePath   As String, _
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
                                            ByVal ROCYearMonth As Long, _
                                            ByVal NUMYearMonth As Long)

    Dim wbOld    As Workbook, wbImp As Workbook
    Dim wbNew    As Workbook, wbDecl As Workbook
    Dim arrOldSh() As String, arrTplSh() As String, arrImpF() As String, arrImpSh() As String
    Dim oldTplRel As String, newTplRel As String
    Dim tplPath   As String, declTplPath As String
    Dim j         As Long

    Dim filterFields As Variant
    Dim filterValues As Variant

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    '— 路徑置換 ——
    oldTplRel = Replace(tplPattern, "YYYYMM", oldMon)
    newTplRel = Replace(tplPattern, "YYYYMM", newMon)
    arrTplSh = Split(tplSheet, ",")
    arrImpF   = Split(Replace(impPattern, "YYYYMM", newMon), ",")
    arrImpSh  = Split(impSheets, ",")

    tplPath     = basePath & "\" & oldTplRel
    declTplPath = basePath & "\" & declTplRel

    Select Case rptId
        Case "AI405"
            filterFields = Array("交易類別")
            filterValues = Array(Array("發行前買斷", "買斷", "附買回", "附賣回", "發行前賣斷", "賣斷", "首購", "附買回續作", "附賣回續作"))
    End Select    

    ' 表23底稿右邊有公式，不要拿掉

    '— 開舊底稿 + 匯入資料 ——
    Set wbOld = Workbooks.Open(tplPath, ReadOnly:=True)

    Select Case rptId
        Case "AI405"
            wbOld.Sheets("AI405").Range("C2").Value = NUMYearMonth
    End Select

    
    For j = LBound(arrImpF) To UBound(arrImpF)

        ' If檔名startwith 會計資料庫 then go through old path
        ' If檔名Startwith 票券庫存明細表日結 then go through new path

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
        End Select        
    Next j

    '— 另存新底稿 ——
    wbOld.SaveCopyAs basePath & "\" & newTplRel
    wbOld.Close False

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
                                                                 ByVal ROCYearMonth As Long, _
                                                                 ByVal NUMYearMonth As Long)

    Dim wbOld    As Workbook, wbImp As Workbook
    Dim wbNew    As Workbook, wbDecl As Workbook
    Dim arrOldSh() As String, arrTplSh() As String, arrImpF() As String, arrImpSh() As String
    Dim oldTplRel As String, newTplRel As String
    Dim tplPath   As String, declTplPath As String
    Dim j         As Long

    Dim filterFields As Variant
    Dim filterValues As Variant

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    '— 路徑置換 ——
    oldTplRel = Replace(tplPattern, "YYYYMM", oldMon)
    newTplRel = Replace(tplPattern, "YYYYMM", newMon)
    arrTplSh = Split(tplSheet, ",")
    arrImpF   = Split(Replace(impPattern, "YYYYMM", newMon), ",")
    arrImpSh  = Split(impSheets, ",")

    tplPath     = basePath & "\" & oldTplRel
    declTplPath = basePath & "\" & declTplRel 

    ' 表23底稿右邊有公式，不要拿掉

    '— 開舊底稿 + 匯入資料 ——
    Set wbOld = Workbooks.Open(tplPath, ReadOnly:=True)

    Select Case rptId
        Case "AI233"
            wbOld.Sheets("AI233").Range("B4").Value = NUMYearMonth
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
                                      ByVal ROCYearMonth As Long, _
                                      ByVal NUMYearMonth As Long)

    Dim wbOld    As Workbook, wbImp As Workbook
    Dim wbNew    As Workbook, wbDecl As Workbook
    Dim arrOldSh() As String, arrTplSh() As String, arrImpF() As String, arrImpSh() As String
    Dim oldTplRel As String, newTplRel As String
    Dim tplPath   As String, declTplPath As String
    Dim j         As Long

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    '— 路徑置換 ——
    oldTplRel = Replace(tplPattern, "YYYYMM", oldMon)
    newTplRel = Replace(tplPattern, "YYYYMM", newMon)
    arrTplSh = Split(tplSheet, ",")
    arrImpF   = Split(Replace(impPattern, "YYYYMM", newMon), ",")
    arrImpSh  = Split(impSheets, ",")

    tplPath     = basePath & "\" & oldTplRel
    declTplPath = basePath & "\" & declTplRel

    ' 表23底稿右邊有公式，不要拿掉

    '— 開舊底稿 + 匯入資料 ——
    Set wbOld = Workbooks.Open(tplPath, ReadOnly:=True)

    Select Case rptId
        Case "AI345"
            wbOld.Sheets("AI345_NEW").Range("A2").Value = NUMYearMonth
    End Select
    
    For j = LBound(arrImpF) To UBound(arrImpF)

        ' If檔名startwith 會計資料庫 then go through old path
        ' If檔名Startwith 票券庫存明細表日結 then go through new path

        Select Case True
            Case Mid(arrImpF(j), InStrRev(arrImpF(j), "\") + 1) Like "CloseRate*"
                wbOld.Sheets(arrTplSh(j)).Range("A:Z").ClearContents
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
                                                                       ByVal ROCYearMonth As Long, _
                                                                       ByVal NUMYearMonth As Long)

    Dim wbOld    As Workbook, wbImp As Workbook
    Dim wbNew    As Workbook, wbDecl As Workbook
    Dim arrOldSh() As String, arrTplSh() As String, arrImpF() As String, arrImpSh() As String
    Dim oldTplRel As String, newTplRel As String
    Dim tplPath   As String, declTplPath As String
    Dim j         As Long

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    '— 路徑置換 ——
    oldTplRel = Replace(tplPattern, "YYYYMM", oldMon)
    newTplRel = Replace(tplPattern, "YYYYMM", newMon)
    arrTplSh = Split(tplSheet, ",")
    arrImpF   = Split(Replace(impPattern, "YYYYMM", newMon), ",")
    arrImpSh  = Split(impSheets, ",")

    tplPath     = basePath & "\" & oldTplRel
    declTplPath = basePath & "\" & declTplRel

    ' 表23底稿右邊有公式，不要拿掉

    '— 開舊底稿 + 匯入資料 ——
    Set wbOld = Workbooks.Open(tplPath, ReadOnly:=True)

    Select Case rptId
        Case "AI601"
            wbOld.Sheets("AI601").Range("H2").Value = NUMYearMonth
    End Select

    
    For j = LBound(arrImpF) To UBound(arrImpF)

        ' If檔名startwith 會計資料庫 then go through old path
        ' If檔名Startwith 票券庫存明細表日結 then go through new path

        Select Case True
            Case Mid(arrImpF(j), InStrRev(arrImpF(j), "\") + 1) Like "CloseRate*"
                wbOld.Sheets(arrTplSh(j)).Range("A:Z").ClearContents
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
                Set wbImp = Workbooks.Open(basePath & "\" & Trim(arrImpF(j)), ReadOnly:=True)

                With wbOld.Sheets(arrTplSh(j))
                    wbImp.Sheets(Trim(arrImpSh(j))).UsedRange.Copy
                    .Range("A1").PasteSpecial xlPasteValues
                End With
                wbImp.Close False

            Case Mid(arrImpF(j), InStrRev(arrImpF(j), "\") + 1) Like "票券評價表*"
                wbOld.Sheets(arrTplSh(j)).Range("A1:V50").ClearContents
                Set wbImp = Workbooks.Open(basePath & "\" & Trim(arrImpF(j)), ReadOnly:=True)

                With wbOld.Sheets(arrTplSh(j))
                    wbImp.Sheets(Trim(arrImpSh(j))).UsedRange.Copy
                    .Range("A1").PasteSpecial xlPasteValues
                End With
                wbImp.Close False
                
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

    '— 貼入申報模板 ——
    Set wbNew  = Workbooks.Open(basePath & "\" & newTplRel)
    Set wbDecl = Workbooks.Open(declTplPath)

    Dim srcSh As String, rngStrings() As String
    Dim rngSrc As Range, rngDst As Range
    Dim srcAddr As Variant

    Dim errNum As Long, errDesc As String
    Dim tmpSh As Worksheet  
    
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
                         ByVal ROCYearMonth As Long, _
                         ByVal NUMYearMonth As Long)

    Dim wbOld    As Workbook, wbImp As Workbook
    Dim wbNew    As Workbook, wbDecl As Workbook
    Dim arrOldSh() As String, arrTplSh() As String, arrImpF() As String, arrImpSh() As String
    Dim oldTplRel As String, newTplRel As String
    Dim tplPath   As String, declTplPath As String
    Dim j         As Long


    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    '— 路徑置換 ——
    oldTplRel = Replace(tplPattern, "YYYYMM", oldMon)
    newTplRel = Replace(tplPattern, "YYYYMM", newMon)
    arrTplSh = Split(tplSheet, ",")
    arrImpF   = Split(Replace(impPattern, "YYYYMM", newMon), ",")
    arrImpSh  = Split(impSheets, ",")

    tplPath     = basePath & "\" & oldTplRel
    declTplPath = basePath & "\" & declTplRel


    ' 表23底稿右邊有公式，不要拿掉

    '— 開舊底稿 + 匯入資料 ——
    Set wbOld = Workbooks.Open(tplPath, ReadOnly:=True)

    Select Case rptId
        Case "表15B"
            wbOld.Sheets("外幣可轉讓定期存單發行、償還及餘額統計表").Range("A2").Value = ROCYearMonth
        Case "表16"
            wbOld.Sheets("可轉讓定期存單每日持有額").Range("B2").Value = ROCYearMonth
        Case "表24"
            wbOld.Sheets("表24").Range("E2").Value = ROCYearMonth
        Case "表27"
            wbOld.Sheets("表27").Range("D3").Value = ROCYearMonth

        Case "AI816"
            wbOld.Sheets("AI816").Range("A2").Value = ROCYearMonth
        Case "AI271"
            wbOld.Sheets("AI271").Range("A3").Value = ROCYearMonth
        Case "AI272"
            wbOld.Sheets("AI272").Range("A3").Value = ROCYearMonth
        Case "AI273"
            wbOld.Sheets("AI273").Range("A3").Value = ROCYearMonth
        Case "AI281"
            wbOld.Sheets("AI281").Range("A3").Value = ROCYearMonth
        Case "AI282"
            wbOld.Sheets("AI282").Range("A3").Value = ROCYearMonth
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
End Sub



' =========================


Private Sub GetMonths(ByRef oldMon As String, ByRef newMon As String)
    Dim ym As String, ry As Integer, m As Integer
    ym = ThisWorkbook.Names("YearMonth").RefersToRange.Value
    newMon = ym
    ry = CInt(Left(ym, Len(ym) - 2))
    m  = CInt(Right(ym, 2)) - 1
    If m = 0 Then ry = ry - 1: m = 12
    oldMon = CStr(ry) & Format(m, "00")
End Sub

Sub ProcessAllReports()
    Dim wbCtl    As Workbook
    Dim wsRpt    As Worksheet, wsMap As Worksheet
    Dim basePath As String
    Dim lastRpt  As Long, lastMap As Long
    Dim oldMon   As String, newMon As String
    Dim i        As Long, caseType As String

    Call GetMonths(oldMon, newMon)

    Set wbCtl    = ThisWorkbook
    Set wsRpt    = wbCtl.Sheets("ReportsConfig")
    Set wsMap    = wbCtl.Sheets("Mappings")
    basePath     = wbCtl.Path
    lastRpt      = wsRpt.Cells(wsRpt.Rows.Count, "A").End(xlUp).Row
    lastMap      = wsMap.Cells(wsMap.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRpt
        caseType = wsRpt.Cells(i, "H").Value
        Select Case LCase(caseType)
            Case "會計資料庫"
                Call Import_會計資料庫( _
                    basePath, oldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, _
                    lastMap)
            Case "PNCDCAL"
                Call Import_PNCDCAL(basePath, _
                                    oldMon, _
                                    newMon, _
                                    wsRpt.Cells(i, "A").Value, _
                                    wsRpt.Cells(i, "B").Value, _
                                    wsRpt.Cells(i, "C").Value, _
                                    wsRpt.Cells(i, "D").Value, _
                                    wsRpt.Cells(i, "E").Value, _
                                    wsRpt.Cells(i, "F").Value, _
                                    wsMap, _
                                    lastMap)                    
            Case "票券交易明細表_交易日"
                Call Import_票券交易明細表_交易日( _
                    basePath, oldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, lastMap)
            Case "會計資料庫及票券庫存明細表日結"
                Call Import_會計資料庫及票券庫存明細表日結( _
                    basePath, oldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, lastMap)
            Case "會計資料庫及債券風險部位餘額表和票券風險部位餘額表"
                Call Import_會計資料庫及債券風險部位餘額表和票券風險部位餘額表( _
                    basePath, oldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, lastMap)
            Case "會計資料庫及關帳匯率"
                Call Import_會計資料庫及關帳匯率( _
                    basePath, oldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, lastMap)
            Case "會計資料庫及債券交易明細表"
                Call Import_會計資料庫及債券交易明細表( _
                    basePath, oldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, lastMap)
            Case "票券交易明細表_交割日"
                Call Import_票券交易明細表_交割日( _
                    basePath, oldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, lastMap)

            Case "會計資料庫及外幣債評估表及債券評價表及票券評價表及匯率表"
                Call Import_會計資料庫及外幣債評估表及債券評價表及票券評價表及匯率表( _
                    basePath, oldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, lastMap)
            Case Else
                MsgBox "未知 CaseType: " & caseType & "（ReportID=" & wsRpt.Cells(i, "A").Value & "）", vbExclamation
        End Select
    Next i

    MsgBox "全部報表處理完成！", vbInformation
End Sub

在以上的代碼中我原本預設 oldMon 和 newMon 是輸入類似這樣的日期格式 11406 11407 11408等這樣的格式，
但我現在需要將這樣的數字轉換成 民國 114 年 06 月、民國 114 年 07 月、民國 114 年 08 月，
請問要怎麼轉換


' =========

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




' ================


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

'— 2. 主流程：依 CaseType 分派 ——
Sub ProcessAllReports()
    Dim wbCtl    As Workbook
    Dim wsRpt    As Worksheet, wsMap As Worksheet
    Dim basePath As String
    Dim lastRpt  As Long, lastMap As Long
    Dim oldMon   As String, newMon As String
    Dim ROCYearMonth As String, NUMYearMonth As String, 
    Dim i        As Long, caseType As String

    Call GetMonths(oldMon, newMon)

    ROCYearMonth = ConvertToROCFormat(newMon, "ROC")
    NUMYearMonth = ConvertToROCFormat(newMon, "NUM")

    Set wbCtl    = ThisWorkbook
    Set wsRpt    = wbCtl.Sheets("ReportsConfig")
    Set wsMap    = wbCtl.Sheets("Mappings")
    basePath     = wbCtl.Path
    lastRpt      = wsRpt.Cells(wsRpt.Rows.Count, "A").End(xlUp).Row
    lastMap      = wsMap.Cells(wsMap.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRpt
        caseType = wsRpt.Cells(i, "H").Value
        Select Case LCase(caseType)
            Case "會計資料庫"
                Call Import_會計資料庫( _
                    basePath, oldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, _
                    lastMap, _
                    ROCYearMonth, _
                    NUMYearMonth)
            Case "PNCDCAL"
                Call Import_PNCDCAL(basePath, _
                                    oldMon, _
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
                                    NUMYearMonth)
            Case "票券交易明細表_交易日"
                Call Import_票券交易明細表_交易日( _
                    basePath, oldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, _
                    lastMap, _
                    ROCYearMonth, _
                    NUMYearMonth)
            Case "會計資料庫及票券庫存明細表日結"
                Call Import_會計資料庫及票券庫存明細表日結( _
                    basePath, oldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, _
                    lastMap, _
                    ROCYearMonth, _
                    NUMYearMonth)
            Case "會計資料庫及債券風險部位餘額表和票券風險部位餘額表"
                Call Import_會計資料庫及債券風險部位餘額表和票券風險部位餘額表( _
                    basePath, oldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, _
                    lastMap, _
                    ROCYearMonth, _
                    NUMYearMonth)
            Case "會計資料庫及關帳匯率"
                Call Import_會計資料庫及關帳匯率( _
                    basePath, oldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, _
                    lastMap, _
                    ROCYearMonth, _
                    NUMYearMonth)
            Case "會計資料庫及債券交易明細表"
                Call Import_會計資料庫及債券交易明細表( _
                    basePath, oldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, _
                    lastMap, _
                    ROCYearMonth, _
                    NUMYearMonth)
            Case "票券交易明細表_交割日"
                Call Import_票券交易明細表_交割日( _
                    basePath, oldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, _
                    lastMap, _
                    ROCYearMonth, _
                    NUMYearMonth)

            Case "會計資料庫及外幣債評估表及債券評價表及票券評價表及匯率表"
                Call Import_會計資料庫及外幣債評估表及債券評價表及票券評價表及匯率表( _
                    basePath, oldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, _
                    lastMap, _
                    ROCYearMonth, _
                    NUMYearMonth)

            Case "無"
                Call Import_無資料( _
                    basePath, oldMon, newMon, _
                    wsRpt.Cells(i, "A").Value, _
                    wsRpt.Cells(i, "B").Value, _
                    wsRpt.Cells(i, "C").Value, _
                    wsRpt.Cells(i, "D").Value, _
                    wsRpt.Cells(i, "E").Value, _
                    wsRpt.Cells(i, "F").Value, _
                    wsMap, _
                    lastMap, _
                    ROCYearMonth, _
                    NUMYearMonth)                    
            Case Else
                MsgBox "未知 CaseType: " & caseType & "（ReportID=" & wsRpt.Cells(i, "A").Value & "）", vbExclamation
        End Select
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
                             ByVal ROCYearMonth As Long, _
                             ByVal NUMYearMonth As Long)

    Dim wbOld    As Workbook, wbImp As Workbook
    Dim wbNew    As Workbook, wbDecl As Workbook
    Dim arrOldSh() As String, arrImpF() As String, arrImpSh() As String
    Dim oldTplRel As String, newTplRel As String
    Dim tplPath   As String, declTplPath As String
    Dim j         As Long

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

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
            wbOld.Sheets("AI430").Range("C2").Value = NUMYearMonth
    End Select    


    For j = LBound(arrImpF) To UBound(arrImpF)
        Set wbImp = Workbooks.Open(basePath & "\" & Trim(arrImpF(j)), ReadOnly:=True)
        With wbOld.Sheets(Trim(tplSheet))
            ' .Cells.Clear
            .Cells.ClearContents

            wbImp.Sheets(Trim(arrImpSh(j))).UsedRange.Copy
            .Range("A1").PasteSpecial xlPasteValues
        End With
        wbImp.Close False
    Next j

    '— 另存新底稿 ——
    wbOld.SaveCopyAs basePath & "\" & newTplRel
    wbOld.Close False

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
End Sub

以上VBA程序中，在Import會計資料庫的case中，
我想要在wbOld所有要做的事情，並建立新的檔案後，
所有事情都做完之後，
將wbOld刪掉，只留下新的檔案，請問要怎麼修改，請給我完整內容，並在修改處清楚標示




