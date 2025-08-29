1.
以下先說明我的VBA專案如下，
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

ReportsConfig 分頁，

其專案流程為透過迴圈逐一讀取Excel的ReportsConfig分頁資料來進行。
ReportsConfig分頁內容中的資料範例如下:

"報表名稱
rptID"	"原始(前月)底稿路徑(YYYYMM代表報表年月份)
TplPath"	"原始底稿需更新分頁名稱
(若多個分頁係以逗號"",""分隔)
TplSheet"	"需匯入報表之財務/會計數據路徑(YYYYMM代表報表年月份)
(若多個檔案係以逗號"",""分隔)
ImpPath"	"匯入報表檔案之分頁名稱
(若多個分頁係以逗號"",""分隔)
ImpSheets"	"申報檔案路徑 (檔名皆不含年月份)
DeclTpl"	申報期限類別	"
需處理資料分類
(不同Case係以所需匯入報表類型區分)
CaseType"
表6	央行\YYYYMM 表6 主要業務概況月報表：Ａ+放款對 (3)2.xlsx	票券庫存(日結),E	批次報表\票券庫存明細表(日結).csv,批次報表\會計資料庫.xlsx	票券庫存明細表(日結),餘額E (2)	無	月報	會計資料庫及票券庫存明細表日結
表7	央行\YYYYMM 表7 主要業務概況月報表：B++放款2.xlsx	票券庫存(日結),E	批次報表\票券庫存明細表(日結).csv,批次報表\會計資料庫.xlsx	票券庫存明細表(日結),餘額E (2)	無	月報	會計資料庫及票券庫存明細表日結
表10	央行\YYYYMM 表10.表40.買入票券及投資簡明表-新版.xls	E	批次報表\會計資料庫.xlsx	餘額E (2)	央行\(010)買入證券及投資統計表.xlsx	月報	會計資料庫
表15A	央行\YYYYMM 表15A.可轉讓定期存單發行償還及餘額統計表.xls	PNCDCAL	批次報表\PNCDCAL.txt	無	央行\(015A)新台幣可轉讓定期存單發行、償還及餘額統計表.xlsx	月報	PNCDCAL

2.現在我需要將專案進行修改，要將原始設定資料儲存在csv檔案或是json檔案，Excel在執行時會透過 powershell管理，
我現在Excel檔案中相關年月日資料都是填寫 YYYYMM，我希望在Excel中看到的是已經將實際填寫的日期轉換過的文字

以下a-c是我隨筆把想到的寫下來，請幫我把所有想法彙整起來，先幫我擬一個比較完整你會比較好回答的問題

a.
Excel click => 透過powershell抓取csv資料，刷新Excel，接著執行後續Excel VBA功能

powershell要定義好每個欄位的可選項目
欄位有以下

報表名稱
原始底稿路徑
原始底稿需更新分頁
匯入報表名稱
匯入報表分頁
申報檔案路徑
申報期限類別
處理類型

以物件的方式去儲存

例如:

{
    報表名稱:value,
    原始底稿路徑:value,
    原始底稿須更新分頁:[],
    匯入報表名稱:[],
    匯入報表分頁:[],
    申報檔案路徑,
    申報期限類別,
    處理類型,
}

需要另外設定不同欄位，母體資料的設定(一個中介的窗口)

票券交易明細表 可篩選條件(逗號分隔)(只有特定匯入報表可選，其他選擇之後一律跳為無)(所以要多一個欄位放 可篩選條件)

powershell 類似用這種方式去儲存，
然後儲存成csv檔案或json檔案(這部分要再思考)，
Json後續其他同仁要理解可能很困難，

Array在csv儲存方式要在問gpt怎樣儲存比較好，

click => call powershell更新欄位資訊(包含取得westernMonth rocMonth(這部分反映在) 等seeting資訊)

另外當更新執行頁面的報表年月份時，要同步更新ReportsConfig中的

- 新增一個分頁設定
    原始底稿路徑
        原始底稿分頁
    匯入報表名稱
        匯入報表分頁

- 這邊要多設定，需匯入報表要出現下拉式選單
    選擇原始底稿1
        原始底稿分頁可選項目為原始底稿1中的分頁
    
- 票券交易明細表篩選條件可自己選


b.
重置按鈕 Click => Call powershell 那隻將資料push到Excel中

c.
UpdateButton Click => Call powershell 更新csv檔案，

更新Excel資料後要，一併更新csv檔案(內含資料結構)




' ===================================================

票券別
CP1
CP2
TA
同業NCD
ABCP
BA
TB1
TB2
央行NCD
一年以上央行NCD
MN

交易類別
首購
續發
買斷
附買回
附賣回
承銷發行
代銷買入
賣斷
附買回履約
附賣回履約
代銷發行
兌償/到期還本
附買回解約
附賣回解約
代銷賣出
攤提
附買回續作
附賣回續作



' ===================================================


"報表名稱
rptID"	"原始(前月)底稿路徑(YYYYMM代表報表年月份)
TplPath"	"原始底稿需更新分頁名稱
(若多個分頁係以逗號"",""分隔)
TplSheet"	"需匯入報表之財務/會計數據路徑(YYYYMM代表報表年月份)
(若多個檔案係以逗號"",""分隔)
ImpPath"	"匯入報表檔案之分頁名稱
(若多個分頁係以逗號"",""分隔)
ImpSheets"	"申報檔案路徑 (檔名皆不含年月份)
DeclTpl"	申報期限類別	"
需處理資料分類
(不同Case係以所需匯入報表類型區分)
CaseType"
表6	央行\YYYYMM 表6 主要業務概況月報表：Ａ+放款對 (3)2.xlsx	票券庫存(日結),E	批次報表\票券庫存明細表(日結).csv,批次報表\會計資料庫.xlsx	票券庫存明細表(日結),餘額E (2)	無	月報	會計資料庫及票券庫存明細表日結
表7	央行\YYYYMM 表7 主要業務概況月報表：B++放款2.xlsx	票券庫存(日結),E	批次報表\票券庫存明細表(日結).csv,批次報表\會計資料庫.xlsx	票券庫存明細表(日結),餘額E (2)	無	月報	會計資料庫及票券庫存明細表日結
表10	央行\YYYYMM 表10.表40.買入票券及投資簡明表-新版.xls	E	批次報表\會計資料庫.xlsx	餘額E (2)	央行\(010)買入證券及投資統計表.xlsx	月報	會計資料庫
表15A	央行\YYYYMM 表15A.可轉讓定期存單發行償還及餘額統計表.xls	PNCDCAL	批次報表\PNCDCAL.txt	無	央行\(015A)新台幣可轉讓定期存單發行、償還及餘額統計表.xlsx	月報	PNCDCAL
表15B	央行\YYYYMM 表15B-外幣可轉讓定期存單發行、償還及餘額統計表.xls	無	無	無	央行\(015B)外幣可轉讓定期存單發行、償還及餘額統計表.xlsx	月報	無
表16	央行\YYYYMM 表16.可轉讓定期存單每日持有額.xls	無	無	無	央行\(016)可轉讓定期存單每日持有額.xlsx	月報	無
表20	央行\YYYYMM 表20.賣出附買回約定票(債)券交易餘額統計表.xls	E	批次報表\會計資料庫.xlsx	餘額E (2)	央行\(020)附買回票券及債券負債餘額統計表.xlsx	月報	會計資料庫
表22	央行\YYYYMM 表22.票券交易統計表.xls	票券交易明細表	批次報表\票券交易明細表_交割日.csv	票券交易明細表_交割日	央行\(022)票券交易統計表.xlsx	月報	票券交易明細表_交割日
表23	央行\YYYYMM 表23.票券利率統計表.xls	承銷交易	批次報表\票券交易明細表_交割日.csv	票券交易明細表_交割日	央行\(023)票券利率統計表.xlsx	月報	票券交易明細表_交割日
表24	央行\YYYYMM 表24.承銷商業本票統計表.xls				央行\(024)承銷商業本票統計表.xlsx	月報	無
表27	央行\YYYYMM 表27.金融債券發行償還及餘額統計表.xls		手動更新		央行\(027)金融債券統計表.xlsx	月報	無
表36	央行\YYYYMM 表36 附賣回債(票)券投資餘額統計表.xls	票券庫存(日結),E	批次報表\票券庫存明細表(日結).csv,批次報表\會計資料庫.xlsx	票券庫存明細表(日結),餘額E (2)	央行\(036)RS餘額統計表.xlsx	月報	會計資料庫及票券庫存明細表日結
AI233	AI233\YYYYMM AI233 新臺幣投資明細之利率變動情境分析表AI233.xlsx	餘額C,債券風險部位餘額表,票券風險部位餘額表	批次報表\會計資料庫.xlsx,批次報表\債券風險部位餘額.csv,批次報表\票券風險部位餘額.csv	餘額C,債券風險部位餘額,票券風險部位餘額	AI233\AI233-新臺幣投資明細之利率變動情境分析表-EXCEL上傳檔案.xls	月報	會計資料庫及債券風險部位餘額表和票券風險部位餘額表
AI345	AI345\YYYYMM AI345 本國銀行不良資產彙總暨評估明細表.xls	餘額C,匯率	批次報表\會計資料庫.xlsx,批次報表\CloseRate.xls	餘額C,Sheet1	AI345\11406 AI345 本國銀行不良資產彙總暨評估明細表.xls	月報	會計資料庫及關帳匯率
AI405	AI405\YYYYMM AI405 首次買入金額交易量及持有餘額申報表.xls	C,債券交易明細,匯率	批次報表\會計資料庫.xlsx,批次報表\債券交易明細表.csv,批次報表\CloseRate.xls	餘額C,債券交易明細表,Sheet1	AI405\AI405債券發行首次買入金額交易量及持有餘額資料表 - 複製.xls	月報	會計資料庫及債券交易明細表及匯率表
AI410	AI410\YYYYMM AI410 票券發行及首次買入金額資料申報表.xls	票券交易明細表選首購、承銷發行)	批次報表\票券交易明細表_交割日.csv	票券交易明細表_交割日	AI410\AI410票券發行及首次買入金額資料申報表.xls	月報	票券交易明細表_交割日
AI415	AI415\YYYYMM AI415 票券交易量資料.xls	Sheet1(不含履約.兌償到期)	批次報表\票券交易明細表_交割日.csv	票券交易明細表_交割日	AI415\AI415票券交易量申報表.xls	月報	票券交易明細表_交割日
AI430	AI430\YYYYMM AI430 票券持有餘額資料.xls	C	批次報表\會計資料庫.xlsx	餘額C	AI430\AI430票券持有餘額資料申報表.xls	月報	會計資料庫
AI601	AI601\YYYYMM AI601 本國銀行投資明細表.xls	餘額C,餘額D,外幣債評估表,債券評價表-AC-AI601,票券評價表-AC-AI601,匯率	批次報表\會計資料庫.xlsx,批次報表\會計資料庫.xlsx,批次報表\YYYYMM外幣債損益評估表(月底)對AC5100B.xlsx,批次報表\債券評價表.csv,批次報表\票券評價表.csv,批次報表\CloseRate.xls	餘額C,餘額D,評估表,債券評價表,票券評價表,Sheet1	AI601\AI601-投資明細表.xls	月報	會計資料庫及外幣債評估表及債券評價表及票券評價表及匯率表
AI816	AI816\YYYYMM AI816 銀行辦理大陸地區法人票券金融業務統計表.xls	無	無	無	AI816\AI816銀行辦理大陸地區法人票券金融業務統計表.xls	月報	無
AI271	季報\YYYYMM-AI271-證券化與結構型商品統計表（一）發行證券化商品.xls	無	無	無	季報\AI271-證券化與結構型商品統計表（一）發行證券化商品.xls	季報	無
AI272	季報\YYYYMM-AI272-證券化與結構型商品統計表（二）辦理結構型商品_附表1.xls	無	無	無	季報\AI272-證券化與結構型商品統計表（二）辦理結構型商品_附表1.xls	季報	無
AI273	季報\YYYYMM-AI273-證券化與結構型商品統計表（二）辦理結構型商品_附表2.xls	無	無	無	季報\AI273-證券化與結構型商品統計表（二）辦理結構型商品_附表2.xls	季報	無
AI281	季報\YYYYMM-AI281-證券化與結構型商品統計表（三）投資證券化商品.xls	無	無	無	季報\AI281-證券化與結構型商品統計表（三）投資證券化商品.xls	季報	無
AI282	季報\YYYYMM-AI282-證券化與結構型商品統計表（四）投資結構型商品.xls	無	無	無	季報\AI282-證券化與結構型商品統計表（四）投資結構型商品.xls	季報	無
AI605	AI605_半年報\YYYYMM AI605按攤銷後成本法衡量之債務工具投資、對其他個體權益及嵌入式衍生工具統計表.xlsx	餘額A,餘額C,匯率,持有到期債券風險餘額表,外幣債公允價值評估	批次報表\會計資料庫.xlsx,批次報表\會計資料庫.xlsx,批次報表\CloseRate.xls,批次報表\債券評價表.csv,批次報表\YYYYMM外幣債損益評估表(月底)對AC5100B.xlsx	餘額A,餘額C,Sheet1,債券評價表,OutputData	AI605_半年報\AI605-按攤銷後成本衡量之債務工具投資、對其他個體權益及嵌入式衍生工具統計表-new.xls	半年報	會計資料庫及匯率及債券評價及外幣債評估表