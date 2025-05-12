' 初始化階段

Option Explicit

'=== Global Config Settings ===
Public gDataMonthString As String         ' 使用者輸入的資料月份
Public gDataMonthStringROC As String      ' 資料月份ROC Format
Public gDataMonthStringROC_NUM As String  ' 資料月份ROC_NUM Format
Public gDataMonthStringROC_F1F2 As String ' 資料月份ROC_F1F2 Format
Public gDBPath As String                  ' 資料庫路徑
Public gReportFolder As String            ' 原始申報報表 Excel 檔所在資料夾
Public gOutputFolder As String            ' 更新後另存新檔的資料夾
Public gReportNames As Variant            ' 報表名稱陣列
Public gReports As Collection             ' Declare Collections that Save all instances of clsReport

'=== 主流程入口 ===
Public Sub Main()
    Dim isInputValid As Boolean
    isInputValid = False
    Do
        gDataMonthString = InputBox("請輸入資料月份 (格式: yyyy/mm):", "輸入資料月份")
        If IsValidDataMonth(gDataMonthString) Then
            isInputValid = True
        ElseIf Trim(gDataMonthString) = "" Then
            MsgBox "請輸入報表資料所屬的年度/月份 (例如: 2024/01)", vbExclamation, "輸入錯誤"
            WriteLog "請輸入報表資料所屬的年度/月份 (例如: 2024/01)"
        Else
            MsgBox "格式錯誤，請輸入正確格式 (yyyy/mm)", vbExclamation, "格式錯誤"
            WriteLog "格式錯誤，請輸入正確格式 (yyyy/mm)"
        End If
    Loop Until isInputValid
    
    '轉換gDataMonthString為ROC Format
    gDataMonthStringROC = ConvertToROCFormat(gDataMonthString, "ROC")
    gDataMonthStringROC_NUM = ConvertToROCFormat(gDataMonthString, "NUM")
    gDataMonthStringROC_F1F2 = ConvertToROCFormat(gDataMonthString, "F1F2")
    ' 設定其他 config 參數（請根據實際環境調整）
    gDBPath = ThisWorkbook.Path & "\" & ThisWorkbook.Sheets("ControlPanel").Range("DBsPathFileName").Value
    ' gDBPath = "\\10.10.122.40\後台作業\99_個人資料夾\8.修豪\" & ThisWorkbook.Sheets("ControlPanel").Range("DBsPathFileName").Value
    ' 空白報表路徑
    gReportFolder = ThisWorkbook.Path & "\" & ThisWorkbook.Sheets("ControlPanel").Range("EmptyReportPath").Value
    ' 產生之申報報表路徑
    gOutputFolder = ThisWorkbook.Path & "\" & ThisWorkbook.Sheets("ControlPanel").Range("OutputReportPath").Value
    ' 製作報表List
    'gReportNames 少FB1 FM5
    gReportNames = Array("CNY1", "FB1", "FB2", "FB3", "FB3A", "FM5", "FM11", "FM13", "AI821", "Table2", "FB5", "FB5A", "FM2", "FM10", "F1_F2", "Table41", "AI602", "AI240")
    
    ' Process A: 初始化所有報表，將初始資料寫入 Access DB with Null Data
    Call InitializeReports
    MsgBox "完成 Process A"
    WriteLog "完成 Process A"
    ' Process B: 製表及更新Access DB Data
    Call Process_CNY1
    Call Process_FB1
    Call Process_FB2
    Call Process_FB3
    Call Process_FB3A
    Call Process_FM5
    Call Process_FM11
    Call Process_FM13
    Call Process_AI821
    Call Process_Table2
    Call Process_FB5
    Call Process_FB5A
    Call Process_FM2
    Call Process_FM10
    Call Process_F1_F2
    Call Process_Table41
    Call Process_AI602
    Call Process_AI240
    MsgBox "完成 Process B"
    WriteLog "完成 Process B"
    ' Process C: 開啟原始Excel報表(EmptyReportPath)，填入Excel報表數據，
    ' 另存新檔(OutputReportPath)
    Call UpdateExcelReports
    MsgBox "完成 Process C (全部處理程序完成)"
    WriteLog "完成 Process C (全部處理程序完成)"
End Sub

'=== A. 初始化所有報表並將初始資料寫入 Access ===
Public Sub InitializeReports()
    Dim rpt As clsReport
    Dim rptName As Variant, key As Variant
    Set gReports = New Collection
    For Each rptName In gReportNames
        Set rpt = New clsReport
        rpt.Init rptName, gDataMonthStringROC, gDataMonthStringROC_NUM, gDataMonthStringROC_F1F2
        gReports.Add rpt, rptName
        ' 將各工作表內每個欄位初始設定寫入 Access DB
        Dim wsPositions As Object
        Dim combinedPositions As Object
        ' 合併所有工作表，Key 格式 "wsName|fieldName"
        Set combinedPositions = rpt.GetAllFieldPositions 
        For Each key In combinedPositions.Keys
            InsertIntoTable gDBPath, "MonthlyDeclarationReport", gDataMonthString, rptName, key, "", combinedPositions(key)
        Next key
    Next rptName
    MsgBox "完成'報表初始欄位資訊儲存'及'初始資料庫資料建立'"
    WriteLog "完成'報表初始欄位資訊儲存'及'初始資料庫資料建立'"
End Sub

'=== B 各報表獨立處理邏輯 ===

Public Sub Process_CNY1()
    '=== Equal Setting ===
    'Fetch Query Access DB table
    Dim dataArr As Variant

    'Declare worksheet and handle data
    Dim xlsht As Worksheet

    Dim i As Integer, j As Integer
    Dim lastRow As Integer

    Dim reportTitle As String
    Dim queryTable As String

    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("CNY1")
    
    reportTitle = "CNY1"
    queryTable = "CNY1_DBU_AC5601"

    ' dataArr = GetAccessDataAsArray(gDBPath, queryTable, gDataMonthString)
    dataArr = GetAccessDataAsArray(gDBPath, queryTable, gDataMonthString)

    ' If UBound(dataArr) < 2 Then
    '     MsgBox "CNY1 查詢資料不完整！", vbExclamation
    ' End If
    
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    xlsht.Range("A:E").ClearContents
    xlsht.Range("T2:T100").ClearContents

    '=== Paste Queyr Table into Excel ===
    If Err.Number <> 0 Or LBound(dataArr) > UBound(dataArr) Then
        MsgBox reportTitle & ": " & queryTable & " 資料表無資料"
        WriteLog reportTitle & ": " & queryTable & " 資料表無資料"
    Else
        For j = 0 To UBound(dataArr, 2)
            For i = 0 To UBound(dataArr, 1)
                xlsht.Cells(i + 1, j + 1).Value = dataArr(i, j)
            Next i
        Next j
    End If

    '--------------
    'Unique Setting
    '--------------
    Dim rngs As Range
    Dim rng As Range

    Dim fxReceive As Double
    Dim fxPay As Double
    
    fxReceive = 0
    fxPay = 0
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row

    If (lastRow > 1) Then
        Set rngs = xlsht.Range("C2:C" & lastRow)

        For Each rng In rngs
            If CStr(rng.Value) = "155930402" Then
                fxReceive = fxReceive + rng.Offset(0, 2).Value
            ElseIf CStr(rng.Value) = "255930402" Then
                fxPay = fxPay + rng.Offset(0, 2).Value
            End If
        Next rng

        fxReceive = ABs(Round(fxReceive / 1000, 0))
        fxPay = ABs(Round(fxPay / 1000, 0))
    End If
    
    xlsht.Range("CNY1_其他金融資產_淨額").Value = fxReceive
    rpt.SetField "CNY1", "CNY1_其他金融資產_淨額", CStr(fxReceive)

    xlsht.Range("CNY1_其他").Value = fxReceive
    rpt.SetField "CNY1", "CNY1_其他", CStr(fxReceive)

    xlsht.Range("CNY1_資產總計").Value = fxReceive
    rpt.SetField "CNY1", "CNY1_資產總計", CStr(fxReceive)

    xlsht.Range("CNY1_其他金融負債").Value = fxPay
    rpt.SetField "CNY1", "CNY1_其他金融負債", CStr(fxPay)

    xlsht.Range("CNY1_其他什項金融負債").Value = fxPay
    rpt.SetField "CNY1", "CNY1_其他什項金融負債", CStr(fxPay)

    xlsht.Range("CNY1_負債總計").Value = fxPay
    rpt.SetField "CNY1", "CNY1_負債總計", CStr(fxPay)
    
    xlsht.Range("T2:T100").NumberFormat = "#,##,##"
    
    ' 1.Validation filled all value (NO Null value exist)
    ' 2.Update Access DB
    If rpt.ValidateFields() Then
        Dim key As Variant
        Dim allValues As Object, allPositions As Object

        ' key 格式 "wsName|fieldName"
        Set allValues = rpt.GetAllFieldValues()  
        Set allPositions = rpt.GetAllFieldPositions()

        For Each key In allValues.Keys
            ' UpdateRecord gDBPath, gDataMonthString, rpt.ReportName, key, allPositions(key), CDbl(allValues(key))
            UpdateRecord gDBPath, gDataMonthString, rpt.ReportName, key, allPositions(key), allValues(key)
        Next key
    End If
End Sub

Public Sub Process_FB1()
    '=== Equal Setting ===
    'Fetch Query Access DB table
    Dim dataArr As Variant

    'Declare worksheet and handle data
    Dim xlsht As Worksheet

    Dim i As Integer, j As Integer

    Dim reportTitle As String
    Dim queryTable As String

    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("FB1")
    
    reportTitle = "FB1"

    queryTable = "FB1_OBU_AC4620B_Subtotal"

    ' dataArr = GetAccessDataAsArray(gDBPath, queryTable, gDataMonthString)
    dataArr = GetAccessDataAsArray(gDBPath, queryTable, gDataMonthString)
    
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    xlsht.Range("A:B").ClearContents

    '=== Paste Queyr Table into Excel ===
    If Err.Number <> 0 Or LBound(dataArr) > UBound(dataArr) Then
        MsgBox reportTitle & ": " & queryTable & " 資料表無資料"
        WriteLog reportTitle & ": " & queryTable & " 資料表無資料"
    Else
        For j = 0 To UBound(dataArr, 2)
            For i = 0 To UBound(dataArr, 1)
                xlsht.Cells(i + 1, j + 1).Value = dataArr(i, j)
            Next i
        Next j
        MsgBox reportTitle & ": " & queryTable & " 資料表有資料，此表單尚無有資料紀錄，尚請確認。"
        WriteLog reportTitle & ": " & queryTable & " 資料表有資料，此表單尚無有資料紀錄，尚請確認。"
    End If

    If rpt.ValidateFields() Then
        Dim key As Variant
        Dim allValues As Object, allPositions As Object

        ' key 格式 "wsName|fieldName"
        Set allValues = rpt.GetAllFieldValues()  
        Set allPositions = rpt.GetAllFieldPositions()

        For Each key In allValues.Keys
            ' UpdateRecord gDBPath, gDataMonthString, rpt.ReportName, key, allPositions(key), CDbl(allValues(key))
            UpdateRecord gDBPath, gDataMonthString, rpt.ReportName, key, allPositions(key), allValues(key)
        Next key
    End If
End Sub




'========================================================================
' 共用Utility

Option Explicit

'=== Connection to Access DBs ===
'=== Fetch 2d Array From Access Query Tables: Return 2d Arrays with Columns And Datas ===
Public Function GetAccessDataAsArray(ByVal DBPath As String, _
                                     ByVal QueryName As String, _
                                     Optional ByVal dataMonthString As String = vbNullString) As Variant
    Dim conn As Object, cmd As Object, rs As Object
    Dim dataArr As Variant
    Dim colCount As Integer, rowCount As Integer
    Dim headerArr() As String, i As Integer, j As Integer
    On Error GoTo ErrHandler
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = conn
    cmd.CommandText = QueryName
    cmd.CommandType = 4 ' 儲存查詢
    If dataMonthString <> vbNullString Then
        cmd.Parameters.Append cmd.CreateParameter("DataMonthParam", 200, 1, 255, dataMonthString)
    End If
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3 ' adUseClient
    rs.Open cmd
    If rs Is Nothing Or rs.EOF Then
        WriteLog QueryName & "查詢結果為空，請檢查資料庫與查詢條件。"
        GetAccessDataAsArray = Array()
        Exit Function
    End If
    colCount = rs.Fields.Count
    ReDim headerArr(0 To colCount - 1)
    For i = 0 To colCount - 1
        headerArr(i) = rs.Fields(i).Name
    Next i
    dataArr = rs.GetRows()
    rowCount = UBound(dataArr, 2) + 1
    Dim resultArr() As Variant
    ReDim resultArr(0 To rowCount, 0 To colCount - 1)
    ' 第一列存放欄位名稱
    For i = 0 To colCount - 1
        resultArr(0, i) = headerArr(i)
    Next i
    ' 後續存放資料
    For i = 0 To colCount - 1
        For j = 1 To rowCount
            resultArr(j, i) = dataArr(i, j - 1)
        Next j
    Next i
    rs.Close
    conn.Close
    Set rs = Nothing
    Set cmd = Nothing
    Set conn = Nothing
    GetAccessDataAsArray = resultArr
    Exit Function
ErrHandler:
    MsgBox "發生錯誤: " & Err.Description, vbCritical
    WriteLog "發生錯誤: " & Err.Description
    GetAccessDataAsArray = Array()
End Function

' Validate UserInput With Form yyyy/mm
Public Function IsValidDataMonth(ByVal userInput As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Pattern = "^\d{4}/(0[1-9]|1[0-2])$"
        .IgnoreCase = True
        .Global = False
    End With
    IsValidDataMonth = regex.Test(Trim(userInput))
End Function

'Insert Record to Access DB (Used for Initial Null Data Creation)
Public Sub InsertIntoTable(ByVal DBPath As String, _
                           ByVal tableName As String, _
                           ByVal dataMonthString As String, _
                           ByVal reportName As String, _
                           ByVal worksheetName_fieldKey As String, _
                           ByVal fieldValue As String, _
                           ByVal fieldAddress As String)
    Dim conn As Object, cmd As Object
    Dim sql As String
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath
    sql = "INSERT INTO " & tableName & " (DataMonthString, ReportName, WorksheetName_FieldKey, FieldValue, FieldAddress, CaseCreatedAt) " & _
          "VALUES ('" & dataMonthString & "', '" & reportName & "', '" & WorksheetName_FieldKey & "', '" & fieldValue & "', '" & fieldAddress & "', Now());"
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = conn
    cmd.CommandText = sql
    cmd.Execute
    conn.Close
    Set conn = Nothing
    Set cmd = Nothing
End Sub

'Update Record for each Process Report
Public Sub UpdateRecord(ByVal DBPath As String, _
                        ByVal dataMonthString As String, _
                        ByVal reportName As String, _
                        ByVal worksheetName_fieldKey As String, _
                        ByVal fieldAddress As String, _
                        ByVal fieldValue As String)
    Dim conn As Object, cmd As Object
    Dim sql As String
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath
    
    ' sql = "UPDATE MonthlyDeclarationReport SET FieldValue = '" & fieldValue & "', CaseCreatedAt = Now() " & _
    '       "WHERE DataMonthString = '" & dataMonthString & "' " & _
    '       "AND ReportName = '" & reportName & "' " & _
    '       "AND WorksheetName_FieldKey = '" & worksheetName_fieldKey & "' " & _
    '       "AND FieldAddress = '" & fieldAddress & "';"

    sql = "UPDATE MonthlyDeclarationReport SET FieldValue = '" & fieldValue & "', CaseCreatedAt = Now() " & _
          "WHERE DataMonthString = '" & dataMonthString & "' " & _
          "AND ReportName = '" & reportName & "' " & _
          "AND WorksheetName_FieldKey = '" & worksheetName_fieldKey & "';"
    
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = conn
    cmd.CommandText = sql
    cmd.Execute
    
    conn.Close
    Set conn = Nothing
    Set cmd = Nothing
End Sub


' Transform to ROC format
Public Function ConvertToROCFormat(ByVal dataMonthString As String, _
                                   ByVal returnType As String) As String
    Dim parts() As String
    Dim rocYear As Integer
    Dim result As String

    parts = Split(dataMonthString, "/")
    rocYear = CInt(parts(0)) - 1911

    If returnType = "ROC" Then
        result = " 民國 " & CStr(rocYear) & " 年 " & parts(1) & " 月"
    ElseIf returnType = "NUM" Then
        result = CStr(rocYear) & parts(1)
    ElseIf returnType = "F1F2" Then
        result = CStr(rocYear) & "年" & parts(1) & "月份"
    End If
    
    ConvertToROCFormat = result
End Function

' Create FieldList for F1F2 Report
Function GenerateFieldList( _
    transactionTypes As Variant, _
    currencies       As Variant, _
    colLetters       As Variant, _
    startRows        As Variant _
) As Variant
    Dim nTypes  As Long, nCurs As Long, total As Long
    Dim result() As Variant
    Dim i As Long, j As Long, index As Long
    
    nTypes = UBound(transactionTypes) - LBound(transactionTypes) + 1
    nCurs  = UBound(currencies)       - LBound(currencies)       + 1
    total  = nTypes * nCurs
    ReDim result(0 To total - 1)
    
    index = 0
    For i = LBound(transactionTypes) To UBound(transactionTypes)
        For j = LBound(currencies) To UBound(currencies)
            result(index) = Array( _
                transactionTypes(i) & "_" & currencies(j), _
                colLetters(i) & (startRows(i) + j), _
                Null _
            )
            index = index + 1
        Next j
    Next i
    
    GenerateFieldList = result
End Function


' LogFile

Function GetLogFileName() As String
    Dim folderPath As String
    Dim uuid As String
    Dim fileName As String
    
    folderPath = ThisWorkbook.Path & "\LogFile_Frontend\"  ' 你也可以指定其他資料夾
    ' folderPath = "\\10.10.122.40\後台作業\99_個人資料夾\8.修豪\LogFile_Frontend\"
    uuid = CreateUUID()
    fileName = "LogFile_" & Format(Now, "yyyymmdd_hhnnss") & "_" & uuid & ".txt"
    
    GetLogFileName = folderPath & fileName
End Function

' 模擬UUID
Public Function CreateUUID() As String
    Randomize
    CreateUUID = Format(Now, "hhmmss") & _
                    Hex(Int(Rnd() * 65536)) & _
                    Hex(Int(Rnd() * 65536))
End Function

Sub WriteLog(logMessage As String, _
             Optional logFilePath As String = "")             
    Static logFile As String
    
    If logFilePath <> "" Then
        logFile = logFilePath
    ElseIf logFile = "" Then
        logFile = GetLogFileName()
    End If

    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open logFile For Append As #fileNum
    Print #fileNum, Format(Now, "yyyy-mm-dd hh:nn:ss") & " - " & logMessage
    Close #fileNum
End Sub


' Fetch ExchangeRate from AccessDB
Function GetExchangeRates(BaseCurrency As String, _
                          DataDate As Date, _
                          Direction As String, _
                          ParamArray QuoteCurrencies() As Variant) As Variant

    ' ===Example===
    ' Excel 365 以上:
    ' 任一儲存格輸入（以橫向為例）：
    ' =GetExchangeRates("USD", DATE(2025,3,31), "v", "TWD","JPY","GBP") 按 Enter 後，會自動「溢出」為多欄
    ' ===Example===
    ' Excel 2019 以前:
    ' 1.選取要填的區塊（例如 1×3或3×1儲存格範圍）
    ' 2.輸入公式 =GetExchangeRates("USD", DATE(2025,3,31), "v", "TWD","JPY","GBP")
    ' 3.同時按下 Ctrl+Shift+Enter，公式即填滿選取區域

    Dim conn As Object
    Dim rs As Object
    Dim DBPath As String
    Dim sql As String
    Dim i As Long
    Dim results() As Variant

    On Error GoTo ErrHandler

    Dim bCurr As String
    bCurr = UCase(BaseCurrency)

    Dim qCurr() As Variant
    ReDim qCurr(LBound(QuoteCurrencies) To UBound(QuoteCurrencies))
    For i = LBound(QuoteCurrencies) To UBound(QuoteCurrencies)
        qCurr(i) = UCase(QuoteCurrencies(i))
    Next i

    ' Access 資料庫路徑
    DBPath = ThisWorkbook.Path & "\" & ThisWorkbook.Sheets("ControlPanel").Range("DBsPathFileName").Value
    ' DBPath = "\\10.10.122.40\後台作業\99_個人資料夾\8.修豪\" & ThisWorkbook.Sheets("ControlPanel").Range("DBsPathFileName").Value

    ' Build connection
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath

    ' Output direction
    If UCase(Direction) = "V" Then
        ReDim results(1 To UBound(qCurr) + 1, 1 To 1)
    Else
        ReDim results(1 To 1, 1 To UBound(qCurr) + 1)
    End If

    ' Connect to AccessDB
    For i = LBound(qCurr) To UBound(qCurr)
        sql = "SELECT Rate FROM CloseRate " & _
              "WHERE BaseCurrency = '" & bCurr  & "' " & _
              "  AND QuoteCurrency = '" & qCurr(i) & "' " & _
              "  AND DataDate = #" & Format(DataDate, "yyyy\/mm\/dd") & "#"

        Set rs = conn.Execute(sql)

        Dim rateValue As Variant
        If Not rs.EOF Then
            ' rateValue = rs.Fields("Rate").Value
            rateValue = rs!Rate
        Else
            rateValue = "找不到匯率: " & bCurr & " 兌換 " & qCurr(i)
        End If

        If UCase(Direction) = "V" Then
            results(i + 1, 1) = rateValue
        Else
            results(1, i + 1) = rateValue
        End If

        rs.Close
    Next i

    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    GetExchangeRates = results
    Exit Function

ErrHandler:
    GetExchangeRates = "資料庫錯誤或參數錯誤"
End Function




'──────────────────────────────────────────
' 主函數：支援 Range 或 逗號分隔字串，
' 第一列為欄位名稱，且新增可選的 GroupFlag 篩選
Function GetAccountCodeMapFlex( _
    CategoryParam As Variant, _
    Optional GroupFlagParam As Variant, _       '【★】新增：GroupFlagParam
    Optional SubTypeParam As Variant, _
    Optional TypeParam As Variant) As Variant

    Dim dbPath As String
    dbPath = "C:\Your\Path\To\Database.accdb"  ' ← 修改為你的 Access 資料庫路徑

    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")

    Dim sql As String
    Dim whereClauses As Collection
    Set whereClauses = New Collection

    '── 1) 必填：Category
    Dim catClause As String
    catClause = BuildInClauseParam(CategoryParam, "Category")
    If catClause = "" Then
        GetAccountCodeMapFlex = CVErr(xlErrValue)
        Exit Function
    End If
    whereClauses.Add catClause

    '── 2) 可選：GroupFlag
    Dim grpClause As String
    grpClause = BuildInClauseParam(GroupFlagParam, "GroupFlag")   '【★】新增
    If grpClause <> "" Then whereClauses.Add grpClause         '【★】新增

    '── 3) 可選：AssetMeasurementSubType
    Dim subClause As String
    subClause = BuildInClauseParam(SubTypeParam, "AssetMeasurementSubType")
    If subClause <> "" Then whereClauses.Add subClause

    '── 4) 可選：AssetMeasurementType
    Dim typeClause As String
    typeClause = BuildInClauseParam(TypeParam, "AssetMeasurementType")
    If typeClause <> "" Then whereClauses.Add typeClause

    '── 組出 SQL，調整回傳欄位順序：AccountCode, GroupFlag, AccountTitle
    sql = "SELECT AccountCode, GroupFlag, AccountTitle FROM AccountCodeMap WHERE "
    Dim part As Variant
    For Each part In whereClauses
        sql = sql & part & " AND "
    Next
    sql = Left(sql, Len(sql) - 5)  ' 去掉最後的 " AND "

    '── 執行查詢
    On Error GoTo ErrHandler
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"
    rs.Open sql, conn

    '── 處理回傳：第一列放欄位名稱，之後才是資料
    If rs.EOF Then
        ' 若無資料，只回欄位名稱
        Dim noDataArr() As Variant
        ReDim noDataArr(0 To 0, 0 To rs.Fields.Count - 1)
        Dim fi As Integer
        For fi = 0 To rs.Fields.Count - 1
            noDataArr(0, fi) = rs.Fields(fi).Name
        Next
        GetAccountCodeMapFlex = noDataArr

    Else
        Dim rawArr As Variant
        rawArr = rs.GetRows()  ' fields × rows

        Dim nFields As Long, nRows As Long
        nFields = UBound(rawArr, 1) + 1  ' 3 個欄位
        nRows   = UBound(rawArr, 2) + 1

        Dim outArr() As Variant
        ReDim outArr(0 To nRows, 0 To nFields - 1)

        ' 第一列：欄位名稱
        Dim f As Long, r As Long
        For f = 0 To nFields - 1
            outArr(0, f) = rs.Fields(f).Name
        Next

        ' 之後各列：資料
        For r = 0 To nRows - 1
            For f = 0 To nFields - 1
                outArr(r + 1, f) = rawArr(f, r)
            Next
        Next

        GetAccountCodeMapFlex = outArr
    End If

    rs.Close: conn.Close
    Set rs = Nothing: Set conn = Nothing
    Exit Function

ErrHandler:
    GetAccountCodeMapFlex = "Error: " & Err.Description
    On Error GoTo 0
End Function

'──────────────────────────────────────────
' 輔助函數：根據參數類型（Range 或 String）建構 IN(...) 子句
Private Function BuildInClauseParam(paramIn As Variant, fieldName As String) As String
    Dim items As Collection
    Set items = New Collection
    Dim v As Variant

    If TypeName(paramIn) = "Range" Then
        For Each v In paramIn.Cells
            If Trim(v.Value & "") <> "" Then items.Add CStr(v.Value)
        Next
    ElseIf VarType(paramIn) = vbString Then
        Dim arr As Variant
        arr = Split(paramIn, ",")
        For Each v In arr
            v = Trim(v)
            If v <> "" Then items.Add v
        Next
    End If

    Dim sqlVals As String, elem As Variant
    For Each elem In items
        sqlVals = sqlVals & "'" & Replace(elem, "'", "''") & "',"
    Next
    If sqlVals <> "" Then
        sqlVals = Left(sqlVals, Len(sqlVals) - 1)
        BuildInClauseParam = fieldName & " IN (" & sqlVals & ")"
    Else
        BuildInClauseParam = ""
    End If
End Function


' ### 範例呼叫

' 假設：

' * **A1\:A2** 放 Category（必填）
' * **B1\:B2** 放 GroupFlag（可選）
' * **C1\:C3** 放 AssetMeasurementSubType（可選）
' * **D1\:D2** 放 AssetMeasurementType（可選）

' ```excel
' =GetAccountCodeMapFlex(A1:A2, B1:B2, C1:C3, D1:D2)
' ```

' 或混用字串方式：

' ```excel
' =GetAccountCodeMapFlex("外幣債,國內股","外幣債,國內債","FVPL_GovBond_Foreign,AC_CompanyBond_Foreign","FVPL")
' ```

' * 不想篩 GroupFlag，直接把第二個參數留空或省略即可：

'   ```excel
'   =GetAccountCodeMapFlex(A1:A2,,C1:C3,D1:D2)
'   ```
' * 第一列會自動顯示 `AccountCode | GroupFlag | AccountTitle`，後續列才是資料。




' =================================================================

Module for Dynamic Cases

Public Sub Process_FB3A()
    '=== Equal Setting ===
    'Fetch Query Access DB table
    Dim dataArr As Variant

    'Declare worksheet and handle data
    Dim xlsht As Worksheet

    Dim i As Integer, j As Integer
    Dim lastRow As Integer

    Dim reportTitle As String
    Dim queryTable As String

    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("FB3A")

    reportTitle = "FB3A"
    queryTable = "FB3A_OBU_MM4901B"

    ' dataArr = GetAccessDataAsArray(gDBPath, queryTable, gDataMonthString)
    dataArr = GetAccessDataAsArray(gDBPath, queryTable, gDataMonthString)

    Set xlsht = ThisWorkbook.Sheets(reportTitle)

    'Clear Excel Data
    xlsht.Range("A:J").ClearContents
    xlsht.Range("K2:Q200").ClearContents


    '=== Paste Queyr Table into Excel ===
    If Err.Number <> 0 Or LBound(dataArr) > UBound(dataArr) Then
        MsgBox reportTitle & ": " & queryTable & " 資料表無資料"
        WriteLog reportTitle & ": " & queryTable & " 資料表無資料"
    Else
        For j = 0 To UBound(dataArr, 2)
            For i = 0 To UBound(dataArr, 1)
                xlsht.Cells(i + 1, j + 1).Value = dataArr(i, j)
            Next i
        Next j
    End If
    

    '--------------
    'Unique Setting
    '--------------
    Dim BankCode As Variant
    Dim CounterParty As String, Category As String
    Dim Amount As Double

    Dim targetRow As Long
    Dim targetCol As String
    
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row

    ' 設定第二部份記錄的起始列（Row 10）
    targetRow = 10

    If (lastRow > 1) Then
        ' 逐列處理原始資料（從第二列開始）
        For i = 2 To lastRow
            ' 讀取原始資料欄位值（依照題目定義的欄位順序）
            ' 原始資料欄位：
            ' A: DataID
            ' B: DataMonthString
            ' C: DealDate
            ' D: DealID
            ' E: CounterParty
            ' F: MaturityDate
            ' G: CurrencyType
            ' H: Amount
            ' I: Category
            ' J: BankCode
            
            ' 銀行代碼
            BankCode = xlsht.Cells(i, "J").Value        
            ' CounterParty
            CounterParty = xlsht.Cells(i, "E").Value
            ' 金額
            Amount = Round(xlsht.Cells(i, "H").value / 1000, 0)
            ' 類別 
            Category = xlsht.Cells(i, "I").Value               
            ' TWTP_MP / OBU_MP / TWTP_MT / OBU_MT
            
            ' K：BankCode
            xlsht.Cells(i, "K").Value = BankCode
            ' L：CounterParty
            xlsht.Cells(i, "L").Value = CounterParty
            
            ' 根據 Category 將金額填入對應分類欄位
            Select Case Category
                Case "DBU_MP"
                    ' M：DBU_MP
                    xlsht.Cells(i, "M").Value = Amount
                Case "OBU_MP"
                    ' N：OBU_MP
                    xlsht.Cells(i, "N").Value = Amount
                Case "DBU_MT"
                    ' O：DBU_MT
                    xlsht.Cells(i, "O").Value = Amount
                Case "OBU_MT"
                    ' P：OBU_MT
                    xlsht.Cells(i, "P").Value = Amount
            End Select
            

            ' 二、記錄儲存格位置和數值（輸出位置由 Row 10 開始）
            ' 這邊假設：BankCode 記錄在 C 欄；金額根據 Category 記錄在 E (TWTP_MP) / F (OBU_MP) / G (TWTP_MT) / H (OBU_MT)

            Select Case Category
                Case "DBU_MP"
                    targetCol = "E"
                Case "OBU_MP"
                    targetCol = "F"
                Case "DBU_MT"
                    targetCol = "G"
                Case "OBU_MT"
                    targetCol = "H"
            End Select

            xlsht.Cells(i, "Q").Value =  targetCol & CStr(targetRow)

            ' rpt.SetField "FOA", "FB3A_BankCode", "C" & CStr(targetRow), BankCode
            ' rpt.SetField "FOA", "FB3A_Amount", targetCol & CStr(targetRow), Amount

            rpt.AddDynamicField "FOA", "FB3A_BankCode_" & Format(BankCode, "0000"), "C" & CStr(targetRow), CStr(Format(BankCode, "0000"))
            rpt.AddDynamicField "FOA", "FB3A_Amount_" & Format(BankCode, "0000"), targetCol & CStr(targetRow), CStr(Amount)
            
            InsertIntoTable gDBPath, "MonthlyDeclarationReport", gDataMonthString, "FB3A", "FOA|FB3A_BankCode_" & Format(BankCode, "0000"), CStr(Format(BankCode, "0000")), "C" & CStr(targetRow)
            InsertIntoTable gDBPath, "MonthlyDeclarationReport", gDataMonthString, "FB3A", "FOA|FB3A_Amount_" & Format(BankCode, "0000"), CStr(Amount), targetCol & CStr(targetRow)

            UpdateRecord gDBPath, gDataMonthString, rpt.ReportName, "FOA|FB3A_申報時間", "C2", gDataMonthStringROC
            
            targetRow = targetRow + 1
        Next i

        xlsht.Range("M2:M100").NumberFormat = "#,##,##"
        xlsht.Range("N2:N100").NumberFormat = "#,##,##"
        xlsht.Range("O2:O100").NumberFormat = "#,##,##"
        xlsht.Range("P2:P100").NumberFormat = "#,##,##"
    End If
End Sub


Public Sub Process_FM2()
    '=== Equal Setting ===
    'Fetch Query Access DB table
    Dim dataArr_1 As Variant
    Dim dataArr_2 As Variant
    Dim dataArr_3 As Variant

    'Declare worksheet and handle data
    Dim xlsht As Worksheet

    Dim i As Integer, j As Integer
    Dim lastRow As Integer

    Dim reportTitle As String
    Dim queryTable_1 As String
    Dim queryTable_2 As String
    Dim queryTable_3 As String

    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("FM2")
    
    reportTitle = "FM2"
    queryTable_1 = "FM2_OBU_MM4901B_LIST"
    queryTable_2 = "FM2_OBU_MM4901B_Subtotal"
    queryTable_3 = "FM2_OBU_MM4901B_Subtotal_BankCode"

    ' dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1, gDataMonthString)
    ' dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2, gDataMonthString)
    ' dataArr_3 = GetAccessDataAsArray(gDBPath, queryTable_3, gDataMonthString)
    dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1, gDataMonthString)
    dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2, gDataMonthString)
    dataArr_3 = GetAccessDataAsArray(gDBPath, queryTable_3, gDataMonthString)

    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    xlsht.Range("A:P").ClearContents
    xlsht.Range("Q2:W200").ClearContents

    '=== Paste Queyr Table into Excel ===
    If Err.Number <> 0 Or LBound(dataArr_1) > UBound(dataArr_1) Then
        MsgBox "資料有誤: " & reportTitle & "| " & queryTable_1 & " 資料表無資料"
        WriteLog "資料有誤: " & reportTitle & "| " & queryTable_1 & " 資料表無資料"
    Else
        For j = 0 To UBound(dataArr_1, 2)
            For i = 0 To UBound(dataArr_1, 1)
                xlsht.Cells(i + 1, j + 1).Value = dataArr_1(i, j)
            Next i
        Next j
    End If


    If Err.Number <> 0 Or LBound(dataArr_2) > UBound(dataArr_2) Then
        MsgBox "資料有誤: " & reportTitle & "| " & queryTable_2 & " 資料表無資料"
        WriteLog "資料有誤: " & reportTitle & "| " & queryTable_2 & " 資料表無資料"
    Else
        For j = 0 To UBound(dataArr_2, 2)
            For i = 0 To UBound(dataArr_2, 1)
                xlsht.Cells(i + 1, j + 10).Value = dataArr_2(i, j)
            Next i
        Next j
    End If


    If Err.Number <> 0 Or LBound(dataArr_3) > UBound(dataArr_3) Then
        MsgBox "資料有誤: " & reportTitle & "| " & queryTable_3 & " 資料表無資料"
        WriteLog "資料有誤: " & reportTitle & "| " & queryTable_3 & " 資料表無資料"
    Else
        For j = 0 To UBound(dataArr_3, 2)
            For i = 0 To UBound(dataArr_3, 1)
                xlsht.Cells(i + 1, j + 13).Value = dataArr_3(i, j)
            Next i
        Next j
    End If

    '--------------
    'Unique Setting
    '--------------
    Dim BankCode As Variant
    Dim CounterParty As String, Category As String
    Dim Amount As Double

    Dim pasteRow As Long
    Dim targetRow As Long
    Dim targetCol As String
    
    ' 設定第二部份記錄的起始列（Row 10）
    pasteRow = 2
    targetRow = 10
    lastRow = xlsht.Cells(xlsht.Rows.Count, "M").End(xlUp).Row
    
    ' 逐列處理原始資料（從第二列開始）
    For i = 2 To lastRow
        ' 讀取原始資料欄位值（依照題目定義的欄位順序）
        ' 原始資料欄位：
        ' A: DataID
        ' B: DataMonthString
        ' C: DealDate
        ' D: DealID
        ' E: CounterParty
        ' F: MaturityDate
        ' G: CurrencyType
        ' H: Amount
        ' I: Category
        ' J: BankCode
        

        If (Not IsEmpty(xlsht.Cells(i, "P").Value)) Then
            '銀行代碼
            BankCode = xlsht.Cells(i, "P").Value        
            'CounterParty
            CounterParty = xlsht.Cells(i, "M").Value
            ' 金額
            Amount = Round(xlsht.Cells(i, "O").value / 1000, 0)
            ' 類別 
            Category = xlsht.Cells(i, "N").Value               
            'TWTP_MP / OBU_MP / TWTP_MT / OBU_MT
            
            ' K：BankCode
            xlsht.Cells(pasteRow, "Q").Value = BankCode
            ' L：CounterParty
            xlsht.Cells(pasteRow, "R").Value = CounterParty

            ' 根據 Category 將金額填入對應分類欄位
            Select Case Category
                Case "DBU_MP"
                    ' M：TWTP_MP
                    xlsht.Cells(pasteRow, "S").Value = Amount
                Case "OBU_MP"
                    ' N：OBU_MP
                    xlsht.Cells(pasteRow, "T").Value = Amount
                Case "DBU_MT"
                    ' O：TWTP_MT
                    xlsht.Cells(pasteRow, "U").Value = Amount
                Case "OBU_MT"
                    ' P：OBU_MT
                    xlsht.Cells(pasteRow, "V").Value = Amount
            End Select
        

            ' 二、記錄儲存格位置和數值（輸出位置由 Row 10 開始）
            ' 這邊假設：BankCode 記錄在 C 欄；金額根據 Category 記錄在 E (TWTP_MP) / F (OBU_MP) / G (TWTP_MT) / H (OBU_MT)
            Select Case Category
                Case "DBU_MP"
                    targetCol = "E"
                Case "OBU_MP"
                    targetCol = "F"
                Case "DBU_MT"
                    targetCol = "G"
                Case "OBU_MT"
                    targetCol = "H"
            End Select
            
            xlsht.Cells(pasteRow, "W").Value =  targetCol & CStr(targetRow)
            ' rpt.SetField "FOA", "FM2_BankCode", "C" & CStr(targetRow), BankCode
            ' rpt.SetField "FOA", "FM2_Amount", targetCol & CStr(targetRow), Amount

            rpt.AddDynamicField "FOA", "FM2_BankCode_" & Format(BankCode, "0000"), "C" & CStr(targetRow), CStr(Format(BankCode, "0000"))
            rpt.AddDynamicField "FOA", "FM2_Amount_" & Format(BankCode, "0000"), targetCol & CStr(targetRow), CStr(Amount) 
            
            InsertIntoTable gDBPath, "MonthlyDeclarationReport", gDataMonthString, "FM2", "FOA|FM2_BankCode_" & Format(BankCode, "0000"), CStr(Format(BankCode, "0000")), "C" & CStr(targetRow)
            InsertIntoTable gDBPath, "MonthlyDeclarationReport", gDataMonthString, "FM2", "FOA|FM2_Amount_" & Format(BankCode, "0000"), CStr(Amount), targetCol & CStr(targetRow)

            UpdateRecord gDBPath, gDataMonthString, rpt.ReportName, "FOA|FM2_申報時間", "C2", gDataMonthStringROC
            
            pasteRow = pasteRow + 1
            targetRow = targetRow + 1
        End If
    Next i

    xlsht.Range("S2:S100").NumberFormat = "#,##,##"
    xlsht.Range("T2:T100").NumberFormat = "#,##,##"
    xlsht.Range("U2:U100").NumberFormat = "#,##,##"
    xlsht.Range("V2:V100").NumberFormat = "#,##,##"
End Sub






'================================================
Answer here

以下是整合後的完整範例，示範如何在全域只跟 Access 要一次最大 Index，並且後續每次插入都自動遞增：

```vb
'=== 模組最上方：全域宣告 ===
Option Explicit

' 使用者輸入與路徑設定
Public gDataMonthString         As String         ' 使用者輸入的資料月份 (yyyy/mm)
Public gDataMonthStringROC      As String         ' ROC 格式
Public gDataMonthStringROC_NUM  As String         ' ROC_NUM 格式
Public gDataMonthStringROC_F1F2 As String         ' ROC_F1F2 格式
Public gDBPath                  As String         ' Access 資料庫路徑
Public gReportFolder            As String         ' 原始空白報表資料夾
Public gOutputFolder            As String         ' 產出報表資料夾
Public gReportNames             As Variant        ' 報表名稱陣列
Public gReports                 As Collection     ' clsReport 物件集合

' 全域 Index 計數器
Public gRecIndex                As Long

'=== 取得當月最大 Index（只執行一次） ===
Public Function GetMaxIndex(ByVal DBPath As String, _
                            ByVal tableName As String, _
                            ByVal dataMonthString As String) As Long
    Dim conn As Object, rs As Object
    Dim sql  As String
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath
    
    sql = "SELECT MAX([Index]) AS MaxIdx FROM " & tableName & _
          " WHERE DataMonthString='" & dataMonthString & "';"
    Set rs = conn.Execute(sql)
    
    If Not rs.EOF And Not IsNull(rs.Fields("MaxIdx").Value) Then
        GetMaxIndex = rs.Fields("MaxIdx").Value
    Else
        GetMaxIndex = 0
    End If
    
    rs.Close:  conn.Close
    Set rs   = Nothing
    Set conn = Nothing
End Function

'=== 修改後的 InsertIntoTable：改用全域 gRecIndex，並自動 +1 ===
Public Sub InsertIntoTable(ByVal DBPath As String, _
                           ByVal tableName As String, _
                           ByVal dataMonthString As String, _
                           ByVal reportName As String, _
                           ByVal worksheetName_fieldKey As String, _
                           ByVal fieldValue As String, _
                           ByVal fieldAddress As String)
    Dim conn      As Object
    Dim cmd       As Object
    Dim sql       As String
    Dim thisIndex As Long
    
    ' 使用全域計數器
    thisIndex = gRecIndex
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath
    
    sql = "INSERT INTO " & tableName & _
          " (DataMonthString, ReportName, WorksheetName_FieldKey, FieldValue, FieldAddress, [Index], CaseCreatedAt) " & _
          "VALUES (" & _
            "'" & dataMonthString & "', " & _
            "'" & reportName & "', " & _
            "'" & worksheetName_fieldKey & "', " & _
            "'" & fieldValue & "', " & _
            "'" & fieldAddress & "', " & _
            thisIndex & ", Now());"
    
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = conn
    cmd.CommandText     = sql
    cmd.Execute
    
    conn.Close
    Set cmd  = Nothing
    Set conn = Nothing
    
    ' 每插一筆就遞增
    gRecIndex = gRecIndex + 1
End Sub

'=== 主流程入口 Main ===
Public Sub Main()
    Dim isInputValid As Boolean
    
    ' 1. 讓使用者輸入並驗證 gDataMonthString
    isInputValid = False
    Do
        gDataMonthString = InputBox("請輸入資料月份 (格式: yyyy/mm):", "輸入資料月份")
        If IsValidDataMonth(gDataMonthString) Then
            isInputValid = True
        ElseIf Trim(gDataMonthString) = "" Then
            MsgBox "請輸入報表資料所屬的年度/月份 (例如: 2024/01)", vbExclamation, "輸入錯誤"
            WriteLog "請輸入報表資料所屬的年度/月份 (例如: 2024/01)"
        Else
            MsgBox "格式錯誤，請輸入正確格式 (yyyy/mm)", vbExclamation, "格式錯誤"
            WriteLog "格式錯誤，請輸入正確格式 (yyyy/mm)"
        End If
    Loop Until isInputValid
    
    ' 2. 轉換 ROC 格式
    gDataMonthStringROC      = ConvertToROCFormat(gDataMonthString, "ROC")
    gDataMonthStringROC_NUM  = ConvertToROCFormat(gDataMonthString, "NUM")
    gDataMonthStringROC_F1F2 = ConvertToROCFormat(gDataMonthString, "F1F2")
    
    ' 3. 設定路徑與報表清單
    gDBPath       = ThisWorkbook.Path & "\" & ThisWorkbook.Sheets("ControlPanel").Range("DBsPathFileName").Value
    gReportFolder = ThisWorkbook.Path & "\" & ThisWorkbook.Sheets("ControlPanel").Range("EmptyReportPath").Value
    gOutputFolder = ThisWorkbook.Path & "\" & ThisWorkbook.Sheets("ControlPanel").Range("OutputReportPath").Value
    gReportNames  = Array("CNY1", "FB1", "FB2", "FB3", "FB3A", "FM5", "FM11", "FM13", _
                          "AI821", "Table2", "FB5", "FB5A", "FM2", "FM10", "F1_F2", _
                          "Table41", "AI602", "AI240")
    
    ' 4. 只呼叫一次，取得目前最大 Index，並設定下一筆要用的 gRecIndex
    gRecIndex = GetMaxIndex(gDBPath, "MonthlyDeclarationReport", gDataMonthString) + 1
    
    ' 5. Process A: 初始化所有報表到 Access（初始 Null Data）
    Call InitializeReports
    MsgBox "完成 Process A"
    WriteLog "完成 Process A"
    
    ' 6. Process B: 各報表製表並更新 DB
    Call Process_CNY1
    Call Process_FB1
    Call Process_FB2
    Call Process_FB3
    Call Process_FB3A
    Call Process_FM5
    Call Process_FM11
    Call Process_FM13
    Call Process_AI821
    Call Process_Table2
    Call Process_FB5
    Call Process_FB5A
    Call Process_FM2
    Call Process_FM10
    Call Process_F1_F2
    Call Process_Table41
    Call Process_AI602
    Call Process_AI240
    MsgBox "完成 Process B"
    WriteLog "完成 Process B"
    
    ' 7. Process C: 填入 Excel 報表並另存
    Call UpdateExcelReports
    MsgBox "完成 Process C (全部處理程序完成)"
    WriteLog "完成 Process C (全部處理程序完成)"
End Sub

'=== A. 初始化所有報表並插入 Null Data ===
Public Sub InitializeReports()
    Dim rpt       As clsReport
    Dim rptName   As Variant, key As Variant
    Dim combined  As Object
    
    Set gReports = New Collection
    
    For Each rptName In gReportNames
        Set rpt = New clsReport
        rpt.Init rptName, gDataMonthStringROC, gDataMonthStringROC_NUM, gDataMonthStringROC_F1F2
        gReports.Add rpt, rptName
        
        Set combined = rpt.GetAllFieldPositions
        For Each key In combined.Keys
            ' 全部呼 InsertIntoTable，內部自動用 gRecIndex
            InsertIntoTable gDBPath, "MonthlyDeclarationReport", _
                            gDataMonthString, rptName, key, "", combined(key)
        Next key
    Next rptName
    
    MsgBox "完成'報表初始欄位資訊儲存'及'初始資料庫資料建立'"
    WriteLog "完成'報表初始欄位資訊儲存'及'初始資料庫資料建立'"
End Sub

'=== 範例：Process_FB3A 只要呼 InsertIntoTable 就好（省略繁複細節） ===
Public Sub Process_FB3A()
    ' … 取 dataArr, 填 Excel, rpt 物件處理 …  
    ' 每次需要寫入 DB，就：
    InsertIntoTable gDBPath, "MonthlyDeclarationReport", _
                    gDataMonthString, "FB3A", _
                    "FOA|FB3A_BankCode_" & bankCodeKey, bankCodeValue, bankCodeAddress
    InsertIntoTable gDBPath, "MonthlyDeclarationReport", _
                    gDataMonthString, "FB3A", _
                    "FOA|FB3A_Amount_" & bankCodeKey, amountValue, amountAddress
    ' … 其餘流程 …
End Sub
```

**重點說明**

1. **全域只執行一次** `GetMaxIndex`，結果存到 `gRecIndex`。
2. **`InsertIntoTable`** 內部都用 `gRecIndex`，並在每次執行後自動 `gRecIndex = gRecIndex + 1`。
3. 之後呼 `InsertIntoTable` 時，完全不用再帶 Index，也不用再跟 Access 查最大值。

這樣即可大幅降低與 Access 互動次數，並確保每筆資料的 `[Index]` 欄位連續、遞增。

https://chatgpt.com/share/68221a61-b738-8010-964f-97ca96618c7a
