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

    Dim RecIndex As Long

    RecIndex = gRecIndex

    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath
    sql = "INSERT INTO " & tableName & " (DataMonthString, ReportName, WorksheetName_FieldKey, FieldValue, FieldAddress, [RecordIndex], CaseCreatedAt) " & _
          "VALUES ('" & dataMonthString & "', '" & reportName & "', '" & WorksheetName_FieldKey & "', '" & fieldValue & "', '" & fieldAddress & "', " & RecIndex & ", Now());"

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
    ' folderPath = "\\10.10.122.40\後台作業\99_個人資料夾\8.修豪\DbsMReport20250513_V1\LogFile_Frontend\"
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

Public Function GetMaxRecordIndex(ByVal DBPath As String, _
                                  ByVal tableName As String, _
                                  ByVal dataMonthString As String) As Long
    Dim conn As Object, rs As Object
    Dim sql  As String
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath
    
    sql = "SELECT MAX([RecordIndex]) AS MaxRecIndex FROM " & tableName & " WHERE DataMonthString='" & dataMonthString & "';"
    Set rs = conn.Execute(sql)
    
    If Not rs.EOF And Not IsNull(rs.Fields("MaxRecIndex").Value) Then
        GetMaxRecordIndex = rs.Fields("MaxRecIndex").Value
    Else
        GetMaxRecordIndex = 0
    End If
    
    rs.Close:  conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Function


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
    ' DBPath = "\\10.10.122.40\後台作業\99_個人資料夾\8.修豪\DbsMReport20250513_V1\" & ThisWorkbook.Sheets("ControlPanel").Range("DBsPathFileName").Value

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

Public Function GetMapData(ByVal DBPath As String, _
                           ByVal reportName As String, _
                           ByVal tableName As String) As Variant
    Dim conn As Object, rs As Object
    Dim sql As String
    Dim results() As Variant
    Dim i As Long
    Dim arr As Variant

    ' 1. 建立 ADO 连接
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath
    
    ' 2. 根据 tableName 组织 SQL
    Select Case LCase(tableName)
        Case "fiedlvaluepositionmap"
            sql = "SELECT fp.TargetSheetName, fp.SourceNameTag, fp.TargetCellAddress " & _
                  "FROM FiedlValuePositionMap AS fp " & _
                  "INNER JOIN Report AS r ON fp.ReportID = r.ReportID " & _
                  "WHERE r.ReportName = '" & reportName & "' " & _
                  "ORDER BY fp.DataId;"
        Case "querytablemap"
            sql = "SELECT qm.QueryTableName, qm.ImportColName, qm.ImportColNumber " & _
                  "FROM QueryTableMap AS qm " & _
                  "INNER JOIN Report AS r ON qm.ReportID = r.ReportID " & _
                  "WHERE r.ReportName = '" & reportName & "' " & _
                  "ORDER BY qm.DataId;"
        Case Else
            GetMapData = Array()
            Exit Function
    End Select
    
    Set rs = conn.Execute(sql)
    If rs.EOF Then
        GetMapData = Array()    ' 没有紀錄
    Else
        arr = rs.GetRows
    
        ' 转置成 arr2(紀錄, 字段)
        Dim recCount As Long, fldCount As Long
        fldCount = UBound(arr, 1) + 1
        recCount = UBound(arr, 2) + 1
        ReDim results(0 To recCount - 1, 0 To fldCount - 1)
    
        Dim r As Long, f As Long
        For f = 0 To fldCount - 1
            For r = 0 To recCount - 1
                results(r, f) = arr(f, r)
            Next r
        Next f
    
        GetMapData = results
    End If
    
    rs.Close: conn.Close
    Set rs = Nothing: Set conn = Nothing
End Function
