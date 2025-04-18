我有定義class clsReport如下，
Option Explicit
' Report Title
Private clsReportName As String
' Dictionary：key = Worksheet Name，value = Dictionary( Keys "Fiedl Values" 與 "Field Addresses" )
Private clsWorksheets As Object
'=== 初始化報表 (根據報表名稱建立各工作表的欄位定義) ===
Public Sub Init(ByVal reportName As String, _
                ByVal dataMonthStringROC As String, _
                ByVal dataMonthStringROC_NUM As String, _
                ByVal dataMonthStringROC_F1F2 As String)
    clsReportName = reportName
    Set clsWorksheets = CreateObject("Scripting.Dictionary")
    
    Select Case reportName
        Case "F1_F2"
            Dim currencies_F1 As Variant, currencies_F2 As Variant
            Dim transactionTypes_F1 As Variant, transactionTypes_F2 As Variant
            Dim colLetters_F1 As Variant, colLetters_F2 As Variant
            Dim startRows As Variant
        
            Dim fieldList_F1 As Variant, fieldList_F2 As Variant
        
            ' F1幣別 and F2幣別 (Row)
            currencies_F1 = Array("JPY", "GBP", "CHF", "CAD", "AUD", "NZD", "SGD", "HKD", "ZAR", "SEK", "THB", "RM", "EUR", "CNY", "OTHER")
        
            currencies_F2 = Array("EUR_JPY", "EUR_GBP", "EUR_CHF", "EUR_CAD", "EUR_AUD", "EUR_SGD", "EUR_HKD", "EUR_CNY", "EUR_OTHER", _
            "GBP_JPY", "GBP_CHF", "GBP_CAD", "GBP_AUD", "GBP_SGD", "GBP_HKD", "GBP_CNY", "GBP_OTHER",  _
            "JPY_CHF", "JPY_CAD", "JPY_AUD", "JPY_SGD", "JPY_HKD", "JPY_CNY", "JPY_OTHER", _
            "CNY_AUD", "CNY_SGD", "CNY_HKD", "CNY_OTHER")
        
            ' F1交易類別 and F2交易類別 (Col)
            transactionTypes_F1 = Array("F1_與國外金融機構及非金融機構間交易_SPOT", _
                                        "F1_與國外金融機構及非金融機構間交易_SWAP", _
                                        "F1_與國內金融機構間交易_SPOT", _
                                        "F1_與國內金融機構間交易_SWAP", _
                                        "F1_與國內顧客間交易_SPOT")
        
            transactionTypes_F2 = Array("F2_與國外金融機構及非金融機構間交易_SPOT", _
                                        "F2_與國外金融機構及非金融機構間交易_SWAP", _
                                        "F2_與國內金融機構間交易_SPOT", _
                                        "F2_與國內金融機構間交易_SWAP")
            
            ' 每組交易對應的欄位(F1 and F2 OutputReport對應欄位)
            colLetters_F1 = Array("O", "Q", "I", "K", "B")
            colLetters_F2 = Array("O", "Q", "I", "K")
        
            ' 每組交易的起始儲存格列數
            startRows = Array(8, 8, 8, 8, 8)
        
            fieldList_F1 = GenerateFieldList(transactionTypes_F1, currencies_F1, colLetters_F1, startRows)
            fieldList_F2 = GenerateFieldList(transactionTypes_F2, currencies_F2, colLetters_F2, startRows)
        
            ' Add to Worksheet Fields for F1
            AddWorksheetFields "f1", fieldList_F1
            AddDynamicField "f1", "F1_申報時間", "A3", dataMonthStringROC_F1F2
        
            ' Add to Worksheet Fields for F2
            AddWorksheetFields "f2", fieldList_F2
            AddDynamicField "f2", "F2_申報時間", "A3", dataMonthStringROC_F1F2
    End Select
End Sub

'=== Private Method：Add Def for Worksheet Field === 
' fieldDefs is array of fields(each field(Array) of fields(Array)),
' for each Index's Form => (FieldName, CellAddress, InitialVAlue(null))
Private Sub AddWorksheetFields(ByVal wsName As String, _
                               ByVal fieldDefs As Variant)
    Dim wsDict As Object, dictValues As Object, dictAddresses As Object

    Dim i As Long, arrField As Variant

    Set dictValues = CreateObject("Scripting.Dictionary")
    Set dictAddresses = CreateObject("Scripting.Dictionary")
    
    For i = LBound(fieldDefs) To UBound(fieldDefs)
        arrField = fieldDefs(i)
        dictValues.Add arrField(0), arrField(2)
        dictAddresses.Add arrField(0), arrField(1)
    Next i
    
    Set wsDict = CreateObject("Scripting.Dictionary")
    wsDict.Add "Values", dictValues
    wsDict.Add "Addresses", dictAddresses
    
    clsWorksheets.Add wsName, wsDict
End Sub

Public Sub AddDynamicField(ByVal wsName As String, _
                           ByVal fieldName As String, _
                           ByVal cellAddress As String, _
                           ByVal initValue As Variant)
    Dim wsDict As Object
    Dim dictValues As Object, dictAddresses As Object
    
    ' 如果該工作表尚未建立，先建立一組新的 Dictionary
    If Not clsWorksheets.Exists(wsName) Then
        Set dictValues = CreateObject("Scripting.Dictionary")
        Set dictAddresses = CreateObject("Scripting.Dictionary")
        
        Set wsDict = CreateObject("Scripting.Dictionary")
        wsDict.Add "Values", dictValues
        wsDict.Add "Addresses", dictAddresses
        
        clsWorksheets.Add wsName, wsDict
    End If
    
    ' 取得該工作表的字典
    Set wsDict = clsWorksheets(wsName)
    Set dictValues = wsDict("Values")
    Set dictAddresses = wsDict("Addresses")
    
    ' 如果欄位已存在，可依需求選擇更新或忽略（此處以加入為例）
    If Not dictValues.Exists(fieldName) Then
        dictValues.Add fieldName, initValue
        dictAddresses.Add fieldName, cellAddress
    Else
        ' 若需要更新，直接賦值：
        dictValues(fieldName) = initValue
        dictAddresses(fieldName) = cellAddress
    End If
End Sub

'=== Set Field Value for one sheetName ===  
Public Sub SetField(ByVal wsName As String, _
                    ByVal fieldName As String, _
                    ByVal value As Variant)
    If Not clsWorksheets.Exists(wsName) Then
        Err.Raise 1002, , "工作表 [" & wsName & "] 尚未定義於報表 " & clsReportName
    End If
    Dim wsDict As Object
    Set wsDict = clsWorksheets(wsName)
    Dim dictValues As Object
    Set dictValues = wsDict("Values")
    If dictValues.Exists(fieldName) Then
        dictValues(fieldName) = value
    Else
        Err.Raise 1001, , "欄位 [" & fieldName & "] 不存在於工作表 [" & wsName & "] 的報表 " & clsReportName
    End If
End Sub

'=== With NO Parma: Get All Field Values ===  
'=== With wsName: Get Field Values within the worksheet Key 格式："wsName|fieldName" ===
Public Function GetAllFieldValues(Optional ByVal wsName As String = "") As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    Dim wsKey As Variant, dictV As Object, fieldKey As Variant
    If wsName <> "" Then
        If Not clsWorksheets.Exists(wsName) Then
            Err.Raise 1002, , "工作表 [" & wsName & "] 尚未定義於報表 " & clsReportName
        End If
        Set result = clsWorksheets(wsName)("Values")
    Else
        For Each wsKey In clsWorksheets.Keys
            Set dictV = clsWorksheets(wsKey)("Values")
            For Each fieldKey In dictV.Keys
                result.Add wsKey & "|" & fieldKey, dictV(fieldKey)
            Next fieldKey
        Next wsKey
    End If
    Set GetAllFieldValues = result
End Function

'=== With No Param: Get All Field Addresses ===  
'=== With wsName: Get Field Addresses within the worksheet Key 格式："wsName|fieldName" ===
Public Function GetAllFieldPositions(Optional ByVal wsName As String = "") As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    Dim wsKey As Variant, dictA As Object, fieldKey As Variant
    If wsName <> "" Then
        If Not clsWorksheets.Exists(wsName) Then
            Err.Raise 1002, , "工作表 [" & wsName & "] 尚未定義於報表 " & clsReportName
        End If
        Set result = clsWorksheets(wsName)("Addresses")
    Else
        For Each wsKey In clsWorksheets.Keys
            Set dictA = clsWorksheets(wsKey)("Addresses")
            For Each fieldKey In dictA.Keys
                result.Add wsKey & "|" & fieldKey, dictA(fieldKey)
            Next fieldKey
        Next wsKey
    End If
    Set GetAllFieldPositions = result
End Function

'=== 驗證是否每個欄位都有填入數值 (若指定 wsName 則驗證該工作表) ===  
Public Function ValidateFields(Optional ByVal wsName As String = "") As Boolean
    Dim msg As String, key As Variant
    msg = ""
    Dim dictValues As Object
    If wsName <> "" Then
        If Not clsWorksheets.Exists(wsName) Then
            Err.Raise 1002, , "工作表 [" & wsName & "] 尚未定義於報表 " & clsReportName
        End If
        Set dictValues = clsWorksheets(wsName)("Values")
        For Each key In dictValues.Keys
            If IsNull(dictValues(key)) Then msg = msg & wsName & " - " & key & vbCrLf
        Next key
    Else
        Dim wsKey As Variant
        For Each wsKey In clsWorksheets.Keys
            Set dictValues = clsWorksheets(wsKey)("Values")
            For Each key In dictValues.Keys
                If IsNull(dictValues(key)) Then msg = msg & wsKey & " - " & key & vbCrLf
            Next key
        Next wsKey
    End If
    If msg <> "" Then
        MsgBox "報表 [" & clsReportName & "] 以下欄位未填入有效資料:" & vbCrLf & msg, vbExclamation
        WriteLog "報表 [" & clsReportName & "] 以下欄位未填入有效資料:" & vbCrLf & msg
        ValidateFields = False
    Else
        ValidateFields = True
    End If
End Function

'=== 將 class 中的數值依據各工作表之欄位設定寫入指定的 Workbook ===  
' 此方法會針對 clsWorksheets 中定義的每個工作表名稱，嘗試在傳入的 Workbook 中找到對應工作表，並更新其欄位
Public Sub ApplyToWorkbook(ByRef wb As Workbook)
    Dim wsKey As Variant, wsDict As Object, dictValues As Object, dictAddresses As Object
    Dim ws As Worksheet, fieldKey As Variant
    For Each wsKey In clsWorksheets.Keys
        On Error Resume Next
        Set ws = wb.Sheets(wsKey)
        On Error GoTo 0
        If Not ws Is Nothing Then
            Set wsDict = clsWorksheets(wsKey)
            Set dictValues = wsDict("Values")
            Set dictAddresses = wsDict("Addresses")
            For Each fieldKey In dictValues.Keys
                If Not IsNull(dictValues(fieldKey)) Then
                    ws.Range(dictAddresses(fieldKey)).Value = dictValues(fieldKey)
                Else
                    MsgBox "工作表 [" & wsKey & "] 中找不到欄位: " & fieldKey , vbExclamation
                    WriteLog "工作表 [" & wsKey & "] 中找不到欄位: " & fieldKey
                End If
            Next fieldKey
        Else
            MsgBox "Workbook 中找不到工作表: " & wsKey, vbExclamation
            WriteLog "Workbook 中找不到工作表: " & wsKey
            Exit Sub
        End If
        Set ws = Nothing
    Next wsKey
End Sub

'=== 報表名稱屬性 ===  
Public Property Get ReportName() As String
    ReportName = clsReportName
End Property


這邊是我的Utility
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
        MsgBox QueryName & "查詢結果為空，請檢查資料庫與查詢條件。", vbExclamation
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

以下是我的主程式


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
Public Sub Process_F1_F2()
    '=== Equal Setting ===
    'Fetch Query Access DB table
    ' F1
    Dim dataArr_1 As Variant
    Dim dataArr_2 As Variant
    Dim dataArr_3 As Variant
    Dim dataArr_4 As Variant
    Dim dataArr_5 As Variant
    Dim dataArr_6 As Variant
    ' F2
    Dim dataArr_7 As Variant
    Dim dataArr_8 As Variant
    Dim dataArr_9 As Variant
    Dim dataArr_10 As Variant

    'Declare worksheet and handle data
    Dim xlsht As Worksheet

    Dim i As Integer, j As Integer
    Dim lastRow As Long

    Dim reportTitle As String
    ' F1
    Dim queryTable_1 As String
    Dim queryTable_2 As String
    Dim queryTable_3 As String
    Dim queryTable_4 As String
    Dim queryTable_5 As String
    Dim queryTable_6 As String
    ' F2
    Dim queryTable_7 As String
    Dim queryTable_8 As String
    Dim queryTable_9 As String
    Dim queryTable_10 As String

    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("F1_F2")

    reportTitle = "F1_F2"
    ' F1
    queryTable_1 = "F1_Foreign_DL6850_FS"
    queryTable_2 = "F1_Foreign_DL6850_SS"
    queryTable_3 = "F1_Domestic_DL6850_FS"
    queryTable_4 = "F1_Domestic_DL6850_SS"
    queryTable_5 = "F1_CM2810_LIST"
    queryTable_6 = "F1_CM2810_Subtotal"
    ' F2
    queryTable_7 = "F2_Foreign_DL6850_FS"
    queryTable_8 = "F2_Foreign_DL6850_SS"
    queryTable_9 = "F2_Domestic_DL6850_FS"
    queryTable_10 = "F2_Domestic_DL6850_SS"

    ' dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1, gDataMonthString)
    ' dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2, gDataMonthString)
    ' dataArr_3 = GetAccessDataAsArray(gDBPath, queryTable_3, gDataMonthString)
    ' dataArr_4 = GetAccessDataAsArray(gDBPath, queryTable_4, gDataMonthString)
    ' dataArr_5 = GetAccessDataAsArray(gDBPath, queryTable_5, gDataMonthString)
    ' dataArr_5 = GetAccessDataAsArray(gDBPath, queryTable_5, gDataMonthString)
    ' F1
    dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1, gDataMonthString)
    dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2, gDataMonthString)
    dataArr_3 = GetAccessDataAsArray(gDBPath, queryTable_3, gDataMonthString)
    dataArr_4 = GetAccessDataAsArray(gDBPath, queryTable_4, gDataMonthString)
    dataArr_5 = GetAccessDataAsArray(gDBPath, queryTable_5, gDataMonthString)
    dataArr_6 = GetAccessDataAsArray(gDBPath, queryTable_6, gDataMonthString)
    ' F2
    dataArr_7 = GetAccessDataAsArray(gDBPath, queryTable_7, gDataMonthString)
    dataArr_8 = GetAccessDataAsArray(gDBPath, queryTable_8, gDataMonthString)
    dataArr_9 = GetAccessDataAsArray(gDBPath, queryTable_9, gDataMonthString)
    dataArr_10 = GetAccessDataAsArray(gDBPath, queryTable_10, gDataMonthString)

    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    xlsht.Range("A:AA").ClearContents
    ' F1
    xlsht.Range("AC2:AC100").ClearContents
    xlsht.Range("AG2:AG100").ClearContents
    xlsht.Range("AK2:AK100").ClearContents
    xlsht.Range("AO2:AO100").ClearContents
    xlsht.Range("AS2:AS100").ClearContents
    ' F2
    xlsht.Range("AW2:AW100").ClearContents
    xlsht.Range("BA2:BA100").ClearContents
    xlsht.Range("BE2:BE100").ClearContents
    xlsht.Range("BI2:BI100").ClearContents

    
    '=== Paste Queyr Table into Excel ===
    ' F1
    If Err.Number <> 0 Or LBound(dataArr_1) > UBound(dataArr_1) Then
        MsgBox "資料有誤: " & reportTitle & ": " & queryTable_1 & "資料表無資料"
        WriteLog "資料有誤: " & reportTitle & ": " & queryTable_1 & "資料表無資料"
        Exit Sub
    Else
        For j = 0 To UBound(dataArr_1, 2)
            For i = 0 To UBound(dataArr_1, 1)
                xlsht.Cells(i + 1, j + 1).Value = dataArr_1(i, j)
            Next i
        Next j
    End If

    If Err.Number <> 0 Or LBound(dataArr_2) > UBound(dataArr_2) Then
        MsgBox "資料有誤: " & reportTitle & ": " & queryTable_2 & "資料表無資料"
        WriteLog "資料有誤: " & reportTitle & ": " & queryTable_2 & "資料表無資料"
        Exit Sub
    Else
        For j = 0 To UBound(dataArr_2, 2)
            For i = 0 To UBound(dataArr_2, 1)
                xlsht.Cells(i + 1, j + 3).Value = dataArr_2(i, j)
            Next i
        Next j
    End If

    If Err.Number <> 0 Or LBound(dataArr_3) > UBound(dataArr_3) Then
        MsgBox "資料有誤: " & reportTitle & ": " & queryTable_3 & "資料表無資料"
        WriteLog "資料有誤: " & reportTitle & ": " & queryTable_3 & "資料表無資料"
        Exit Sub
    Else
        For j = 0 To UBound(dataArr_3, 2)
            For i = 0 To UBound(dataArr_3, 1)
                xlsht.Cells(i + 1, j + 5).Value = dataArr_3(i, j)
            Next i
        Next j
    End If

    If Err.Number <> 0 Or LBound(dataArr_4) > UBound(dataArr_4) Then
        MsgBox reportTitle & ": " & queryTable_4 & "資料表無資料"
        WriteLog reportTitle & ": " & queryTable_4 & "資料表無資料"
    Else
        For j = 0 To UBound(dataArr_4, 2)
            For i = 0 To UBound(dataArr_4, 1)
                xlsht.Cells(i + 1, j + 7).Value = dataArr_4(i, j)
            Next i
        Next j
    End If

    If Err.Number <> 0 Or LBound(dataArr_5) > UBound(dataArr_5) Then
        MsgBox "資料有誤: " & reportTitle & ": " & queryTable_5 & "資料表無資料"
        WriteLog "資料有誤: " & reportTitle & ": " & queryTable_5 & "資料表無資料"
        Exit Sub
    Else
        For j = 0 To UBound(dataArr_5, 2)
            For i = 0 To UBound(dataArr_5, 1)
                xlsht.Cells(i + 1, j + 9).Value = dataArr_5(i, j)
            Next i
        Next j
    End If

    If Err.Number <> 0 Or LBound(dataArr_6) > UBound(dataArr_6) Then
        MsgBox "資料有誤: " & reportTitle & ": " & queryTable_6 & "資料表無資料"
        WriteLog "資料有誤: " & reportTitle & ": " & queryTable_6 & "資料表無資料"
        Exit Sub
    Else
        For j = 0 To UBound(dataArr_6, 2)
            For i = 0 To UBound(dataArr_6, 1)
                xlsht.Cells(i + 1, j + 17).Value = dataArr_6(i, j)
            Next i
        Next j
    End If

    ' F2
    If Err.Number <> 0 Or LBound(dataArr_7) > UBound(dataArr_7) Then
        MsgBox "資料有誤: " & reportTitle & ": " & queryTable_7 & "資料表無資料"
        WriteLog "資料有誤: " & reportTitle & ": " & queryTable_7 & "資料表無資料"
        Exit Sub
    Else
        For j = 0 To UBound(dataArr_7, 2)
            For i = 0 To UBound(dataArr_7, 1)
                xlsht.Cells(i + 1, j + 20).Value = dataArr_7(i, j)
            Next i
        Next j
    End If

    If Err.Number <> 0 Or LBound(dataArr_8) > UBound(dataArr_8) Then
        MsgBox "資料有誤: " & reportTitle & ": " & queryTable_8 & "資料表無資料"
        WriteLog "資料有誤: " & reportTitle & ": " & queryTable_8 & "資料表無資料"
        Exit Sub
    Else
        For j = 0 To UBound(dataArr_8, 2)
            For i = 0 To UBound(dataArr_8, 1)
                xlsht.Cells(i + 1, j + 22).Value = dataArr_8(i, j)
            Next i
        Next j
    End If

    If Err.Number <> 0 Or LBound(dataArr_9) > UBound(dataArr_9) Then
        MsgBox "資料有誤: " & reportTitle & ": " & queryTable_9 & "資料表無資料"
        WriteLog "資料有誤: " & reportTitle & ": " & queryTable_9 & "資料表無資料"
        Exit Sub
    Else
        For j = 0 To UBound(dataArr_9, 2)
            For i = 0 To UBound(dataArr_9, 1)
                xlsht.Cells(i + 1, j + 24).Value = dataArr_9(i, j)
            Next i
        Next j
    End If

    If Err.Number <> 0 Or LBound(dataArr_10) > UBound(dataArr_10) Then
        MsgBox "資料有誤: " & reportTitle & ": " & queryTable_10 & "資料表無資料"
        WriteLog "資料有誤: " & reportTitle & ": " & queryTable_10 & "資料表無資料"
        Exit Sub
    Else
        For j = 0 To UBound(dataArr_10, 2)
            For i = 0 To UBound(dataArr_10, 1)
                xlsht.Cells(i + 1, j + 26).Value = dataArr_10(i, j)
            Next i
        Next j
    End If

    Dim currencies_F1 As Variant
    Dim currencies_F2 As Variant

    currencies_F1 = Array("JPY", "GBP", "CHF", "CAD", "AUD", "NZD", "SGD", "HKD", "ZAR", "SEK", "THB", "RM", "EUR", "CNY", "OTHER")

    currencies_F2 = Array("EUR_JPY", "EUR_GBP", "EUR_CHF", "EUR_CAD", "EUR_AUD", "EUR_SGD", "EUR_HKD", "EUR_CNY", "EUR_OTHER", _
    "GBP_JPY", "GBP_CHF", "GBP_CAD", "GBP_AUD", "GBP_SGD", "GBP_HKD", "GBP_CNY", "GBP_OTHER",  _
    "JPY_CHF", "JPY_CAD", "JPY_AUD", "JPY_SGD", "JPY_HKD", "JPY_CNY", "JPY_OTHER", _
    "CNY_AUD", "CNY_SGD", "CNY_HKD", "CNY_OTHER")
    
    ' 定義交易名稱，對應到不同資料表
    Dim transactionTypes_F1 As Variant
    Dim transactionTypes_F2 As Variant

    transactionTypes_F1 = Array("F1_與國外金融機構及非金融機構間交易_SPOT", _
                            "F1_與國外金融機構及非金融機構間交易_SWAP", _
                            "F1_與國內金融機構間交易_SPOT", _
                            "F1_與國內金融機構間交易_SWAP", _
                            "F1_與國內顧客間交易_SPOT")

    transactionTypes_F2 = Array("F2_與國外金融機構及非金融機構間交易_SPOT", _
                            "F2_與國外金融機構及非金融機構間交易_SWAP", _
                            "F2_與國內金融機構間交易_SPOT", _
                            "F2_與國內金融機構間交易_SWAP")
    
    ' 對應每個交易類型在 Excel 中的欄位範圍
    Dim dataRanges_F1 As Variant
    Dim dataRanges_F2 As Variant
    ' Cur 在前一欄, Value 在後一欄
    dataRanges_F1 = Array("A:B", "C:D", "E:F", "G:H", "Q:R") 
    dataRanges_F2 = Array("T:U", "V:W", "X:Y", "Z:AA")
    
    Dim curDict As Object
    Dim currCol As Integer
    For i = LBound(transactionTypes_F1) To UBound(transactionTypes_F1)
        ' 建立字典儲存貨幣數值，並初始化為 0
        Set curDict = CreateObject("Scripting.Dictionary")
        For j = LBound(currencies_F1) To UBound(currencies_F1)
            curDict.Add currencies_F1(j), 0
        Next j
        
        ' 確定當前交易的資料範圍
        ' 取得起始欄位（Cur欄）
        currCol = xlsht.Range(dataRanges_F1(i)).Column
        lastRow = xlsht.Cells(xlsht.Rows.Count, currCol).End(xlUp).Row

        For j = 2 To lastRow ' 假設第1列是標題，從第2列開始
            Dim curCode As String, curValue As Variant
            ' 貨幣名稱
            curCode = xlsht.Cells(j, currCol).Value 
            ' 貨幣數值 百萬元，四捨五入小數第一位
            curValue = Round(xlsht.Cells(j, currCol + 1).Value / 1000000, 1) 
            
            ' 確保 Value 為數字，且 Cur 是已定義的貨幣
            If IsNumeric(curValue) And curDict.Exists(curCode) Then
                ' 若累加改成 curDict(curCode) = curDict(curCode) + curValue
                curDict(curCode) = curValue 
            End If
        Next j
        
        ' 依照固定貨幣順序填入 Excel 和報表
        For j = LBound(currencies_F1) To UBound(currencies_F1)
            Dim fieldName As String, valueNum As Variant
            ' 產生field名稱
            fieldName = transactionTypes_F1(i) & "_" & currencies_F1(j) 
            valueNum = curDict(currencies_F1(j))
        
            ' 設定 Excel 的 Range 值
            xlsht.Range(fieldName).Value = valueNum
            
            ' 設定報表欄位
            rpt.SetField "f1", fieldName, CStr(valueNum)
        Next j
    Next i


    For i = LBound(transactionTypes_F2) To UBound(transactionTypes_F2)
        ' 建立字典儲存貨幣數值，並初始化為 0
        Set curDict = CreateObject("Scripting.Dictionary")
        For j = LBound(currencies_F2) To UBound(currencies_F2)
            curDict.Add currencies_F2(j), 0
        Next j
        
        ' 確定當前交易的資料範圍
        ' 取得起始欄位（Cur欄）
        currCol = xlsht.Range(dataRanges_F2(i)).Column
        lastRow = xlsht.Cells(xlsht.Rows.Count, currCol).End(xlUp).Row

        For j = 2 To lastRow ' 假設第1列是標題，從第2列開始
            ' 貨幣名稱
            curCode = xlsht.Cells(j, currCol).Value 
            ' 貨幣數值 百萬元，四捨五入小數第一位
            curValue = Round(xlsht.Cells(j, currCol + 1).Value / 1000000, 1) 
            
            ' 確保 Value 為數字，且 Cur 是已定義的貨幣
            If IsNumeric(curValue) And curDict.Exists(curCode) Then
                ' 若累加改成 curDict(curCode) = curDict(curCode) + curValue
                curDict(curCode) = curValue 
            End If
        Next j
        
        ' 依照固定貨幣順序填入 Excel 和報表
        For j = LBound(currencies_F2) To UBound(currencies_F2)
            ' 產生field名稱
            fieldName = transactionTypes_F2(i) & "_" & currencies_F2(j) 
            valueNum = curDict(currencies_F2(j))
        
            ' 設定 Excel 的 Range 值
            xlsht.Range(fieldName).Value = valueNum
            
            ' 設定報表欄位
            rpt.SetField "f2", fieldName, CStr(valueNum)
        Next j
    Next i
    
    ' F1
    xlsht.Range("AC2:AC100").NumberFormat = "#,##,##.0"
    xlsht.Range("AG2:AG100").NumberFormat = "#,##,##.0"
    xlsht.Range("AK2:AK100").NumberFormat = "#,##,##.0"
    xlsht.Range("AO2:AO100").NumberFormat = "#,##,##.0"
    xlsht.Range("AS2:AS100").NumberFormat = "#,##,##.0"
    
    ' F2
    xlsht.Range("AW2:AW100").NumberFormat = "#,##,##.0"
    xlsht.Range("BA2:BA100").NumberFormat = "#,##,##.0"
    xlsht.Range("BE2:BE100").NumberFormat = "#,##,##.0"
    xlsht.Range("BI2:BI100").NumberFormat = "#,##,##.0"

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

' Process C 更新原始申報檔案欄位數值及另存新檔
Public Sub UpdateExcelReports()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Dim rpt As clsReport
    Dim rptName As Variant
    Dim wb As Workbook
    Dim emptyFilePath As String, outputFilePath As String
    For Each rptName In gReportNames
        Set rpt = gReports(rptName)
        ' 開啟原始 Excel 檔（檔名以報表名稱命名）
        emptyFilePath = gReportFolder & "\" & ThisWorkbook.Sheets("ControlPanel").Range(rptName & "_FileName").Value & "." & ThisWorkbook.Sheets("ControlPanel").Range(rptName & "_ExtensionName").Value
        outputFilePath = gOutputFolder & "\" & ThisWorkbook.Sheets("ControlPanel").Range(rptName & "_FileName").Value & Replace(gDataMonthString, "/", "") & "." & ThisWorkbook.Sheets("ControlPanel").Range(rptName & "_ExtensionName").Value

        Set wb = Workbooks.Open(emptyFilePath)
        If wb Is Nothing Then
            MsgBox "無法開啟檔案: " & emptyFilePath, vbExclamation
            WriteLog "無法開啟檔案: " & emptyFilePath
            GoTo CleanUp
            ' Eixt Sub
        End If
        ' 報表內有多個工作表，呼叫 ApplyToWorkbook 讓 clsReport 自行依各工作表更新
        rpt.ApplyToWorkbook wb
        wb.SaveAs Filename:=outputFilePath
        wb.Close SaveChanges:=False
        Set wb = Nothing   ' Release Workbook Object
    Next rptName
    MsgBox "完成申報報表更新"
    WriteLog "完成申報報表更新"

CleanUp:
    ' 還原警示
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True    
End Sub

但是在執行 ApplyToWorkbook 時出現很多下面錯誤，我只摘錄每個類別的第一個，請告訴我是出了什麼問題，照理說我已經在初始化過程中都先建好所有欄位，不應該會有
2025-04-18 16:43:34 - 工作表 [f1] 中找不到欄位: F1_與國外金融機構及非金融機構間交易_SPOT_JPY
....
2025-04-18 16:43:49 - 工作表 [f1] 中找不到欄位: F1_與國外金融機構及非金融機構間交易_SWAP_JPY
....
2025-04-18 16:43:55 - 工作表 [f1] 中找不到欄位: F1_與國內金融機構間交易_SPOT_JPY
....
2025-04-18 16:44:01 - 工作表 [f1] 中找不到欄位: F1_與國內金融機構間交易_SWAP_JPY
....
2025-04-18 16:44:06 - 工作表 [f1] 中找不到欄位: F1_與國內顧客間交易_SPOT_JPY
....
2025-04-18 16:44:14 - 工作表 [f2] 中找不到欄位: F2_與國外金融機構及非金融機構間交易_SPOT_EUR_JPY
....
2025-04-18 16:44:21 - 工作表 [f2] 中找不到欄位: F2_與國外金融機構及非金融機構間交易_SWAP_EUR_JPY
....
2025-04-18 16:44:28 - 工作表 [f2] 中找不到欄位: F2_與國內金融機構間交易_SPOT_EUR_JPY
....
2025-04-18 16:44:35 - 工作表 [f2] 中找不到欄位: F2_與國內金融機構間交易_SWAP_EUR_JPY
....
