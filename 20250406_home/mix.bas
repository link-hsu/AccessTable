' Utility.bas

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
        MsgBox "查詢結果為空，請檢查資料庫與查詢條件。", vbExclamation
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
                           ByVal fieldValue As Double, _
                           ByVal fieldAddress As String)
    Dim conn As Object, cmd As Object
    Dim sql As String
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath
    sql = "INSERT INTO " & tableName & " (DataMonthString, ReportName, WorksheetName_FieldKey, FieldValue, FieldAddress, CaseCreatedAt) " & _
          "VALUES ('" & dataMonthString & "', '" & reportName & "', '" & WorksheetName_FieldKey & "', " & fieldValue & ", '" & fieldAddress & "', Now());"
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
                        ByVal fieldValue As Double)
    Dim conn As Object, cmd As Object
    Dim sql As String
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath
    
    sql = "UPDATE MonthlyDeclarationReport SET FieldValue = " & fieldValue & ", CaseUpdatedAt = Now() " & _
          "WHERE DataMonthString = '" & dataMonthString & "' " & _
          "AND ReportName = '" & reportName & "' " & _
          "AND WorksheetName_FieldKey = '" & worksheetName_fieldKey & "' " & _
          "AND FieldAddress = '" & fieldAddress & "';"
    
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = conn
    cmd.CommandText = sql
    cmd.Execute
    
    conn.Close
    Set conn = Nothing
    Set cmd = Nothing
End Sub


' clsReport.bas

Option Explicit

' Report Title
Private clsReportName As As String

' Dictionary：key = Worksheet Name，value = Dictionary( Keys "Fiedl Values" 與 "Field Addresses" )
Private clsWorksheets As Object

'=== 初始化報表 (根據報表名稱建立各工作表的欄位定義) ===
Public Sub Init(ByVal reportName As String)
    clsReportName = reportName
    Set clsWorksheets = CreateObject("Scripting.Dictionary")
    
    Select Case reportName
        Case "CNY1"
            ' 假設 CNY1 報表有三個工作表：X、Y、Z  
            ' 工作表 X 定義：  
            '   - "其他金融資產_淨額" 儲存格地址 "B2"  
            '   - "其他" 儲存格地址 "C2"  
            '   - "CNY1_資產總計" 儲存格地址 "D2"
            AddWorksheetFields "CNY1", Array( _
                Array("CNY1_其他金融資產_淨額", "G98", Null), _
                Array("CNY1_其他", "G100", Null), _
                Array("CNY1_資產總計", "G116", Null), _
                Array("CNY1_其他金融負債", "G170", Null), _
                Array("CNY1_其他什項金融負債", "G172", Null), _
                Array("CNY1_負債總計", "G184", Null) )



            ' ' 工作表 Y 定義：  
            ' '   - "其他金融負債" 儲存格地址 "E2"  
            ' '   - "其他什項金融負債" 儲存格地址 "F2"
            ' AddWorksheetFields "Y", Array( _
            '     Array("其他金融負債", "E2", Null), _
            '     Array("其他什項金融負債", "F2", Null) )
            ' ' 工作表 Z 定義：  
            ' '   - "CNY1_負債總計" 儲存格地址 "G2"
            ' AddWorksheetFields "Z", Array( _
            '     Array("CNY1_負債總計", "G2", Null) )
        Case "FB2"
            AddWorksheetFields "FOA", Array( _
                Array("FB2_存放及拆借同業", "F9", Null), _
                Array("FB2_拆放銀行同業", "F13", Null), _
                Array("FB2_應收款項_淨額", "F36", Null), _
                Array("FB2_應收利息", "F41", Null), _
                Array("FB2_資產總計", "F85", Null) )
        Case "FB3"
            AddWorksheetFields "FOA", Array( _
                Array("FB3_存放及拆借同業_資產面_台灣地區", "D9", Null), _
                Array("FB3_同業存款及拆放_負債面_台灣地區", "D10", Null) )
        Case "FB3A"
            AddWorksheetFields "Sheet1", Array( _
                Array("總收入", "B5", Null), _
                Array("總支出", "C5", Null) )
        Case "FM5"
            ' No Data
        Case "FM11"
            AddWorksheetFields "FOA", Array( _
                Array("FM11_一利息股息收入_利息", "E15", Null), _
                Array("FM11_五證券投資評價及減損損失_一年期以上之債權證券", "I25", Null), _
                Array("FM11_一利息收入_自中華民國境內其他客戶", "E36", Null) )
        Case "FM13"
            AddWorksheetFields "FOA", Array( _
                Array("FM13_OBU_香港_債票券投資", "D9", Null), _
                Array("FM13_OBU_韓國_債票券投資", "F9", Null), _
                Array("FM13_OBU_泰國_債票券投資", "H9", Null), _
                Array("FM13_OBU_馬來西亞_債票券投資", "J9", Null), _
                Array("FM13_OBU_菲律賓_債票券投資", "L9", Null), _
                Array("FM13_OBU_印尼_債票券投資", "N9", Null), _
                Array("FM13_OBU_債票券投資_評價調整", "T9", Null), _
                Array("FM13_OBU_債票券投資_累計減損", "U9", Null) )
        Case "AI821"
            AddWorksheetFields "Table1", Array( _
                Array("AI821_本國銀行", "D61", Null), _
                Array("AI821_陸銀在臺分行", "D62", Null), _
                Array("AI821_外商銀行在臺分行", "D63", Null), _
                Array("AI821_大陸地區銀行", "D64", Null), _
                Array("AI821_其他", "D65", Null) )
        Case "Table2"
            AddWorksheetFields "FOA", Array( _
                Array("Table2_其他", "D17", Null), _
                Array("Table_美元_F1", "L7", Null), _
                Array("Table2_美元_F3", "N7", Null), _
                Array("Table2_美元_F4", "O7", Null) )
        Case "FB5_FB5A"
            AddWorksheetFields "FOA", Array( _
                Array("FB5_外匯交易_即期外匯_DBU", "G11", Null) )
        Case "FM2"
            '這邊要動態ADDWORKSHEETfIELDS
            AddWorksheetFields "FOA", Array( _
                Array("總收入", "B5", Null), _
                Array("總支出", "C5", Null) )
        Case "FM10"
            AddWorksheetFields "FOA", Array( _
                Array("FM10_FVOCI_總額C", "F20", Null), _
                Array("FM10_FVOCI_淨額D", "G20", Null), _
                Array("FM10_AC_總額E", "H20", Null), _
                Array("FM10_AC_淨額F", "I20", Null) )
        Case "F1_F2"
            Dim currencies As Variant, transactions As Variant, startRows As Variant, colLetters As Variant
            Dim i As Integer, j As Integer
            currencies = Array("日圓", "英鎊", "瑞士法郎", "加拿大幣", "澳幣", "紐西蘭幣", "新加坡幣", "港幣", "南非幣", "瑞典幣", "泰幣", "馬來幣", "歐元", "人民幣")
            transactions = Array("與國內顧客間交易_即期", "與國內金融機構間交易_即期", "與國內金融機構間交易_換匯", "與國外金融機構及非金融機構間交易_即期", "與國外金融機構及非金融機構間交易_換匯")
            startRows = Array(5, 5, 5, 5, 5) ' 每組交易的起始儲存格列數
            colLetters = Array("B", "C", "D", "H", "J") ' 每組交易對應的欄位

            Dim fieldList() As Variant
            Dim index As Integer
            index = 0
            ReDim fieldList(UBound(transactions) * UBound(currencies))

            For i = LBound(transactions) To UBound(transactions)
                For j = LBound(currencies) To UBound(currencies)
                    fieldList(index) = Array(transactions(i) & "_" & currencies(j), colLetters(i) & (startRows(i) + j), Null)
                    index = index + 1
                Next j
            Next i

            ' 加入到 Worksheet Fields
            AddWorksheetFields "Sheet1", fieldList
        Case "Table41"
            AddWorksheetFields "FOA", Array( _
                Array("Table41_四衍生工具處分利益", "D25", Null), _
                Array("Table41_四衍生工具處分損失", "G25", Null) )
        Case "AI602"
            AddWorksheetFields "Sheet1", Array( _
                Array("總收入", "B5", Null), _
                Array("總支出", "C5", Null) )
        Case "AI240"
            AddWorksheetFields "Sheet1", Array( _
                Array("總收入", "B5", Null), _
                Array("總支出", "C5", Null) )
        ' 如有其他報表，依需求加入不同工作表及欄位定義
    End Select
End Sub


'=== Private Method：Add Def for Worksheet Field === 
' fieldDefs is array of fields(each field(Array) of fields(Array)), for each Index's Form => (FieldName, CellAddress, InitialVAlue(null))
Private Sub AddWorksheetFields(ByVal wsName As String,
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

'=== Set Field Value for one sheetName ===  
Public Sub SetField(ByVal wsName As String,
                    ByVal fieldName As String,
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
                End If
            Next fieldKey
        Else
            MsgBox "Workbook 中找不到工作表: " & wsKey, vbExclamation
        End If
        Set ws = Nothing
    Next wsKey
End Sub

'=== 報表名稱屬性 ===  
Public Property Get ReportName() As String
    ReportName = clsReportName
End Property


' Module.bas

Option Explicit

'=== Global Config Settings ===
Public gDataMonthString As String    ' 使用者輸入的資料月份
Public gDBPath As String               ' 資料庫路徑
Public gReportFolder As String         ' 原始申報報表 Excel 檔所在資料夾
Public gOutputFolder As String         ' 更新後另存新檔的資料夾
Public gReportNames As Variant         ' 報表名稱陣列
Public gReports As Collection          ' 存放所有報表 (clsReport) 的 Collection

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
        Else
            MsgBox "格式錯誤，請輸入正確格式 (yyyy/mm)", vbExclamation, "格式錯誤"
        End If
    Loop Until isInputValid
    
    ' 設定其他 config 參數（請根據實際環境調整）
    gDBPath = ThisWorkbook.Path & "\" & Sheets("ControlPanel").Range("DBsPathFileName").Value
    gReportFolder = "C:\申報報表\原始檔\"      ' 調整為實際路徑
    gOutputFolder = "C:\申報報表\Processed\"      ' 調整為實際路徑
    gReportNames = Array("CNY1", "MM4901B", "AC5601", "AC5602")
    
    ' (a) 先初始化所有報表，並將初始資料寫入 Access（例如寫入 ReportConfig 資料表）
    Call InitializeReports
    ' (b) 各報表分別進行資料處理（各自邏輯分離）
    Call ProcessCNY1
    Call ProcessMM4901B
    Call ProcessAC5601
    Call ProcessAC5602
    ' (c) 最後更新申報 Excel 檔案並另存新檔
    Call UpdateExcelReports
    MsgBox "完成全部流程處理"
End Sub

'=== (a) 初始化所有報表並將初始資料寫入 Access ===
Public Sub InitializeReports()
    Dim rpt As clsReport
    Dim rptName As Variant, key As Variant
    Set gReports = New Collection
    For Each rptName In gReportNames
        Set rpt = New clsReport
        rpt.Init rptName
        gReports.Add rpt, rptName
        ' 將各工作表內每個欄位初始設定寫入 Access（資料表名稱例如 ReportConfig）
        Dim wsPositions As Object
        Dim combinedPositions As Object
        Set combinedPositions = rpt.GetAllFieldPositions ' 合併所有工作表，Key 格式 "wsName|fieldName"
        For Each key In combinedPositions.Keys
            InsertIntoTable gDBPath, "MonthlyDeclarationReport", gDataMonthString, rptName, key, 0, combinedPositions(key)
        Next key
    Next rptName
    MsgBox "報表初始化及初始資料建立完成"
End Sub

'=== (b) 各報表獨立處理邏輯 ===

'【CNY1】資料處理（示範：假設 Query 資料分別對應於工作表 X、Y、Z 的欄位）
Public Sub ProcessCNY1()
    Dim rpt As clsReport
    Dim dataArr As Variant, tempValue As Double
    Set rpt = gReports("CNY1")
    ' 取得 Query 資料 (查詢名稱 "CNY1_Query")
    dataArr = GetAccessDataAsArray(gDBPath, "CNY1_Query", gDataMonthString)
    If UBound(dataArr) < 1 Then
        MsgBox "CNY1 查詢資料不完整！", vbExclamation
        Exit Sub
    End If
    ' 假設 dataArr(1,0) 為工作表 X 的 "其他金融資產_淨額"、"其他"、"CNY1_資產總計" 的共同值：
    tempValue = CDbl(dataArr(1, 0))
    rpt.SetField "X", "其他金融資產_淨額", tempValue
    rpt.SetField "X", "其他", tempValue
    rpt.SetField "X", "CNY1_資產總計", tempValue
    ' 假設 dataArr(1,1) 為工作表 Y 的 "其他金融負債" 與 "其他什項金融負債" 的值：
    tempValue = CDbl(dataArr(1, 1))
    rpt.SetField "Y", "其他金融負債", tempValue
    rpt.SetField "Y", "其他什項金融負債", tempValue
    ' 假設 dataArr(1,2) 為工作表 Z 的 "CNY1_負債總計" 的值：
    tempValue = CDbl(dataArr(1, 2))
    rpt.SetField "Z", "CNY1_負債總計", tempValue
    ' 驗證並更新資料庫（以 Insert 示範，若記錄已存在則應用 UPDATE）
    If rpt.ValidateFields() Then
        Dim key As Variant
        Dim allValues As Object, allPositions As Object

        Set allValues = rpt.GetAllFieldValues()  ' key 格式 "wsName|fieldName"
        Set allPositions = rpt.GetAllFieldPositions()

        For Each key In allValues.Keys
            UpdateRecord gDBPath, gDataMonthString, rpt.ReportName, key, allPositions(key), CDbl(allValues(key))
        Next key
    End If
End Sub

'【MM4901B】資料處理（示範數值代入）
Public Sub ProcessMM4901B()
    Dim rpt As clsReport
    Set rpt = gReports("MM4901B")
    Dim dataArr As Variant
    dataArr = GetAccessDataAsArray(gDBPath, "MM4901B_Query", gDataMonthString)
    ' 假設 MM4901B 報表只有一個工作表 "Sheet1"
    rpt.SetField "Sheet1", "短期負債", 100
    rpt.SetField "Sheet1", "長期負債", 200
    Dim key As Variant
    If rpt.ValidateFields("Sheet1") Then
        Dim allValues As Object
        Set allValues = rpt.GetAllFieldValues("Sheet1")
        For Each key In allValues.Keys
            UpdateRecord gDBPath, gDataMonthString, rpt.ReportName, key, _
                CDbl(allValues(key))
        Next key
    End If
End Sub

'【AC5601】資料處理（示範）
Public Sub ProcessAC5601()
    Dim rpt As clsReport
    Set rpt = gReports("AC5601")
    Dim dataArr As Variant
    dataArr = GetAccessDataAsArray(gDBPath, "AC5601_Query", gDataMonthString)
    ' 示範：依照實際運算邏輯設定
    rpt.SetField "Sheet1", "資產總計", 300
    rpt.SetField "Sheet1", "負債總計", 150
    Dim key As Variant
    If rpt.ValidateFields("Sheet1") Then
        Dim allValues As Object
        Set allValues = rpt.GetAllFieldValues("Sheet1")
        For Each key In allValues.Keys
            UpdateRecord gDBPath, gDataMonthString, rpt.ReportName, key, _
                CDbl(allValues(key))
        Next key
    End If
End Sub

'【AC5602】資料處理（示範）
Public Sub ProcessAC5602()
    Dim rpt As clsReport
    Set rpt = gReports("AC5602")
    Dim dataArr As Variant
    dataArr = GetAccessDataAsArray(gDBPath, "AC5602_Query", gDataMonthString)
    rpt.SetField "Sheet1", "總收入", 500
    rpt.SetField "Sheet1", "總支出", 400
    Dim key As Variant
    If rpt.ValidateFields("Sheet1") Then
        Dim allValues As Object
        Set allValues = rpt.GetAllFieldValues("Sheet1")
        For Each key In allValues.Keys
            UpdateRecord gDBPath, gDataMonthString, rpt.ReportName, key, _
                CDbl(allValues(key))
        Next key
    End If
End Sub

'=== (c) 更新申報 Excel 檔案，將各報表物件數值寫入對應儲存格（各工作表），並另存新檔 ===
Public Sub UpdateExcelReports()
    Dim rpt As clsReport
    Dim rptName As Variant
    Dim wb As Workbook
    Dim reportFilePath As String, outputFilePath As String
    For Each rptName In gReportNames
        Set rpt = gReports(rptName)
        ' 開啟原始 Excel 檔（檔名以報表名稱命名）
        reportFilePath = gReportFolder & rptName & ".xlsx"
        Set wb = Workbooks.Open(reportFilePath)
        If wb Is Nothing Then
            MsgBox "無法開啟檔案: " & reportFilePath, vbExclamation
        End If
        ' 由於報表內可能有多個工作表，呼叫 ApplyToWorkbook 讓 clsReport 自行依各工作表更新
        rpt.ApplyToWorkbook wb
        outputFilePath = gOutputFolder & rptName & "_Processed.xlsx"
        wb.SaveAs Filename:=outputFilePath
        wb.Close SaveChanges:=False
        Set wb = Nothing   ' Release Workbook Object
    Next rptName
    MsgBox "所有 Excel 申報報表已更新並另存！"
End Sub