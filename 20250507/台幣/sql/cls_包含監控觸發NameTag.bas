Q:
' ===============
這是我的classReport
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
        Case "CNY1"
            AddWorksheetFields "CNY1", Array( _
                Array("CNY1_申報時間", "C2", dataMonthStringROC), _
                Array("CNY1_負債總計", "G184", Null) )
        Case "FB1"
            'No Data
            AddWorksheetFields "FOA", Array( _
                Array("FB1_申報時間", "C2", dataMonthStringROC) )
        Case "FB2"
            AddWorksheetFields "FOA", Array( _
                Array("FB2_申報時間", "D2", dataMonthStringROC), _
                Array("FB2_存放及拆借同業", "F9", Null), _
                Array("FB2_拆放銀行同業", "F13", Null), _
                Array("FB2_應收款項_淨額", "F36", Null), _
                Array("FB2_應收利息", "F41", Null), _
                Array("FB2_資產總計", "F85", Null) )
        Case "FB3"
            AddWorksheetFields "FOA", Array( _
                Array("FB3_申報時間", "C2", dataMonthStringROC), _
                Array("FB3_存放及拆借同業_資產面_台灣地區", "D9", Null), _
                Array("FB3_同業存款及拆放_負債面_台灣地區", "D10", Null) )
        Case "FB3A"
            ' Dynamically create in following Process Processdure
            AddWorksheetFields "FOA", Array( _
                Array("FB3A_申報時間", "C2", dataMonthStringROC) )
        Case "FM5"
            ' No Data
            AddWorksheetFields "FOA", Array( _
                Array("FM5_申報時間", "C2", dataMonthStringROC) )
        Case "FM11"
            AddWorksheetFields "FOA", Array( _
                Array("FM11_申報時間", "D2", dataMonthStringROC), _
                Array("FM11_一利息股息收入_利息_其他", "E15", Null), _
                Array("FM11_一利息收入_自中華民國境內其他客戶", "E36", Null) )
        Case "FM13"
            AddWorksheetFields "FOA", Array( _
                Array("FM13_申報時間", "D2", dataMonthStringROC), _
                Array("FM13_OBU_債票券投資_累計減損", "U9", Null) )
        Case "AI821"
            AddWorksheetFields "Table1", Array( _
                Array("AI821_申報時間", "B3", dataMonthStringROC_NUM), _
                Array("AI821_其他", "D65", Null) )
        Case "Table2"
            AddWorksheetFields "FOA", Array( _
                Array("Table2_申報時間", "E3", dataMonthStringROC), _
                Array("Table2_B_01_F4_合計", "O29", Null) )
        Case "FB5"
            AddWorksheetFields "FOA", Array( _
                Array("FB5_申報時間", "C2", dataMonthStringROC), _
                Array("FB5_外匯交易_即期外匯_DBU", "G11", Null) )
        Case "FB5A"
            'No Data
            AddWorksheetFields "FOA", Array( _
                Array("FB5A_申報時間", "C2", dataMonthStringROC) )
        Case "FM2"
            ' Dynamically create in following Process Processdure
            AddWorksheetFields "FOA", Array( _
                Array("FM2_申報時間", "C2", dataMonthStringROC) )
        Case "FM10"
            AddWorksheetFields "FOA", Array( _
                Array("FM10_申報時間", "C2", dataMonthStringROC), _
                Array("FM10_FVPL_總額A", "D20", Null), _
                Array("FM10_FVPL_淨額B", "E20", Null), _
                Array("FM10_FVOCI_總額C", "F20", Null), _
                Array("FM10_FVOCI_淨額D", "G20", Null), _
                Array("FM10_AC_總額E", "H20", Null), _
                Array("FM10_AC_淨額F", "I20", Null), _
                Array("FM10_四其他_境內_總額H", "K28", Null), _
                Array("FM10_四其他_境內_淨額I", "L28", Null) ) 
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

        Case "Table41"
            AddWorksheetFields "FOA", Array( _
                Array("Table41_申報時間", "A3", dataMonthStringROC), _
                Array("Table41_四衍生工具處分利益", "D25", Null), _
                Array("Table41_四衍生工具處分損失", "G25", Null) )
        Case "AI602"
            AddWorksheetFields "Table1", Array( _
                Array("AI602_申報時間", "B3", dataMonthStringROC_NUM), _
                Array("AI602_公司債_帳面價值_合計_F10", "L11", Null) )
            AddWorksheetFields "Table2", Array( _
                Array("AI602_金融債_投資成本_FVPL_F1", "C10", Null), _
                Array("AI602_金融債_帳面價值_合計_F5", "G11", Null) )
        Case "AI240"
            AddWorksheetFields "工作表1", Array( _
                Array("AI240_申報時間", "A2", dataMonthStringROC_NUM), _
                Array("AI240_其他到期資金流入項目_10天", "C5", Null), _
                Array("AI240_其他到期資金流出項目_1年以上", "H6", Null) )
        ' 如有其他報表，依需求加入不同工作表及欄位定義
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
        If ws Is Nothing Then
            MsgBox "Workbook 中找不到工作表: " & wsKey, vbExclamation
            WriteLog "Workbook 中找不到工作表: " & wsKey
            Exit Sub
        End If
        
        Set wsDict = clsWorksheets(wsKey)
        Set dictValues = wsDict("Values")
        Set dictAddresses = wsDict("Addresses")
        For Each fieldKey In dictValues.Keys
            If Not IsNull(dictValues(fieldKey)) Then
                On Error Resume Next
                ws.Range(dictAddresses(fieldKey)).Value = dictValues(fieldKey)
                If Err.Number <> 0 Then
                    MsgBox "工作表 [" & wsKey & "] 找不到儲存格 " & _
                           dictAddresses(fieldKey) & " （欄位：" & fieldKey & "）", vbExclamation
                    WriteLog "工作表 [" & wsKey & "] 找不到儲存格 " & _
                             dictAddresses(fieldKey) & " （欄位：" & fieldKey & "）"
                    Err.Clear
                End If
                On Error GoTo 0
            Else
                ' 沒呼叫 SetField 的欄位 (值還是 Null)
                MsgBox "工作表 [" & wsKey & "] 欄位尚未設定值: " & fieldKey, vbExclamation
                WriteLog "工作表 [" & wsKey & "] 欄位尚未設定值: " & fieldKey
            End If
        Next fieldKey
        Set ws = Nothing
    Next wsKey
End Sub

'=== 報表名稱屬性 ===  
Public Property Get ReportName() As String
    ReportName = clsReportName
End Property




這是我的Module.bas為主執行序

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
    ....其他報表省略
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
    ...

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

現在我有其他報表，因為有太多field需要設定，相關欄位名稱，我有紀錄在工作表儲存格的命名名稱中，假設Table50中的R2:R30的儲存格名稱有紀錄Array("FM5_申報時間", "C2", dataMonthStringROC)中的第一個欄位名稱，S2:S30有紀錄Array("FM5_申報時間", "C2", dataMonthStringROC)中的第二個要紀錄的儲存格位置，請問我有辦法修改流程，事先定義好不同分頁和所在連續的儲存格，這是一個共用的項目，可以用在clsReport中初始化欄位，以及我後續Process階段要設置數值使用，要怎麼修改，請告訴我修改那些內容，請用較小的幅度去修改，不要修改我的大框架


' ===========================
A:
你可以維持原有的 `clsReport` 架構，只在少數地方加一個「從工作表讀取欄位定義」的方法，然後在 `Init` 裡的某些 `Case`（像是你的 Table50）改成呼叫它。整體步驟如下：

---

## 1. 在 `clsReport` 類別裡，新增一個「從 Excel 範圍讀取欄位定義」的方法

```vb
'--- 新增於 clsReport 類別中 ---
' 從指定工作表、兩段連續儲存格 (欄位名稱清單、儲存格地址清單) 讀取欄位定義
Public Sub AddWorksheetFieldsFromRanges( _
        ByVal wsName As String, _
        ByVal configSheetName As String, _
        ByVal namesRange As String, _
        ByVal addressesRange As String, _
        ByVal Optional initValue As Variant = Null)

    Dim cfg As Worksheet
    Dim arrNames As Variant, arrAddrs As Variant
    Dim i As Long, n As Long
    Set cfg = ThisWorkbook.Sheets(configSheetName)

    arrNames = cfg.Range(namesRange).Value     ' e.g. "R2:R30"
    arrAddrs = cfg.Range(addressesRange).Value ' e.g. "S2:S30"

    ' 找實際有幾列
    n = UBound(arrNames, 1)
    For i = 1 To n
        If Trim(arrNames(i, 1) & "") <> "" Then
            ' 每一行讀一組 [FieldName, Address, initValue]
            Call AddWorksheetFields(wsName, Array( _
                Array(arrNames(i, 1), arrAddrs(i, 1), initValue) _
            ))
        End If
    Next i
End Sub
```

這個方法就只做一件事：把 `namesRange`、`addressesRange` 的值各取一次，逐行呼叫你已經寫好的 `AddWorksheetFields`。

---

## 2. 在 `Init` 裡，取代某個 `Case` 的手動定義，改成「讀設定表」

假設你在某張工作表（例如叫 `"FieldConfig"`）裡：

* R2\:R30：存放欄位名稱陣列，
* S2\:S30：存放對應的儲存格位址

那麼只要把原先手寫的：

```vb
Case "Table50"
    AddWorksheetFields "Table50", Array( _
        Array("FM5_申報時間", "C2", dataMonthStringROC), _
        ...  ' 一大堆欄位
    )
```

改成：

```vb
Case "Table50"
    ' 由 FieldConfig 工作表的 R2:R30、S2:S30 一次匯入
    Call AddWorksheetFieldsFromRanges( _
         "Table50",           ' 要寫入的目標工作表
         "FieldConfig",       ' 存設定的工作表
         "R2:R30",            ' 儲存格名稱清單
         "S2:S30",            ' 儲存格地址清單
         dataMonthStringROC   ' 可指定所有初始值（或改成 Null，再分批 SetField）
    )
```

這樣，當你往後新增一整組欄位時，只要在 `FieldConfig!R2:R30`、`FieldConfig!S2:S30` 依序填好欄位名稱與對應位址，`Init` 就會自動抓進來。

---

## 3. 在 Process 階段，一樣用舊有的 `rpt.SetField wsName, fieldName, value` 來填值

因為你 `clsReport` 的結構沒變，`GetAllFieldPositions`、`SetField`、`ApplyToWorkbook` 都不用改。甚至可以只在需要時，呼叫

```vb
rpt.SetField "Table50", "欄位名稱", 你想寫的值
```

一切照舊。

---

### 小結

1. **不改現有大框架**，只在 `clsReport` 多加一個輔助的 `AddWorksheetFieldsFromRanges` 方法。
2. 在 `Init` 針對那些欄位多、想改成批次讀設定表的 `Case`，改成呼叫這個方法。
3. 其他程式碼（Process CNY1、Process\_FB2…、`SetField`、`ApplyToWorkbook` 等）完全不動。

這樣未來只要填好一張「欄位設定表」，就能一鍵搞定，不需要再手動在 VBA 裡打十幾組陣列。



' =====================

Q:
另外有一個問題是假設我要讀取的是R2:R50的欄位，實際上我設定的是儲存格在R2:R50的名稱，而不是其中的value，你有留意到這部份嗎，請告訴我你修改哪邊，請你標示出來，給我完整版本

' =====================
A:
你說得非常關鍵，我明白你的意思是：**你在欄位設定表 `FieldConfig` 中，是使用命名儲存格（Named Ranges），而非實際的值。** 也就是說，在 `"R2:R50"` 這些儲存格中，存的是儲存格名稱，例如 `"FM5_申報時間"` 指的是命名儲存格，而非字串 `"FM5_申報時間"`。

### ✅ 修改點說明：

原本的這段：

```vb
arrNames = cfg.Range(namesRange).Value
```

這樣取得的是「儲存格中的值」，但你要的其實是儲存格本身的「名稱」，這就要用：

```vb
cfg.Range(namesRange).Cells(i, 1).Name.Name
```

---

### ✅ 完整修正版程式碼如下（已標示變動行）：

```vb
'--- 新增於 clsReport 類別中 ---
' 從指定工作表、兩段連續儲存格 (欄位名稱儲存格的「名稱」清單、儲存格地址清單) 讀取欄位定義
Public Sub AddWorksheetFieldsFromRanges( _
        ByVal wsName As String, _
        ByVal configSheetName As String, _
        ByVal namesRange As String, _
        ByVal addressesRange As String, _
        Optional ByVal initValue As Variant = Null)

    Dim cfg As Worksheet
    Dim nameCell As Range
    Dim i As Long, n As Long
    Dim arrAddrs As Variant
    Dim fieldName As String, fieldAddress As String

    Set cfg = ThisWorkbook.Sheets(configSheetName)

    arrAddrs = cfg.Range(addressesRange).Value ' 一次抓出所有儲存格位址

    n = cfg.Range(namesRange).Rows.Count        ' ← 計算要處理幾列

    For i = 1 To n
        ' --- [修改點1] 取得第 i 個儲存格的名稱（不是 Value，而是 Name）
        Set nameCell = cfg.Range(namesRange).Cells(i, 1)
        On Error Resume Next
        fieldName = nameCell.Name.Name          ' ← e.g. "FM5_申報時間"
        On Error GoTo 0

        fieldAddress = arrAddrs(i, 1)

        If Len(Trim(fieldName)) > 0 And Len(Trim(fieldAddress)) > 0 Then
            ' --- 使用已經存在的 AddWorksheetFields 方法加入欄位定義
            Call AddWorksheetFields(wsName, Array( _
                Array(fieldName, fieldAddress, initValue) _
            ))
        End If
    Next i
End Sub
```

---

### ✅ 示意資料範例：

| R欄（R2\:R50） = 儲存格名稱指標（不要填值） | S欄（S2\:S50） = 對應位址 |
| --------------------------- | ------------------ |
| `FM5_申報時間`（這是命名儲存格）         | `C2`               |
| `FM5_統一編號`                  | `C3`               |
| `FM5_名稱`                    | `C4`               |
| …                           | …                  |

只要你把工作表上的 `R2:R50` 儲存格都設定好命名（透過 Excel 的「公式」>「名稱管理員」建立），這段程式碼就能正確讀取並套用。

---

### ✅ 使用方式

```vb
Case "Table50"
    ' 從設定表中讀取命名欄位 + 對應位址
    Call AddWorksheetFieldsFromRanges( _
        "Table50", _
        "FieldConfig", _
        "R2:R50", _
        "S2:S50", _
        dataMonthStringROC)
```

---

如還需要動態篩選命名或做驗證機制（例如檢查名稱是否存在），我可以幫你再強化。是否要再幫你加上「無效命名儲存格的警告」呢？


' =============================
Q:

延續上面問題，
我希望
AddWorksheetFieldsFromRanges
可以回傳我要儲存的那些欄位，
然後我再
Public Sub Init(ByVal reportName As String, _
                ByVal dataMonthStringROC As String, _
                ByVal dataMonthStringROC_NUM As String, _
                ByVal dataMonthStringROC_F1F2 As String)
把回傳的值放到            AddWorksheetFields 中的參數

並且讓我在主程序中的處理階段中，也可以透過clsReport取得，這些欄位，
我想將這些欄位上的值，逐一設定在那個儲存格上面的值


' ======================================
A:

下面範例示範最小幅度的改動，讓你：

1. **`clsReport`** 多一個 `GetFieldDefsFromRanges` 方法，**回傳**一個 `Variant` 二維陣列，每筆都是 `{欄位名稱, 儲存格位址, initValue}`
2. `Init` 裡面呼叫它，拿到回傳的 `fieldDefs`，再丟給原本的 `AddWorksheetFields`
3. 在主程序的 Process 階段，可以透過既有的 `GetAllFieldPositions(wsName)` 或 `GetAllFieldValues(wsName)`，取得那張表所有欄位的「名稱→位址」或「名稱→值」對應，再跑迴圈一一 `Range(addr).Value = val`

---

## 1. `clsReport` 裡新增函式

```vb
'=== 新增於 clsReport 類別中 ===
' 從命名儲存格清單(名稱)、位址清單(值) 取得欄位定義陣列並回傳
Public Function GetFieldDefsFromRanges( _
        ByVal configSheetName As String, _
        ByVal namesRange     As String, _
        ByVal addressesRange As String, _
        Optional ByVal initValue As Variant = Null _
    ) As Variant

    Dim cfg As Worksheet
    Dim arrAddrs As Variant
    Dim fieldDefs() As Variant
    Dim n As Long, i As Long
    Dim nameCell As Range, fieldName As String, fieldAddr As String

    Set cfg = ThisWorkbook.Sheets(configSheetName)
    arrAddrs = cfg.Range(addressesRange).Value
    n = cfg.Range(namesRange).Rows.Count

    ReDim fieldDefs(0 To n - 1)   ' 每一列一筆

    For i = 1 To n
        Set nameCell = cfg.Range(namesRange).Cells(i, 1)
        On Error Resume Next
        fieldName = nameCell.Name.Name    ' ← 取「命名儲存格」的名稱，而非 Value  ### Modified
        On Error GoTo 0
        fieldAddr = CStr(arrAddrs(i, 1))

        If Len(Trim(fieldName)) > 0 And Len(Trim(fieldAddr)) > 0 Then
            fieldDefs(i - 1) = Array(fieldName, fieldAddr, initValue)
        Else
            fieldDefs(i - 1) = Empty
        End If
    Next i

    GetFieldDefsFromRanges = fieldDefs
End Function
```

> **標示重點**
>
> * `fieldName = nameCell.Name.Name`：由 `.Name.Name` 取「儲存格的命名」
> * 最後把 `fieldDefs` 回傳

---

## 2. `Init` 裡呼叫並傳給 `AddWorksheetFields`

```vb
Public Sub Init(ByVal reportName As String, _
                ByVal dataMonthStringROC      As String, _
                ByVal dataMonthStringROC_NUM  As String, _
                ByVal dataMonthStringROC_F1F2 As String)

    clsReportName = reportName
    Set clsWorksheets = CreateObject("Scripting.Dictionary")

    Select Case reportName

        ' …… 你原有的 Case ……

        Case "Table50"
            Dim defs As Variant
            ' ← 取得設定表 FieldConfig!R2:R50 (命名儲存格清單)、S2:S50 (對應位址)
            defs = Me.GetFieldDefsFromRanges( _
                        "FieldConfig", "R2:R50", "S2:S50", dataMonthStringROC _
                   )   ' ### Modified
            ' ← 把它傳給既有的 AddWorksheetFields，一次加入所有欄位
            AddWorksheetFields "Table50", defs    ' ### Modified

        ' …… 其他 Case ……

    End Select
End Sub
```

---

## 3. Process 階段取回欄位，逐一寫進工作表

假設你的 Process\_Table50 程序要把已取到或計算好的值，寫回到 Excel：

```vb
Public Sub Process_Table50()
    Dim rpt As clsReport
    Set rpt = gReports("Table50")

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Table50")

    ' 取得「欄位名稱 → 儲存格位址」字典
    Dim posDict As Object
    Set posDict = rpt.GetAllFieldPositions("Table50")

    ' 取得「欄位名稱 → 欄位值」字典（事先透過 SetField 填好的）
    Dim valDict As Object
    Set valDict = rpt.GetAllFieldValues("Table50")

    Dim fld As Variant, addr As String, val As Variant

    For Each fld In posDict.Keys
        addr = posDict(fld)
        If valDict.Exists(fld) Then
            val = valDict(fld)
            ws.Range(addr).Value = val
        End If
    Next fld

    ' … 其餘更新 DB、驗證流程照舊 …
End Sub
```

---

### 小結

1. **新增 `GetFieldDefsFromRanges`**（回傳 `Variant` 陣列）
2. **`Init` 改成先呼叫它，再把回傳值丟給 `AddWorksheetFields`**
3. **Process 階段直接用 `GetAllFieldPositions`＋`GetAllFieldValues`，迴圈寫入**

這樣：

* 不動你的大架構
* `Init` 只多 5 行改動
* Process 也只多一段迴圈示範

如果有需要再幫你加上錯誤檢查或「命名不存在」警告，再告訴我！








很好，你的理解是正確的，也提出了進一步的要求，下面我來詳細說明：

---

## ✅ 回答你的問題

### 🔸 Q1：是只要命名被更改都會提醒嗎？

是的。只要：

* **命名被刪除**
* **命名的範圍（RefersTo）被改變**（例如從 `=Sheet1!$A$1` 改成 `=Sheet1!$B$1`）

這個 VBA 就會發出警告。

---

### 🔸 Q2：儲存格參照的欄位變動也會偵測嗎？

是的。只要你在「名稱管理員」中去更改命名所對應的儲存格（例如命名 `MyCell` 原本對應 `A1`，你改成 `B1`），這段程式都會偵測出來，因為 RefersTo 會不同。

---

## 🔧 接下來：加上「自動復原被更動的命名」

以下是完整版程式碼，會：

1. 在開啟檔案時記錄命名
2. 每次使用者切換工作表就檢查命名是否被更動
3. 發現異常 → 警告 → 自動復原命名（或重新建立）

請將以下程式碼貼在 **ThisWorkbook** 模組中：

```vba
Private OriginalNames As Collection

Private Sub Workbook_Open()
    Call 記錄原始命名
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    Call 檢查命名是否被更動
End Sub

Private Sub 記錄原始命名()
    Dim nm As Name
    Set OriginalNames = New Collection
    For Each nm In ThisWorkbook.Names
        ' 儲存命名名稱與對應範圍
        OriginalNames.Add nm.RefersTo, nm.Name
    Next nm
End Sub

Private Sub 檢查命名是否被更動()
    Dim nmName As Variant
    Dim expectedRef As String
    On Error Resume Next
    
    For Each nmName In OriginalNames
        expectedRef = OriginalNames(nmName)
        
        ' 名稱被刪除
        If ThisWorkbook.Names(nmName) Is Nothing Then
            MsgBox "命名 """ & nmName & """ 已被刪除，系統將自動復原！", vbCritical
            ThisWorkbook.Names.Add Name:=nmName, RefersTo:=expectedRef
        Else
            ' 名稱存在但範圍被改了
            If ThisWorkbook.Names(nmName).RefersTo <> expectedRef Then
                MsgBox "命名 """ & nmName & """ 的參照已被更改，系統將自動復原！", vbExclamation
                ' 刪除錯誤命名並重新建立正確的
                ThisWorkbook.Names(nmName).Delete
                ThisWorkbook.Names.Add Name:=nmName, RefersTo:=expectedRef
            End If
        End If
    Next
    On Error GoTo 0
End Sub
```

---

## 💡 補充建議

* 若你希望這段監控程式在「儲存時」或「工作表變更時」也觸發，可以額外補上：

  * `Workbook_SheetChange`
  * `Workbook_BeforeSave`

我也可以幫你加上這些額外觸發點，如果你有需要。

---

需要我幫你轉成繁體變數/註解、或讓你選擇是否復原嗎？我也可以強化這段程式。只要說一聲！

