Option Explicit

' Report Title
Private clsReportName As String

' Dictionary：key = Worksheet Name，value = Dictionary( Keys "Fiedl Values" 與 "Field Addresses" )
Private clsWorksheets As Object

'=== 初始化報表 (根據報表名稱建立各工作表的欄位定義) ===
Public Sub Init(ByVal reportName As String,
                ByVal dataMonthStringROC As String,
                ByVal dataMonthStringROC_NUM As String,
                ByVal dataMonthStringROC_F1F2 As String)
    clsReportName = reportName
    Set clsWorksheets = CreateObject("Scripting.Dictionary")
    
    Select Case reportName
        ' Case Example
        ' 假設 CNY1 報表有三個工作表：X、Y、Z  
            ' 工作表 X 定義：  
            '   - "其他金融資產_淨額" 儲存格地址 "B2"  
            '   - "其他" 儲存格地址 "C2"  
            '   - "CNY1_資產總計" 儲存格地址 "D2"
            ' 工作表 Y 定義：  
            '   - "其他金融負債" 儲存格地址 "E2"  
            '   - "其他什項金融負債" 儲存格地址 "F2"
            ' 工作表 Z 定義：  
            '   - "CNY1_負債總計" 儲存格地址 "G2"
        Case "CNY1"
            AddWorksheetFields "CNY1", Array( _
                Array("CNY1_申報時間", "C2", dataMonthStringROC), _
                Array("CNY1_其他金融資產_淨額", "G98", Null), _
                Array("CNY1_其他", "G100", Null), _
                Array("CNY1_資產總計", "G116", Null), _
                Array("CNY1_其他金融負債", "G170", Null), _
                Array("CNY1_其他什項金融負債", "G172", Null), _
                Array("CNY1_負債總計", "G184", Null) )
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
                Array("FM11_五證券投資評價及減損迴轉利益_一年期以上之債權證券", "E25", Null), _
                Array("FM11_五證券投資評價及減損損失_一年期以上之債權證券", "I25", Null), _
                Array("FM11_一利息收入_自中華民國境內其他客戶", "E36", Null) )
        Case "FM13"
            AddWorksheetFields "FOA", Array( _
                Array("FM13_申報時間", "D2", dataMonthStringROC), _
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
                Array("AI821_申報時間", "B3", dataMonthStringROC_NUM), _
                Array("AI821_本國銀行", "D61", Null), _
                Array("AI821_陸銀在臺分行", "D62", Null), _
                Array("AI821_外商銀行在臺分行", "D63", Null), _
                Array("AI821_大陸地區銀行", "D64", Null), _
                Array("AI821_其他", "D65", Null) )
        Case "Table2"
            AddWorksheetFields "FOA", Array( _
                Array("Table2_申報時間", "E3", dataMonthStringROC), _
                Array("Table2_其他", "D17", Null), _
                Array("Table_美元_F1", "L7", Null), _
                Array("Table2_美元_F3", "N7", Null), _
                Array("Table2_美元_F4", "O7", Null) )
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
                Array("FM10_FVOCI_總額C", "F20", Null), _
                Array("FM10_FVOCI_淨額D", "G20", Null), _
                Array("FM10_AC_總額E", "H20", Null), _
                Array("FM10_AC_淨額F", "I20", Null), _
                Array("FM10_四其他_境內_總額H", "K28", Null), _
                Array("FM10_四其他_境內_淨額I", "L28", Null) )
        Case "F1_F2"
            Dim currencies As Variant, transactionTypes As Variant, startRows As Variant, colLetters As Variant
            Dim i As Integer, j As Integer
            currencies = Array("JPY", "GBP", "CHF", "CAD", "AUD", "NZD", "SGD", "HKD", "ZAR", "SEK", "THB", "RM", "EUR", "CNY", "OTHER")
            transactionTypes = Array("F1_與國外金融機構及非金融機構間交易_SPOT", _
                                     "F1_與國外金融機構及非金融機構間交易_SWAP", _
                                     "F1_與國內金融機構間交易_SPOT", _
                                     "F1_與國內金融機構間交易_SWAP", _
                                     "F1_與國內顧客間交易_SPOT")
            ' 每組交易的起始儲存格列數
            startRows = Array(8, 8, 8, 8, 8)
            ' 每組交易對應的欄位
            colLetters = Array("O", "Q", "I", "K", "B") 

            Dim fieldList() As Variant
            Dim index As Integer
            index = 0

            '修改這裡
            'ReDim fieldList(UBound(transactionTypes) * UBound(currencies))
            ReDim fieldList((UBound(transactionTypes) + 1) * (UBound(currencies) + 1) - 1)

            For i = LBound(transactionTypes) To UBound(transactionTypes)
                For j = LBound(currencies) To UBound(currencies)
                    fieldList(index) = Array(transactionTypes(i) & "_" & currencies(j), colLetters(i) & (startRows(i) + j), Null)
                    index = index + 1
                Next j
            Next i
            ' Add to Worksheet Fields
            AddWorksheetFields "f1", fieldList
            AddDynamicField "f1", "F1_申報時間", "A3", dataMonthStringROC_F1F2
            AddWorksheetFields "f2", Array( _
                Array("F2_申報時間", "A3", dataMonthStringROC_F1F2) )
        Case "Table41"
            AddWorksheetFields "FOA", Array( _
                Array("Table41_申報時間", "A3", dataMonthStringROC), _
                Array("Table41_四衍生工具處分利益", "D25", Null), _
                Array("Table41_四衍生工具處分損失", "G25", Null) )
        Case "AI602"
            AddWorksheetFields "Table1", Array( _
                Array("AI602_申報時間", "B3", dataMonthStringROC_NUM), _
                Array("AI602_政府公債_投資成本_FVOCI_F2", "D10", Null), _
                Array("AI602_政府公債_投資成本_AC_F3", "E10", Null), _
                Array("AI602_政府公債_投資成本_合計_F5", "G10", Null), _
                Array("AI602_公司債_投資成本_FVOCI_F7", "I10", Null), _
                Array("AI602_公司債_投資成本_AC_F8", "J10", Null), _
                Array("AI602_公司債_投資成本_合計_F10", "L10", Null), _
                Array("AI602_政府公債_帳面價值_FVOCI_F2", "D11", Null), _
                Array("AI602_政府公債_帳面價值_AC_F3", "E11", Null), _
                Array("AI602_政府公債_帳面價值_合計_F5", "G11", Null), _
                Array("AI602_公司債_帳面價值_FVOCI_F7", "I11", Null), _
                Array("AI602_公司債_帳面價值_AC_F8", "J11", Null), _
                Array("AI602_公司債_帳面價值_合計_F10", "L11", Null) )
            AddWorksheetFields "Table2", Array( _
                Array("AI602_金融債_投資成本_FVOCI_F2", "D10", Null), _
                Array("AI602_金融債_投資成本_AC_F3", "E10", Null), _
                Array("AI602_金融債_投資成本_合計_F5", "G10", Null), _
                Array("AI602_金融債_帳面價值_FVOCI_F2", "D11", Null), _
                Array("AI602_金融債_帳面價值_AC_F3", "E11", Null), _
                Array("AI602_金融債_帳面價值_合計_F5", "G11", Null) )
        Case "AI240"
            AddWorksheetFields "工作表1", Array( _
                Array("AI240_申報時間", "A2", dataMonthStringROC_NUM), _
                Array("AI240_其他到期資金流入項目_10天", "B5", Null), _
                Array("AI240_其他到期資金流入項目_30天", "B5", Null), _
                Array("AI240_其他到期資金流入項目_90天", "B5", Null), _
                Array("AI240_其他到期資金流入項目_180天", "B5", Null), _
                Array("AI240_其他到期資金流入項目_1年", "B5", Null), _
                Array("AI240_其他到期資金流入項目_1年以上", "B5", Null), _
                Array("AI240_其他到期資金流出項目_10天", "B5", Null), _
                Array("AI240_其他到期資金流出項目_30天", "B5", Null), _
                Array("AI240_其他到期資金流出項目_90天", "B5", Null), _
                Array("AI240_其他到期資金流出項目_180天", "B5", Null), _
                Array("AI240_其他到期資金流出項目_1年", "B5", Null), _
                Array("AI240_其他到期資金流出項目_1年以上", "B5", Null) )
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
                End If
            Next fieldKey
        Else
            MsgBox "Workbook 中找不到工作表: " & wsKey, vbExclamation
            Exit Sub
        End If
        Set ws = Nothing
    Next wsKey
End Sub

'=== 報表名稱屬性 ===  
Public Property Get ReportName() As String
    ReportName = clsReportName
End Property
