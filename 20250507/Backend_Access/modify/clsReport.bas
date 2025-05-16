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
                Array("FM11_三證券投資處分利益_一年期以上之債權證券", "E20", Null), _
                Array("FM11_三證券投資處分損失_一年期以上之債權證券", "I20", Null), _
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
        Case "TABLE2"
            AddWorksheetFields "FOA", Array( _
                Array("Table2_申報時間", "E3", dataMonthStringROC), _
                Array("Table2_A_1011100_其他", "D17", Null), _
                Array("Table2_A_1010000_合計", "D20", Null), _
                Array("Table2_B_01_F1_原幣國外資產", "L7", Null), _
                Array("Table2_B_01_F3_折合率", "N7", Null), _
                Array("Table2_B_01_F4_折合新台幣國外資產", "O7", Null), _
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

        Case "TABLE41"
            AddWorksheetFields "FOA", Array( _
                Array("Table41_申報時間", "A3", dataMonthStringROC), _
                Array("Table41_四衍生工具處分利益", "D25", Null), _
                Array("Table41_四衍生工具處分損失", "G25", Null), _
                Array("Table41_一利息收入", "D9", Null), _
                Array("Table41_一利息收入_利息", "D10", Null), _
                Array("Table41_一利息收入_利息_存放銀行同業", "D12", Null), _
                Array("Table41_二金融服務收入", "D19", Null), _
                Array("Table41_一利息支出", "G9", Null), _
                Array("Table41_一利息支出_利息", "G10", Null), _
                Array("Table41_一利息支出_利息_外國人新台幣存款", "G14", Null), _
                Array("Table41_一利息支出_利息_外國人外匯存款", "G15", Null), _
                Array("Table41_二金融服務支出", "G19", Null) )
        Case "AI602"
            AddWorksheetFields "Table1", Array( _
                Array("AI602_申報時間", "B3", dataMonthStringROC_NUM), _
                Array("AI602_政府公債_投資成本_FVPL_F1", "C10", Null), _
                Array("AI602_政府公債_投資成本_FVOCI_F2", "D10", Null), _
                Array("AI602_政府公債_投資成本_AC_F3", "E10", Null), _
                Array("AI602_政府公債_投資成本_合計_F5", "G10", Null), _
                Array("AI602_公司債_投資成本_FVPL_F6", "H10", Null), _
                Array("AI602_公司債_投資成本_FVOCI_F7", "I10", Null), _
                Array("AI602_公司債_投資成本_AC_F8", "J10", Null), _
                Array("AI602_公司債_投資成本_合計_F10", "L10", Null), _
                Array("AI602_政府公債_帳面價值_FVPL_F1", "C11", Null), _
                Array("AI602_政府公債_帳面價值_FVOCI_F2", "D11", Null), _
                Array("AI602_政府公債_帳面價值_AC_F3", "E11", Null), _
                Array("AI602_政府公債_帳面價值_合計_F5", "G11", Null), _
                Array("AI602_公司債_帳面價值_FVPL_F6", "H11", Null), _
                Array("AI602_公司債_帳面價值_FVOCI_F7", "I11", Null), _
                Array("AI602_公司債_帳面價值_AC_F8", "J11", Null), _
                Array("AI602_公司債_帳面價值_合計_F10", "L11", Null) )
            AddWorksheetFields "Table2", Array( _
                Array("AI602_金融債_投資成本_FVPL_F1", "C10", Null), _
                Array("AI602_金融債_投資成本_FVOCI_F2", "D10", Null), _
                Array("AI602_金融債_投資成本_AC_F3", "E10", Null), _
                Array("AI602_金融債_投資成本_合計_F5", "G10", Null), _
                Array("AI602_金融債_帳面價值_FVPL_F1", "C11", Null), _
                Array("AI602_金融債_帳面價值_FVOCI_F2", "D11", Null), _
                Array("AI602_金融債_帳面價值_AC_F3", "E11", Null), _
                Array("AI602_金融債_帳面價值_合計_F5", "G11", Null) )
        Case "AI240"
            AddWorksheetFields "工作表1", Array( _
                Array("AI240_申報時間", "A2", dataMonthStringROC_NUM), _
                Array("AI240_其他到期資金流入項目_10天", "C5", Null), _
                Array("AI240_其他到期資金流入項目_30天", "D5", Null), _
                Array("AI240_其他到期資金流入項目_90天", "E5", Null), _
                Array("AI240_其他到期資金流入項目_180天", "F5", Null), _
                Array("AI240_其他到期資金流入項目_1年", "G5", Null), _
                Array("AI240_其他到期資金流入項目_1年以上", "H5", Null), _
                Array("AI240_其他到期資金流出項目_10天", "C6", Null), _
                Array("AI240_其他到期資金流出項目_30天", "D6", Null), _
                Array("AI240_其他到期資金流出項目_90天", "E6", Null), _
                Array("AI240_其他到期資金流出項目_180天", "F6", Null), _
                Array("AI240_其他到期資金流出項目_1年", "G6", Null), _
                Array("AI240_其他到期資金流出項目_1年以上", "H6", Null) )
        ' 
        Case "AI822"
            AddWorksheetFields "Table1", Array( _
                Array("AI822_申報時間", "B3", dataMonthStringROC_NUM), _
                Array("AI822_授信、投資及資金拆存總額度", "C9", Null), _
                Array("AI822_上年度決算後淨值", "C10", Null), _
                Array("AI822_對大陸地區之授信、投資及資金拆存總額度占上年度決算後淨值之倍數", "C11", Null) )
                

            AddWorksheetFields "Table2", Array( _
                Array("AI822_授信", "C9", Null), _
                Array("AI822_直接往來之授信", "C10", Null), _
                Array("AI822_間接往來之授信", "C11", Null), _
                Array("AI822_減短期貿易融資", "C12", Null) )

            AddWorksheetFields "Table4", Array( _
                Array("AI822_資金拆存_小計", "E9", Null), _
                Array("AI822_債權債務剩餘期限不足3個月且交易對手之長期債信或短期債信符合投資等級以上者_帳列小計C3", "C10", Null), _
                Array("AI822_債權債務剩餘期限不足3個月且交易對手之長期債信或短期債信符合投資等級以上者_小計", "E10", Null), _
                Array("AI822_債權債務剩餘期限不足3個月且交易對手之長期債信或短期債信符合投資等級以上者_資金拆借帳列金額", "F10", Null), _
                Array("AI822_債權債務剩餘期限不足3個月且交易對手之長期債信或短期債信符合投資等級以上者_存放銀行同業帳列金額", "G10", Null), _
                Array("AI822_債權債務剩餘期限不足3個月且交易對手之長期債信或短期債信符合投資等級以上者_帳列小計D5", "H10", Null), _
                Array("AI822_大陸地區銀行_資金拆借帳列金額", "F11", Null), _
                Array("AI822_大陸地區銀行_存放銀行同業帳列金額", "G11", Null), _
                Array("AI822_大陸地區銀行_帳列小計", "H11", Null), _
                Array("AI822_中國人民銀行_資金拆借帳列金額", "F12", Null), _
                Array("AI822_中國人民銀行_存放銀行同業帳列金額", "G12", Null), _
                Array("AI822_中國人民銀行_帳列小計", "H12", Null), _
                Array("AI822_政策性及國有商業銀行_資金拆借帳列金額", "F13", Null), _
                Array("AI822_政策性及國有商業銀行_存放銀行同業帳列金額", "G13", Null), _
                Array("AI822_政策性及國有商業銀行_帳列小計", "H13", Null), _
                Array("AI822_股份制商業銀行_資金拆借帳列金額", "F14", Null), _
                Array("AI822_股份制商業銀行_存放銀行同業帳列金額", "G14", Null), _
                Array("AI822_股份制商業銀行_帳列小計", "H14", Null), _
                Array("AI822_其他_資金拆借帳列金額", "F15", Null), _
                Array("AI822_其他_存放銀行同業帳列金額", "G15", Null), _
                Array("AI822_其他_帳列小計", "H15", Null) )
                ' Array("AI822_債權債務剩餘期限不足3個月且交易對手之長期債信或短期債信符合投資等級以上者_適用權數", "D10", Null), _

            AddWorksheetFields "Table5", Array( _
                Array("AI822_保證_減風險移轉", "C9", Null), _
                Array("AI822_擔保品_減風險移轉", "D9", Null), _
                Array("AI822_小計_減風險移轉", "E9", Null), _
                Array("AI822_保證_授信", "C10", Null), _
                Array("AI822_擔保品_授信", "D10", Null), _
                Array("AI822_小計_授信", "E10", Null), _
                Array("AI822_保證_投資", "C11", Null), _
                Array("AI822_擔保品_投資", "D11", Null), _
                Array("AI822_小計_投資", "E11", Null) )

            AddWorksheetFields "Table6", Array( _
                Array("AI822_資金拆存予陸資銀行在台分行_資金拆借帳列金額", "C9", Null), _
                Array("AI822_資金拆存予陸資銀行在台分行_存放銀行同業帳列金額", "D9", Null), _
                Array("AI822_資金拆存予陸資銀行在台分行_帳列小計", "E9", Null), _
                Array("AI822_授信予陸資銀行在台分行", "E10", Null), _
                Array("AI822_投資陸資銀行在台分行發行之債券及可轉讓定期存單等", "E11", Null), _
                Array("AI822_當月授信轉銷呆帳金額", "E12", Null) )
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
