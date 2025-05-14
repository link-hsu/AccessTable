' ==================Question:
這是我的clsReport

Option Explicit

' Report Title
Private clsReportName As String

' Dictionary：key = Worksheet Name，value = Dictionary( Keys "Fiedl Values" 與 "Field Addresses" )
Private clsWorksheets As Object

'=== 初始化報表 (根據報表名稱建立各工作表的欄位定義) ===
Public Sub Init(ByVal reportName As String, \_
ByVal dataMonthStringROC As String, \_
ByVal dataMonthStringROC\_NUM As String, \_
ByVal dataMonthStringROC\_F1F2 As String)
clsReportName = reportName
Set clsWorksheets = CreateObject("Scripting.Dictionary")

```
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
    Case "Table2"
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

    Case "Table41"
        AddWorksheetFields "FOA", Array( _
            Array("Table41_申報時間", "A3", dataMonthStringROC), _
            Array("Table41_四衍生工具處分利益", "D25", Null), _
            Array("Table41_四衍生工具處分損失", "G25", Null) )
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
    ' 如有其他報表，依需求加入不同工作表及欄位定義
End Select
```

End Sub

'=== Private Method：Add Def for Worksheet Field ===
' fieldDefs is array of fields(each field(Array) of fields(Array)),
' for each Index's Form => (FieldName, CellAddress, InitialVAlue(null))
Private Sub AddWorksheetFields(ByVal wsName As String, \_
ByVal fieldDefs As Variant)
Dim wsDict As Object, dictValues As Object, dictAddresses As Object

```
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
```

End Sub

Public Sub AddDynamicField(ByVal wsName As String, \_
ByVal fieldName As String, \_
ByVal cellAddress As String, \_
ByVal initValue As Variant)
Dim wsDict As Object
Dim dictValues As Object, dictAddresses As Object

```
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
```

End Sub

'=== Set Field Value for one sheetName ===
Public Sub SetField(ByVal wsName As String, \_
ByVal fieldName As String, \_
ByVal value As Variant)
If Not clsWorksheets.Exists(wsName) Then
Err.Raise 1002, , "工作表 \[" & wsName & "] 尚未定義於報表 " & clsReportName
End If
Dim wsDict As Object
Set wsDict = clsWorksheets(wsName)
Dim dictValues As Object
Set dictValues = wsDict("Values")
If dictValues.Exists(fieldName) Then
dictValues(fieldName) = value
Else
Err.Raise 1001, , "欄位 \[" & fieldName & "] 不存在於工作表 \[" & wsName & "] 的報表 " & clsReportName
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
Err.Raise 1002, , "工作表 \[" & wsName & "] 尚未定義於報表 " & clsReportName
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
Err.Raise 1002, , "工作表 \[" & wsName & "] 尚未定義於報表 " & clsReportName
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
Err.Raise 1002, , "工作表 \[" & wsName & "] 尚未定義於報表 " & clsReportName
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
MsgBox "報表 \[" & clsReportName & "] 以下欄位未填入有效資料:" & vbCrLf & msg, vbExclamation
WriteLog "報表 \[" & clsReportName & "] 以下欄位未填入有效資料:" & vbCrLf & msg
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

```
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
```

End Sub

'=== 報表名稱屬性 ===
Public Property Get ReportName() As String
ReportName = clsReportName
End Property

以下是我的執行Module

Option Explicit

'=== Global Config Settings ===
Public gDataMonthString As String         ' 使用者輸入的資料月份
Public gDataMonthStringROC As String      ' 資料月份ROC Format
Public gDataMonthStringROC\_NUM As String  ' 資料月份ROC\_NUM Format
Public gDataMonthStringROC\_F1F2 As String ' 資料月份ROC\_F1F2 Format
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

```
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
```

End Sub

'=== A. 初始化所有報表並將初始資料寫入 Access ===
Public Sub InitializeReports()
Dim rpt As clsReport
Dim rptName As Variant, key As Variant
Set gReports = New Collection
For Each rptName In gReportNames
Set rpt = New clsReport
rpt.Init rptName, gDataMonthStringROC, gDataMonthStringROC\_NUM, gDataMonthStringROC\_F1F2
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

Public Sub Process\_CNY1()
'=== Equal Setting ===
'Fetch Query Access DB table
Dim dataArr As Variant

```
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
```

End Sub

Public Sub Process\_FB1()
'=== Equal Setting ===
'Fetch Query Access DB table
Dim dataArr As Variant

```
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
```

End Sub

Public Sub Process\_FB2()
'=== Equal Setting ===
'Fetch Query Access DB table
Dim dataArr As Variant

```
'Declare worksheet and handle data
Dim xlsht As Worksheet

Dim i As Integer, j As Integer
Dim lastRow As Integer

Dim reportTitle As String
Dim queryTable As String

'Setting class clsReport
Dim rpt As clsReport
Set rpt = gReports("FB2")

reportTitle = "FB2"
queryTable = "FB2_OBU_AC4620B"

' dataArr = GetAccessDataAsArray(gDBPath, queryTable, gDataMonthString)
dataArr = GetAccessDataAsArray(gDBPath, queryTable, gDataMonthString)

Set xlsht = ThisWorkbook.Sheets(reportTitle)

'Clear Excel Data
xlsht.Range("A:F").ClearContents
xlsht.Range("T2:T100").ClearContents

'=== Paste Queyr Table into Excel ===
If Err.Number <> 0 Or LBound(dataArr) > UBound(dataArr) Then
    MsgBox "資料有誤: " & reportTitle & "| " & queryTable & " 資料表無資料"
    WriteLog "資料有誤: " & reportTitle & "| " & queryTable & " 資料表無資料"
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

Dim loanAmount As Double
Dim loanInterest As Double
Dim totalAsset As Double

loanAmount = 0
loanInterest = 0
totalAsset = 0
lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
Set rngs = xlsht.Range("C2:C" & lastRow)

'
For Each rng In rngs
    If CStr(rng.Value) = "115037101" Then
        loanAmount = loanAmount + rng.Offset(0, 2).Value
    ElseIf CStr(rng.Value) = "115037105" Then
        loanAmount = loanAmount + rng.Offset(0, 2).Value
    ElseIf CStr(rng.Value) = "115037115" Then
        loanAmount = loanAmount + rng.Offset(0, 2).Value
    ElseIf CStr(rng.Value) = "130152771" Then
        loanInterest = loanInterest + rng.Offset(0, 2).Value
    ElseIf CStr(rng.Value) = "130152773" Then
        loanInterest = loanInterest + rng.Offset(0, 2).Value
    ElseIf CStr(rng.Value) = "130152777" Then
        loanInterest = loanInterest + rng.Offset(0, 2).Value
    End If
Next rng

loanAmount = Round(loanAmount / 1000, 0)
loanInterest = Round(loanInterest / 1000, 0)
totalAsset = loanAmount + loanInterest

xlsht.Range("FB2_存放及拆借同業").Value = loanAmount
rpt.SetField "FOA", "FB2_存放及拆借同業", CStr(loanAmount)

xlsht.Range("FB2_拆放銀行同業").Value = loanAmount
rpt.SetField "FOA", "FB2_拆放銀行同業", CStr(loanAmount)

xlsht.Range("FB2_應收款項_淨額").Value = loanInterest
rpt.SetField "FOA", "FB2_應收款項_淨額", CStr(loanInterest)

xlsht.Range("FB2_應收利息").Value = loanInterest
rpt.SetField "FOA", "FB2_應收利息", CStr(loanInterest)

xlsht.Range("FB2_資產總計").Value = totalAsset
rpt.SetField "FOA", "FB2_資產總計", CStr(totalAsset)

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
```

End Sub

.....以下Process\_Report省略
直到下面是最後的程序

' Process C 更新原始申報檔案欄位數值及另存新檔
Public Sub UpdateExcelReports()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False

```
Dim rpt As clsReport
Dim rptName As Variant
Dim wb As Workbook
Dim emptyFilePath As String, outputFilePath As String
For Each rptName In gReportNames
    Set rpt = gReports(rptName)
    ' 開啟原始 Excel 檔（檔名以報表名稱命名）
    emptyFilePath = gReportFolder & "\" & ThisWorkbook.Sheets("ControlPanel").Range(rptName & "_FileName").Value & "." & ThisWorkbook.Sheets("ControlPanel").Range(rptName & "_ExtensionName").Value

    If rptName = "F1_F2" Then
        outputFilePath = gOutputFolder & "\" & ThisWorkbook.Sheets("ControlPanel").Range(rptName & "_FileName").Value & "." & ThisWorkbook.Sheets("ControlPanel").Range(rptName & "_ExtensionName").Value
    Else
        outputFilePath = gOutputFolder & "\" & ThisWorkbook.Sheets("ControlPanel").Range(rptName & "_FileName").Value & Replace(gDataMonthString, "/", "") & "." & ThisWorkbook.Sheets("ControlPanel").Range(rptName & "_ExtensionName").Value
    End If

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
```

CleanUp:
' 還原警示
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True
End Sub

我希望將我的程式處理流程修改成這樣，
以上現在的版本是直接將
gReportNames
中的所有實體化實例都跑一次，其中每一個實例都有對應的Sub，例如AI240對應Sub Process\_AI240，
我現在希望提供使用者其他的選擇，所以在使用者點擊後，
在input輸入 gDataMonthString之後，
我希望跳出視窗，讓使用者可以選擇兩個按鈕，
第一個是要跑全部的報表，第二個是只要執行特定名稱的報表，
如果點擊要跑全部的報表，則就執行全部的執行程序，
如果點擊要執行特定名稱的報表，則跳出Input讓使用者輸入要執行的報表，
針對使用者輸入的內容，使用者可以輸入 CNY1, FB2,FB3, FM11,AI821,Table2，
，此時要針對這個字串進行處理，replace中間的 " " space以""取代，並且使用,來分割，
所以分別需要執行 CNY1及FB2及FB3及FM11及AI821及Table2這幾個報表，
請問要怎麼修改

' ==================Answer:

以下範例示範如何在 `Main` 入口新增「全部報表／指定報表」的選擇，並根據使用者輸入動態決定要跑哪幾個報表。

> **思路**
>
> 1. 先請使用者輸入資料月份並做基本設定
> 2. 跳出訊息盒（`MsgBox`）詢問：要執行「全部報表」還是「指定報表」
> 3. 如果選「全部」，就沿用既有的 `gReportNames`；若選「指定」，再跳出 `InputBox`，請使用者輸入以逗號分隔的報表名稱
> 4. 把使用者的字串中所有空白去除、`Split` 成陣列後，覆寫掉 `gReportNames`
> 5. 底下的 `InitializeReports`、`Process_…`、`UpdateExcelReports` 這三大段流程都以新的 `gReportNames` 為基礎

---

```vb
Public Sub Main()
    Dim isInputValid As Boolean
    isInputValid = False
    
    '—— 1. 要求使用者輸入資料月份 —— 
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
    
    '—— 2. 轉換並初始化全域變數 —— 
    gDataMonthStringROC      = ConvertToROCFormat(gDataMonthString, "ROC")
    gDataMonthStringROC_NUM  = ConvertToROCFormat(gDataMonthString, "NUM")
    gDataMonthStringROC_F1F2 = ConvertToROCFormat(gDataMonthString, "F1F2")
    gDBPath       = ThisWorkbook.Path & "\" & ThisWorkbook.Sheets("ControlPanel").Range("DBsPathFileName").Value
    gReportFolder = ThisWorkbook.Path & "\" & ThisWorkbook.Sheets("ControlPanel").Range("EmptyReportPath").Value
    gOutputFolder = ThisWorkbook.Path & "\" & ThisWorkbook.Sheets("ControlPanel").Range("OutputReportPath").Value
    
    '—— 原本的完整報表清單 —— 
    Dim allReportNames As Variant
    allReportNames = Array("CNY1", "FB1", "FB2", "FB3", "FB3A", _
                           "FM5", "FM11", "FM13", "AI821", _
                           "Table2", "FB5", "FB5A", "FM2", _
                           "FM10", "F1_F2", "Table41", "AI602", "AI240")
    
    '—— 3. 詢問要跑「全部」或「指定」 —— 
    Dim resp As VbMsgBoxResult
    resp = MsgBox("要執行全部報表嗎？" & vbCrLf & _
                  "【是】→ 全部報表" & vbCrLf & _
                  "【否】→ 指定報表", _
                  vbQuestion + vbYesNo, "執行報表選擇")
                  
    Dim userInput As String
    If resp = vbYes Then
        gReportNames = allReportNames
    Else
        userInput = InputBox("請輸入要執行的報表名稱，用逗號分隔 (例如：CNY1,FB2,FM11)：", "指定報表")
        ' 移除空白並拆成陣列
        userInput = Replace(userInput, " ", "")
        gReportNames = Split(userInput, ",")
        ' （可加驗證：檢查每個名稱是否存在於 allReportNames 中，若無則提示錯誤並結束或重新輸入）
    End If
    
    '—— 4. 初始化、處理、更新 —— 
    Call InitializeReports
    MsgBox "完成 Process A"
    WriteLog "完成 Process A"
    
    ' 依 gReportNames 動態呼叫各 Process_Sub
    Dim rptName As Variant
    For Each rptName In gReportNames
        Select Case rptName
            Case "CNY1":    Call Process_CNY1
            Case "FB1":     Call Process_FB1
            Case "FB2":     Call Process_FB2
            Case "FB3":     Call Process_FB3
            Case "FB3A":    Call Process_FB3A
            Case "FM5":     Call Process_FM5
            Case "FM11":    Call Process_FM11
            Case "FM13":    Call Process_FM13
            Case "AI821":   Call Process_AI821
            Case "Table2":  Call Process_Table2
            Case "FB5":     Call Process_FB5
            Case "FB5A":    Call Process_FB5A
            Case "FM2":     Call Process_FM2
            Case "FM10":    Call Process_FM10
            Case "F1_F2":   Call Process_F1_F2
            Case "Table41": Call Process_Table41
            Case "AI602":   Call Process_AI602
            Case "AI240":   Call Process_AI240
            Case Else
                MsgBox "未知的報表名稱: " & rptName, vbExclamation
        End Select
    Next rptName
    
    MsgBox "完成 Process B"
    WriteLog "完成 Process B"
    
    ' 最後更新 Excel 檔案
    Call UpdateExcelReports
    MsgBox "完成 Process C (全部處理程序完成)"
    WriteLog "完成 Process C (全部處理程序完成)"
End Sub
```

### 說明

1. **MsgBox + vbYesNo** 作為「全部／指定」選項。
2. 若選「否」，再用 `InputBox` 請使用者輸入以「,」分隔的報表名稱，並以 `Replace` 抹去所有空白，再用 `Split` 生成陣列。
3. 直接把新的陣列丟給 `gReportNames`，後面所有處理都以此為依據。
4. 在迴圈中用 `Select Case` 配對名稱呼叫對應 `Process_XXX` 子程式。

你也可以再加一段「檢查輸入的報表名稱是否都合法」的程式，若使用者打錯名稱則 `MsgBox` 提醒並結束或重新輸入。這樣整體走起來就能讓使用者選擇要全部執行還是只執行特定報表了。





' ==================Question:

請幫我標示出修改的地方在哪邊，並幫我加入 「檢查輸入的報表名稱是否都合法」的程式


' ==================Answer:

以下是針對你的 `Main` 程式所作的完整修改，包含：

---

### ✅ **新增或修改的部分標示說明**

我會用以下方式標示：

* `► 新增`：新增的程式碼區段
* `★ 修改`：原程式碼有變動的部分

---

### ✅ **加入的功能**

1. 提供「全部報表 / 指定報表」選項（MsgBox）
2. 使用者輸入報表清單後，檢查是否存在不合法的報表名稱（回報錯誤）

---

### 📌 **修改後的完整程式碼含標註**

```vb
Public Sub Main()
    Dim isInputValid As Boolean
    isInputValid = False

    '—— 1. 要求使用者輸入資料月份 —— 
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

    '—— 2. 轉換並初始化全域變數 —— 
    gDataMonthStringROC      = ConvertToROCFormat(gDataMonthString, "ROC")
    gDataMonthStringROC_NUM  = ConvertToROCFormat(gDataMonthString, "NUM")
    gDataMonthStringROC_F1F2 = ConvertToROCFormat(gDataMonthString, "F1F2")
    gDBPath       = ThisWorkbook.Path & "\" & ThisWorkbook.Sheets("ControlPanel").Range("DBsPathFileName").Value
    gReportFolder = ThisWorkbook.Path & "\" & ThisWorkbook.Sheets("ControlPanel").Range("EmptyReportPath").Value
    gOutputFolder = ThisWorkbook.Path & "\" & ThisWorkbook.Sheets("ControlPanel").Range("OutputReportPath").Value

    '► 新增：所有合法報表名稱清單
    Dim allReportNames As Variant
    allReportNames = Array("CNY1", "FB1", "FB2", "FB3", "FB3A", _
                           "FM5", "FM11", "FM13", "AI821", _
                           "Table2", "FB5", "FB5A", "FM2", _
                           "FM10", "F1_F2", "Table41", "AI602", "AI240")

    '► 新增：詢問是否執行全部報表
    Dim resp As VbMsgBoxResult
    resp = MsgBox("要執行全部報表嗎？" & vbCrLf & _
                  "【是】→ 全部報表" & vbCrLf & _
                  "【否】→ 指定報表", _
                  vbQuestion + vbYesNo, "執行報表選擇")

    Dim userInput As String
    If resp = vbYes Then
        gReportNames = allReportNames
    Else
        userInput = InputBox("請輸入要執行的報表名稱，用逗號分隔 (例如：CNY1,FB2,FM11)：", "指定報表")
        userInput = Replace(userInput, " ", "")
        gReportNames = Split(userInput, ",")

        '► 新增：檢查不合法的報表名稱
        Dim i As Integer
        Dim invalidReports As String
        Dim found As Boolean

        For i = LBound(gReportNames) To UBound(gReportNames)
            found = False
            Dim j As Integer
            For j = LBound(allReportNames) To UBound(allReportNames)
                If UCase(gReportNames(i)) = UCase(allReportNames(j)) Then
                    found = True
                    Exit For
                End If
            Next j
            If Not found Then
                invalidReports = invalidReports & gReportNames(i) & ", "
            End If
        Next i

        If Len(invalidReports) > 0 Then
            invalidReports = Left(invalidReports, Len(invalidReports) - 2)
            MsgBox "下列報表名稱不合法，請重新確認：" & vbCrLf & invalidReports, vbCritical, "報表名稱錯誤"
            Exit Sub
        End If
    End If

    '—— 4. 初始化、處理、更新 —— 
    Call InitializeReports
    MsgBox "完成 Process A"
    WriteLog "完成 Process A"

    '★ 修改：依 gReportNames 執行
    Dim rptName As Variant
    For Each rptName In gReportNames
        Select Case UCase(rptName)
            Case "CNY1":    Call Process_CNY1
            Case "FB1":     Call Process_FB1
            Case "FB2":     Call Process_FB2
            Case "FB3":     Call Process_FB3
            Case "FB3A":    Call Process_FB3A
            Case "FM5":     Call Process_FM5
            Case "FM11":    Call Process_FM11
            Case "FM13":    Call Process_FM13
            Case "AI821":   Call Process_AI821
            Case "TABLE2":  Call Process_Table2
            Case "FB5":     Call Process_FB5
            Case "FB5A":    Call Process_FB5A
            Case "FM2":     Call Process_FM2
            Case "FM10":    Call Process_FM10
            Case "F1_F2":   Call Process_F1_F2
            Case "TABLE41": Call Process_Table41
            Case "AI602":   Call Process_AI602
            Case "AI240":   Call Process_AI240
            Case Else
                MsgBox "未知的報表名稱: " & rptName, vbExclamation
        End Select
    Next rptName

    MsgBox "完成 Process B"
    WriteLog "完成 Process B"

    Call UpdateExcelReports
    MsgBox "完成 Process C (全部處理程序完成)"
    WriteLog "完成 Process C (全部處理程序完成)"
End Sub
```

---

### ✅ **效果總結**

* 使用者在選擇「指定報表」後，如輸入了不存在的報表名稱，會立即收到錯誤訊息，並結束程式。
* 報表處理邏輯全面採用 `gReportNames`，動態依據使用者選擇執行。

---

若你希望「錯誤報表可重新輸入」而非直接 `Exit Sub`，我也可以幫你改成 `Do Until valid` 的結構。要嗎？

' ==================Question:

1.使用Excel vba，請問要如何讓Excel開啟時讓視窗停留在ControlPanel分頁，

2.我想在這個程序中的最後加入將該分頁的顏色變成其他顏色，例如 黃色 或 紅色 請給我一些顏色選擇

Public Sub Process\_AI240()
'=== Equal Setting ===
'Fetch Query Access DB table
Dim dataArr\_1 As Variant
Dim dataArr\_2 As Variant

```
'Declare worksheet and handle data
Dim xlsht As Worksheet

Dim i As Integer, j As Integer
Dim lastRow As Integer

Dim reportTitle As String
Dim queryTable_1 As String
Dim queryTable_2 As String

'Setting class clsReport
Dim rpt As clsReport
Set rpt = gReports("AI240")

reportTitle = "AI240"
queryTable_1 = "AI240_DBU_DL6850_LIST"
queryTable_2 = "AI240_DBU_DL6850_Subtoal"

' dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1, gDataMonthString)
' dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2, gDataMonthString)
dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1, gDataMonthString)
dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2, gDataMonthString)

Set xlsht = ThisWorkbook.Sheets(reportTitle)

'Clear Excel Data
xlsht.Range("A:L").ClearContents
xlsht.Range("T2:T100").ClearContents

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

'--------------
'Unique Setting
'--------------
Dim rngs As Range
Dim rng As Range

Dim buyAmountTWD_0to10 As Double
Dim buyAmountTWD_11to30 As Double
Dim buyAmountTWD_31to90 As Double
Dim buyAmountTWD_91to180 As Double
Dim buyAmountTWD_181to365 As Double
Dim buyAmountTWD_over365 As Double

Dim sellAmountTWD_0to10 As Double
Dim sellAmountTWD_11to30 As Double
Dim sellAmountTWD_31to90 As Double
Dim sellAmountTWD_91to180 As Double
Dim sellAmountTWD_181to365 As Double
Dim sellAmountTWD_over365 As Double

buyAmountTWD_0to10 = 0
buyAmountTWD_11to30 = 0
buyAmountTWD_31to90 = 0
buyAmountTWD_91to180 = 0
buyAmountTWD_181to365 = 0
buyAmountTWD_over365 = 0

sellAmountTWD_0to10 = 0
sellAmountTWD_11to30 = 0
sellAmountTWD_31to90 = 0
sellAmountTWD_91to180 = 0
sellAmountTWD_181to365 = 0
sellAmountTWD_over365 = 0

lastRow = xlsht.Cells(xlsht.Rows.Count, "J").End(xlUp).Row
Set rngs = xlsht.Range("J2:J" & lastRow)

For Each rng In rngs
    If CStr(rng.Value) = "基準日後0-10天" Then
        buyAmountTWD_0to10 = buyAmountTWD_0to10 + rng.Offset(0, 1).Value
        sellAmountTWD_0to10 = sellAmountTWD_0to10 + rng.Offset(0, 2).Value
    ElseIf CStr(rng.Value) = "基準日後11-30天" Then
        buyAmountTWD_11to30 = buyAmountTWD_11to30 + rng.Offset(0, 1).Value
        sellAmountTWD_11to30 = sellAmountTWD_11to30 + rng.Offset(0, 2).Value
    ElseIf CStr(rng.Value) = "基準日後31-90天" Then
        buyAmountTWD_31to90 = buyAmountTWD_31to90 + rng.Offset(0, 1).Value
        sellAmountTWD_31to90 = sellAmountTWD_31to90 + rng.Offset(0, 2).Value
    ElseIf CStr(rng.Value) = "基準日後91-180天" Then
        buyAmountTWD_91to180 = buyAmountTWD_91to180 + rng.Offset(0, 1).Value
        sellAmountTWD_91to180 = sellAmountTWD_91to180 + rng.Offset(0, 2).Value
    ElseIf CStr(rng.Value) = "基準日後181天-1年" Then
        buyAmountTWD_181to365 = buyAmountTWD_181to365 + rng.Offset(0, 1).Value
        sellAmountTWD_181to365 = sellAmountTWD_181to365 + rng.Offset(0, 2).Value
    ElseIf CStr(rng.Value) = "超過基準日後一年" Then
        buyAmountTWD_over365 = buyAmountTWD_over365 + rng.Offset(0, 1).Value
        sellAmountTWD_over365 = sellAmountTWD_over365 + rng.Offset(0, 2).Value
    End If
Next rng


xlsht.Range("AI240_其他到期資金流入項目_10天").Value = buyAmountTWD_0to10
rpt.SetField "工作表1", "AI240_其他到期資金流入項目_10天", CStr(buyAmountTWD_0to10)

xlsht.Range("AI240_其他到期資金流入項目_30天").Value = buyAmountTWD_11to30
rpt.SetField "工作表1", "AI240_其他到期資金流入項目_30天", CStr(buyAmountTWD_11to30)

xlsht.Range("AI240_其他到期資金流入項目_90天").Value = buyAmountTWD_31to90
rpt.SetField "工作表1", "AI240_其他到期資金流入項目_90天", CStr(buyAmountTWD_31to90)

xlsht.Range("AI240_其他到期資金流入項目_180天").Value = buyAmountTWD_91to180
rpt.SetField "工作表1", "AI240_其他到期資金流入項目_180天", CStr(buyAmountTWD_91to180)

xlsht.Range("AI240_其他到期資金流入項目_1年").Value = buyAmountTWD_181to365
rpt.SetField "工作表1", "AI240_其他到期資金流入項目_1年", CStr(buyAmountTWD_181to365)

xlsht.Range("AI240_其他到期資金流入項目_1年以上").Value = buyAmountTWD_over365
rpt.SetField "工作表1", "AI240_其他到期資金流入項目_1年以上", CStr(buyAmountTWD_over365)


xlsht.Range("AI240_其他到期資金流出項目_10天").Value = sellAmountTWD_0to10
rpt.SetField "工作表1", "AI240_其他到期資金流出項目_10天", CStr(sellAmountTWD_0to10)

xlsht.Range("AI240_其他到期資金流出項目_30天").Value = sellAmountTWD_11to30
rpt.SetField "工作表1", "AI240_其他到期資金流出項目_30天", CStr(sellAmountTWD_11to30)

xlsht.Range("AI240_其他到期資金流出項目_90天").Value = sellAmountTWD_31to90
rpt.SetField "工作表1", "AI240_其他到期資金流出項目_90天", CStr(sellAmountTWD_31to90)

xlsht.Range("AI240_其他到期資金流出項目_180天").Value = sellAmountTWD_91to180
rpt.SetField "工作表1", "AI240_其他到期資金流出項目_180天", CStr(sellAmountTWD_91to180)

xlsht.Range("AI240_其他到期資金流出項目_1年").Value = sellAmountTWD_181to365
rpt.SetField "工作表1", "AI240_其他到期資金流出項目_1年", CStr(sellAmountTWD_181to365)

xlsht.Range("AI240_其他到期資金流出項目_1年以上").Value = sellAmountTWD_over365
rpt.SetField "工作表1", "AI240_其他到期資金流出項目_1年以上", CStr(sellAmountTWD_over365)

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
```

End Sub

3.我想要在開啟Excel時，將所有分頁的顏色都變成沒有顏色，

請問以上問題怎麼使用vba實現



' ==================Answer:


下面範例示範如何用 VBA 同時達成三個需求：

1. **開啟 Excel 時，自動切到 ControlPanel 工作表**
2. **在 `Process_AI240` 最後，將該工作表分頁設成您想要的顏色（例如黃色或紅色）**
3. **開啟 Excel 時，先清除所有分頁顏色**

---

## 一. 在 `ThisWorkbook` 模組加入 Workbook\_Open 事件

打開 VBA 編輯器 (Alt+F11)，在左側找到 **ThisWorkbook**，貼上：

```vb
'── ThisWorkbook 模組 ──
Private Sub Workbook_Open()
    Dim ws As Worksheet
    
    ' 1. 清除所有分頁顏色
    For Each ws In Me.Worksheets
        ws.Tab.ColorIndex = xlColorIndexNone
    Next ws
    
    ' 2. 切到 ControlPanel 分頁
    On Error Resume Next
    Me.Worksheets("ControlPanel").Activate
    On Error GoTo 0
End Sub
```

* `ws.Tab.ColorIndex = xlColorIndexNone`：移除所有分頁的底色。
* `Me.Worksheets("ControlPanel").Activate`：啟動時自動切到 ControlPanel 頁面。

---

## 二. 在 `Process_AI240` 最後加入分頁著色

```vb
Public Sub Process_AI240()
    ' ...（原有程式略）...

    ' 2. Update Access DB
    If rpt.ValidateFields() Then
        ' ...（更新資料庫略）...
    End If

    '► 新增：將 AI240 該工作表分頁設色
    With xlsht.Tab
        ' 以下提供兩種做法，請擇一啟用：
        
        ' 方法一：用 ColorIndex（內建色碼，快速簡單）
        '.ColorIndex = 6    ' 6 = 黃色
        '.ColorIndex = 3    ' 3 = 紅色
        
        ' 方法二：用 RGB 指定精確色彩
        .Color = RGB(255, 255, 0)   ' 黃色
        '.Color = RGB(255,   0,   0) ' 紅色
    End With

End Sub
```

### 常用分頁色彩選項

| 顏色 | ColorIndex | RGB 值            |
| -- | ---------- | ---------------- |
| 黃色 | 6          | `RGB(255,255,0)` |
| 紅色 | 3          | `RGB(255,0,0)`   |
| 綠色 | 4          | `RGB(0,255,0)`   |
| 藍色 | 5          | `RGB(0,0,255)`   |
| 橙色 | 46         | `RGB(255,165,0)` |
| 紫色 | 13         | `RGB(128,0,128)` |

* **ColorIndex**：Excel 內定 56 色索引值，使用 `.ColorIndex = n` 最簡單。
* **.Color = RGB(...)**：若要更精確的自訂色，用 RGB 三原色指定。

---

完成以上設定後：

* **每次打開活頁簿**：所有分頁底色先被清除，然後畫面游標自動跳到 ControlPanel。
* **執行到 `Process_AI240`**：AI240 那頁的分頁底色就會變成你指定的黃色或紅色。



' ==================Question:

以下是我的執行Module

Option Explicit

'=== Global Config Settings ===
Public gDataMonthString As String         ' 使用者輸入的資料月份
Public gDataMonthStringROC As String      ' 資料月份ROC Format
Public gDataMonthStringROC\_NUM As String  ' 資料月份ROC\_NUM Format
Public gDataMonthStringROC\_F1F2 As String ' 資料月份ROC\_F1F2 Format
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

```
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
```

End Sub

'=== A. 初始化所有報表並將初始資料寫入 Access ===
Public Sub InitializeReports()
Dim rpt As clsReport
Dim rptName As Variant, key As Variant
Set gReports = New Collection
For Each rptName In gReportNames
Set rpt = New clsReport
rpt.Init rptName, gDataMonthStringROC, gDataMonthStringROC\_NUM, gDataMonthStringROC\_F1F2
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

Public Sub Process\_CNY1()
'=== Equal Setting ===
'Fetch Query Access DB table
Dim dataArr As Variant

```
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
```

End Sub

' Process C 更新原始申報檔案欄位數值及另存新檔
Public Sub UpdateExcelReports()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False

```
Dim rpt As clsReport
Dim rptName As Variant
Dim wb As Workbook
Dim emptyFilePath As String, outputFilePath As String
For Each rptName In gReportNames
    Set rpt = gReports(rptName)
    ' 開啟原始 Excel 檔（檔名以報表名稱命名）
    emptyFilePath = gReportFolder & "\" & ThisWorkbook.Sheets("ControlPanel").Range(rptName & "_FileName").Value & "." & ThisWorkbook.Sheets("ControlPanel").Range(rptName & "_ExtensionName").Value

    If rptName = "F1_F2" Then
        outputFilePath = gOutputFolder & "\" & ThisWorkbook.Sheets("ControlPanel").Range(rptName & "_FileName").Value & "." & ThisWorkbook.Sheets("ControlPanel").Range(rptName & "_ExtensionName").Value
    Else
        outputFilePath = gOutputFolder & "\" & ThisWorkbook.Sheets("ControlPanel").Range(rptName & "_FileName").Value & Replace(gDataMonthString, "/", "") & "." & ThisWorkbook.Sheets("ControlPanel").Range(rptName & "_ExtensionName").Value
    End If

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
```

CleanUp:
' 還原警示
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True
End Sub

在像以上的處理程序中，因為其中有些儲存格的值，需要使用者自行填入，
例如 xlsht.Range("CNY1\_國外部").Value 或 xlsht.Range("CNY1\_信託部").Value 沒有填入資料，
則會詢問使用者是不是仍然要製作這份報表，
如果回答是，則繼續執行程序，將那些欄位的值填為 0
如果回答否，則將現在那個報表從 gReportNames 移除，以及那個報表從 gReports 中移除，後續就不會製作這個報表，接著再中斷這個Sub
這是用來管理當使用者沒有輸入製作報表所需要的值，所採行的處理，請問要怎麼修改，在修改的地方幫我標示出來




' ==================Answer:

以下範例示範以 **Process\_CNY1** 為例，說明如何在 VBA 中：

1. 檢查「必要欄位」是否填入
2. 若未填，跳出 `MsgBox` 詢問是否要繼續

   * **是** → 自動將這些欄位填 0
   * **否** → 從 `gReportNames` 與 `gReports` 中移除此報表，並 `Exit Sub`

> **標示說明**
>
> * `► 新增`：全新加入的程式段
> * `★ 修改`：對你原有程式有調整的地方

---

```vb
'=== 全域：移除陣列元素的輔助 Function ► 新增 ===
Private Sub RemoveFromArray(ByRef arr As Variant, ByVal val As String)
    Dim tmp As Collection
    Set tmp = New Collection
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If UCase(arr(i)) <> UCase(val) Then
            tmp.Add arr(i)
        End If
    Next i
    
    ReDim arr(0 To tmp.Count - 1)
    For i = 1 To tmp.Count
        arr(i - 1) = tmp(i)
    Next i
End Sub

'=== B 各報表獨立處理邏輯（以 Process_CNY1 為例） ===
Public Sub Process_CNY1()
    ' ...（前面程式不變）...

    '--------------
    ' Unique Setting
    '--------------
    ' ...（你的 fxReceive/fxPay 處理）...

    '► 新增：檢查「必要欄位」是否有使用者手動填入的欄位
    '    這裡假設「CNY1_國外部」與「CNY1_信託部」為必要欄位
    If xlsht.Range("CNY1_國外部").Value = "" Or _
       xlsht.Range("CNY1_信託部").Value = "" Then

        Dim resp As VbMsgBoxResult
        resp = MsgBox( _
            "有必要欄位尚未填入：" & vbCrLf & _
            "・CNY1_國外部" & vbCrLf & _
            "・CNY1_信託部" & vbCrLf & vbCrLf & _
            "是否仍要繼續製作此報表？", _
            vbQuestion + vbYesNo, "必要欄位未填")

        If resp = vbYes Then
            ' 使用者願意繼續 → 將缺值欄位設為 0
            xlsht.Range("CNY1_國外部").Value = 0
            rpt.SetField "CNY1", "CNY1_國外部", 0

            xlsht.Range("CNY1_信託部").Value = 0
            rpt.SetField "CNY1", "CNY1_信託部", 0
        Else
            ' 使用者取消 → 從 gReportNames 與 gReports 移除，然後結束此 Sub
            ► 新增 RemoveFromArray(gReportNames, rpt.ReportName)
            ► 新增 gReports.Remove rpt.ReportName
            Exit Sub
        End If
    End If

    '★ 修改：原本在這裡呼叫 ValidateFields 並更新 DB
    If rpt.ValidateFields() Then
        ' ...（更新資料庫）...
    End If
End Sub
```

### 步驟說明

1. **RemoveFromArray**

   * 先將陣列轉進 `Collection`，跳過要刪除的值，再重建陣列。
2. **Process\_CNY1**

   * 在「Unique Setting」與「ValidateFields」之間插入必要欄位檢查。
   * 若任一必要欄位為空，跳出詢問。
   * **是** → 自動補 0 並繼續後續流程。
   * **否** → 呼叫 `RemoveFromArray` & `gReports.Remove`，把此報表從待處理名單移除，並 `Exit Sub`。

如此一來，若使用者不想繼續製作某張報表，就可以動態把它從清單中拿掉，後續不會再被執行或匯出。你可以仿照這個範例，把同樣邏輯套用到其他 `Process_XXX` 子程式裡。




' ==================Question:

以下是我的執行Module

Option Explicit

'=== Global Config Settings ===
Public gDataMonthString As String         ' 使用者輸入的資料月份
Public gDataMonthStringROC As String      ' 資料月份ROC Format
Public gDataMonthStringROC\_NUM As String  ' 資料月份ROC\_NUM Format
Public gDataMonthStringROC\_F1F2 As String ' 資料月份ROC\_F1F2 Format
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

```
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
```

End Sub

'=== A. 初始化所有報表並將初始資料寫入 Access ===
Public Sub InitializeReports()
Dim rpt As clsReport
Dim rptName As Variant, key As Variant
Set gReports = New Collection
For Each rptName In gReportNames
Set rpt = New clsReport
rpt.Init rptName, gDataMonthStringROC, gDataMonthStringROC\_NUM, gDataMonthStringROC\_F1F2
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

Public Sub Process\_CNY1()
'=== Equal Setting ===
'Fetch Query Access DB table
Dim dataArr As Variant

```
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
```

End Sub

' Process C 更新原始申報檔案欄位數值及另存新檔
Public Sub UpdateExcelReports()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False

```
Dim rpt As clsReport
Dim rptName As Variant
Dim wb As Workbook
Dim emptyFilePath As String, outputFilePath As String
For Each rptName In gReportNames
    Set rpt = gReports(rptName)
    ' 開啟原始 Excel 檔（檔名以報表名稱命名）
    emptyFilePath = gReportFolder & "\" & ThisWorkbook.Sheets("ControlPanel").Range(rptName & "_FileName").Value & "." & ThisWorkbook.Sheets("ControlPanel").Range(rptName & "_ExtensionName").Value

    If rptName = "F1_F2" Then
        outputFilePath = gOutputFolder & "\" & ThisWorkbook.Sheets("ControlPanel").Range(rptName & "_FileName").Value & "." & ThisWorkbook.Sheets("ControlPanel").Range(rptName & "_ExtensionName").Value
    Else
        outputFilePath = gOutputFolder & "\" & ThisWorkbook.Sheets("ControlPanel").Range(rptName & "_FileName").Value & Replace(gDataMonthString, "/", "") & "." & ThisWorkbook.Sheets("ControlPanel").Range(rptName & "_ExtensionName").Value
    End If

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
```

CleanUp:
' 還原警示
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True
End Sub

在像以上的處理程序中，因為其中有些儲存格的值，需要使用者自行填入，
我需要將所有需要自行填入的工作表名稱(也就是上面的gReportNames陣列中都是分頁名稱)，
以及我定義好的儲存格名稱，我想要逐一遍歷那些分頁取得儲存格中的數值，
我想要逐一使用inputvalue詢問使用者，那些分頁的儲存格名稱的數值是不是 XXX，其中儲存格已經填的數值為inputvalue的預設值，
如果使用者輸入其他數值，則儲存起來修改該儲存格中的數值，
如果使用者沒有輸入數值，則詢問使用者是否還要製作該報表，
如果回答是，則繼續執行程序，將那些欄位的值填為 0
如果回答否，也就是從gReportName中移除，後續不製作該報表

特定儲存格名稱例如
xlsht.Range("CNY1\_國外部") 或 xlsht.Range("CNY1\_信託部")

請告訴我怎麼修改，並標示修改的地方在哪邊，
以上程序要在初始化之前就先做了，才不用初始化多餘的報表，請融入原來的架構中，不要另外再開一個Sub




' ==================Answer:

以下示範如何把「一開始就依各報表、指定的儲存格詢問使用者輸入／確認值，並在必要時把該報表從清單移除」的邏輯，**融入到原有的 Main 流程、且在 InitializeReports 之前執行**。

> **標示說明**
>
> * `► 新增`：新增的程式碼段。
> * `★ 修改`：對原程式有變動的地方。

```vb
Public Sub Main()
    Dim isInputValid As Boolean
    isInputValid = False

    '—— 1. 要求使用者輸入資料月份 —— 
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

    '—— 2. 初始化各種 Format 與全域路徑 —— 
    gDataMonthStringROC      = ConvertToROCFormat(gDataMonthString, "ROC")
    gDataMonthStringROC_NUM  = ConvertToROCFormat(gDataMonthString, "NUM")
    gDataMonthStringROC_F1F2 = ConvertToROCFormat(gDataMonthString, "F1F2")
    gDBPath       = ThisWorkbook.Path & "\" & ThisWorkbook.Sheets("ControlPanel").Range("DBsPathFileName").Value
    gReportFolder = ThisWorkbook.Path & "\" & ThisWorkbook.Sheets("ControlPanel").Range("EmptyReportPath").Value
    gOutputFolder = ThisWorkbook.Path & "\" & ThisWorkbook.Sheets("ControlPanel").Range("OutputReportPath").Value

    '—— 3. 預設全部報表清單 —— 
    gReportNames = Array("CNY1", "FB1", "FB2", "FB3", "FB3A", "FM5", _
                         "FM11", "FM13", "AI821", "Table2", "FB5", _
                         "FB5A", "FM2", "FM10", "F1_F2", "Table41", _
                         "AI602", "AI240")

    '► 新增：定義每張報表必需由使用者填入／確認的儲存格名稱
    Dim req As Object
    Set req = CreateObject("Scripting.Dictionary")
    req.Add "CNY1", Array("CNY1_國外部", "CNY1_信託部")
    ' ► 若有其他報表也要此流程，可依同模式再加入：
    ' req.Add "FB2", Array("FB2_必要欄1", "FB2_必要欄2")
    ' req.Add "FM11", Array("FM11_必要欄1", ...)

    '► 新增：準備一個 Collection 暫存要「移除」的報表
    Dim toRemove As Collection
    Set toRemove = New Collection

    '► 新增：逐一詢問使用者每張報表、每個必要欄位的值
    Dim rptName As Variant, ws As Worksheet
    Dim fields As Variant, fld As Variant
    Dim defaultVal As Variant, userVal As String
    Dim resp As VbMsgBoxResult

    For Each rptName In gReportNames
        If req.Exists(rptName) Then
            Set ws = ThisWorkbook.Sheets(rptName)
            fields = req(rptName)
            For Each fld In fields
                defaultVal = ws.Range(fld).Value
                userVal = InputBox( _
                    "請確認報表 " & rptName & " 的 [" & fld & "]" & vbCrLf & _
                    "目前值：" & defaultVal & vbCrLf & _
                    "若要修改，請輸入新數值；若維持，直接按「確定」。", _
                    "確認欄位值", CStr(defaultVal) _
                )
                If userVal = "" Then
                    '► 空白表示使用者不輸入
                    resp = MsgBox("您未輸入任何數值，是否仍要製作報表 " & rptName & "？", _
                                 vbQuestion + vbYesNo, "繼續製作？")
                    If resp = vbYes Then
                        ws.Range(fld).Value = 0
                    Else
                        toRemove.Add rptName
                        Exit For   ' 跳出該報表的欄位迴圈
                    End If
                Else
                    If IsNumeric(userVal) Then
                        ws.Range(fld).Value = CDbl(userVal)
                    Else
                        MsgBox "您輸入的不是數字，將保留原值：" & defaultVal, vbExclamation
                        ws.Range(fld).Value = defaultVal
                    End If
                End If
            Next fld
        End If
    Next rptName

    '► 新增：把使用者取消的報表，從 gReportNames 中移除
    If toRemove.Count > 0 Then
        Dim i As Long
        Dim tmpArr As Variant
        tmpArr = gReportNames
        ReDim gReportNames(0 To UBound(tmpArr) - toRemove.Count)
        
        Dim idx As Long
        idx = 0
        Dim keep As Boolean
        Dim name As Variant
        
        For Each name In tmpArr
            keep = True
            For i = 1 To toRemove.Count
                If UCase(name) = UCase(toRemove(i)) Then
                    keep = False
                    Exit For
                End If
            Next i
            If keep Then
                gReportNames(idx) = name
                idx = idx + 1
            End If
        Next name
    End If

    '★ 修改：以上結束後，才呼叫 InitializeReports
    Call InitializeReports
    MsgBox "完成 Process A"
    WriteLog "完成 Process A"

    ' Process B: 依 gReportNames 執行各報表
    For Each rptName In gReportNames
        Select Case UCase(rptName)
            Case "CNY1":    Call Process_CNY1
            Case "FB1":     Call Process_FB1
            Case "FB2":     Call Process_FB2
            Case "FB3":     Call Process_FB3
            Case "FB3A":    Call Process_FB3A
            Case "FM5":     Call Process_FM5
            Case "FM11":    Call Process_FM11
            Case "FM13":    Call Process_FM13
            Case "AI821":   Call Process_AI821
            Case "TABLE2":  Call Process_Table2
            Case "FB5":     Call Process_FB5
            Case "FB5A":    Call Process_FB5A
            Case "FM2":     Call Process_FM2
            Case "FM10":    Call Process_FM10
            Case "F1_F2":   Call Process_F1_F2
            Case "TABLE41": Call Process_Table41
            Case "AI602":   Call Process_AI602
            Case "AI240":   Call Process_AI240
            Case Else
                MsgBox "未知的報表名稱: " & rptName, vbExclamation
        End Select
    Next rptName

    MsgBox "完成 Process B"
    WriteLog "完成 Process B"

    ' Process C
    Call UpdateExcelReports
    MsgBox "完成 Process C (全部處理程序完成)"
    WriteLog "完成 Process C (全部處理程序完成)"
End Sub
```

---

### 主要變動點

1. **在 `gReportNames` 定義之後、`InitializeReports` 之前，新增一段**

   * 用 `Dictionary` (`req`) 先定義「哪些報表需要哪些使用者填欄位」。
   * 迴圈逐一 `InputBox` 詢問，並視使用者回應做三件事：

     * **輸入空白** → 再問一次「是否取消此報表？」

       * **是**：將該報表名稱加入 `toRemove`；
       * **否**：自動把該欄位設為 0。
     * **輸入數字** → 取代該儲存格的預設值。
     * **輸入非數字** → 提示錯誤並保留原值。
2. **用一小段邏輯將 `toRemove` 中的報表從 `gReportNames` 陣列剔除**，之後就不會再初始化或執行。
3. **只在更新完使用者欄位之後**，才呼叫 `InitializeReports`，避免初始化不需要跑的報表。

如此即能在最一開始，就先完成使用者必填欄位的管理，並能動態決定哪些報表要跳過。


https://chatgpt.com/share/6824c07b-eef8-8010-943f-c9deec4e41af

' ==================Question:


' ==================Answer:





' ==================Question:


' ==================Answer:





' ==================Question:


' ==================Answer:




' ==================Question:


' ==================Answer:
