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
    Dim rptFieldInfo As Object
    Dim rptToInit As Variant
    Dim rptArray As Variant
    Set rptFieldInfo = CreateObject("Scripting.Dictionary")

    With rptFieldInfo
        .Add "TABLE10", _
        Array(
            Array("FOA", "S2:S21,W2:W21,AA2:AA21,AE2:AE21,AI2:AI21,AM2:AM21,AQ2:AQ21", "U2:U21,Y2:Y21,AC2:AC21,AG2:AG21,AK2:AK21,AO2:AO21,AS2:AS21")            
        )

        .Add "TABLE15A", _
        Array(
            Array("FOA", "S2:S10,W2:W10,AA2:AA10,AE2:AE10,AI2:AI10", "U2:U10,Y2:Y10,AC2:AC10,AG2:AG10,AK2:AK10")
        )

        .Add "TABLE15B", _
        Array()

        .Add "TABLE16", _
        Array(
            Array("FOA", "", "B2")
        )

        .Add "TABLE20", _
        Array(
            Array("FOA", "S2:S5,W2:W5,AA2:AA5", "U2:U5,Y2:Y5,AC2:AC5")
        )

        .Add "TABLE22", _
        Array(
            Array("FOA", "S2:S5,W2:W5", "U2:U5,Y2:Y5")
        )

        .Add "TABLE23", _
        Array(
            Array("FOA", "S2:S6", "U2:U6")
        )

        .Add "TABLE24", _
        Array(
            Array("FOA", "S2:S14,W2:W14,AA2:AA14,AE2:AE14,AI2:AI14,AM2:AM14,AQ2:AQ14,AU2:AU14", "U2:U14,Y2:Y14,AC2:AC14,AG2:AG14,AK2:AK14,AO2:AO14,AS2:AS14,AW2:AW14")
        )

        .Add "TABLE27", _
        Array(
            Array("FOA", "S2:S7,W2:W7,AA2:AA7,AE2:AE7,AI2:AI7", "U2:U7,Y2:Y7,AC2:AC7,AG2:AG7,AK2:AK7")
        )

        .Add "TABLE36", _
        Array(
            Array("FOA", "S2:S4,W2:W4,AA2:AA4", "U2:U4,Y2:Y4,AC2:AC4")
        )

        .Add "AI233", _
        Array(
            Array("Table1", "S2:S5,W2:W5,AA2:AA5,AE2:AE9,AI2:AI5,AM2:AM5,AQ2:AQ5,AU2:AU9", "U2:U5,Y2:Y5,AC2:AC5,AG2:AG9,AK2:AK5,AO2:AO5,AS2:AS5,AW2:AW9"),
            Array("Table2", "S10:S11,W10:W11,AA10:AA11", "U10:U11,Y10:Y11,AC10:AC11"),
            Array("Table4", "S12:S15,W12:W15,AA12:AA15,AE12:AE17", "U12:U15,Y12:Y15,AC12:AC15,AG12:AG17")
        )

        .Add "AI345", _
        Array(
            Array("", "", "")
        )

        .Add "AI405", _
        Array(
            Array("Table1", "S2:S5,W2:W5", "U2:U5,Y2:Y5")
        )

        .Add "AI410", _
        Array(
            Array("Table1", "S2:S8,W2:W8", "U2:U8,Y2:Y8")
        )

        .Add "AI430", _
        Array(
            Array("Table1", "S2:S8", "U2:U8")
        )

        .Add "AI601", _
        Array(
            Array("Table1", "S2:S48,W2:W48,AA2:AA48,AE2:AE48,AI2:AI48", "U2:U48,Y2:Y48,AC2:AC48,AG2:AG48,AK2:AK48"),
            Array("Table2", "AM2:AM48,AQ2:AQ48,AU2:AU48,AY2:AY48,BC2:BC48,BG2:BG48,BK2:BK48", "AO2:AO48,AS2:AS48,AW2:AW48,BA2:BA48,BE2:BE48,BI2:BI48,BM2:BM48"),
            Array("Table3", "S49:S65", "U49:U65")
        )

        .Add "AI605", _
        Array(
            Array("Table1", "S2:S3,W2:W3,AA2:AA3,AE2:AE3,AI2:AI3,AM2:AM3,AQ2:AQ3,AU2:AU3", "U2:U3,Y2:Y3,AC2:AC3,AG2:AG3,AK2:AK3,AO2:AO3,AS2:AS3,AW2:AW3"),
            Array("Table3", "S5:S6,W5:W6,", "U5:U6,Y5:Y6")
        )
    End With

    clsReportName = reportName
    Set clsWorksheets = CreateObject("Scripting.Dictionary")

    If rptFieldInfo.Exists(reportName) Then
        rptToInit = rptFieldInfo(reportName)
        For i = LBound(rptToInit) To UBound(rptToInit)
            Dim rptSheet As Variant
            rptSheet = rptToInit(i)
            
            Dim initSheetName As String
            Dim nameTagRng As String
            Dim addrRng As String
            Dim initValue As Variant
            
            initSheetName = rptSheet(0)
            nameTagRng = rptSheet(1)
            addrRng = rptSheet(2)
            initValue = Null
            
            ' 跳過空的 range 定義
            If Trim(nameTagRng) = "" Or Trim(addrRng) = "" Then
                WriteLog "Init 跳過 [" & reportName & "] 的 [" & initSheetName & "]：range 定義為空"
            Else
                rptArray = Me.GetFieldFromXlRanges(reportName, nameTagRng, addrRng, initValue)
                ' 呼叫 AddWorksheetFields，第一參數用 initSheetName
                AddWorksheetFields initSheetName, rptArray
            End If
        Next i        
    Else
        WriteLog "Init未定義報表：" & reportName
    End If
    
    Select Case reportName
        Case "TABLE10"
            AddDynamicField reportName, "TABLE10_申報時間", "D2", dataMonthStringROC
        Case "TABLE15A"
            AddDynamicField reportName, "TABLE15A_申報時間", "D2", dataMonthStringROC
        Case "TABLE15B"
            AddDynamicField reportName, "TABLE15B_申報時間", "D2", dataMonthStringROC
        Case "TABLE16"
            AddDynamicField reportName, "TABLE16_申報時間", "B2", dataMonthStringROC
        Case "TABLE20"
            AddDynamicField reportName, "TABLE20_申報時間", "I3", dataMonthStringROC
        Case "TABLE22"
            AddDynamicField reportName, "TABLE22_申報時間", "E2", dataMonthStringROC
        Case "TABLE23"
            AddDynamicField reportName, "TABLE23_申報時間", "E2", dataMonthStringROC
        Case "TABLE24"
            AddDynamicField reportName, "TABLE24_申報時間", "G2", dataMonthStringROC
        Case "TABLE27"
            AddDynamicField reportName, "TABLE27_申報時間", "E3", dataMonthStringROC
        Case "TABLE36"
            AddDynamicField reportName, "TABLE36_申報時間", "E2", dataMonthStringROC
        Case "AI233"
            AddDynamicField reportName, "AI233_申報時間", "B3", dataMonthStringROC_NUM
        Case "AI345"
            AddDynamicField reportName, "AI345_申報時間", "A2", dataMonthStringROC_NUM
        Case "AI405"
            AddDynamicField reportName, "AI405_申報時間", "B3", dataMonthStringROC_NUM
        Case "AI410"
            AddDynamicField reportName, "AI410_申報時間", "B3", dataMonthStringROC_NUM
        Case "AI430"
            AddDynamicField reportName, "AI430_申報時間", "B3", dataMonthStringROC_NUM
        Case "AI601"
            AddDynamicField reportName, "AI601_申報時間", "B3", dataMonthStringROC_NUM
        Case "AI605"
            AddDynamicField reportName, "AI605_申報時間", "B3", dataMonthStringROC_NUM
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

Public Function GetFieldFromXlRanges(ByVal rptSheetName As String, _
                                     ByVal namesRange As String, _
                                     ByVal addressesRange As String, _
                                     Optional ByVal initValue As Variant = Null) As Variant
    Dim rptSheet As Worksheet
    Dim nameAreas As Range, addrAreas As Range
    Dim area As Range, cell As Range
    Dim nameList As Collection, addrList As Collection
    Dim fieldDefs() As Variant
    Dim i As Long

    ' 1. 取得工作表
    On Error Resume Next
    Set rptSheet = ThisWorkbook.Worksheets(rptSheetName)
    On Error GoTo 0
    If rptSheet Is Nothing Then
        Err.Raise vbObjectError, , "找不到工作表：" & rptSheetName
    End If

    ' 2. 取得多段範圍
    On Error Resume Next
    Set nameAreas = rptSheet.Range(namesRange)
    Set addrAreas = rptSheet.Range(addressesRange)
    On Error GoTo 0
    If nameAreas Is Nothing Then _
        Err.Raise vbObjectError, , "namesRange 無效：" & namesRange
        WriteLog "namesRange 無效：" & namesRange
    If addrAreas Is Nothing Then _
        Err.Raise vbObjectError, , "addressesRange 無效：" & addressesRange
        WriteLog "addressesRange 無效：" & addressesRange
    ' 3. 把所有 namesRange 的每一個 cell 依照遍歷順序收進 nameList
    Set nameList = New Collection
    For Each area In nameAreas.Areas
        For Each cell In area.Cells
            nameList.Add cell
        Next
    Next

    ' 4. 把所有 addressesRange 的每一個 cell.value 依照遍歷順序收進 addrList
    Set addrList = New Collection
    For Each area In addrAreas.Areas
        For Each cell In area.Cells
            addrList.Add CStr(cell.Value)
        Next
    Next

    ' 5. 確認「名稱清單」與「位址清單」筆數相同
    If nameList.Count <> addrList.Count Then
        Err.Raise vbObjectError, , _
          "名稱筆數 (" & nameList.Count & ") 與位址筆數 (" & addrList.Count & ") 不一致。"
        WriteLog "名稱筆數 (" & nameList.Count & ") 與位址筆數 (" & addrList.Count & ") 不一致。"
    End If

    ' 6. 建立回傳陣列
    ReDim fieldDefs(0 To nameList.Count - 1)
    For i = 1 To nameList.Count
        ' 先清除 Err
        Err.Clear

        On Error Resume Next

        Dim nm As String
        nm = nameList(i).Name.Name
        On Error GoTo 0
        
        ' 如果沒取到名稱，就主動拋錯並中斷
        If nm = "" Then
            Err.Raise vbObjectError + 513, "GetFieldFromXlRanges", _
                      "第 " & i & " 筆 nameList 中的儲存格沒有名稱。"
            WriteLog "第 " & i & " 筆 nameList 中的儲存格沒有名稱。"
        End If

        ' cell.Name.Name 取的是「該儲存格所屬的定義名稱」；錯誤就回空字串
        fieldDefs(i - 1) = Array( _
            nm, _
            addrList(i), _
            initValue _
        )
        Err.Clear
        On Error GoTo 0
    Next

    GetFieldFromXlRanges = fieldDefs
End Function

'=== 報表名稱屬性 ===  
Public Property Get ReportName() As String
    ReportName = clsReportName
End Property
