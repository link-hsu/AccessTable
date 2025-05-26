Q:

1.這是我的clsReport
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

這是我的主執行序module

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
Public gRecIndex As Long                  ' RecordIndex 計數器

'=== UserForm 新增全域 allReportNames
Public allReportNames As Variant

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

    ThisWorkbook.Sheets("ControlPanel").Range("gDataMonthString").Value = "'" & gDataMonthString
    
    '轉換gDataMonthString為ROC Format
    gDataMonthStringROC = ConvertToROCFormat(gDataMonthString, "ROC")
    gDataMonthStringROC_NUM = ConvertToROCFormat(gDataMonthString, "NUM")
    gDataMonthStringROC_F1F2 = ConvertToROCFormat(gDataMonthString, "F1F2")
    ' 設定其他 config 參數（請根據實際環境調整）
    gDBPath = ThisWorkbook.Path & "\" & ThisWorkbook.Sheets("ControlPanel").Range("DBsPathFileName").Value
    ' gDBPath = "\\10.10.122.40\後台作業\99_個人資料夾\8.修豪\DbsMReport20250513_V1\" & ThisWorkbook.Sheets("ControlPanel").Range("DBsPathFileName").Value
    ' 空白報表路徑
    gReportFolder = ThisWorkbook.Path & "\" & ThisWorkbook.Sheets("ControlPanel").Range("EmptyReportPath").Value
    ' 產生之申報報表路徑
    gOutputFolder = ThisWorkbook.Path & "\" & ThisWorkbook.Sheets("ControlPanel").Range("OutputReportPath").Value

    ' ========== 宣告所有報表 ==========
    ' 製作報表List
    ' gReportNames 少FB1 FM5
    allReportNames = Array("TABLE10", "TABLE15A", "TABLE15B", "TABLE16", "TABLE20", "TABLE22", "TABLE23", "TABLE24", "TABLE27", "TABLE36", "AI233", "AI345", "AI405", "AI410", "AI430", "AI601", "AI605")

    ' =====testArray=====
    ' allReportNames = Array("AI822")

    ' ========== 選擇產生全部或部分報表 ==========
    Dim respRunAll As VbMsgBoxResult
    Dim userInput As String
    Dim i As Integer, j As Integer
    respRunAll = MsgBox("要執行全部報表嗎？" & vbCrLf & _
                  "【是】→ 全部報表" & vbCrLf & _
                  "【否】→ 指定報表", _
                  vbQuestion + vbYesNo, "選擇產生全部或部分報表")    
    If respRunAll = vbYes Then
        gReportNames = allReportNames
        For i = LBound(gReportNames) To UBound(gReportNames)
            gReportNames(i) = UCase(gReportNames(i))
        Next i
    Else
        ' UserForm 勾選清單
        Dim frm As ReportSelector
        Set frm = New ReportSelector
        frm.Show vbModal
        ' 若 gReportNames 未被填（使用者未選任何項目），則中止
        If Not IsArray(gReportNames) Or UBound(gReportNames) < 0 Then
            MsgBox "未選擇任何報表，程序結束", vbInformation
            Exit Sub
        End If
        ' 轉大寫（保留原邏輯）
        For i = LBound(gReportNames) To UBound(gReportNames)
            gReportNames(i) = UCase(gReportNames(i))
        Next i
    End If
    
    ' 檢查不符合的報表名稱
    Dim invalidReports As String
    Dim found As Boolean

    For i = LBound(gReportNames) To UBound(gReportNames)
        found = False
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
        MsgBox "報表名稱錯誤，請重新確認：" & vbCrLf & invalidReports, vbCritical, "報表名稱錯誤"
        WriteLog "報表名稱錯誤，請重新確認：" & vbCrLf & invalidReports
        Exit Sub
    End If
    
    ' ========== 處理其他部門提供數據欄位 ==========
    ' 定義每張報表必需由使用者填入／確認的儲存格名稱
    Dim req As Object
    Set req = CreateObject("Scripting.Dictionary")
    ' req.Add "TABLE41", Array("Table41_國外部_一利息收入", _
    '                          "Table41_國外部_一利息收入_利息", _
    '                          "Table41_國外部_一利息收入_利息_存放銀行同業", _
    '                          "Table41_國外部_二金融服務收入", _
    '                          "Table41_國外部_一利息支出", _
    '                          "Table41_國外部_一利息支出_利息", _
    '                          "Table41_國外部_一利息支出_利息_外國人外匯存款", _
    '                          "Table41_國外部_二金融服務支出", _
    '                          "Table41_企銷處_一利息支出", _
    '                          "Table41_企銷處_一利息支出_利息", _
    '                          "Table41_企銷處_一利息支出_利息_外國人新台幣存款")
                            
    ' req.Add "AI822", Array("AI822_會計科_上年度決算後淨值", _
    '                        "AI822_國外部_直接往來之授信", _
    '                        "AI822_國外部_間接往來之授信", _
    '                        "AI822_授管處_直接往來之授信")

    ' 暫存要移除的報表
    Dim toRemove As Collection
    Set toRemove = New Collection

    ' 逐一詢問使用者每張報表、每個必要欄位的值
    Dim ws As Worksheet
    Dim rptName As Variant 
    Dim fields As Variant, fld As Variant
    Dim defaultVal As Variant, userVal As String
    Dim respToContinue As VbMsgBoxResult
    Dim respHasInput As VbMsgBoxResult

    For Each rptName In gReportNames
        If req.Exists(rptName) Then
            Set ws = ThisWorkbook.Sheets(rptName)
            fields = req(rptName)

            ' --- 新增：先問一次是否已自行填入該報表所有資料 ---
            respHasInput = MsgBox( _
                "是否已填入 " & rptName & " 報表資料？", _
                vbQuestion + vbYesNo, "確認是否填入資料")
            If respHasInput = vbYes Then
                ' --- 已填入：只檢查「空白」的必要欄位 ---
                For Each fld In fields
                    If Trim(CStr(ws.Range(fld).Value)) = "" Then
                        defaultVal = 0
                        userVal = InputBox( _
                            "報表 " & rptName & " 的欄位 [" & fld & "] 尚未輸入，請填入數值：", _
                            "請填入必要欄位", "")
                        If userVal = "" Then
                            respToContinue = MsgBox("未輸入任何數值，是否仍要製作報表 " & rptName & "？", _
                                         vbQuestion + vbYesNo, "繼續製作？")
                            If respToContinue = vbYes Then
                                ws.Range(fld).Value = 0
                            Else
                                toRemove.Add rptName
                                Exit For
                            End If
                        ElseIf IsNumeric(userVal) Then
                            ws.Range(fld).Value = CDbl(userVal)
                        Else
                            ws.Range(fld).Value = 0
                            MsgBox "您輸入的不是數字，將保留為 0", vbExclamation
                            WriteLog "您輸入的不是數字，將保留為 0"
                        End If
                    End If
                Next fld
            Else
                For Each fld In fields
                    defaultVal = ws.Range(fld).Value
                    userVal = InputBox( _
                        "請確認報表 " & rptName & " 的 [" & fld & "]" & vbCrLf & _
                        "目前值：" & defaultVal & vbCrLf & _
                        "若要修改，請輸入新數值；若已更改，請直接點擊「確定」。", _
                        "欄位值", CStr(defaultVal) _
                    )
                    If userVal = "" Then
                        ' 空白表示使用者沒有輸入
                        respToContinue = MsgBox("未輸入任何數值，是否仍要製作報表 " & rptName & "？", _
                                    vbQuestion + vbYesNo, "繼續製作？")
                        If respToContinue = vbYes Then
                            If IsNumeric(defaultVal) Then
                                ws.Range(fld).Value = CDbl(defaultVal)
                            Else
                                ws.Range(fld).Value = 0
                            End If
                        Else
                            toRemove.Add rptName
                            Exit For   ' 跳出該報表的欄位迴圈
                        End If
                    ElseIf IsNumeric(userVal) Then
                        ws.Range(fld).Value = CDbl(userVal)
                    Else
                        If IsNumeric(defaultVal) Then
                            ws.Range(fld).Value = CDbl(defaultVal)
                        Else
                            ws.Range(fld).Value = 0
                        End If
                        MsgBox "您輸入的不是數字，將保留原值：" & defaultVal, vbExclamation
                        WriteLog "您輸入的不是數字，將保留原值：" & defaultVal
                    End If
                Next fld
            End If
        End If
    Next rptName

    ' 把使用者取消的報表，從 gReportNames 中移除
    If toRemove.Count > 0 Then
        Dim tmpArr As Variant
        Dim idx As Long
        Dim keep As Boolean
        Dim name As Variant

        tmpArr = gReportNames
        ReDim gReportNames(0 To UBound(tmpArr) - toRemove.Count)
    
        idx = 0    
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
        If idx = 0 Then
            MsgBox "所有報表均取消，程序結束", vbInformation
            WriteLog "所有報表均取消，程序結束", vbInformation
            Exit Sub
        End If
    End If

    ' ========== 取得第幾次寫入資料庫年月資料之RecordIndex ==========
    gRecIndex = GetMaxRecordIndex(gDBPath, "MonthlyDeclarationReport", gDataMonthString) + 1

    ' ========== 報表初始化 ==========
    ' Process A: 初始化所有報表，將初始資料寫入 Access DB with Null Data
    Call InitializeReports
    ' MsgBox "完成 Process A"
    WriteLog "完成 Process A"
    
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
            Case "AI822":   Call Process_AI822
            ' 
            Case "TABLE10":    Call Process_TABLE10
            Case "TABLE15A":    Call Process_TABLE15A
            Case "TABLE15B":    Call Process_TABLE15B
            Case "TABLE16":    Call Process_TABLE16
            Case "TABLE20":    Call Process_TABLE20
            Case "TABLE22":    Call Process_TABLE22
            Case "TABLE23":    Call Process_TABLE23
            Case "TABLE24":    Call Process_TABLE24
            Case "TABLE27":    Call Process_TABLE27
            Case "TABLE36":    Call Process_TABLE36
            Case "AI233":    Call Process_AI233
            Case "AI345":    Call Process_AI345
            Case "AI405":    Call Process_AI405
            Case "AI410":    Call Process_AI410
            Case "AI430":    Call Process_AI430
            Case "AI601":    Call Process_AI601
            Case "AI605":    Call Process_AI605
            Case Else
                MsgBox "未知的報表名稱: " & rptName, vbExclamation
                WriteLog "未知的報表名稱: " & rptName
        End Select
    Next rptName    
    ' MsgBox "完成 Process B"
    WriteLog "完成 Process B"

    ' ========== 產生新報表 ==========
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
    ' MsgBox "完成'報表初始欄位資訊儲存'及'初始資料庫資料建立'"
    WriteLog "完成'報表初始欄位資訊儲存'及'初始資料庫資料建立'"
End Sub

2.以上是我現在的規劃，我在clsReport中的Init定義好要從Excel抓取的Range位置，預先儲存在 rptFieldInfo，例如Table10
.Add "TABLE10", _
Array(
    Array("FOA", "S2:S21,W2:W21,AA2:AA21,AE2:AE21,AI2:AI21,AM2:AM21,AQ2:AQ21", "U2:U21,Y2:Y21,AC2:AC21,AG2:AG21,AK2:AK21,AO2:AO21,AS2:AS21")            
)
要將相關數據從
rptArray = Me.GetFieldFromXlRanges(reportName, nameTagRng, addrRng, initValue)
' 呼叫 AddWorksheetFields，第一參數用 initSheetName
AddWorksheetFields initSheetName, rptArray
去call AddWorksheetFields，
其中rptArray 是像這樣的構造 Array( _
        Array("FB2_申報時間", "D2", dataMonthStringROC), _
        Array("FB2_存放及拆借同業", "F9", Null), _
        Array("FB2_拆放銀行同業", "F13", Null), _
        Array("FB2_應收款項_淨額", "F36", Null), _
        Array("FB2_應收利息", "F41", Null), _
        Array("FB2_資產總計", "F85", Null) )

2.以上是我原本已經設計的規劃，但我現在打算加入以下Aceess資料庫中的資料表欄位資料如下，其中Report中的ReportName就是1.中的rptName，TargetSheetName就是上面例子中的"TABLE10"，SourceNameTag就是上面中的"FB2_申報時間" "FB2_存放及拆借同業" "FB2_拆放銀行同業"...等，TargetCellAddress就是上面的"D2" "F9" "F13" "F36" "F41"...
，我打算不用rptFieldInfo，改成將相關資料登錄在access中，Excel直接抓取access中的資料，請問我要怎麼修改代碼，以及有什麼其他需要修改的地方也一併告訴我，
請不要修改我的主要架構，我直覺是需要先建立一個從access抓取資料的函數，然後從clsReport中取得資料後修改init中的相關程序，將從Access中提取的資料取代我原本定義rptFieldInfo提取的資料，請將函數寫的有彈性一點，才可以再更多地方可以共同使用，且函數可以直接回傳比較具體的資料

Table: Report
  • ReportID          (Autonumber, PK)
  • ReportName        (Short Text)
  • ReportType        (Short Text or Lookup)

Table: PositionMap
  • DataId     (Autonumber, PK)
  • ReportID          (Number, FK → Report.ReportID)
  • SourceSheetName   (Short Text)
  • SourceTitle       (Short Text)
  • SourceNameTag     (Short Text)
  • SourceCellAddress (Short Text)
  • TargetSheetName   (Short Text)
  • TargetCellAddress (Short Text)
  • DataType          (Short Text)

Table: QueryMap
  • DataId        (Autonumber, PK)
  • ReportID          (Number, FK → Report.ReportID)
  • QueryTableName         (Number, FK → Report.ReportID)
  • ImportColName  (Short Text)
  • ImportColNumber  (Number)


以上是我一部分的問題，請不要回答我，我會在下一個commnad繼續貼上我的問題，請接收到全部的資訊後再一次回覆我


延續上面的問題
3.以下是其中的一個Process Report，其中2.中QueryMap的資料表就是用來儲存Query資料以及紀錄相關要存放資料的欄位設定數據，
請不要更改GetAccessDataAsArray函數，在現有架構基礎下幫我將QueryTable名稱，及access抓取資料要放的欄位等數據都從資料表中抓取，
例如ImportColName就紀錄要置放哪些連續欄位，ImportColNumber紀錄欄位數有多少，請幫我修改

Public Sub Process_FM11()
    '=== Equal Setting ===
    'Fetch Query Access DB table
    Dim dataArr_1 As Variant
    Dim dataArr_2 As Variant

    'Declare worksheet and handle data
    Dim xlsht As Worksheet

    Dim i As Integer, j As Integer
    Dim lastRow As Integer

    Dim reportTitle As String
    Dim queryTable_1 As String
    Dim queryTable_2 As String

    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("FM11")
    
    reportTitle = rpt.ReportName
    queryTable_1 = "FM11_OBU_AC5411B"
    queryTable_2 = "FM11_OBU_AC5411B_Subtotal"

    ' dataArr = GetAccessDataAsArray(gDBPath, queryTable, gDataMonthString)
    dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1, gDataMonthString)
    dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2, gDataMonthString)

    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    xlsht.Range("A:G").ClearContents
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
                xlsht.Cells(i + 1, j + 6).Value = dataArr_2(i, j)
            Next i
        Next j
    End If

    '--------------
    'Unique Setting
    '--------------
    Dim rngs As Range
    Dim rng As Range

    Dim foreignInterestRevenue As Double
    Dim gainOnSecuritiesDisposal As Double
    Dim lossOnSecuritiesDisposal As Double
    Dim reversalImpairmentPL As Double
    Dim valuationImpairmentLoss As Double
    Dim domesticInterestRevenue As Double

    foreignInterestRevenue = 0
    gainOnSecuritiesDisposal = 0
    lossOnSecuritiesDisposal = 0
    reversalImpairmentPL = 0
    valuationImpairmentLoss = 0
    domesticInterestRevenue = 0
    
    lastRow = xlsht.Cells(xlsht.Rows.Count, "F").End(xlUp).Row
    Set rngs = xlsht.Range("F2:F" & lastRow)

    For Each rng In rngs
        If CStr(rng.Value) = "InterestRevenue" Then
            foreignInterestRevenue = foreignInterestRevenue + rng.Offset(0, 1).Value
        
        ElseIf CStr(rng.Value) = "GainOnDisposal" Then
            gainOnSecuritiesDisposal = gainOnSecuritiesDisposal + rng.Offset(0, 1).Value

        ElseIf CStr(rng.Value) = "LossOnDisposal" Then
            lossOnSecuritiesDisposal = lossOnSecuritiesDisposal + rng.Offset(0, 1).Value

        ElseIf CStr(rng.Value) = "ValuationProfit" Then
            reversalImpairmentPL = reversalImpairmentPL + rng.Offset(0, 1).Value

        ' FVPL 金融資產評價損失
        ElseIf CStr(rng.Value) = "ValuationLoss" Then
            valuationImpairmentLoss = valuationImpairmentLoss + rng.Offset(0, 1).Value

        ' 拆放證券公司息 OSU
        ElseIf CStr(rng.Value) = "OSU息" Then
            domesticInterestRevenue = domesticInterestRevenue + rng.Offset(0, 1).Value
        End If
    Next rng

    foreignInterestRevenue = Round(foreignInterestRevenue / 1000, 0)
    gainOnSecuritiesDisposal = Round(gainOnSecuritiesDisposal / 1000, 0)
    lossOnSecuritiesDisposal = Round(lossOnSecuritiesDisposal / 1000, 0)
    reversalImpairmentPL = Round(reversalImpairmentPL / 1000, 0)
    valuationImpairmentLoss = Round(valuationImpairmentLoss / 1000, 0)
    domesticInterestRevenue = Round(domesticInterestRevenue / 1000, 0)
    
    xlsht.Range("FM11_一利息股息收入_利息_其他").Value = foreignInterestRevenue
    rpt.SetField "FOA", "FM11_一利息股息收入_利息_其他", CStr(foreignInterestRevenue)

    xlsht.Range("FM11_三證券投資處分利益_一年期以上之債權證券").Value = gainOnSecuritiesDisposal
    rpt.SetField "FOA", "FM11_三證券投資處分利益_一年期以上之債權證券", CStr(gainOnSecuritiesDisposal)

    xlsht.Range("FM11_三證券投資處分損失_一年期以上之債權證券").Value = lossOnSecuritiesDisposal
    rpt.SetField "FOA", "FM11_三證券投資處分損失_一年期以上之債權證券", CStr(lossOnSecuritiesDisposal)

    xlsht.Range("FM11_五證券投資評價及減損迴轉利益_一年期以上之債權證券").Value = reversalImpairmentPL
    rpt.SetField "FOA", "FM11_五證券投資評價及減損迴轉利益_一年期以上之債權證券", CStr(reversalImpairmentPL)

    xlsht.Range("FM11_五證券投資評價及減損損失_一年期以上之債權證券").Value = valuationImpairmentLoss
    rpt.SetField "FOA", "FM11_五證券投資評價及減損損失_一年期以上之債權證券", CStr(valuationImpairmentLoss)

    xlsht.Range("FM11_一利息收入_自中華民國境內其他客戶").Value = domesticInterestRevenue
    rpt.SetField "FOA", "FM11_一利息收入_自中華民國境內其他客戶", CStr(domesticInterestRevenue)
    
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
    ' 更改分頁顏色為黃色(6)
    xlsht.Tab.ColorIndex = 6
End Sub

其中
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


4.最後我想將 PositionMap 及 QueryMap中的資料，設定一個可以更改相關設定的機制，
例如我需要在Excel中有下拉式的選單，可能有很多階層的選單，先是select 報表名稱，
之後是該報表在資料中包含的NameTag有哪些，以及Range是什麼等資訊，
並且我需要一個可以填入數據的地方，去更改那些數據，
更改之後同時要更新資料庫的資料，也要更新excel欄位NameTag的相關資訊，請告訴我要怎麼建置

5.以上是我第一部分的問題，請考量全部的資訊後回覆我
以上的問題請不用一次回答我，可以分批仔細的回覆我的問題，我會跟你說請繼續


A:

好的，我们先从第 1 点着手，分批深入讨论。

---

## 1. 用 Access 的 PositionMap 取代 `rptFieldInfo`

### 核心思路

1. **新增一个通用的函数** `GetPositionMapData`

   * 负责从 Access 的 `PositionMap` 表中，根据 `ReportName` 读取所有对应的资料：

     * `TargetSheetName`
     * `SourceNameTag`
     * `TargetCellAddress`
   * 回传一个二维 `Variant` 陣列或 `Collection`，格式与 `rptFieldInfo(reportName)` 一致，方便后续共用。

2. **修改 `clsReport.Init`**

   * 移除原先 hard-code 的 `rptFieldInfo` 定义。
   * 在初始化之初调用 `GetPositionMapData reportName` 得到所有映射，然后对可用记录依序：

     ```vb
     For Each mapping In mappings
         AddWorksheetFields mapping.TargetSheetName, Array( _
             Array(mapping.SourceNameTag, mapping.TargetCellAddress, Null) _
         )
     Next
     ```
   * 后续写入 Access DB、加上动态字段等流程都无需更动。

### 示例代码

```vb
' === 模块：modDataAccess ===
' 通用：从 PositionMap 取出指定报表的所有映射
Public Function GetPositionMapData( _
        ByVal DBPath As String, _
        ByVal reportName As String _
    ) As Variant
    Dim conn As Object, rs As Object
    Dim sql As String
    Dim results() As Variant
    Dim i As Long
    
    ' 1. 建立连接
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath
    
    ' 2. SQL: 先 INNER JOIN Report 找到 ReportID，再取 PositionMap
    sql = "SELECT pm.TargetSheetName, pm.SourceNameTag, pm.TargetCellAddress " & _
          "FROM PositionMap AS pm " & _
          "INNER JOIN Report AS r ON pm.ReportID = r.ReportID " & _
          "WHERE r.ReportName = '" & reportName & "' " & _
          "ORDER BY pm.DataId;"
    
    Set rs = conn.Execute(sql)
    If rs.EOF Then
        GetPositionMapData = Array()  ' 没有记录
    Else
        ' 3. 把结果装进二维数组：每行一笔 mapping
        rs.MoveLast: rs.MoveFirst
        ReDim results(0 To rs.RecordCount - 1, 0 To 2)
        i = 0
        Do Until rs.EOF
            results(i, 0) = rs.Fields("TargetSheetName").Value
            results(i, 1) = rs.Fields("SourceNameTag").Value
            results(i, 2) = rs.Fields("TargetCellAddress").Value
            i = i + 1
            rs.MoveNext
        Loop
        GetPositionMapData = results
    End If
    
    rs.Close: conn.Close
    Set rs = Nothing: Set conn = Nothing
End Function
```

```vb
' === 类别：clsReport 中的 Init 调整 ===
Public Sub Init(ByVal reportName As String, _
                ByVal dataMonthStringROC As String, _
                ByVal dataMonthStringROC_NUM As String, _
                ByVal dataMonthStringROC_F1F2 As String)
    
    Dim mappings As Variant
    Dim i As Long
    Dim sheetName As String, nameTag As String, addr As String
    
    clsReportName = reportName
    Set clsWorksheets = CreateObject("Scripting.Dictionary")
    
    ' —— 1. 从 Access 读出该报表所有映射定义 —— 
    mappings = GetPositionMapData(gDBPath, reportName)
    
    If IsArray(mappings) And UBound(mappings) >= 0 Then
        For i = 0 To UBound(mappings, 1)
            sheetName = mappings(i, 0)
            nameTag   = mappings(i, 1)
            addr      = mappings(i, 2)
            
            ' 跳过空定义
            If Len(Trim(nameTag)) > 0 And Len(Trim(addr)) > 0 Then
                ' 每笔只包含单一 cell，可自行扩展支持多 cell
                AddWorksheetFields sheetName, Array( _
                    Array(nameTag, addr, Null) _
                )
            End If
        Next i
    Else
        WriteLog "Init 未在 PositionMap 找到报表：" & reportName
    End If
    
    ' —— 2. 保持原有的动态字段逻辑不变 —— 
    Select Case reportName
        Case "TABLE10"
            AddDynamicField reportName, "TABLE10_申報時間", "D2", dataMonthStringROC
        ' … 其余 Case 不变
    End Select
End Sub
```

### 说明

* `GetPositionMapData` 足够通用，后续若要改取 `QueryMap`、或支持多目标 cell，只要改此函数即能复用。
* `Init` 中只把 PositionMap 的资料带入 `AddWorksheetFields`，其它流程如写 Access、`ValidateFields`、`ApplyToWorkbook` 均可保留。

---

以上是第 1 部分的解决方案，若理解 OK 或有细节要讨论，我们再进到第 2 部分！请告诉我继续哪个问题。


' ****************************
' ****************************
以上第一部分修改

非常好，感謝你提供原本的用法。你使用 `AddWorksheetFields` 的方式是一次傳入：

```vb
AddWorksheetFields "工作表名稱", Array( _
    Array("NameTag1", "儲存格位置1", 預設值1), _
    Array("NameTag2", "儲存格位置2", 預設值2), _
    ...
)
```

而我上面寫的版本是多次呼叫 `AddWorksheetFields`，每次只有一組欄位：

```vb
AddWorksheetFields sheetName, Array( _
    Array(nameTag, addr, Null) _
)
```

### ✅ 相容性檢查與調整

其實你的一次傳多組是完全正確且更有效率的方式，我會根據這個做法修改，讓我們的程式邏輯可以**從 PositionMap 抓多筆資料**，然後**一次傳入 AddWorksheetFields**，這樣才和你現有的架構一致。

---

## 🔧 修正後的完整版本

### 📍**修改 1：`GetPositionMapData` 函數**（**保持不變**）

這部分我們上面提供的程式可以直接使用，回傳一個 `Variant` 的 2D 陣列，格式為：

```vb
results(i, 0) = TargetSheetName  
results(i, 1) = SourceNameTag  
results(i, 2) = TargetCellAddress  
```

這部分不用修改。

---

### 📍**修改 2：clsReport.Init 函數中「載入 PositionMap 設定」的邏輯**

改成將**同一工作表的所有欄位先蒐集起來，最後再用一次 `AddWorksheetFields` 傳入。**

#### ✅ 修正後程式：

```vb
Public Sub Init(ByVal reportName As String, _
                ByVal dataMonthStringROC As String, _
                ByVal dataMonthStringROC_NUM As String, _
                ByVal dataMonthStringROC_F1F2 As String)
    
    Dim mappings As Variant
    Dim i As Long
    Dim sheetName As String, nameTag As String, addr As String
    
    clsReportName = reportName
    Set clsWorksheets = CreateObject("Scripting.Dictionary")

    ' === 從 Access 讀取 PositionMap 資料 ===
    mappings = GetPositionMapData(gDBPath, reportName)

    Dim wsFields As Object
    Set wsFields = CreateObject("Scripting.Dictionary")

    If IsArray(mappings) And UBound(mappings) >= 0 Then
        For i = 0 To UBound(mappings, 1)
            sheetName = mappings(i, 0)
            nameTag = mappings(i, 1)
            addr = mappings(i, 2)

            If Len(Trim(nameTag)) > 0 And Len(Trim(addr)) > 0 Then
                If Not wsFields.exists(sheetName) Then
                    wsFields.Add sheetName, Array()
                End If

                ' 先取出目前該 sheet 的陣列
                Dim tmpList As Variant
                tmpList = wsFields(sheetName)

                ' 將新欄位加入陣列
                ReDim Preserve tmpList(0 To UBound(tmpList) + 1)
                tmpList(UBound(tmpList)) = Array(nameTag, addr, Null)

                ' 寫回 dictionary
                wsFields(sheetName) = tmpList
            End If
        Next i

        ' 最後，依每個工作表將欄位群一次性加到報表設定中
        Dim key As Variant
        For Each key In wsFields.Keys
            AddWorksheetFields key, wsFields(key)
        Next key

    Else
        WriteLog "Init 未在 PositionMap 找到報表：" & reportName
    End If

    ' === 原有的動態欄位設定邏輯保留 ===
    Select Case reportName
        Case "TABLE10"
            AddDynamicField reportName, "TABLE10_申報時間", "D2", dataMonthStringROC
        Case "TABLE11"
            AddDynamicField reportName, "TABLE11_申報時間", "D2", dataMonthStringROC
        ' ... 其他報表
    End Select
End Sub
```

---

## ✅ 這次修改項目總整理

| 區塊                   | 修改內容                                                  | 位置                           |
| -------------------- | ----------------------------------------------------- | ---------------------------- |
| `GetPositionMapData` | 保留不動，回傳 `2D Array` 結構                                 | `modDataAccess`              |
| `clsReport.Init`     | ✅ 改成先蒐集所有欄位後，一次性呼叫 `AddWorksheetFields`<br>✅ 保留動態欄位設定 | `clsReport.cls` 中的 `Init` 函數 |

---

接下來如果你確認這段 OK，我們可以處理第 2 部分：「Process\_FM11 中如何從 QueryMap 撈取查詢名稱與 Excel 放置欄位」。

請告訴我是否繼續 👍





' ****************************
' ****************************
' =========================


接下來我們針對「第 2、3 點：如何從 QueryMap 讀取 QueryTableName／ImportColName／ImportColNumber，並改寫 `Process_FM11`」做說明。

---

## 一、在 VBA 中新增──通用函數 `GetQueryMapData`

先在一個標準模組（如 `modDataAccess`）裡，加入一支跟 `GetPositionMapData` 類似的函式，用來讀取 `QueryMap` 中該報表所有的 Query 設定：

```vb
' === 模組：modDataAccess ===

' 取出指定報表的所有 QueryMap 設定：
' 回傳格式：二维陣列 (0 To N-1, 0 To 2)，
'   [i,0] = QueryTableName
'   [i,1] = ImportColName
'   [i,2] = ImportColNumber
Public Function GetQueryMapData( _
        ByVal DBPath As String, _
        ByVal reportName As String _
    ) As Variant

    Dim conn As Object, rs As Object
    Dim sql  As String
    Dim results() As Variant
    Dim i As Long
    
    ' 1. 建立 ADODB 連線
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath
    
    ' 2. JOIN Report 取出該報表的 QueryMap
    sql = "SELECT qm.QueryTableName, qm.ImportColName, qm.ImportColNumber " & _
          "FROM QueryMap AS qm " & _
          "INNER JOIN Report AS r ON qm.ReportID = r.ReportID " & _
          "WHERE r.ReportName = '" & reportName & "' " & _
          "ORDER BY qm.DataId;"
    
    Set rs = conn.Execute(sql)
    
    If rs.EOF Then
        GetQueryMapData = Array()    ' 無設定
    Else
        rs.MoveLast: rs.MoveFirst
        ReDim results(0 To rs.RecordCount - 1, 0 To 2)
        i = 0
        Do Until rs.EOF
            results(i, 0) = rs!QueryTableName
            results(i, 1) = rs!ImportColName
            results(i, 2) = rs!ImportColNumber
            i = i + 1
            rs.MoveNext
        Loop
        GetQueryMapData = results
    End If
    
    rs.Close:  conn.Close
    Set rs = Nothing:  Set conn = Nothing
End Function
```

這支函式同樣回傳一個 `(N × 3)` 的 `Variant` 陣列，讓你可以在任何需要讀取 QueryMap 的流程裡重複使用。

---

## 二、改寫 `Process_FM11`，動態抓取 QueryMap

把原本硬寫的兩個 `queryTable_1/queryTable_2`，改成一律走 `GetQueryMapData` 回來的清單，然後依序呼叫 `GetAccessDataAsArray`，再把回傳的資料放到對應的起始欄位。以下示範如何改：

```vb
Public Sub Process_FM11()
    Dim rpt As clsReport
    Set rpt = gReports("FM11")
    
    Dim mappings As Variant
    mappings = GetQueryMapData(gDBPath, rpt.ReportName)
    
    Dim xlsht As Worksheet
    Set xlsht = ThisWorkbook.Sheets(rpt.ReportName)
    
    Dim i As Long, j As Long
    Dim startCol As Long
    Dim arr As Variant
    
    ' 1. 清空該報表上的所有區塊（依需要調整）
    xlsht.Cells.ClearContents
    
    ' 2. 依照 QueryMap 設定，動態讀取並貼上
    If Not IsArray(mappings) Or UBound(mappings) < 0 Then
        MsgBox "QueryMap 無任何設定：" & rpt.ReportName, vbExclamation
        Exit Sub
    End If
    
    For i = 0 To UBound(mappings, 1)
        Dim qtName   As String
        Dim importAt As String
        Dim nCols    As Long
        
        qtName   = mappings(i, 0)
        importAt = mappings(i, 1)       ' e.g. "A1" 或 "C2"
        nCols    = CLng(mappings(i, 2)) ' 幾個欄位
        
        ' 2.1 從 Access 拿資料
        arr = GetAccessDataAsArray(gDBPath, qtName, gDataMonthString)
        If Not IsArray(arr) Then
            MsgBox "資料有誤: " & rpt.ReportName & " | " & qtName, vbExclamation
            WriteLog "資料有誤: " & rpt.ReportName & " | " & qtName
            GoTo NextMapping
        End If
        
        ' 2.2 轉換 ImportColName ("C2") → 起始 row / col 座標
        startCol = Range(importAt).Column
        Dim startRow As Long
        startRow = Range(importAt).Row
        
        ' 2.3 將 arr (欄位數 = nCols) 貼進 Excel
        '     arr 的格式：第一維是列、第二維是欄；上界可動態抓
        For j = 0 To UBound(arr, 2)      ' 用第 2 維跑欄
            ' 確保不超出 nCols
            If j < nCols Then
                Dim r As Long, c As Long
                For r = 0 To UBound(arr, 1)
                    c = startCol + j
                    xlsht.Cells(startRow + r, c).Value = arr(r, j)
                Next r
            End If
        Next j
        
NextMapping:
    Next i
    
    
    ' 3.（以下保持原本的 Unique Setting／累加邏輯不動）…
    '    外匯營業收入、利息、損益彙總等運算
    '    最後再寫回 clsReport、ValidateFields、UpdateRecord…
    
    ' 範例：
    ' rpt.SetField "FOA", "FM11_一利息股息收入_利息_其他", CStr(foreignInterestRevenue)
    ' …
End Sub
```

### 重點說明

1. **彈性**：

   * 任何新加的 query 都只要在 `QueryMap` 填入一筆設定，就能自動被拿到並貼上，毋須改程式。

2. **ImportColName**：

   * 建議在 `QueryMap.ImportColName` 中存放「貼上起始儲存格地址」（例如 `"A1"`、`"C2"`），方便做 `Range(importAt).Column/Row` 轉換。
   * `ImportColNumber` 則限制一次最多貼幾個欄位。

3. **錯誤處理**：

   * 如某筆 Query 回傳空，會跳過並寫 Log，但不會完全中斷整個流程。

---

如以上說明有疑問，或要調整「貼上前要先清哪些範圍」等細節，再跟我說；接著我們可以再針對第 4 點：如何在 Excel 畫面上建置下拉選單、即時編輯並同步更新 Access 與 NameTag。




' ================


下面先就第 4 點──「在 Excel 介面上透過下拉與編輯機制，同時更新資料庫與 Excel NameTag」──提出一個可行的整體設計方案，包含畫面架構與核心 VBA 流程。你的主架構不必大幅修改，只要新增一張設定表或一個 UserForm，以及配套的存取／更新函式即可。

---

## 方案概覽

1. **在 Excel 增設「設定」工作表**（或使用 UserForm）

   * 左上角放一個「報表名稱」下拉清單
   * 右上角放一個「欄位名稱」下拉清單（依據所選報表動態載入）
   * 下方放一個清單格區塊，顯示該報表所有的 SourceNameTag、TargetSheetName、TargetCellAddress，並可直接在該表格修改
   * 下方再提供兩個按鈕：

     * **\[載入設定]**：讀取 `PositionMap` 資料並填入表格
     * **\[儲存變更]**：把使用者在表格的修改同步寫回 Access，並更新 Workbook.Names

2. **動態下拉清單**

   * 「報表名稱」：Data Validation 來源指向 `Report` 資料表的 ReportName 清單（可先用 `GetReportNames` 讀取並寫到隱藏範圍）
   * 當使用者在「報表名稱」 select 之後，Worksheet\_Change 事件觸發，呼叫 `GetPositionMapData`（或自訂函式）把這個報表的所有 Mapping 顯示在下方清單格，也更新「欄位名稱」的 Data Validation (依序列出 `SourceNameTag`)

3. **儲存變更流程**

   * 掃描下方表格中每一列：

     * 讀出 DataId（可先把 DataId 隱藏在某一欄）
     * 讀出最新的 TargetSheetName、SourceNameTag、TargetCellAddress
   * 對每筆變更執行 `UPDATE PositionMap SET … WHERE DataId = …`（寫一個 `UpdatePositionMapRecord(DataId, newSheet, newTag, newAddr)`）
   * 同時：若使用者修改了 NameTag (SourceNameTag) 或 TargetCellAddress，還要更新 Excel 中對應的 Name 物件：

     ```vb
     ThisWorkbook.Names(SourceNameTag).RefersTo = "=" & newSheet & "!" & newAddr
     ```

     （可包在同一個更新迴圈裡）

4. **必要的輔助 VBA 函式**

   * `GetAllReportNames()`：讀 `Report` 表回傳陣列, 填充「報表名稱」下拉
   * `LoadPositionMapToSheet(reportName As String)`：把 `PositionMap` 讀出後，寫入工作表範圍；同時隱藏 DataId
   * `UpdatePositionMapRecord(DataId As Long, sheetName As String, tag As String, addr As String)`：執行 ADODB UPDATE
   * `SyncExcelName(tag As String, sheetName As String, addr As String)`：更新 Workbook.Names

5. **UserForm 版（可選）**
   若偏好視窗操作，也可以改用一個 `UserForm`，其元件包括：

   * Combobox `cboReport`（報表）
   * ListBox `lstMappings`（多欄顯示：DataId、SourceTag、Sheet、Addr）
   * TextBox `txtSheet`、`txtTag`、`txtAddr`（編輯選中列）
   * CommandButton `btnLoad`、`btnSave`

   只要在 `cboReport_Change` 調用 `GetPositionMapData` 幫 `lstMappings` 賦值；在 `btnSave_Click` 掃回修改並呼叫上面那幾支更新函式即可。

---

### 範例：在「Config」工作表做 Data Validation + 按鈕

1. **準備名稱範圍**

   ```vb
   ' 模組：modConfig
   Public Sub SetupConfigValidation()
       Dim reports As Variant
       reports = GetAllReportNames(gDBPath)
       ' 假設把它寫到隱藏工作表 Hidden!A1:A{N}
       Worksheets("Hidden").Range("A1").Resize(UBound(reports)+1).Value = Application.Transpose(reports)
       ' 然後在 Config!B2 設定 Data Validation → List → =Hidden!$A$1:$A${最後一列}
       With Worksheets("Config").Range("B2").Validation
           .Delete
           .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                Operator:=xlBetween, Formula1:="=Hidden!$A$1:$A$" & UBound(reports)+1
       End With
   End Sub
   ```

2. **當 ReportName 變更時，載入 Mapping**

   ```vb
   ' 放在 Config 工作表的 code module
   Private Sub Worksheet_Change(ByVal Target As Range)
       If Target.Address = "$B$2" Then  ' ReportName 儲存格
           Dim rptName As String
           rptName = Target.Value
           If rptName <> "" Then LoadPositionMapToSheet rptName
       End If
   End Sub
   ```

3. **將 Mapping 寫到 Config!A5\:D100**

   ```vb
   Public Sub LoadPositionMapToSheet(ByVal reportName As String)
       Dim mappings As Variant
       mappings = GetPositionMapData(gDBPath, reportName)
       Dim ws As Worksheet: Set ws = Worksheets("Config")
       ws.Range("A5:D100").ClearContents
       Dim i As Long
       For i = 0 To UBound(mappings,1)
           ws.Cells(5+i, 1).Value = mappings(i,0)          ' TargetSheetName
           ws.Cells(5+i, 2).Value = mappings(i,1)          ' SourceNameTag
           ws.Cells(5+i, 3).Value = mappings(i,2)          ' TargetCellAddress
           ' DataId 可先讀出後隱藏在第 4 欄，供 Update 用
           ws.Cells(5+i, 4).Value = GetPositionMapData_Id(gDBPath, reportName, mappings(i,1))
       Next i
       ' 同時在欄位名稱下拉（例如 E2）設定 Data Validation，清單來源 = ws.Range("B5:B" & 5+UBound(mappings,1))
   End Sub
   ```

4. **儲存變更時更新 DB & Excel Names**

   ```vb
   Public Sub SaveConfigChanges()
       Dim ws As Worksheet: Set ws = Worksheets("Config")
       Dim lastRow As Long
       lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
       Dim i As Long
       For i = 5 To lastRow
           Dim dataId As Long
           Dim sht As String, tag As String, addr As String
           dataId = ws.Cells(i, 4).Value
           sht      = ws.Cells(i, 1).Value
           tag      = ws.Cells(i, 2).Value
           addr     = ws.Cells(i, 3).Value
           ' 1) 更新 Access
           UpdatePositionMapRecord dataId, sht, tag, addr
           ' 2) 更新 Excel Name
           SyncExcelName tag, sht, addr
       Next i
       MsgBox "設定已儲存並同步！", vbInformation
   End Sub
   ```

---

以上是一整套「表單＋Data Validation＋同步更新」的架構，核心在於：

1. **Data Access 函式**（`GetPositionMapData`／`GetQueryMapData`／`UpdatePositionMapRecord`／`GetAllReportNames`）
2. **畫面互動**（工作表 Change 事件或 UserForm）
3. **同步更新 Excel Names**（`ThisWorkbook.Names(…)`）

如需更詳細的程式碼範本（或改用 UserForm 實作），或想要進一步調整欄位顯示格式，請告訴我！



' ================


Q:
1.請給我更詳細的程式碼範本
2.也給我使用Userform實作的範本

A:
下面提供两套完整的 VBA 范本：

1. **直接在「Config」工作表上** 用 Data Validation + 按钮互动
2. **UserForm 实作**

> 所有代码均假设已引用 `modDataAccess` 模块内的通用存取与更新函数：
>
> ```vb
> Public Function GetPositionMapData(DBPath As String, reportName As String) As Variant
>     ' …（前述范例）…
> End Function
> Public Function GetQueryMapData(DBPath As String, reportName As String) As Variant
>     ' …（前述范例）…
> End Function
> Public Function GetAllReportNames(DBPath As String) As Variant
>     ' SELECT ReportName FROM Report…
> End Function
> Public Sub UpdatePositionMapRecord(DataId As Long, sheetName As String, tag As String, addr As String)
>     ' 使用 ADODB 执行 UPDATE PositionMap … WHERE DataId=…
> End Sub
> Public Sub SyncExcelName(tag As String, sheetName As String, addr As String)
>     ThisWorkbook.Names(tag).RefersTo = "=" & sheetName & "!" & addr
> End Sub
> ```

---

## 一、Config 工作表 + Data Validation + 按钮

### 1.1 隐藏页与名称范围准备（`modConfig` 模块）

```vb
Option Explicit

' 将所有 ReportName 写到 隐藏页 Hidden!A1:A{N}
Public Sub SetupReportNameList()
    Dim reports As Variant
    Dim wsH As Worksheet
    Dim i As Long
    
    reports = GetAllReportNames(gDBPath)  ' e.g. Array("TABLE10","TABLE15A",…)
    Set wsH = ThisWorkbook.Worksheets("Hidden")
    wsH.Cells.Clear
    
    For i = LBound(reports) To UBound(reports)
        wsH.Cells(i + 1, 1).Value = reports(i)
    Next i
    
    ' 建立名称范围：Name="ReportList", RefersTo=Hidden!$A$1:$A${最后一行}
    ThisWorkbook.Names.Add Name:="ReportList", _
        RefersTo:="=Hidden!$A$1:$A$" & UBound(reports) + 1
End Sub
```

### 1.2 Config 页上的 Data Validation & 事件

1. 在 `Config` 工作表里，

   * B2 放「报表名称」下拉，Data Validation → List → `=ReportList`
   * C2 放「栏位名称」下拉，留空（后面动态填）
   * A5\:C100 预留给映射表格：A=TargetSheetName, B=SourceNameTag, C=TargetCellAddress
   * 在 B2 右侧挂一个「载入设置」按钮（指向 `LoadPositionMapToSheet`）
   * 在 B3 挂一个「储存变更」按钮（指向 `SaveConfigChanges`）

2. 在 `Config` 的 Worksheet 代码区，添加：

```vb
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address = "$B$2" Then
        If Target.Value <> "" Then
            LoadPositionMapToSheet Target.Value
            ' 同步更新 C2 的下拉清单：=Config!$B$5:$B${5+N}
            Dim last As Long
            last = 5 + WorksheetFunction.CountA(Range("A5:A100")) - 1
            With Range("C2").Validation
                .Delete
                .Add Type:=xlValidateList, _
                     AlertStyle:=xlValidAlertStop, _
                     Operator:=xlBetween, _
                     Formula1:="=$B$5:$B$" & last
            End With
        End If
    End If
End Sub
```

### 1.3 `modConfig` 模块中的操作例程

```vb
Option Explicit

' 读取并显示 PositionMap 到 Config!A5:C...
Public Sub LoadPositionMapToSheet(reportName As String)
    Dim mappings As Variant
    Dim ws As Worksheet
    Dim i As Long
    
    mappings = GetPositionMapData(gDBPath, reportName)
    Set ws = ThisWorkbook.Worksheets("Config")
    ws.Range("A5:C100").ClearContents
    
    If Not IsArray(mappings) Then Exit Sub
    
    For i = 0 To UBound(mappings, 1)
        ws.Cells(5 + i, 1).Value = mappings(i, 0)  ' TargetSheetName
        ws.Cells(5 + i, 2).Value = mappings(i, 1)  ' SourceNameTag
        ws.Cells(5 + i, 3).Value = mappings(i, 2)  ' TargetCellAddress
        ' 如果需要 DataId，可额外从 DB 查出并放到 D 列
    Next i
End Sub

' 掃描 Config!A5:C...，更新到 Access & Excel Names
Public Sub SaveConfigChanges()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dataId As Long
    Dim sht As String, tag As String, addr As String
    
    Set ws = ThisWorkbook.Worksheets("Config")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For i = 5 To lastRow
        ' 假设 DataId 存放在 D 列：
        dataId = CLng(ws.Cells(i, 4).Value)
        sht    = ws.Cells(i, 1).Value
        tag    = ws.Cells(i, 2).Value
        addr   = ws.Cells(i, 3).Value
        
        ' 1) 更新 Access
        UpdatePositionMapRecord dataId, sht, tag, addr
        ' 2) 同步 Excel Name
        SyncExcelName tag, sht, addr
    Next i
    
    MsgBox "Config 已更新并同步完成", vbInformation
End Sub
```

---

## 二、UserForm 版：`frmConfig`

### 2.1 设计表单元件

| 控件名称        | 类型            | 说明                               |
| ----------- | ------------- | -------------------------------- |
| `cboReport` | ComboBox      | 列出所有 ReportName                  |
| `lstMap`    | ListBox       | 多列显示：DataId、SourceTag、Sheet、Addr |
| `txtSheet`  | TextBox       | 编辑用：TargetSheetName              |
| `txtTag`    | TextBox       | 编辑用：SourceNameTag                |
| `txtAddr`   | TextBox       | 编辑用：TargetCellAddress            |
| `btnLoad`   | CommandButton | 载入 Map                           |
| `btnSave`   | CommandButton | 存回 DB & 同步 Excel                 |

> 请在 Properties 中把 `lstMap` 的 `ColumnCount` 设为 4，`ColumnWidths` 设为 `"40;100;80;60"`（可调整）。

### 2.2 UserForm 代码

```vb
Option Explicit

Private mapData As Variant

' 表单初始化：填 Report 下拉
Private Sub UserForm_Initialize()
    Dim reports As Variant
    Dim i As Long
    reports = GetAllReportNames(gDBPath)
    Me.cboReport.Clear
    For i = LBound(reports) To UBound(reports)
        Me.cboReport.AddItem reports(i)
    Next i
End Sub

' 选定 Report
Private Sub cboReport_Change()
    Dim rptName As String
    rptName = Me.cboReport.Value
    If rptName <> "" Then
        btnLoad_Click
    End If
End Sub

' 载入 Mapping 到 ListBox
Private Sub btnLoad_Click()
    Dim i As Long
    Me.lstMap.Clear
    mapData = GetPositionMapData(gDBPath, Me.cboReport.Value)
    If Not IsArray(mapData) Then Exit Sub
    
    For i = 0 To UBound(mapData, 1)
        ' 假设第一行第0=Sheet，1=Tag，2=Addr；第3可写 DataId (若有)
        Me.lstMap.AddItem
        Me.lstMap.List(i, 0) = ""                ' DataId 列
        Me.lstMap.List(i, 1) = mapData(i, 1)     ' SourceTag
        Me.lstMap.List(i, 2) = mapData(i, 0)     ' Sheet
        Me.lstMap.List(i, 3) = mapData(i, 2)     ' Addr
    Next i
End Sub

' 当 ListBox 行被点击时，把值载到编辑框
Private Sub lstMap_Click()
    If Me.lstMap.ListIndex < 0 Then Exit Sub
    With Me.lstMap
        Me.txtSheet.Value = .List(.ListIndex, 2)
        Me.txtTag.Value   = .List(.ListIndex, 1)
        Me.txtAddr.Value  = .List(.ListIndex, 3)
    End With
End Sub

' 存回 DB & 同步 Excel 名称
Private Sub btnSave_Click()
    Dim idx As Long
    For idx = 0 To Me.lstMap.ListCount - 1
        Dim dataId  As Long
        Dim sht, tag, addr As String
        
        dataId = CLng(Me.lstMap.List(idx, 0))    ' 若未显示 DataId，可额外查
        tag    = Me.lstMap.List(idx, 1)
        sht    = Me.lstMap.List(idx, 2)
        addr   = Me.lstMap.List(idx, 3)
        
        UpdatePositionMapRecord dataId, sht, tag, addr
        SyncExcelName tag, sht, addr
    Next idx
    MsgBox "所有 Mapping 已存回！", vbInformation
End Sub
```

> **提示**：如果需要将 `DataId` 一并显示与编辑，可在 `GetPositionMapData` 查询时一并取出，再把它放到 `mapData(i,3)`，并将 `lstMap.ColumnCount=5`，对应填充。

---

这样，一套是直接在工作表上操作；另一套是使用 UserForm。你可以根据喜好，以及组织使用规模来选用。任何细节问题都欢迎再讨论！



' ====================

Q:
如果是使用UserForm的話會變得像是怎麼樣，請問有具體的圖像化可以讓我知道嗎?所以是不會在原本的Excel表格中嗎?

A:

上面這張圖就是 UserForm 的示意：

* **最上方** 是下拉式 `ComboBox` (`cboReport`)，讓您選擇要設定的報表。
* **中央大區域** 是多欄 `ListBox` (`lstMap`)，顯示目前載入的 Mapping（DataId、SourceTag、SheetName、CellAddress）。
* **下方三個 TextBox** (`txtSheet`、`txtTag`、`txtAddr`)，點選 ListBox 上的某列之後，就把該列值回填到這裡供您編輯。
* **最底部兩顆按鈕**，`Load`（載入 Mapping）、`Save`（儲存並同步更新 Access & Excel Name）。

因為它是獨立視窗 (UserForm)，不會直接出現在原本的工作表中。這樣可以避免干擾您現有的報表版面，也更專注於「設定」功能。您只要在 VBA IDE 中加入這個 UserForm，再把程式碼貼上，就能以視窗方式管理 PositionMap 了！
