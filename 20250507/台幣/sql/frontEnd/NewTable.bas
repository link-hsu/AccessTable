
Table: Report
  • ReportID          (Autonumber, PK)
  • ReportName        (Short Text)
  • ReportType        (Short Text or Lookup)

Table: PositionMap
  • DataId     (Autonumber, PK)
  • ReportID          (Number, FK → Report.ReportID)
  • SourceSheetName   (Short Text)
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














- ============================================================
Question1.
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
請不要修改我的主要架構，我直覺是需要先建立一個從access抓取資料的函數，然後從clsReport中取得資料後修改init中的相關程序，將從Access中提取的資料取代我原本定義rptFieldInfo提取的資料

Table: Report
  • ReportID          (Autonumber, PK)
  • ReportName        (Short Text)
  • ReportType        (Short Text or Lookup)

Table: PositionMap
  • DataId     (Autonumber, PK)
  • ReportID          (Number, FK → Report.ReportID)
  • SourceSheetName   (Short Text)
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

- ===================================================
Answer1.
下面示範一種做法，保留你現在 `clsReport`／`Main` 的流程架構（尤其是 `Init → AddWorksheetFields → ApplyToWorkbook` 等），只是把原本硬寫在 `Init` 裡的 `rptFieldInfo` 全部改成「從 Access 撈 `PositionMap` 表」來動態帶入。

---

## 1. 建立一支 Access 取資料的共用函式

先在任何標準 Module（例如 `modDB`）裡，加上這兩個函式：

```vba
'--- 在 VBA Editor 裡：工具→參考→打勾 Microsoft ActiveX Data Objects x.x Library
Option Explicit

' 回傳 ADODB.Recordset：篩選出指定 ReportName 的定位設定
Public Function GetPositionMapRS( _
    ByVal dbPath As String, _
    ByVal reportName As String _
) As ADODB.Recordset
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sql As String
    
    sql = "SELECT SourceSheetName, SourceNameTag, SourceCellAddress, " & _
          "TargetSheetName, TargetCellAddress, DataType " & _
          "FROM PositionMap AS P " & _
          "INNER JOIN Report AS R ON P.ReportID = R.ReportID " & _
          "WHERE R.ReportName = '" & Replace(reportName, "'", "''") & "' " & _
          "ORDER BY P.DataId"
    
    Set cn = New ADODB.Connection
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"
    
    Set rs = New ADODB.Recordset
    rs.Open sql, cn, adOpenForwardOnly, adLockReadOnly
    
    Set GetPositionMapRS = rs
End Function
```

> **說明**：
>
> * `dbPath` 就傳 `gDBPath`，針對你 `PositionMap` 表跑一次 `SELECT`。
> * 回傳一個 `Recordset`，每一筆就是一個欄位的來源／目標設定。

---

## 2. 修改 `clsReport.Init`：改用這支函式動態填 `clsWorksheets`

把原本那段巨長的硬寫 `.Add "TABLE10", Array(...)` … 全部拿掉，改成：

```vba
Public Sub Init(ByVal reportName As String, _
                ByVal dataMonthStringROC As String, _
                ByVal dataMonthStringROC_NUM As String, _
                ByVal dataMonthStringROC_F1F2 As String)

    Dim rs As ADODB.Recordset
    Dim mappings As Object
    Dim sheetName As String
    Dim arrFields As Collection
    Dim fldDef As Variant
    
    clsReportName = reportName
    Set clsWorksheets = CreateObject("Scripting.Dictionary")
    
    ' 1) 從 Access 撈出所有該報表的定位
    Set rs = GetPositionMapRS(gDBPath, reportName)
    
    ' 2) 用 Dictionary 先以 SourceSheetName 分組
    Set mappings = CreateObject("Scripting.Dictionary")
    Do While Not rs.EOF
        sheetName = rs!SourceSheetName
        If Not mappings.Exists(sheetName) Then
            Set arrFields = New Collection
            mappings.Add sheetName, arrFields
        Else
            Set arrFields = mappings(sheetName)
        End If
        
        ' 每筆 record 組成一個三元素陣列：NameTag, SourceCell, TargetCell
        fldDef = Array( _
            rs!SourceNameTag, _
            rs!SourceCellAddress, _
            rs!TargetCellAddress _
        )
        arrFields.Add fldDef
        
        rs.MoveNext
    Loop
    rs.Close
    
    ' 3) 依每個工作表名稱，呼叫 AddWorksheetFields
    Dim key As Variant
    For Each key In mappings.Keys
        ' 把 Collection 轉成 VBA 陣列格式給 AddWorksheetFields
        Dim v() As Variant
        Dim i As Long
        ReDim v(0 To mappings(key).Count - 1)
        For i = 1 To mappings(key).Count
            v(i - 1) = mappings(key)(i)
        Next i
        
        ' 這裡：第一參數是 SourceSheetName，也是你原本的 initSheetName
        AddWorksheetFields key, v
    Next key
    
    ' 4) 其他你現有的 AddDynamicField 邏輯不變
    Select Case reportName
        Case "TABLE10"
            AddDynamicField reportName, "TABLE10_申報時間", "D2", dataMonthStringROC
        ' … 其餘保持原樣
    End Select
End Sub
```

### 重點解說

1. **拿掉**：

   ```vba
   With rptFieldInfo
      .Add "TABLE10", … 
      ' … 全部硬寫 mapping
   End With
   ```

2. **改用** `GetPositionMapRS` 撈資料：

   * `SourceSheetName` → 你的 `rptSheet(0)`
   * `SourceNameTag`   → `rptSheet(1)`（原本 arrField(0)）
   * `SourceCellAddress` → `rptSheet(2)`（原本 arrField(1)）
   * `TargetCellAddress` → 新增的第三要素（原本 arrField(2)）

3. **分組**：先把同一張 Sheet 的多筆欄位放到一個 `Collection`，最後轉成 二維 `Variant` 陣列呼叫 `AddWorksheetFields`。

4. **不動**：`AddDynamicField`、`SetField`、`ApplyToWorkbook`… 主流程完全不改，依舊把 `clsWorksheets` 裡的 `Values`/`Addresses` 寫到新的 Workbook。

---

## 3. 其他注意事項

* **參考設定**：VBA → 工具 → 參考 → 勾「Microsoft ActiveX Data Objects x.x Library」。
* **gDBPath**：確保在 `Main` 裡 `gDBPath` 已設好、那支函式才能正確連到你的 `.accdb`。
* **PositionMap 欄位**：一定要包含 `SourceSheetName, SourceNameTag, SourceCellAddress, TargetSheetName, TargetCellAddress`，上面 SQL 才抓得到。
* **Error Handling**：可依需要把 `rs.Close`、`cn.Close` 包在 `Finally` 區段或 `On Error` 裡，避免連線沒關閉。
* **DataType**：若你還想把 `DataType` 也帶進去（目前 `AddWorksheetFields` 只取前三欄），可以把 `fldDef` 擴成四元素陣列，並改 `AddWorksheetFields` 的迴圈去存這個第四欄位。

---

這樣一來，整個系統就只要你在 Access 的 `PositionMap` 表裡「新增一筆報表定位」，按 `Main` → `Init` 就能自動把它撈出、塞到對應的 Excel Sheet/Cell，省去手動改程式碼的麻煩。
- ================================
Question2.
延續上面議題，我想將函數寫的更有彈性一點，
方法則用類似下面函數的那種方式   
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath
，
另外函數回傳recordset是好的做法嗎?會不會有更好的做法?
請幫我修改完整版本，並告訴我到底修改哪邊(和最原始版本相比)

-============================================
Answer2.
下面是一個更「無痕」且彈性更高的作法，將原本直接回傳 `Recordset` 的 `GetPositionMapRS` 改成：

1. **用 `CreateObject("ADODB.Connection")`** 自行開啟／關閉連線
2. 把查到的資料**組成 Scripting.Dictionary** （Key=工作表名稱、Value=儲存欄位定義陣列），回傳給呼叫端
3. 呼叫端（`clsReport.Init`）只要拿到這個 Dictionary，照原邏輯分群／呼 `AddWorksheetFields` 就好

---

## 一、修改前後差異

| 項目                | 修改前                    | 修改後                                 |
| ----------------- | ---------------------- | ----------------------------------- |
| 函式名稱              | `GetPositionMapRS`     | `GetPositionMap`                    |
| 連線建立              | `New ADODB.Connection` | `CreateObject("ADODB.Connection")`  |
| 回傳型態              | `ADODB.Recordset`      | `Scripting.Dictionary`              |
| 呼叫端處理             | 直接用 `rs.MoveNext`      | 用 `For Each sheetName In dict.Keys` |
| 資源（Connection/RS） | 由呼叫端決定何時關閉             | 在函式內自動關 `rs.Close` / `conn.Close`   |

---

## 二、完整新版：`modDB` 模組

```vba
' === 模組：modDB ===
Option Explicit
' 需在【工具→參考】勾選 Microsoft ActiveX Data Objects x.x Library

'--- 新增：回傳 Dictionary 而不是 Recordset ---'
'   Dictionary 結構：  
'     dict("SheetA") = Array( Array("NameTag1","A1","B1"), Array("NameTag2","A2","B2") )  
Public Function GetPositionMap( _
    ByVal dbPath As String, _
    ByVal reportName As String _
) As Object
    Dim conn      As Object       ' ADODB.Connection
    Dim rs        As Object       ' ADODB.Recordset
    Dim sql       As String
    Dim dictMap   As Object       ' Scripting.Dictionary
    Dim arrFields As Collection
    Dim sheetName As String
    Dim fldDef    As Variant

    ' --- 1) SQL 撈資料，不指定 ORDER BY 也可依需求加 ---'
    sql = "SELECT P.SourceSheetName, P.SourceNameTag, " & _
          "P.SourceCellAddress, P.TargetSheetName, P.TargetCellAddress, P.DataType " & _
          "FROM PositionMap AS P " & _
          "INNER JOIN Report AS R ON P.ReportID = R.ReportID " & _
          "WHERE R.ReportName = '" & Replace(reportName, "'", "''") & "';"

    ' --- 2) 建立並開啟連線 ---'
    Set conn = CreateObject("ADODB.Connection")                 ' << 改用 CreateObject >>
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 0, 1                                     ' adOpenForwardOnly, adLockReadOnly

    ' --- 3) 轉成 Dictionary 分組 ---'
    Set dictMap = CreateObject("Scripting.Dictionary")
    Do While Not rs.EOF
        sheetName = CStr(rs.Fields("SourceSheetName").Value)
        If Not dictMap.Exists(sheetName) Then
            Set arrFields = New Collection
            dictMap.Add sheetName, arrFields
        Else
            Set arrFields = dictMap(sheetName)
        End If

        ' 三元素陣列 = {NameTag, SourceAddress, TargetAddress}
        fldDef = Array( _
            CStr(rs.Fields("SourceNameTag").Value), _
            CStr(rs.Fields("SourceCellAddress").Value), _
            CStr(rs.Fields("TargetCellAddress").Value) _
        )
        arrFields.Add fldDef

        rs.MoveNext
    Loop

    ' --- 4) 關閉資源，回傳 Dictionary ---'
    rs.Close                                                  ' << 在函式內關閉 Recordset >>
    conn.Close                                                ' << 在函式內關閉 Connection >>
    Set GetPositionMap = dictMap
End Function
```

---

## 三、`clsReport.Init` 的對應修改

```diff
Public Sub Init(ByVal reportName As String, _
                ByVal dataMonthStringROC As String, _
                ByVal dataMonthStringROC_NUM As String, _
                ByVal dataMonthStringROC_F1F2 As String)
-   Dim rs As ADODB.Recordset
+   Dim dictMap As Object            ' Scripting.Dictionary
    Dim mappings As Object
    Dim sheetName As String
    Dim arrFields As Collection

    clsReportName = reportName
    Set clsWorksheets = CreateObject("Scripting.Dictionary")

-   ' 1) 從 Access 撈出 Recordset
-   Set rs = GetPositionMapRS(gDBPath, reportName)
+   ' 1) 從 Access 撈回 Dictionary
+   Set dictMap = GetPositionMap(gDBPath, reportName)

-   ' 2) 用 Dictionary 先以 SourceSheetName 分組
-   Set mappings = CreateObject("Scripting.Dictionary")
-   Do While Not rs.EOF
-       ' ... 建立 mappings ...
-       rs.MoveNext
-   Loop
-   rs.Close
+   ' 2) dictMap 已經是「每張 SheetName 分好組」的 Collection
+   Set mappings = dictMap

    ' 3) 依每個工作表名稱，呼叫 AddWorksheetFields
    Dim key As Variant
    For Each key In mappings.Keys
        ' 把 Collection 轉成 VBA 陣列格式給 AddWorksheetFields
        Dim v() As Variant
        Dim i As Long
        ReDim v(0 To mappings(key).Count - 1)
        For i = 1 To mappings(key).Count
            v(i - 1) = mappings(key)(i)
        Next i

        AddWorksheetFields key, v
    Next key

    ' 4) 其餘 AddDynamicField 邏輯不動...
    Select Case reportName
       ' ... 你的原有程式 ...
    End Select
End Sub
```

---

### 為什麼這樣更好？

* **無外露 Recordset/Connection**：呼叫端不必操心何時關閉 DB 物件，也不用載入／管理 ADODB 參考。
* **回傳「純資料」**：`Dictionary` 一旦取出就和資料庫完全脫鉤，執行後關閉連線釋放資源。
* **彈性更高**：若未來想多抓 `DataType`（第 4 個元素），只要在 `fldDef` Array 裡多加一欄即可，呼叫端完全不用改。
* **統一用 `CreateObject`**：不需要在每台機器都設定 VBA 參考，直接 late-bind 最省事。

你只要 **新增上面那支 `GetPositionMap`**，並在 `clsReport.Init` 中把舊的 `GetPositionMapRS` 換成 `GetPositionMap`、並去掉 `Recordset` 分組那段，就能無縫把硬寫的 `rptFieldInfo` 換成由 Access 表驅動的動態設定。對比原始版本，上述 `GetPositionMap` 和 `Init` 裡的那兩處標記 `<< 改用… >>` 就是整個修改重點。
-=========================================
Question3.
有關告訴我更改的地方，請以我最原始檔案來比較，請重新對照註記地方給我一版，詳細標記和我最原始版本修正的地方

-===============================
Answer3.

以下以「最原始版」與「修改後」並列的方式，詳細標記出所有改動。請將下列註記對應到你原本的 VBA 專案中。

---

## 一、在 `modDB` 模組中，新增下列函式

> **位置：** 新增一個標準模組（如 `modDB`），貼上並保留原有程式，新增這支函式即可。

```vb
' === 新增：modDB 內用來從 Access 抓 PositionMap 的函式 ===
Public Function GetPositionMap( _
    ByVal dbPath As String, _
    ByVal reportName As String _
) As Object
    Dim conn      As Object       ' ADODB.Connection
    Dim rs        As Object       ' ADODB.Recordset
    Dim sql       As String
    Dim dictMap   As Object       ' Scripting.Dictionary
    Dim arrFields As Collection
    Dim sheetName As String
    Dim fldDef    As Variant

    ' --- 1) SQL 撈資料 ---
    sql = "SELECT P.SourceSheetName, P.SourceNameTag, " & _
          "P.SourceCellAddress, P.TargetSheetName, " & _
          "P.TargetCellAddress, P.DataType " & _
          "FROM PositionMap AS P " & _
          "INNER JOIN Report AS R ON P.ReportID = R.ReportID " & _
          "WHERE R.ReportName = '" & Replace(reportName, "'", "''") & "';"

    ' --- 2) 建立並開啟連線 (late-bind) ---
    Set conn = CreateObject("ADODB.Connection")     ' <<-- 與原版改為 CreateObject
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 0, 1                         ' adOpenForwardOnly, adLockReadOnly

    ' --- 3) 轉成 Dictionary 分組 (SheetName 為 Key) ---
    Set dictMap = CreateObject("Scripting.Dictionary")
    Do While Not rs.EOF
        sheetName = rs.Fields("SourceSheetName").Value
        If Not dictMap.Exists(sheetName) Then
            Set arrFields = New Collection
            dictMap.Add sheetName, arrFields
        Else
            Set arrFields = dictMap(sheetName)
        End If
        fldDef = Array( _
            rs.Fields("SourceNameTag").Value, _
            rs.Fields("SourceCellAddress").Value, _
            rs.Fields("TargetCellAddress").Value _
        )
        arrFields.Add fldDef
        rs.MoveNext
    Loop

    ' --- 4) 關閉資源並回傳 Dictionary ---
    rs.Close                                        ' <<-- 原版無此
    conn.Close                                      ' <<-- 原版無此
    Set GetPositionMap = dictMap
End Function
```

---

## 二、在 `clsReport` 類別的 `Init` 方法中，標記並改寫

以下左右兩邊分別是「原始程式」與「修改後程式」，以 `<<< 刪除`、`>>> 新增` 註記。

### 1. 原始版 `Init`（節錄關鍵段落）

```vb
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
            Array("FOA", "S2:S21,...", "U2:U21,...")            
        )
        ' … 這裡硬寫數十筆 mapping …
    End With

    clsReportName = reportName
    Set clsWorksheets = CreateObject("Scripting.Dictionary")

    If rptFieldInfo.Exists(reportName) Then
        rptToInit = rptFieldInfo(reportName)
        For i = LBound(rptToInit) To UBound(rptToInit)
            ' … 取出 nameTagRng, addrRng, 呼叫 GetFieldFromXlRanges / AddWorksheetFields …
        Next i        
    Else
        WriteLog "Init未定義報表：" & reportName
    End If

    ' … Rest of your AddDynamicField Select Case  邏輯 …
End Sub
```

### 2. 修改後 `Init`（完整貼出）

```vb
Public Sub Init(ByVal reportName As String, _
                ByVal dataMonthStringROC As String, _
                ByVal dataMonthStringROC_NUM As String, _
                ByVal dataMonthStringROC_F1F2 As String)
    Dim dictMap   As Object    ' <<< 新增：接收 GetPositionMap 回傳值
    Dim mappings  As Object
    Dim sheetName As String
    Dim arrFields As Collection

    clsReportName = reportName
    Set clsWorksheets = CreateObject("Scripting.Dictionary")

    ' >>> 新增：從 Access 讀取 PositionMap，回傳 Dictionary
    Set dictMap = GetPositionMap(gDBPath, reportName)

    ' >>> 新增：直接用 dictMap 當作 mappings 分組資料
    Set mappings = dictMap

    ' 依每個工作表名稱，呼叫 AddWorksheetFields
    Dim key As Variant, v() As Variant
    For Each key In mappings.Keys
        ' 把 Collection 轉成 VBA 2D Variant Array
        ReDim v(0 To mappings(key).Count - 1)
        For i = 1 To mappings(key).Count
            v(i - 1) = mappings(key)(i)
        Next i
        ' 這裡 key 就是 SourceSheetName，相當於你原本的 initSheetName
        AddWorksheetFields key, v
    Next key

    ' >>> 以下 AddDynamicField 與原本 Select Case 全部不變 <<<
    Select Case reportName
        Case "TABLE10"
            AddDynamicField reportName, "TABLE10_申報時間", "D2", dataMonthStringROC
        Case "TABLE15A"
            AddDynamicField reportName, "TABLE15A_申報時間", "D2", dataMonthStringROC
        ' … 其他 case 保持原樣 …
    End Select
End Sub
```

---

### 三、對照重點

| 刪除（−）                                                                                               | 新增（＋）                                                                           |
| --------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------- |
| `Dim rptFieldInfo As Object`                                                                        | `Dim dictMap As Object`                                                         |
| `Set rptFieldInfo = CreateObject("Scripting.Dictionary")`<br>整個 `With rptFieldInfo ... End With` 區塊 | `Set dictMap = GetPositionMap(gDBPath, reportName)`<br>`Set mappings = dictMap` |
| `If rptFieldInfo.Exists(reportName) Then`                                                           | 完全不需要判斷 Exists，因 `GetPositionMap` 回來若沒資料，`mappings.Keys` 直接空迴圈                  |
| 迴圈內呼 `GetFieldFromXlRanges`、`AddWorksheetFields` 的舊邏輯                                               | 統一用 `For Each key In mappings.Keys` + `AddWorksheetFields key, v`               |
| 原本的 `rs`、`rs.Close`、`conn` 管理                                                                       | 全部移入 `GetPositionMap` 裡面，呼叫端不用看到 ADO 物件                                         |

---

這樣，**你唯一要做的**：

1. **貼入**上面「modDB 內用」的 `GetPositionMap` 函式。
2. **替換** `clsReport.Init` 開頭那一大段硬寫 `rptFieldInfo` 的程式，改成呼 `GetPositionMap` + 迴圈。

剩下 **AddDynamicField / SetField / ApplyToWorkbook** 等，不用改動。完成後即可直接由 Access 的 `PositionMap` 表驅動 Excel 各欄位的來源與目標設定。





https://chatgpt.com/share/68333d56-7ae4-8010-9071-70c8d223956d
