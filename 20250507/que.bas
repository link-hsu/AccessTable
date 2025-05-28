這是我的clsReport
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

    Dim addressMap As Variant
    addressMap = GetMapData(gDBPath, reportName, "FiedlValuePositionMap")
    If IsNull(addressMap) Then
        WriteLog "Init階段FieldMap回傳Null，無法取得報表 " & reportName & " 初始化資料。" 
        Exit Sub
    End If    
    
    Dim wsFields As Object
    Set wsFields = CreateObject("Scripting.Dictionary")

    If IsArray(addressMap) And UBound(addressMap) >= 0 Then
        Dim i As Long
        For i = 0 To UBound(addressMap, 1)
            Dim sheetName As String, nameTag As String, addr As String
            ' 這邊可能要設一個檢核是否拿到Empty String的Log
            sheetName = addressMap(i, 0) &  ""
            nameTag = addressMap(i, 1) &  ""
            addr = addressMap(i, 2) &  ""

            
            If Len(Trim(nameTag)) > 0 And Len(Trim(addr)) > 0 Then
                If Not wsFields.Exists(sheetName) Then
                    wsFields.Add sheetName, Array()
                End If
                Dim tmpArray As Variant
                tmpArray = wsFields(sheetName)
                ReDim Preserve tmpArray(0 To UBound(tmpArray) + 1)
                tmpArray(UBound(tmpArray)) = Array(nameTag, addr, Null)
                wsFields(sheetName) = tmpArray
            End If
        Next i

        Dim key As Variant
        For Each key In wsFields.Keys
            AddWorksheetFields key, wsFields(key)
        Next key
    Else
        WriteLog "未找到需初始化之報表： " & reportName
    End If

    WriteLog "dataMonthStringROC" & dataMonthStringROC
    WriteLog "dataMonthStringROC_NUM" & dataMonthStringROC_NUM
    
    Select Case reportName
        Case "TABLE10"
            AddDynamicField "FOA", "TABLE10_申報時間", "D2", dataMonthStringROC
        Case "TABLE15A"
            AddDynamicField "FOA", "TABLE15A_申報時間", "D2", dataMonthStringROC
        Case "TABLE15B"
            AddDynamicField "FOA", "TABLE15B_申報時間", "D2", dataMonthStringROC
        Case "TABLE16"
            AddDynamicField "FOA", "TABLE16_申報時間", "B2", dataMonthStringROC
        Case "TABLE20"
            AddDynamicField "FOA", "TABLE20_申報時間", "I3", dataMonthStringROC
        Case "TABLE22"
            AddDynamicField "FOA", "TABLE22_申報時間", "E2", dataMonthStringROC
        Case "TABLE23"
            AddDynamicField "FOA", "TABLE23_申報時間", "E2", dataMonthStringROC
        Case "TABLE24"
            AddDynamicField "FOA", "TABLE24_申報時間", "G2", dataMonthStringROC
        Case "TABLE27"
            AddDynamicField "FOA", "TABLE27_申報時間", "E3", dataMonthStringROC
        Case "TABLE36"
            AddDynamicField "FOA", "TABLE36_申報時間", "E2", dataMonthStringROC
        Case "AI233"
            AddDynamicField "Table1", "AI233_申報時間", "B3", dataMonthStringROC_NUM
        Case "AI345"
            AddDynamicField "AI345_NEW", "AI345_申報時間", "A2", dataMonthStringROC_NUM
        Case "AI405"
            AddDynamicField "Table1", "AI405_申報時間", "B3", dataMonthStringROC_NUM
        Case "AI410"
            AddDynamicField "Table1", "AI410_申報時間", "B3", dataMonthStringROC_NUM
        Case "AI430"
            AddDynamicField "Table1", "AI430_申報時間", "B3", dataMonthStringROC_NUM
        Case "AI601"
            AddDynamicField "Table1", "AI601_申報時間", "B3", dataMonthStringROC_NUM
        Case "AI605"
            AddDynamicField "Table1", "AI605_申報時間", "B3", dataMonthStringROC_NUM
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


這是我的主執行序

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
            Case "TABLE10":    Call Process_TABLE10
            ' Case "TABLE15A":    Call Process_TABLE15A
            ' Case "TABLE15B":    Call Process_TABLE15B
            ' Case "TABLE16":    Call Process_TABLE16
            ' Case "TABLE20":    Call Process_TABLE20
            ' Case "TABLE22":    Call Process_TABLE22
            ' Case "TABLE23":    Call Process_TABLE23
            ' Case "TABLE24":    Call Process_TABLE24
            ' Case "TABLE27":    Call Process_TABLE27
            ' Case "TABLE36":    Call Process_TABLE36
            ' Case "AI233":    Call Process_AI233
            ' Case "AI345":    Call Process_AI345
            ' Case "AI405":    Call Process_AI405
            ' Case "AI410":    Call Process_AI410
            ' Case "AI430":    Call Process_AI430
            ' Case "AI601":    Call Process_AI601
            ' Case "AI605":    Call Process_AI605
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
    ' Call UpdateExcelReports
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

Public Sub Process_TABLE10()
    '=== Equal Setting ===
    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("TABLE10")

    Dim reportTitle As String
    reportTitle = rpt.ReportName    

    Dim xlsht As Worksheet
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    ' xlsht.Range("A:G").ClearContents
    ' xlsht.Range("T2:T100").ClearContents

    ' xlsht.Cells.ClearContents  ' 或只清指定区
    
    ' === 【修改】 改为调用通用函数 GetMapData(..., "Query") ===
    Dim queryMap As Variant
    queryMap = GetMapData(gDBPath, reportTitle, "QueryTableMap")

    '【修改】改為正確判斷 Array 且至少有一筆
    If Not IsArray(queryMap) Or UBound(queryMap, 1) < 0 Then
        WriteLog "未在 QueryTableMap 找到 " & reportTitle & " 的任何配置"
        Exit Sub
    End If
    
    Dim iMap As Long
    For iMap = 0 To UBound(queryMap, 1)
        Dim tblName As String
        Dim startColLetter As String
        Dim numCols As Long
        Dim dataArr As Variant
        Dim r As Long, c As Long

        tblName = queryMap(iMap, 0)
        startColLetter = queryMap(iMap, 1)
        numCols = CLng(queryMap(iMap, 2))
        
        dataArr = GetAccessDataAsArray(gDBPath, tblName, gDataMonthString)

        '【修改】正確判斷是否有回傳陣列，且至少要有 header（row 0）跟一筆資料（row 1）
        If Not IsArray(dataArr) Or UBound(dataArr, 1) < 1 Then
            WriteLog "資料有誤: " & reportTitle & " | " & tblName & " 無法取得資料或只有欄位名稱"
            GoTo NextMap
        End If

        ' 取得欄位起始欄號        
        Dim startCol As Long
        startCol = xlsht.Range(startColLetter & "1").Column

        ' 將整個 dataArr（含 header）貼到 Excel，上下：0..UBound(dataArr,1)，左右：0..UBound(dataArr,2)
        '【修改】改用 UBound(..., 1)/(..., 2) 以符合 VB 陣列維度
        For r = 0 To UBound(dataArr, 1)
            For c = 0 To UBound(dataArr, 2)
                xlsht.Cells(r + 1, startCol + c).Value = dataArr(r, c)
            Next c
        Next r

NextMap:
    Next iMap
    '--------------
    'Unique Setting
    '--------------
    '--------------
    'Unique Setting
    '--------------
    Dim rngs As Range
    Dim rng As Range

    Dim AC_CompanyBond_Domestic_Cost As Double
    Dim AC_CompanyBond_Domestic_ImpairmentLoss As Double

    Dim AC_GovBond_Domestic_Cost As Double
    Dim AC_GovBond_Domestic_ImpairmentLoss As Double

    Dim AC_NCD_CentralBank_Cost As Double
    Dim AC_NCD_CentralBank_ImpairmentLoss As Double

    Dim AFS_FinancialBond_Domestic_Cost As Double
    Dim AFS_FinancialBond_Domestic_ValuationAdjust As Double

    Dim EquityMethod_Other_Cost As Double
    
    Dim FVOCI_CompanyBond_Domestic_Cost As Double
    Dim FVOCI_CompanyBond_Domestic_ValuationAdjust As Double

    Dim FVOCI_GovBond_Domestic_Cost As Double
    Dim FVOCI_GovBond_Domestic_ValuationAdjust As Double

    Dim FVOCI_NCD_CentralBank_Cost As Double
    Dim FVOCI_NCD_CentralBank_ValuationAdjust As Double

    ' FVOCI_Stock_特別股_上市_ValuationAdjust
    Dim FVOCI_Stock_PreferredStock_Listed_ValuationAdjust As Double

    ' FVOCI_Stock_普通股_上市_Cost
    Dim FVOCI_Stock_CommonStock_Listed_Cost As Double
    ' FVOCI_Stock_普通股_上市_ValuationAdjust
    Dim FVOCI_Stock_CommonStock_Listed_ValuationAdjust As Double

    ' FVOCI_Stock_普通股_上櫃_Cost
    Dim FVOCI_Stock_CommonStock_OTC_Cost As Double
    ' FVOCI_Stock_普通股_上櫃_ValuationAdjust
    Dim FVOCI_Stock_CommonStock_OTC_ValuationAdjust As Double

    ' FVOCI_Stock_普通股_興櫃_Cost
    Dim FVOCI_Stock_CommonStock_Emergin_Cost As Double
    ' FVOCI_Stock_普通股_興櫃_ValuationAdjust
    Dim FVOCI_Stock_CommonStock_Emergin_ValuationAdjust As Double

    Dim FVOCI_Equity_Other_Cost As Double
    Dim FVOCI_Equity_Other_ValuationAdjust As Double

    Dim FVPL_AssetCertificate_Cost As Double
    Dim FVPL_AssetCertificate_ValuationAdjust As Double

    Dim FVPL_CompanyBond_Domestic_Cost As Double
    Dim FVPL_CompanyBond_Domestic_ValuationAdjust As Double

    Dim FVPL_CP_Cost As Double
    Dim FVPL_CP_ValuationAdjust As Double

    公債原始成本
    FVPL_GovBond_Domestic_Cost + FVOCI_GovBond_Domestic_Cost + AC_GovBond_Domestic_Cost

    公債
    透過損益按公允價值衡量之金融資產2 A
    FVPL_GovBond_Domestic_Cost + FVPL_GovBond_Domestic_ValuationAdjust

    公債
    透過其他綜合損益按公允價值衡量之金融資產2 B

    FVOCI_GovBond_Domestic_Cost + FVOCI_GovBond_Domestic_ValuationAdjust

    公債
    ac
    AC_GovBond_Domestic_Cost + AC_GovBond_Domestic_ImpairmentLoss


    2.公司債		
    2.1.公營事業		
        原始取得成本1		
    120050121		強制FVPL金融資產-普通公司債(公營)                 
    121110121		FVOCI債務工具-普通公司債（公營）                  
    122010121		AC債務工具投資-普通公司債(公營)

    FVPL_CompanyBond_Domestic_Cost + FVOCI_CompanyBond_Domestic_Cost + AC_CompanyBond_Domestic_Cost
            
        透過損益按公允價值衡量之金融資產2 A		
    120050121		強制FVPL金融資產-普通公司債(公營)                 
    120070121		強制FVPL金融資產評價調整-普通公司債(公營)   
    
    FVPL_CompanyBond_Domestic_Cost + 
            
        透過其他綜合損益按公允價值衡量之金融資產2 B		
    121110121		FVOCI債務工具-普通公司債（公營）                  
            
        按攤銷後成本衡量之債務工具投資2 C		
    122010121		AC債務工具投資-普通公司債(公營)                   
            
    2.2.民營企業-國內公司債		
        原始取得成本1		
    120050123		強制FVPL金融資產-普通公司債(民營)                 
            
        透過損益按公允價值衡量之金融資產2 A		
    120050123		強制FVPL金融資產-普通公司債(民營)                 
    120070123		強制FVPL金融資產評價調整-普通公司債(民營)         
            
        透過其他綜合損益按公允價值衡量之金融資產2 B		
    121110123		FVOCI債務工具-普通公司債（民營）                  
            
        按攤銷後成本衡量之債務工具投資2 C		
    122010123		AC債務工具投資-普通公司債(民營)                   
    
    



    Dim FVPL_GovBond_Domestic_Cost As Double
    Dim FVPL_GovBond_Domestic_ValuationAdjust As Double

    ' FVPL_Stock_特別股_上市_Cost
    Dim FVPL_Stock_PreferredStock_Listed_Cost As Double
    ' FVPL_Stock_特別股_上市_ValuationAdjust
    Dim FVPL_Stock_PreferredStock_Listed_ValuationAdjust As Double

    ' FVPL_Stock_普通股_上市_Cost
    Dim FVPL_Stock_CommonStock_Listed_Cost As Double
    ' FVPL_Stock_普通股_上市_ValuationAdjust
    Dim FVPL_Stock_CommonStock_Listed_ValuationAdjust As Double

    ' FVPL_Stock_普通股_上櫃_Cost
    Dim FVPL_Stock_CommonStock_OTC_Cost As Double
    ' FVPL_Stock_普通股_上櫃_ValuationAdjust
    Dim FVPL_Stock_CommonStock_OTC_ValuationAdjust As Double

    ' FVPL_Stock_普通股_興櫃_Cost
    Dim FVPL_Stock_CommonStock_Emergin_Cost As Double
    ' FVPL_Stock_普通股_興櫃_ValuationAdjust
    Dim FVPL_Stock_CommonStock_Emergin_ValuationAdjust As Double
    
    lastRow = xlsht.Cells(xlsht.Rows.Count, "A").End(xlUp).Row
    Set rngs = xlsht.Range("A2:A" & lastRow)

    For Each rng In rngs
        If CStr(rng.Value) = "AC_CompanyBond_Domestic_Cost" Then
            AC_CompanyBond_Domestic_Cost = AC_CompanyBond_Domestic_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "AC_CompanyBond_Domestic_ImpairmentLoss" Then
            AC_CompanyBond_Domestic_ImpairmentLoss = AC_CompanyBond_Domestic_ImpairmentLoss + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "AC_GovBond_Domestic_Cost" Then
            AC_GovBond_Domestic_Cost = AC_GovBond_Domestic_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "AC_GovBond_Domestic_ImpairmentLoss" Then
            AC_GovBond_Domestic_ImpairmentLoss = AC_GovBond_Domestic_ImpairmentLoss + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "AC_NCD_CentralBank_Cost" Then
            AC_NCD_CentralBank_Cost = AC_NCD_CentralBank_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "AC_NCD_CentralBank_ImpairmentLoss" Then
            AC_NCD_CentralBank_ImpairmentLoss = AC_NCD_CentralBank_ImpairmentLoss + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "AFS_FinancialBond_Domestic_Cost" Then
            AFS_FinancialBond_Domestic_Cost = AFS_FinancialBond_Domestic_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "AFS_FinancialBond_Domestic_ValuationAdjust" Then
            AFS_FinancialBond_Domestic_ValuationAdjust = AFS_FinancialBond_Domestic_ValuationAdjust + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "EquityMethod_Other_Cost" Then
            EquityMethod_Other_Cost = EquityMethod_Other_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVOCI_CompanyBond_Domestic_Cost" Then
            FVOCI_CompanyBond_Domestic_Cost = FVOCI_CompanyBond_Domestic_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVOCI_CompanyBond_Domestic_ValuationAdjust" Then
            FVOCI_CompanyBond_Domestic_ValuationAdjust = FVOCI_CompanyBond_Domestic_ValuationAdjust + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVOCI_GovBond_Domestic_Cost" Then
            FVOCI_GovBond_Domestic_Cost = FVOCI_GovBond_Domestic_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVOCI_GovBond_Domestic_ValuationAdjust" Then
            FVOCI_GovBond_Domestic_ValuationAdjust = FVOCI_GovBond_Domestic_ValuationAdjust + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVOCI_NCD_CentralBank_Cost" Then
            FVOCI_NCD_CentralBank_Cost = FVOCI_NCD_CentralBank_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVOCI_NCD_CentralBank_ValuationAdjust" Then
            FVOCI_NCD_CentralBank_ValuationAdjust = FVOCI_NCD_CentralBank_ValuationAdjust + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVOCI_Stock_特別股_上市_ValuationAdjust" Then
            FVOCI_Stock_PreferredStock_Listed_ValuationAdjust = FVOCI_Stock_PreferredStock_Listed_ValuationAdjust + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVOCI_Stock_普通股_上市_Cost" Then
            FVOCI_Stock_CommonStock_Listed_Cost = FVOCI_Stock_CommonStock_Listed_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVOCI_Stock_普通股_上市_ValuationAdjust" Then
            FVOCI_Stock_CommonStock_Listed_ValuationAdjust = FVOCI_Stock_CommonStock_Listed_ValuationAdjust + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVOCI_Stock_普通股_上櫃_Cost" Then
            FVOCI_Stock_CommonStock_OTC_Cost = FVOCI_Stock_CommonStock_OTC_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVOCI_Stock_普通股_上櫃_ValuationAdjust" Then
            FVOCI_Stock_CommonStock_OTC_ValuationAdjust = FVOCI_Stock_CommonStock_OTC_ValuationAdjust + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVOCI_Stock_普通股_興櫃_Cost" Then
            FVOCI_Stock_CommonStock_Emergin_Cost = FVOCI_Stock_CommonStock_Emergin_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVOCI_Stock_普通股_興櫃_ValuationAdjust" Then
            FVOCI_Stock_CommonStock_Emergin_ValuationAdjust = FVOCI_Stock_CommonStock_Emergin_ValuationAdjust + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVOCI_Equity_Other_Cost" Then
            FVOCI_Equity_Other_Cost = FVOCI_Equity_Other_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVOCI_Equity_Other_ValuationAdjust" Then
            FVOCI_Equity_Other_ValuationAdjust = FVOCI_Equity_Other_ValuationAdjust + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVPL_AssetCertificate_Cost" Then
            FVPL_AssetCertificate_Cost = FVPL_AssetCertificate_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVPL_AssetCertificate_ValuationAdjust" Then
            FVPL_AssetCertificate_ValuationAdjust = FVPL_AssetCertificate_ValuationAdjust + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVPL_CompanyBond_Domestic_Cost" Then
            FVPL_CompanyBond_Domestic_Cost = FVPL_CompanyBond_Domestic_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVPL_CompanyBond_Domestic_ValuationAdjust" Then
            FVPL_CompanyBond_Domestic_ValuationAdjust = FVPL_CompanyBond_Domestic_ValuationAdjust + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVPL_CP_Cost" Then
            FVPL_CP_Cost = FVPL_CP_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVPL_CP_ValuationAdjust" Then
            FVPL_CP_ValuationAdjust = FVPL_CP_ValuationAdjust + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVPL_GovBond_Domestic_Cost" Then
            FVPL_GovBond_Domestic_Cost = FVPL_GovBond_Domestic_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVPL_GovBond_Domestic_ValuationAdjust" Then
            FVPL_GovBond_Domestic_ValuationAdjust = FVPL_GovBond_Domestic_ValuationAdjust + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVPL_Stock_特別股_上市_Cost" Then
            FVPL_Stock_PreferredStock_Listed_Cost = FVPL_Stock_PreferredStock_Listed_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVPL_Stock_特別股_上市_ValuationAdjust" Then
            FVPL_Stock_PreferredStock_Listed_ValuationAdjust = FVPL_Stock_PreferredStock_Listed_ValuationAdjust + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVPL_Stock_普通股_上市_Cost" Then
            FVPL_Stock_CommonStock_Listed_Cost = FVPL_Stock_CommonStock_Listed_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVPL_Stock_普通股_上市_ValuationAdjust" Then
            FVPL_Stock_CommonStock_Listed_ValuationAdjust = FVPL_Stock_CommonStock_Listed_ValuationAdjust + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVPL_Stock_普通股_上櫃_Cost" Then
            FVPL_Stock_CommonStock_OTC_Cost = FVPL_Stock_CommonStock_OTC_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVPL_Stock_普通股_上櫃_ValuationAdjust" Then
            FVPL_Stock_CommonStock_OTC_ValuationAdjust = FVPL_Stock_CommonStock_OTC_ValuationAdjust + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVPL_Stock_普通股_興櫃_Cost" Then
            FVPL_Stock_CommonStock_Emergin_Cost = FVPL_Stock_CommonStock_Emergin_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVPL_Stock_普通股_興櫃_ValuationAdjust" Then
            FVPL_Stock_CommonStock_Emergin_ValuationAdjust = FVPL_Stock_CommonStock_Emergin_ValuationAdjust + rng.Offset(0, 1).Value
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
    ' 更改分頁顏色為黃色(6)
    xlsht.Tab.ColorIndex = 6
End Sub



Public Function GetMapData(ByVal DBPath As String, _
                           ByVal reportName As String, _
                           ByVal tableName As String) As Variant
    Dim conn As Object, rs As Object
    Dim sql As String
    Dim results() As Variant
    Dim i As Long
    Dim arr As Variant

    WriteLog "DBPath: " & DBPath
    WriteLog "tableName: " & tableName

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

    WriteLog "sql: " & sql
    
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

1.以上那是有關取得FiedlValuePositionMap表中資料的Function，
看是要另外寫一隻function或是修改這隻function和其他程序併用，
你在評估，function用途還是取得資料表資料

主要迴圈等程序還是必須在Process那邊執行
主要目的是在
rpt.SetField "FOA", "FB2_資產總計", CStr(totalAsset)
這個部分，我不想手動逐一去進行寫入，
我會在事先先在Range()中的NameTag把值輸入進去，
請修改下列函數，取得fiedlvaluepositionmap資料表中的，
SourceNameTag及TargetSheetName
逐一使用Range("SourceNameTag")將資料表中所有紀錄要填的數值
逐一執行rpt.SetField TargetSheetName, SourceNameTag, Range("SourceNameTag")，
請問要怎麼修改，請給我完整版本，並告訴我修改了哪邊


2.另外有一個問題是，我在ProcessTABLE10中上半段的有從資料庫動態抓取資料丟入工作表中，
我想要先把Equal Setting階段startColLetter資料紀錄起來(或是有其他需要紀錄的資料也請一併修改)，
在下面的Unique階段我想更動態的取得欄位，更簡潔的將相關資料放入變數中，儲存在相對應的儲存格中，
請問要怎麼修改，請給我完整版，並詳細標記修改哪些地方
