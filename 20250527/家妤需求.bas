Que:

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
    allReportNames = Array("CNY1", "FB1", "FB2", "FB3", "FB3A", "FM5", "FM11", "FM13", "AI821", "Table2", "FB5", "FB5A", "FM2", "FM10", "F1_F2", "Table41", "AI602", "AI240", "AI822")

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
    req.Add "TABLE41", Array("Table41_國外部_一利息收入", _
                             "Table41_國外部_一利息收入_利息", _
                             "Table41_國外部_一利息收入_利息_存放銀行同業", _
                             "Table41_國外部_二金融服務收入", _
                             "Table41_國外部_一利息支出", _
                             "Table41_國外部_一利息支出_利息", _
                             "Table41_國外部_一利息支出_利息_外國人外匯存款", _
                             "Table41_國外部_二金融服務支出", _
                             "Table41_企銷處_一利息支出", _
                             "Table41_企銷處_一利息支出_利息", _
                             "Table41_企銷處_一利息支出_利息_外國人新台幣存款")
                            
    req.Add "AI822", Array("AI822_會計科_上年度決算後淨值", _
                           "AI822_國外部_直接往來之授信", _
                           "AI822_國外部_間接往來之授信", _
                           "AI822_授管處_直接往來之授信")

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
    
    reportTitle = rpt.ReportName
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
        ' MsgBox reportTitle & ": " & queryTable & " 資料表無資料"
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
    ' 更改分頁顏色為黃色(6)
    xlsht.Tab.ColorIndex = 6
End Sub

......略



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
    ' MsgBox "完成申報報表更新"
    WriteLog "完成申報報表更新"

CleanUp:
    ' 還原警示
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True    
End Sub



以上是我的Sub，我有兩點需要修改
1.
在這邊的處理程序中
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


    例如像這樣，
        userVal = InputBox( _
        "請確認報表 " & rptName & " 的 [" & fld & "]" & vbCrLf & _
        "目前值：" & defaultVal & vbCrLf & _
        "若要修改，請輸入新數值；若已更改，請直接點擊「確定」。", _
        "欄位值", CStr(defaultVal) _
    )
    在使用者那邊會看到一連串的數字 445646546513，
    我希望可以改成使用者友善的呈現方式，
    例如 123154456.555 改成 123,154,456.555，
    另外也幫我檢查有沒有地方有需要改成類似這樣友善使用者的陳列方式，

    另外在使用者輸入的時候，當輸入3位數時，我也希望可以讓使用者看到,分隔，請問有辦法做到嗎

2.
我希望在最後的
    MsgBox "完成 Process C (全部處理程序完成)"
這邊還加上已經完成的報表名稱，
一張報表一行


請幫我修改完整版本，並告訴我到底修改哪邊，標示清楚哪邊有修改


Ans:

下面是針對你兩個需求所做的完整修改。為了方便你比對，我在修改的地方都加上了註解 `’*** Modified`。請將以下程式碼取代原本的對應區段。

---

### 1. 在「逐報表必要欄位詢問」區段，新增千分位顯示＆允許使用者帶入千分號

```vb
    ' ========== 處理其他部門提供數據欄位 ==========
    ' 定義每張報表必需由使用者填入／確認的儲存格名稱
    Dim req As Object
    Set req = CreateObject("Scripting.Dictionary")
    ' ... (省略 req.Add 內容)

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
                        ' 呼叫 InputBox，預設不顯示（因為原本是空白要輸入）
                        userVal = InputBox( _
                            "報表 " & rptName & " 的欄位 [" & fld & "] 尚未輸入，請填入數值：", _
                            "請填入必要欄位", "")
                        ' 將使用者輸入的千分號去掉再判斷
                        Dim userValClean As String                                           ’*** Modified
                        userValClean = Replace(userVal, ",", "")                              ’*** Modified
                        If userVal = "" Then
                            respToContinue = MsgBox("未輸入任何數值，是否仍要製作報表 " & rptName & "？", _
                                         vbQuestion + vbYesNo, "繼續製作？")
                            If respToContinue = vbYes Then
                                ws.Range(fld).Value = 0
                            Else
                                toRemove.Add rptName
                                Exit For
                            End If
                        ElseIf IsNumeric(userValClean) Then                                     ’*** Modified
                            ws.Range(fld).Value = CDbl(userValClean)                            ’*** Modified
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
                    ' 使用 Format 顯示千分位及小數（最多留三位）                            ’*** Modified
                    Dim sDefault As String                                              ’*** Modified
                    sDefault = Format(defaultVal, "#,##0.###")                             ’*** Modified

                    userVal = InputBox( _
                        "請確認報表 " & rptName & " 的 [" & fld & "]" & vbCrLf & _
                        "目前值：" & sDefault & vbCrLf & _                                ’*** Modified
                        "若要修改，請輸入新數值；若已更改，請直接點擊「確定」。", _
                        "欄位值", sDefault)                                               ’*** Modified

                    userValClean = Replace(userVal, ",", "")                              ’*** Modified

                    If userVal = "" Then
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

                    ElseIf IsNumeric(userValClean) Then                                     ’*** Modified
                        ws.Range(fld).Value = CDbl(userValClean)                            ’*** Modified

                    Else
                        If IsNumeric(defaultVal) Then
                            ws.Range(fld).Value = CDbl(defaultVal)
                        Else
                            ws.Range(fld).Value = 0
                        End If
                        MsgBox "您輸入的不是數字，將保留原值：" & sDefault, vbExclamation      ’*** Modified
                        WriteLog "您輸入的不是數字，將保留原值：" & sDefault                 ’*** Modified
                    End If
                Next fld
            End If
        End If
    Next rptName
```

**說明：**

1. **顯示格式**

   * 先用 `Format(defaultVal, "#,##0.###")` 把預設數值格式化（千分位、最多三位小數）。
   * 把這個字串當做 InputBox 的預設值與提示顯示。

2. **允許使用者帶入千分號**

   * 將 `userVal` 裡的所有 `","` 用 `Replace` 去掉，變成 `userValClean` 再做 `IsNumeric` 及 `CDbl`。

---

### 2. 在主流程最後的 MsgBox，顯示所有完成的報表名稱

找到你的 `Main` Sub 底部這一段：

```vb
    ' Process C: 開啟原始Excel報表(EmptyReportPath)，填入Excel報表數據，
    ' 另存新檔(OutputReportPath)
    Call UpdateExcelReports
    MsgBox "完成 Process C (全部處理程序完成)"                         '*** 原本只有這一句
    WriteLog "完成 Process C (全部處理程序完成)"
End Sub
```

將它改成：

```vb
    ' Process C: 開啟原始Excel報表(EmptyReportPath)，填入Excel報表數據，
    ' 另存新檔(OutputReportPath)
    Call UpdateExcelReports

    '*** Modified：彙整已完成報表清單
    Dim doneList As String
    For Each rptName In gReportNames
        doneList = doneList & "- " & rptName & vbCrLf
    Next rptName

    MsgBox "完成 Process C (全部處理程序完成)：" & vbCrLf & doneList, vbInformation, "處理完成"   ’*** Modified
    WriteLog "完成 Process C (全部處理程序完成)"                                    ’保持原 log
End Sub
```

這樣在訊息盒裡就會一行一張報表地列出所有已完成的報表名稱。

---

#### 小結

* **第一處修改**：`InputBox` 顯示部分，增加 `Format(...)` 和 `Replace(..., ",", "")`，並標示所有被修改的行。
* **第二處修改**：在最後的 `MsgBox` 之前，加入迴圈組成 `doneList`，並把它接到訊息盒。

把上面兩段程式碼取代你原本對應區段後，就可以達成「千分號顯示＋使用者輸入支援千分號」以及「最後一個 MsgBox 顯示完成報表列表」兩個需求。若有任何問題再跟我說！


Que:

有關第一個問題你是怎麼修改的請詳細跟我說明一下

Ans:

下面針對「第一個問題」也就是「在提示使用者時，讓數字顯示千分位格式，且允許使用者輸入帶有千分號的數字」，做一步步的說明：

---

## 1. 先把原本的 `defaultVal` 用 Format 轉成「千分位字串」

原本我們直接用：

```vb
defaultVal = ws.Range(fld).Value
```

然後在 `InputBox` 顯示是：

```vb
"目前值：" & defaultVal
```

所以一串大數字就會像 `445646546513`。

**修改**
在呼叫 `InputBox` 之前，先做：

```vb
Dim sDefault As String
sDefault = Format(defaultVal, "#,##0.###")
```

* `#,##0.###`：

  * 整數部分以千分位分隔；
  * 小數最多顯示 3 位（不夠就少顯示）。

這樣 `sDefault` 就會變成 `"123,154,456.555"` 這種易讀格式。

接著把 `InputBox` 改為：

```vb
userVal = InputBox( _
  "目前值：" & sDefault & vbCrLf & "…", _
  "欄位值", _
  sDefault)
```

* 這裡不只讓提示文字用千分位，預設值也改成 `sDefault`，方便使用者直接按「確定」帶入帶逗號的顯示字串。

---

## 2. 移除使用者輸入字串裡的千分號再當成數字處理

當使用者在 `InputBox` 裡輸入帶逗號的數字，例如 `1,234,567`，直接做 `IsNumeric("1,234,567")` 會回傳 False，因為 VBA 不認逗號。

**修改**
在接收 `userVal` 之後，立刻做：

```vb
Dim userValClean As String
userValClean = Replace(userVal, ",", "")
```

* `Replace`：把所有 `","` 都去掉
* 例如 `"1,234,567"` → `"1234567"`

接著之後所有的 `IsNumeric(userVal)`、`CDbl(userVal)` 都改成 `IsNumeric(userValClean)`、`CDbl(userValClean)`。

---

## 3. 整理後的流程大致是這樣

1. **讀原值**

   ```vb
   defaultVal = ws.Range(fld).Value
   ```
2. **格式化顯示字串**

   ```vb
   sDefault = Format(defaultVal, "#,##0.###")
   ```
3. **詢問使用者，顯示 & 預設都是千分位字串**

   ```vb
   userVal = InputBox( "目前值：" & sDefault, "欄位值", sDefault )
   ```
4. **清理逗號**

   ```vb
   userValClean = Replace(userVal, ",", "")
   ```
5. **驗證數字並寫回儲存格**

   ```vb
   If IsNumeric(userValClean) Then
       ws.Range(fld).Value = CDbl(userValClean)
   Else
       ' 非數字就保留原值或 0
       ws.Range(fld).Value = CDbl(defaultVal)
   End If
   ```

這樣一來：

* **畫面上**：既能讓使用者看見漂亮的千分位、最多三位小數；
* **輸入時**：使用者按「確定」就把帶逗號的字串回填，或自己輸入帶逗號也會自動去除再轉成正確數值。

希望這個步驟說明能讓你完全理解有哪幾行是怎麼動的！如果還有細節想再確認，隨時告訴我。


Que:

那有辦法讓使用者輸入123456的時候畫面會自動顯示123,456嗎?

Ans:

標準的 VBA `InputBox` 無法在使用者輸入時動態地做千分位格式化──它只是一個簡單的對話框，不支援輸入中即時修改畫面內容。要達到「使用者輸入 `123456`，畫面即時顯示 `123,456`」這種效果，就必須改用自訂的 UserForm，並在 TextBox 的事件中寫點程式，像這樣：

---

## 1. 建立一個專門的 UserForm

1. 在 VBA 編輯器中，右鍵 → `插入` → `UserForm`，改名叫 `frmNumberInput`。
2. 在上面放一個 `Label`（顯示提示文字）、一個 `TextBox`（命名 `txtValue`）、兩個 `CommandButton`（分別命名 `cmdOK`、`cmdCancel`，Caption 設為「確定」、「取消」）。

---

## 2. 在這個 Form 的程式碼區貼上：

```vb
' UserForm 變數，用來在外部讀取結果
Public UserNumberResult As Double
Public UserClickedOK As Boolean

Private Sub txtValue_Change()
    ' 1. 先記住目前游標位置
    Dim pos As Long
    pos = txtValue.SelStart
    
    ' 2. 把裡面的「,」先去掉
    Dim s As String
    s = Replace(txtValue.Text, ",", "")
    
    ' 3. 如果是純數字，格式化；否則就不動
    If IsNumeric(s) Then
        txtValue.Text = Format(CDbl(s), "#,##0")
        ' 4. 恢復游標到文字末端（或依需要微調）
        txtValue.SelStart = Len(txtValue.Text)
    End If
End Sub

Private Sub cmdOK_Click()
    ' 確定按鈕，存結果並關閉
    If IsNumeric(Replace(txtValue.Text, ",", "")) Then
        UserNumberResult = CDbl(Replace(txtValue.Text, ",", ""))
        UserClickedOK = True
    Else
        UserNumberResult = 0
        UserClickedOK = False
    End If
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    ' 取消按鈕
    UserClickedOK = False
    Me.Hide
End Sub
```

* **txtValue\_Change**:

  * 每次文字改變就把所有逗號去掉、判斷是不是數字，再用 `Format(..., "#,##0")` 加上千分號。
  * 記得把游標移到末尾，才能有「越打越順」的體驗。

---

## 3. 在你的主程式裡，用這個 Form 取代原本的 `InputBox`

把這段原本的：

```vb
userVal = InputBox("請輸入數值：", "欄位值", sDefault)
' ……後面還要做 Replace、IsNumeric、CDbl
```

改成：

```vb
' 先設定預設值
With frmNumberInput
    .txtValue.Text = sDefault         ' 預設帶入千分位字串
    .UserClickedOK = False
    .Show vbModal
    If .UserClickedOK Then
        userValClean = CStr(.UserNumberResult)
    Else
        userValClean = ""               ' 模擬按取消
    End If
End With
```

1. **.txtValue.Text = sDefault**：把前面做好的千分位字串放到 TextBox
2. **.Show vbModal**：顯示表單，使用者在上面輸入，`txtValue_Change` 就會自動格式化
3. **讀取 `UserNumberResult`**：這是經過去逗號、`CDbl` 後的數值

---

### 彙整

1. **無法**用原生 `InputBox` 即時做千分位格式化；
2. **解法**：改用 UserForm + TextBox 的 `Change` 事件自行格式化；
3. **介面友善**：不但程式碼寫一次就能套用所有欄位，使用者在輸入中還能看到漂亮的 `123,456`、`1,234,567`。

這樣就能完美達成「輸入時自動補上千分號」的需求。若需要更完整的範本或有其他問題，再跟我說！
