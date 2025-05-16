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
    ' gDBPath = "\\10.10.122.40\後台作業\99_個人資料夾\8.修豪\DbsMReport20250513_V1\" & ThisWorkbook.Sheets("ControlPanel").Range("DBsPathFileName").Value
    ' 空白報表路徑
    gReportFolder = ThisWorkbook.Path & "\" & ThisWorkbook.Sheets("ControlPanel").Range("EmptyReportPath").Value
    ' 產生之申報報表路徑
    gOutputFolder = ThisWorkbook.Path & "\" & ThisWorkbook.Sheets("ControlPanel").Range("OutputReportPath").Value

    ' ========== 宣告所有報表 ==========
    ' 製作報表List
    ' gReportNames 少FB1 FM5
    Dim allReportNames As Variant
    allReportNames = Array("CNY1", "FB1", "FB2", "FB3", "FB3A", "FM5", "FM11", "FM13", "AI821", "Table2", "FB5", "FB5A", "FM2", "FM10", "F1_F2", "Table41", "AI602", "AI240", "AI822")
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
        userInput = InputBox("請輸入要執行的報表名稱，用逗號分隔(例如：CNY1,FB2,FM11)：", "指定要產生的報表")
        userInput = Replace(userInput, " ", "")
        gReportNames = Split(userInput, ",")

        ' 把使用者輸入的報表名稱轉成大寫
        For i = LBound(gReportNames) To UBound(gReportNames)
            gReportNames(i) = UCase(gReportNames(i))
        Next i

        ' 檢查不符合的報表名稱
        Dim invalidReports As String
        Dim found As Boolean

        For i = LBound(gReportNames) To UBound(gReportNames)
            WriteLog "gReportNames(i): " & gReportNames(i)
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
    ' ► 若有其他報表也要此流程，可依同模式再加入：
    ' req.Add "FB2", Array("FB2_必要欄1", "FB2_必要欄2")
    ' req.Add "FM11", Array("FM11_必要欄1", ...)

    ' 暫存要移除的報表
    Dim toRemove As Collection
    Set toRemove = New Collection

    ' 逐一詢問使用者每張報表、每個必要欄位的值
    Dim ws As Worksheet
    Dim rptName As Variant 
    Dim fields As Variant, fld As Variant
    Dim defaultVal As Variant, userVal As String
    Dim respToContinue As VbMsgBoxResult

    For Each rptName In gReportNames
        If req.Exists(rptName) Then
            Set ws = ThisWorkbook.Sheets(rptName)
            fields = req(rptName)
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
    Next rptName

    '► 新增：把使用者取消的報表，從 gReportNames 中移除
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
    WriteLog "完成 Process B"

    ' ========== 產生新報表 ==========
    ' Process C: 開啟原始Excel報表(EmptyReportPath)，填入Excel報表數據，
    ' 另存新檔(OutputReportPath)
    Call UpdateExcelReports
    MsgBox "完成 Process C (全部處理程序完成)"
    WriteLog "完成 Process C (全部處理程序完成)"
End Sub

請問在上面的程序中，我使用
只產生部分報表
我輸入
CNY1,FB2,FM11,AI821
為什麼到上面的都顯示未知的報表名稱，
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

有關這地方的Log WriteLog "gReportNames(i): " & gReportNames(i)如下，
看起來字串是不是正確的嗎?
請幫我分析為什麼會顯示未知的報表名稱
2025-05-16 16:29:05 - gReportNames(i): CNY1
2025-05-16 16:29:05 - gReportNames(i): FB2
2025-05-16 16:29:05 - gReportNames(i): FM11
2025-05-16 16:29:05 - gReportNames(i): AI821
