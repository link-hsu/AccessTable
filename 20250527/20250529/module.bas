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

'=== B 各報表獨立處理邏輯 ===

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
    
    ' 呼叫通用函式，並取得 importCols
    Dim importCols As Collection
    Set importCols = ImportQueryTables(gDBPath, xlsht, reportTitle, gDataMonthString)

    If importCols Is Nothing Or importCols.Count = 0 Then
        Exit Sub
    End If    

    '--------------
    'Unique Setting
    '--------------

    ' AC債務工具投資-普通公司債(公營)
    Dim AC_CompanyBond_Public_Domestic_Cost As Double
    ' 累積減損-累積減損-AC債務工具投資-普通公司(公營)
    Dim AC_CompanyBond_Public_Domestic_ImpairmentLoss As Double

    ' AC債務工具投資-普通公司債(民營)
    Dim AC_CompanyBond_Private_Domestic_Cost As Double
    ' 累積減損-AC債務工具投資-普通公司(民營)
    Dim AC_CompanyBond_Private_Domestic_ImpairmentLoss As Double

    Dim AC_GovBond_Domestic_Cost As Double
    Dim AC_GovBond_Domestic_ImpairmentLoss As Double

    Dim AC_NCD_CentralBank_Cost As Double
    Dim AC_NCD_CentralBank_ImpairmentLoss As Double

    Dim AFS_FinancialBond_Domestic_Cost As Double
    Dim AFS_FinancialBond_Domestic_ValuationAdjust As Double

    Dim EquityMethod_Cost As Double
    Dim EquityMethod_ValuationAdjust As Double

    Dim EquityMethod_Other_Cost As Double
    
    ' FVOCI債務工具-普通公司債（公營）
    Dim FVOCI_CompanyBond_Public_Domestic_Cost As Double
    ' FVOCI債務工具評價調整-普通公司債（公營)
    Dim FVOCI_CompanyBond_Public_Domestic_ValuationAdjust As Double

    ' FVOCI債務工具-普通公司債（民營）
    Dim FVOCI_CompanyBond_Private_Domestic_Cost As Double
    ' FVOCI債務工具評價調整-普通公司債（民營)
    Dim FVOCI_CompanyBond_Private_Domestic_ValuationAdjust As Double

    Dim FVOCI_GovBond_Domestic_Cost As Double
    Dim FVOCI_GovBond_Domestic_ValuationAdjust As Double

    Dim FVOCI_NCD_CentralBank_Cost As Double
    Dim FVOCI_NCD_CentralBank_ValuationAdjust As Double

    ' FVOCI_Stock_特別股_Cost
    Dim FVOCI_Stock_PreferredStock_Cost As Double
    ' FVOCI_Stock_特別股_上市_Cost
    Dim FVOCI_Stock_PreferredStock_Listed_Cost As Double
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

    ' 強制FVPL金融資產-普通公司債(公營)
    Dim FVPL_CompanyBond_Public_Domestic_Cost As Double
    ' 強制FVPL金融資產評價調整-普通公司債(公營)
    Dim FVPL_CompanyBond_Public_Domestic_ValuationAdjust As Double

    ' 強制FVPL金融資產-普通公司債(民營)
    Dim FVPL_CompanyBond_Private_Domestic_Cost As Double
    ' 強制FVPL金融資產評價調整-普通公司債(民營)
    Dim FVPL_CompanyBond_Private_Domestic_ValuationAdjust As Double

    Dim FVPL_CP_Cost As Double
    Dim FVPL_CP_ValuationAdjust As Double

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

    Dim lastRow1 As Long, lastRow2 As Long
    Dim rngs1 As Range, rngs2 As Range, rng As Range    

    ' --- 【修改1.2】不迴圈，直接取用 importCols(1) 和 importCols(2) ---
    If importCols.Count >= 1 Then
        lastRow1 = xlsht.Cells(xlsht.Rows.Count, importCols(1)).End(xlUp).Row
        Set rngs1 = xlsht.Range(xlsht.Cells(2, importCols(1)), xlsht.Cells(lastRow1, importCols(1)))
        
        For Each rng In rngs1
            Select Case CStr(rng.Value)
                Case "AC_CompanyBond_Domestic_Cost"

                Case "AC_CompanyBond_Domestic_ImpairmentLoss"

                Case "AC_GovBond_Domestic_Cost"
                    AC_GovBond_Domestic_Cost = AC_GovBond_Domestic_Cost + rng.Offset(0, 1).Value
                Case "AC_GovBond_Domestic_ImpairmentLoss"
                    AC_GovBond_Domestic_ImpairmentLoss = AC_GovBond_Domestic_ImpairmentLoss + rng.Offset(0, 1).Value
                Case "AC_NCD_CentralBank_Cost"
                    AC_NCD_CentralBank_Cost = AC_NCD_CentralBank_Cost + rng.Offset(0, 1).Value
                Case "AC_NCD_CentralBank_ImpairmentLoss"
                    AC_NCD_CentralBank_ImpairmentLoss = AC_NCD_CentralBank_ImpairmentLoss + rng.Offset(0, 1).Value
                Case "AFS_FinancialBond_Domestic_Cost"
                    AFS_FinancialBond_Domestic_Cost = AFS_FinancialBond_Domestic_Cost + rng.Offset(0, 1).Value
                Case "AFS_FinancialBond_Domestic_ValuationAdjust"
                    AFS_FinancialBond_Domestic_ValuationAdjust = AFS_FinancialBond_Domestic_ValuationAdjust + rng.Offset(0, 1).Value
                Case "EquityMethod_Cost"
                    EquityMethod_Cost = EquityMethod_Cost + rng.Offset(0, 1).Value
                Case "EquityMethod_ValuationAdjust"
                    EquityMethod_ValuationAdjust = EquityMethod_ValuationAdjust + rng.Offset(0, 1).Value            
                Case "EquityMethod_Other_Cost"
                    EquityMethod_Other_Cost = EquityMethod_Other_Cost + rng.Offset(0, 1).Value
                Case "FVOCI_CompanyBond_Domestic_Cost"

                Case "FVOCI_CompanyBond_Domestic_ValuationAdjust"

                Case "FVOCI_GovBond_Domestic_Cost"
                    FVOCI_GovBond_Domestic_Cost = FVOCI_GovBond_Domestic_Cost + rng.Offset(0, 1).Value
                Case "FVOCI_GovBond_Domestic_ValuationAdjust"
                    FVOCI_GovBond_Domestic_ValuationAdjust = FVOCI_GovBond_Domestic_ValuationAdjust + rng.Offset(0, 1).Value
                Case "FVOCI_NCD_CentralBank_Cost"
                    FVOCI_NCD_CentralBank_Cost = FVOCI_NCD_CentralBank_Cost + rng.Offset(0, 1).Value
                Case "FVOCI_NCD_CentralBank_ValuationAdjust"
                    FVOCI_NCD_CentralBank_ValuationAdjust = FVOCI_NCD_CentralBank_ValuationAdjust + rng.Offset(0, 1).Value
                Case "FVOCI_Stock_特別股_Cost"
                    FVOCI_Stock_PreferredStock_Cost = FVOCI_Stock_PreferredStock_Cost + rng.Offset(0, 1).Value
                Case "FVOCI_Stock_特別股_上市_Cost"
                    FVOCI_Stock_PreferredStock_Listed_Cost = FVOCI_Stock_PreferredStock_Listed_Cost + rng.Offset(0, 1).Value                        
                Case "FVOCI_Stock_特別股_上市_ValuationAdjust"
                    FVOCI_Stock_PreferredStock_Listed_ValuationAdjust = FVOCI_Stock_PreferredStock_Listed_ValuationAdjust + rng.Offset(0, 1).Value
                Case "FVOCI_Stock_普通股_上市_Cost"
                    FVOCI_Stock_CommonStock_Listed_Cost = FVOCI_Stock_CommonStock_Listed_Cost + rng.Offset(0, 1).Value
                Case "FVOCI_Stock_普通股_上市_ValuationAdjust"
                    FVOCI_Stock_CommonStock_Listed_ValuationAdjust = FVOCI_Stock_CommonStock_Listed_ValuationAdjust + rng.Offset(0, 1).Value
                Case "FVOCI_Stock_普通股_上櫃_Cost"
                    FVOCI_Stock_CommonStock_OTC_Cost = FVOCI_Stock_CommonStock_OTC_Cost + rng.Offset(0, 1).Value
                Case "FVOCI_Stock_普通股_上櫃_ValuationAdjust"
                    FVOCI_Stock_CommonStock_OTC_ValuationAdjust = FVOCI_Stock_CommonStock_OTC_ValuationAdjust + rng.Offset(0, 1).Value
                Case "FVOCI_Stock_普通股_興櫃_Cost"
                    FVOCI_Stock_CommonStock_Emergin_Cost = FVOCI_Stock_CommonStock_Emergin_Cost + rng.Offset(0, 1).Value
                Case "FVOCI_Stock_普通股_興櫃_ValuationAdjust"
                    FVOCI_Stock_CommonStock_Emergin_ValuationAdjust = FVOCI_Stock_CommonStock_Emergin_ValuationAdjust + rng.Offset(0, 1).Value
                Case "FVOCI_Equity_Other_Cost"
                    FVOCI_Equity_Other_Cost = FVOCI_Equity_Other_Cost + rng.Offset(0, 1).Value
                Case "FVOCI_Equity_Other_ValuationAdjust"
                    FVOCI_Equity_Other_ValuationAdjust = FVOCI_Equity_Other_ValuationAdjust + rng.Offset(0, 1).Value
                Case "FVPL_AssetCertificate_Cost"
                    FVPL_AssetCertificate_Cost = FVPL_AssetCertificate_Cost + rng.Offset(0, 1).Value
                Case "FVPL_AssetCertificate_ValuationAdjust"
                    FVPL_AssetCertificate_ValuationAdjust = FVPL_AssetCertificate_ValuationAdjust + rng.Offset(0, 1).Value
                Case "FVPL_CompanyBond_Domestic_Cost"

                Case "FVPL_CompanyBond_Domestic_ValuationAdjust"

                Case "FVPL_CP_Cost"
                    FVPL_CP_Cost = FVPL_CP_Cost + rng.Offset(0, 1).Value
                Case "FVPL_CP_ValuationAdjust"
                    FVPL_CP_ValuationAdjust = FVPL_CP_ValuationAdjust + rng.Offset(0, 1).Value
                Case "FVPL_GovBond_Domestic_Cost"
                    FVPL_GovBond_Domestic_Cost = FVPL_GovBond_Domestic_Cost + rng.Offset(0, 1).Value
                Case "FVPL_GovBond_Domestic_ValuationAdjust"
                    FVPL_GovBond_Domestic_ValuationAdjust = FVPL_GovBond_Domestic_ValuationAdjust + rng.Offset(0, 1).Value
                Case "FVPL_Stock_特別股_上市_Cost"
                    FVPL_Stock_PreferredStock_Listed_Cost = FVPL_Stock_PreferredStock_Listed_Cost + rng.Offset(0, 1).Value
                Case "FVPL_Stock_特別股_上市_ValuationAdjust"
                    FVPL_Stock_PreferredStock_Listed_ValuationAdjust = FVPL_Stock_PreferredStock_Listed_ValuationAdjust + rng.Offset(0, 1).Value
                Case "FVPL_Stock_普通股_上市_Cost"
                    FVPL_Stock_CommonStock_Listed_Cost = FVPL_Stock_CommonStock_Listed_Cost + rng.Offset(0, 1).Value
                Case "FVPL_Stock_普通股_上市_ValuationAdjust"
                    FVPL_Stock_CommonStock_Listed_ValuationAdjust = FVPL_Stock_CommonStock_Listed_ValuationAdjust + rng.Offset(0, 1).Value
                Case "FVPL_Stock_普通股_上櫃_Cost"
                    FVPL_Stock_CommonStock_OTC_Cost = FVPL_Stock_CommonStock_OTC_Cost + rng.Offset(0, 1).Value
                Case "FVPL_Stock_普通股_上櫃_ValuationAdjust"
                    FVPL_Stock_CommonStock_OTC_ValuationAdjust = FVPL_Stock_CommonStock_OTC_ValuationAdjust + rng.Offset(0, 1).Value
                Case "FVPL_Stock_普通股_興櫃_Cost"
                    FVPL_Stock_CommonStock_Emergin_Cost = FVPL_Stock_CommonStock_Emergin_Cost + rng.Offset(0, 1).Value
                Case "FVPL_Stock_普通股_興櫃_ValuationAdjust"
                    FVPL_Stock_CommonStock_Emergin_ValuationAdjust = FVPL_Stock_CommonStock_Emergin_ValuationAdjust + rng.Offset(0, 1).Value
            End Select
        Next rng
    End If    
    
    If importCols.Count >= 2 Then
        lastRow2 = xlsht.Cells(xlsht.Rows.Count, importCols(2)).End(xlUp).Row
        Set rngs2 = xlsht.Range(xlsht.Cells(2, importCols(2)), xlsht.Cells(lastRow2, importCols(2)))
        
        For Each rng In rngs2
            ' 如果第二筆表也有需要累計的 tag，可以在這裡加
            Select Case CStr(rng.Value)
                Case "強制FVPL金融資產-普通公司債(公營)"
                    FVPL_CompanyBond_Public_Domestic_Cost = FVPL_CompanyBond_Public_Domestic_Cost + rng.Offset(0, 1).Value
                Case "強制FVPL金融資產-普通公司債(民營)"
                    FVPL_CompanyBond_Private_Domestic_Cost = FVPL_CompanyBond_Private_Domestic_Cost + rng.Offset(0, 1).Value
                Case "強制FVPL金融資產評價調整-普通公司債(公營)"
                    FVPL_CompanyBond_Public_Domestic_ValuationAdjust = FVPL_CompanyBond_Public_Domestic_ValuationAdjust + rng.Offset(0, 1).Value
                Case "強制FVPL金融資產評價調整-普通公司債(民營)"
                    FVPL_CompanyBond_Private_Domestic_ValuationAdjust = FVPL_CompanyBond_Private_Domestic_ValuationAdjust + rng.Offset(0, 1).Value
                Case "FVOCI債務工具-普通公司債（公營）"
                    FVOCI_CompanyBond_Public_Domestic_Cost = FVOCI_CompanyBond_Public_Domestic_Cost + rng.Offset(0, 1).Value
                Case "FVOCI債務工具-普通公司債（民營）"
                    FVOCI_CompanyBond_Private_Domestic_Cost = FVOCI_CompanyBond_Private_Domestic_Cost + rng.Offset(0, 1).Value
                Case "FVOCI債務工具評價調整-普通公司債（公營)"
                    FVOCI_CompanyBond_Public_Domestic_ValuationAdjust = FVOCI_CompanyBond_Public_Domestic_ValuationAdjust + rng.Offset(0, 1).Value
                Case "FVOCI債務工具評價調整-普通公司債（民營)"
                    FVOCI_CompanyBond_Private_Domestic_ValuationAdjust = FVOCI_CompanyBond_Private_Domestic_ValuationAdjust + rng.Offset(0, 1).Value
                Case "AC債務工具投資-普通公司債(公營)"
                    AC_CompanyBond_Public_Domestic_Cost = AC_CompanyBond_Public_Domestic_Cost + rng.Offset(0, 1).Value
                Case "AC債務工具投資-普通公司債(民營)"
                    AC_CompanyBond_Private_Domestic_Cost = AC_CompanyBond_Private_Domestic_Cost + rng.Offset(0, 1).Value
                Case "累積減損-累積減損-AC債務工具投資-普通公司(公營)"
                    AC_CompanyBond_Public_Domestic_ImpairmentLoss = AC_CompanyBond_Public_Domestic_ImpairmentLoss + rng.Offset(0, 1).Value
                Case "累積減損-AC債務工具投資-普通公司(民營)"
                    AC_CompanyBond_Private_Domestic_ImpairmentLoss = AC_CompanyBond_Private_Domestic_ImpairmentLoss + rng.Offset(0, 1).Value
            End Select
        Next rng
    End If

    ' HANDLE方式
    
    公債原始成本
    GovBond_Domestic_Cost

    FVPL_GovBond_Domestic_Cost + FVOCI_GovBond_Domestic_Cost + AC_GovBond_Domestic_Cost

    公債
    透過損益按公允價值衡量之金融資產2 A
    FVPL_GovBond_Domestic
    FVPL_GovBond_Domestic_Cost + FVPL_GovBond_Domestic_ValuationAdjust

    公債
    透過其他綜合損益按公允價值衡量之金融資產2 B
    FVOCI_GovBond_Domestic
    FVOCI_GovBond_Domestic_Cost + FVOCI_GovBond_Domestic_ValuationAdjust

    公債
    ac
    AC_GovBond_Domestic
    AC_GovBond_Domestic_Cost + AC_GovBond_Domestic_ImpairmentLoss

    2.公司債		
    2.1.公營事業		
        原始取得成本1		
    CompanyBond_Public_Domestic_Cost
    120050121		強制FVPL金融資產-普通公司債(公營)                 
    121110121		FVOCI債務工具-普通公司債（公營）                  
    122010121		AC債務工具投資-普通公司債(公營)
    
    FVPL_CompanyBond_Public_Domestic_Cost + FVOCI_CompanyBond_Public_Domestic_Cost + AC_CompanyBond_Public_Domestic_Cost

            
        透過損益按公允價值衡量之金融資產2 A		
    FVPL_CompanyBond_Public_Domestic
    120050121		強制FVPL金融資產-普通公司債(公營)                 
    120070121		強制FVPL金融資產評價調整-普通公司債(公營)   
    FVPL_CompanyBond_Public_Domestic_Cost + FVPL_CompanyBond_Public_Domestic_ValuationAdjust
    
            
        透過其他綜合損益按公允價值衡量之金融資產2 B		
    FVOCI_CompanyBond_Public_Domestic
    121110121
    121130121
    ' 重新勾稽
    ' --    

    121110121		FVOCI債務工具-普通公司債（公營）                  
    121130121		FVOCI債務工具評價調整-普通公司債（公營)           
            
    FVOCI_CompanyBond_Public_Domestic_Cost + FVOCI_CompanyBond_Public_Domestic_ValuationAdjust


        按攤銷後成本衡量之債務工具投資2 C		
    AC_CompanyBond_Public_Domestic
    122010121		AC債務工具投資-普通公司債(公營)                   
            

    AC_CompanyBond_Public_Domestic_Cost +


    2.2.民營企業-國內公司債		
        原始取得成本1		
    CompanyBond_Private_Domestic_Cost
    120050123		強制FVPL金融資產-普通公司債(民營)                 
    121110123		FVOCI債務工具-普通公司債（民營）                  
    122010123		AC債務工具投資-普通公司債(民營)                   
    
    FVPL_CompanyBond_Private_Domestic_Cost + FVOCI_CompanyBond_Private_Domestic_Cost + AC_CompanyBond_Private_Domestic_Cost


        透過損益按公允價值衡量之金融資產2 A		
    FVPL_CompanyBond_Private_Domestic
    120050123		強制FVPL金融資產-普通公司債(民營)                 
    120070123		強制FVPL金融資產評價調整-普通公司債(民營)         

    FVPL_CompanyBond_Private_Domestic_Cost + FVPL_CompanyBond_Private_Domestic_ValuationAdjust

        透過其他綜合損益按公允價值衡量之金融資產2 B		
    FVOCI_CompanyBond_Private_Domestic
    121110123		FVOCI債務工具-普通公司債（民營）                  
    121130123		FVOCI債務工具評價調整-普通公司債（民營)               

    FVOCI_CompanyBond_Private_Domestic_Cost + FVOCI_CompanyBond_Private_Domestic_ValuationAdjust

        按攤銷後成本衡量之債務工具投資2 C		
    AC_CompanyBond_Private_Domestic
    122010123		AC債務工具投資-普通公司債(民營)                   
    AC_CompanyBond_Private_Domestic_Cost +
    

    3.股票及股權投資-民營企業
    

        原始取得成本1		    
    Stock_Cost
    1200503
    1210103
    15501
    121019901
    150019901
    ' 重新勾稽
    ' --
    1200503
    FVPL_Stock_PreferredStock_Listed_Cost + FVPL_Stock_CommonStock_Listed_Cost + FVPL_Stock_CommonStock_OTC_Cost + FVPL_Stock_CommonStock_Emergin_Cost
    1210103
    FVOCI_Stock_PreferredStock_Cost + FVOCI_Stock_PreferredStock_Listed_Cost + FVOCI_Stock_CommonStock_Listed_Cost + FVOCI_Stock_CommonStock_OTC_Cost + FVOCI_Stock_CommonStock_Emergin_Cost
    ' * 原來公式寫 15501，實際上這是 15001
    15501 
    EquityMethod_Cost +
    121019901
    FVOCI_Equity_Other_Cost +
    150019901
    EquityMethod_Other_Cost +
            
        透過損益按公允價值衡量之金融資產2 A		
    FVPL_Stock
    1200503
    1200703
    ' 重新勾稽
    ' --

    1200503		強制FVPL金融資產-股票    
    FVPL_Stock_PreferredStock_Listed_Cost + FVPL_Stock_CommonStock_Listed_Cost + FVPL_Stock_CommonStock_OTC_Cost + FVPL_Stock_CommonStock_Emergin_Cost                         
    1200703
    FVPL_Stock_CommonStock_Listed_ValuationAdjust + FVPL_Stock_CommonStock_OTC_ValuationAdjust + FVPL_Stock_CommonStock_Emergin_ValuationAdjust + FVPL_Stock_PreferredStock_Listed_ValuationAdjust
    
        透過其他綜合損益按公允價值衡量之金融資產2 B		
    FVOCI_Stock
    1210103
    1210303
    1210199
    1210399
    ' 重新勾稽
    ' --

    1210103		FVOCI權益工具-股票
    FVOCI_Stock_PreferredStock_Cost + FVOCI_Stock_PreferredStock_Listed_Cost + FVOCI_Stock_CommonStock_Listed_Cost + FVOCI_Stock_CommonStock_OTC_Cost + FVOCI_Stock_CommonStock_Emergin_Cost                          
    1210303		FVOCI權益工具評價調整-股票    
    FVOCI_Stock_PreferredStock_Listed_ValuationAdjust + FVOCI_Stock_CommonStock_Listed_ValuationAdjust + FVOCI_Stock_CommonStock_OTC_ValuationAdjust + FVOCI_Stock_CommonStock_Emergin_ValuationAdjust                    
    1210199		FVOCI權益工具-其他        
    FVOCI_Equity_Other_Cost +
    1210399		FVOCI權益工具評價調整-其他                        
    FVOCI_Equity_Other_ValuationAdjust +
            
        按攤銷後成本衡量之債務工具投資2 C		
        		
        採用權益法之投資-淨額2 E		
    EquityMethod_Stock
    
    15001		採用權益法之投資成本 
    EquityMethod_Other_Cost +                              
    15003		加（減）：採用權益法認列之投資權益調整            
    EquityMethod_ValuationAdjust +            
    4.受益憑證-其他		
            
        原始取得成本1		
    AssetCertificate_Cost
    1200505		強制FVPL金融資產-受益憑證              
    
    FVPL_AssetCertificate_Cost +
            
        透過損益按公允價值衡量之金融資產2 A		
    FVPL_AssetCertificate
    1200505		強制FVPL金融資產-受益憑證                         
    1200705		強制FVPL金融資產評價調整-受益憑證                 
    FVPL_AssetCertificate_Cost + FVPL_AssetCertificate_ValuationAdjust

        透過其他綜合損益按公允價值衡量之金融資產2 B		
            
        按攤銷後成本衡量之債務工具投資2 C		
            
    5.新台幣可轉讓定期存單-中央銀行發行		
            
        原始取得成本1		
    NCD_CentralBank_Cost
    121110911		FVOCI債務工具-央行NCD                             
    122010911		AC債務工具投資-央行NCD   
    FVOCI_NCD_CentralBank_Cost + AC_NCD_CentralBank_Cost
            
        透過損益按公允價值衡量之金融資產2 A		
            
        透過其他綜合損益按公允價值衡量之金融資產2 B		
    FVOCI_NCD_CentralBank
    121110911		FVOCI債務工具-央行NCD                             
    121130911		FVOCI債務工具評價調整-央行NCD                     
    FVOCI_NCD_CentralBank_Cost + FVOCI_NCD_CentralBank_ValuationAdjust
            
        按攤銷後成本衡量之債務工具投資2 C		
    AC_NCD_CentralBank
    122010911		AC債務工具投資-央行NCD                            
    122030911		累積減損-AC債務工具投資-央行NCD                   

    AC_NCD_CentralBank_Cost + AC_NCD_CentralBank_ImpairmentLoss
            
    6.商業本票-民營企業		
            
        原始取得成本1		
    CP_Cost
    120050903		強制FVPL金融資產-商業本票                         
    FVPL_CP_Cost + 
    
            
        透過損益按公允價值衡量之金融資產2 A		
    FVPL_CP
    120050903		強制FVPL金融資產-商業本票                         
    120070903		強制FVPL金融資產評價調整-商業本票                 
    FVPL_CP_Cost + FVPL_CP_ValuationAdjust
            
        透過其他綜合損益按公允價值衡量之金融資產2 B		
            
        按攤銷後成本衡量之債務工具投資2 C		
            
    7.國外機構發行-在國外發行-長期債票券6		
            
        原始取得成本1		
    FinancialBond_Domestic_Cost
    140010147		備供出售-金融債券-海外                  
    AFS_FinancialBond_Domestic_Cost +
    
            
        透過損益按公允價值衡量之金融資產2 A		
            
        透過其他綜合損益按公允價值衡量之金融資產2 B		
    AFS_FinancialBond_Domestic
    140010147		備供出售-金融債券-海外                            
    140030147		備供出售評價調整-金融債券-海外                    

    AFS_FinancialBond_Domestic_Cost + AFS_FinancialBond_Domestic_ValuationAdjust
            
        按攤銷後成本衡量之債務工具投資2 C		

    ' END HANDLE
    
    domesticInterestRevenue = Round(domesticInterestRevenue / 1000, 0)
    
    xlsht.Range("FM11_一利息股息收入_利息_其他").Value = foreignInterestRevenue

    xlsht.Range("FM11_三證券投資處分利益_一年期以上之債權證券").Value = gainOnSecuritiesDisposal

    xlsht.Range("FM11_三證券投資處分損失_一年期以上之債權證券").Value = lossOnSecuritiesDisposal

    xlsht.Range("FM11_五證券投資評價及減損迴轉利益_一年期以上之債權證券").Value = reversalImpairmentPL

    xlsht.Range("FM11_五證券投資評價及減損損失_一年期以上之債權證券").Value = valuationImpairmentLoss

    xlsht.Range("FM11_一利息收入_自中華民國境內其他客戶").Value = domesticInterestRevenue
    
    xlsht.Range("T2:T100").NumberFormat = "#,##,##"

    rpt.SetField "FOA", "FM11_一利息股息收入_利息_其他", CStr(foreignInterestRevenue)

    ' 資料一開始要先清掉

    ' --- 【修改2】用 FieldValuePositionMap 批次填 rpt ---
    Dim SetFieldMap As Variant
    Dim tgtSheet As String, srcTag As String, srcVal As Variant
    
    SetFieldMap = GetMapData(gDBPath, reportTitle, "FieldValuePositionMap")
    If Not IsNull(SetFieldMap) And IsArray(SetFieldMap) Then
        For iMap = 0 To UBound(SetFieldMap, 1)
            tgtSheet = CStr(SetFieldMap(iMap, 0))
            srcTag    = CStr(SetFieldMap(iMap, 1))
            
            On Error Resume Next
            srcVal = ThisWorkbook.Sheets(tgtSheet).Range(srcTag).Value
            On Error GoTo 0
            
            rpt.SetField tgtSheet, srcTag, srcVal
        Next iMap
    Else
        WriteLog "無法取得 FieldValuePositionMap for " & reportTitle
    End If    

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

Public Sub Process_FB1()
    '=== Equal Setting ===
    'Fetch Query Access DB table
    Dim dataArr As Variant

    'Declare worksheet and handle data
    Dim xlsht As Worksheet

    Dim i As Integer, j As Integer

    Dim reportTitle As String
    Dim queryTable As String

    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("FB1")
    
    reportTitle = rpt.ReportName

    queryTable = "FB1_OBU_AC4620B_Subtotal"

    ' dataArr = GetAccessDataAsArray(gDBPath, queryTable, gDataMonthString)
    dataArr = GetAccessDataAsArray(gDBPath, queryTable, gDataMonthString)
    
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    xlsht.Range("A:B").ClearContents

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
    ' 更改分頁顏色為黃色(6)
    xlsht.Tab.ColorIndex = 6
End Sub


Public Sub Process_TABLE15A()
    '=== Equal Setting ===
    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("TABLE15A")

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

    Dim importCols As New Collection
    Dim iMap As Long
    For iMap = 0 To UBound(queryMap, 1)
        Dim tblName As String
        Dim startColLetter As String
        Dim dataArr As Variant
        Dim r As Long, c As Long

        tblName = queryMap(iMap, 0)
        startColLetter = queryMap(iMap, 1)
        
        '【修改1】把欄位字母轉成數字並存入 importCols
        Dim startCol As Long
        startCol = xlsht.Range(startColLetter & "1").Column
        importCols.Add startCol        
        
        dataArr = GetAccessDataAsArray(gDBPath, tblName, gDataMonthString)

        '【修改】正確判斷是否有回傳陣列，且至少要有 header（row 0）跟一筆資料（row 1）
        If Not IsArray(dataArr) Or UBound(dataArr, 1) < 1 Then
            WriteLog "資料有誤: " & reportTitle & " | " & tblName & " 無法取得資料或只有欄位名稱"
            GoTo NextMap
        End If

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

    '=== 【新增 Named Range for TABLE15A 區塊】 ===
    ' Row 名稱 & Col 名稱
    Dim rowNames As Variant, colNames As Variant
    rowNames = Array("TABLE15A_001_30天", _
                     "TABLE15A_002_60天", _
                     "TABLE15A_003_90天", _
                     "TABLE15A_004_120天", _
                     "TABLE15A_005_150天", _
                     "TABLE15A_006_180天", _
                     "TABLE15A_007_270天", _
                     "TABLE15A_008_271以上", _
                     "TABLE15A_合計")
    colNames = Array("上月餘額", _
                     "利率", _
                     "金額", _
                     "本月償還", _
                     "本月餘額")

    Dim rIdx As Long, cIdx As Long
    Dim nameTag As String    
    ' rowNames(0) 對應 dataArr row 1，依此類推
    For rIdx = LBound(rowNames) To UBound(rowNames)
        For cIdx = LBound(colNames) To UBound(colNames)
            nameTag = rowNames(rIdx) & "_" & colNames(cIdx)
            
            On Error Resume Next
            xlsht.Range(nameTag).Value = dataArr(rIdx + 1, cIdx)
            If Err.Number <> 0 Then
                WriteLog "Named Range 不存在或設定錯誤: " & nameTag
            End If
            On Error GoTo 0
        Next cIdx
    Next rIdx
    '=== 【新增 Named Range 結束】 ===

    ' 資料一開始要先清掉

    ' --- 【修改2】用 FieldValuePositionMap 批次填 rpt ---
    Dim SetFieldMap As Variant
    Dim tgtSheet As String, srcTag As String, srcVal As Variant
    
    SetFieldMap = GetMapData(gDBPath, reportTitle, "FieldValuePositionMap")
    If Not IsNull(SetFieldMap) And IsArray(SetFieldMap) Then
        For iMap = 0 To UBound(SetFieldMap, 1)
            tgtSheet = CStr(SetFieldMap(iMap, 0))
            srcTag    = CStr(SetFieldMap(iMap, 1))
            
            On Error Resume Next
            srcVal = ThisWorkbook.Sheets(tgtSheet).Range(srcTag).Value
            On Error GoTo 0
            
            rpt.SetField tgtSheet, srcTag, srcVal
        Next iMap
    Else
        WriteLog "無法取得 FieldValuePositionMap for " & reportTitle
    End If

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

Public Sub Process_TABLE15B()
    ' 無資料
    '=== Equal Setting ===
    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("TABLE15B")

    Dim reportTitle As String
    reportTitle = rpt.ReportName    

    Dim xlsht As Worksheet
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    ' xlsht.Range("A:G").ClearContents
    ' xlsht.Range("T2:T100").ClearContents

    ' xlsht.Cells.ClearContents  ' 或只清指定区
    
    ' 呼叫通用函式，並取得 importCols
    Dim importCols As Collection
    Set importCols = ImportQueryTables(gDBPath, xlsht, reportTitle, gDataMonthString)

    If importCols Is Nothing Or importCols.Count = 0 Then
        Exit Sub
    End If

    '--------------
    'Unique Setting
    '--------------    

    Dim lastRow1 As Long, lastRow2 As Long
    Dim rngs1 As Range, rngs2 As Range, rng As Range

    ' --- 【修改1.2】不迴圈，直接取用 importCols(1) 和 importCols(2) ---
    If importCols.Count >= 1 Then
        lastRow1 = xlsht.Cells(xlsht.Rows.Count, importCols(1)).End(xlUp).Row
        Set rngs1 = xlsht.Range(xlsht.Cells(2, importCols(1)), xlsht.Cells(lastRow1, importCols(1)))
        
        For Each rng In rngs1
            Select Case CStr(rng.Value)

            End Select
        Next rng
    End If

    ' 資料一開始要先清掉

    ' --- 【修改2】用 FieldValuePositionMap 批次填 rpt ---
    Dim SetFieldMap As Variant
    Dim tgtSheet As String, srcTag As String, srcVal As Variant
    
    SetFieldMap = GetMapData(gDBPath, reportTitle, "FieldValuePositionMap")
    If Not IsNull(SetFieldMap) And IsArray(SetFieldMap) Then
        For iMap = 0 To UBound(SetFieldMap, 1)
            tgtSheet = CStr(SetFieldMap(iMap, 0))
            srcTag    = CStr(SetFieldMap(iMap, 1))
            
            On Error Resume Next
            srcVal = ThisWorkbook.Sheets(tgtSheet).Range(srcTag).Value
            On Error GoTo 0
            
            rpt.SetField tgtSheet, srcTag, srcVal
        Next iMap
    Else
        WriteLog "無法取得 FieldValuePositionMap for " & reportTitle
    End If    

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

Public Sub Process_TABLE16()
    '=== Equal Setting ===
    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("TABLE16")

    Dim reportTitle As String
    reportTitle = rpt.ReportName    

    Dim xlsht As Worksheet
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    ' xlsht.Range("A:G").ClearContents
    ' xlsht.Range("T2:T100").ClearContents

    ' xlsht.Cells.ClearContents  ' 或只清指定区
    
    ' 呼叫通用函式，並取得 importCols
    Dim importCols As Collection
    Set importCols = ImportQueryTables(gDBPath, xlsht, reportTitle, gDataMonthString)

    If importCols Is Nothing Or importCols.Count = 0 Then
        Exit Sub
    End If

    '--------------
    'Unique Setting
    '--------------
    
    Dim lastRow1 As Long, lastRow2 As Long
    Dim rngs1 As Range, rngs2 As Range, rng As Range

    ' --- 【修改1.2】不迴圈，直接取用 importCols(1) 和 importCols(2) ---
    If importCols.Count >= 1 Then
        lastRow1 = xlsht.Cells(xlsht.Rows.Count, importCols(1)).End(xlUp).Row
        Set rngs1 = xlsht.Range(xlsht.Cells(2, importCols(1)), xlsht.Cells(lastRow1, importCols(1)))
        
        For Each rng In rngs1
            Select Case CStr(rng.Value)
                Case "RP_GovBond_Cost"
                    RP_GovBond_Cost = RP_GovBond_Cost + rng.Offset(0, 1).Value
                Case "AC_CompanyBond_Domestic_ImpairmentLoss"
                    RP_CompanyBond_Cost = RP_CompanyBond_Cost + rng.Offset(0, 1).Value
            End Select
        Next rng
    End If

    ' --- 【修改2】用 FieldValuePositionMap 批次填 rpt ---
    Dim SetFieldMap As Variant
    Dim tgtSheet As String, srcTag As String, srcVal As Variant
    
    SetFieldMap = GetMapData(gDBPath, reportTitle, "FieldValuePositionMap")
    If Not IsNull(SetFieldMap) And IsArray(SetFieldMap) Then
        For iMap = 0 To UBound(SetFieldMap, 1)
            tgtSheet = CStr(SetFieldMap(iMap, 0))
            srcTag    = CStr(SetFieldMap(iMap, 1))
            
            On Error Resume Next
            srcVal = ThisWorkbook.Sheets(tgtSheet).Range(srcTag).Value
            On Error GoTo 0
            
            rpt.SetField tgtSheet, srcTag, srcVal
        Next iMap
    Else
        WriteLog "無法取得 FieldValuePositionMap for " & reportTitle
    End If

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

Public Sub Process_TABLE20()
    '=== Equal Setting ===
    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("TABLE20")

    Dim reportTitle As String
    reportTitle = rpt.ReportName

    Dim xlsht As Worksheet
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    ' xlsht.Range("A:G").ClearContents
    ' xlsht.Range("T2:T100").ClearContents

    ' xlsht.Cells.ClearContents  ' 或只清指定区

    ' 呼叫通用函式，並取得 importCols
    Dim importCols As Collection
    Set importCols = ImportQueryTables(gDBPath, xlsht, reportTitle, gDataMonthString)

    If importCols Is Nothing Or importCols.Count = 0 Then
        Exit Sub
    End If
    
    '--------------
    'Unique Setting
    '--------------

    Dim RP_GovBond_Cost As Double: RP_GovBond_Cost = 0
    Dim RP_CompanyBond_Cost As Double: RP_CompanyBond_Cost = 0
    Dim RP_CP_Cost As Double: RP_CP_Cost = 0

    Dim lastRow1 As Long, lastRow2 As Long
    Dim rngs1 As Range, rngs2 As Range, rng As Range

    ' --- 【修改1.2】不迴圈，直接取用 importCols(1) 和 importCols(2) ---
    If importCols.Count >= 1 Then
        lastRow1 = xlsht.Cells(xlsht.Rows.Count, importCols(1)).End(xlUp).Row
        Set rngs1 = xlsht.Range(xlsht.Cells(2, importCols(1)), xlsht.Cells(lastRow1, importCols(1)))
        
        For Each rng In rngs1
            Select Case CStr(rng.Value)
                Case "RP_GovBond_Cost"
                    RP_GovBond_Cost = RP_GovBond_Cost + rng.Offset(0, 1).Value
                Case "AC_CompanyBond_Domestic_ImpairmentLoss"
                    RP_CompanyBond_Cost = RP_CompanyBond_Cost + rng.Offset(0, 1).Value
            End Select
        Next rng
    End If
    
    ' If importCols.Count >= 2 Then
    '     lastRow2 = xlsht.Cells(xlsht.Rows.Count, importCols(2)).End(xlUp).Row
    '     Set rngs2 = xlsht.Range(xlsht.Cells(2, importCols(2)), xlsht.Cells(lastRow2, importCols(2)))
        
    '     For Each rng In rngs2
    '         ' 如果第二筆表也有需要累計的 tag，可以在這裡加
    '         Select Case CStr(rng.Value)
    '             Case "RP_GovBond_Cost"
    '                 RP_GovBond_Cost = RP_GovBond_Cost + rng.Offset(0, 1).Value
    '             Case "AC_CompanyBond_Domestic_ImpairmentLoss"
    '                 RP_CompanyBond_Cost = RP_CompanyBond_Cost + rng.Offset(0, 1).Value
    '         End Select
    '     Next rng
    ' End If

    ' HANDLE方式
    - 民營企業		
        - 其他到期日		
            - 公債		
    225010101		附買回票券及債券負債-公債  
    RP_GovBond_Cost +                       
            - 公司債		
    225010105		附買回票券及債券負債-公司債    
    RP_CompanyBond_Cost +                   
            - 商業本票		
    225010303		#N/A
    找不到這個ACCOUNT CODE

    Table20_0200_二公債_民營企業_其他到期日
    Table20_0300_三公司債_民營企業_其他到期日
    Table20_0400_四商業本票_民營企業_其他到期日

    ' END HANDLE
    
    RP_GovBond_Cost = Round(RP_GovBond_Cost / 1000, 0)
    RP_CompanyBond_Cost = Round(RP_CompanyBond_Cost / 1000, 0)
    RP_CP_Cost = Round(RP_CP_Cost / 1000, 0)
    
    xlsht.Range("Table20_0200_二公債_民營企業_其他到期日").Value = RP_GovBond_Cost
    xlsht.Range("Table20_0300_三公司債_民營企業_其他到期日").Value = RP_CompanyBond_Cost
    xlsht.Range("Table20_0400_四商業本票_民營企業_其他到期日").Value = RP_CP_Cost
    
    ' 資料一開始要先清掉

    ' --- 【修改2】用 FieldValuePositionMap 批次填 rpt ---
    Dim SetFieldMap As Variant
    Dim tgtSheet As String, srcTag As String, srcVal As Variant
    
    SetFieldMap = GetMapData(gDBPath, reportTitle, "FieldValuePositionMap")
    If Not IsNull(SetFieldMap) And IsArray(SetFieldMap) Then
        For iMap = 0 To UBound(SetFieldMap, 1)
            tgtSheet = CStr(SetFieldMap(iMap, 0))
            srcTag    = CStr(SetFieldMap(iMap, 1))
            
            On Error Resume Next
            srcVal = ThisWorkbook.Sheets(tgtSheet).Range(srcTag).Value
            On Error GoTo 0
            
            rpt.SetField tgtSheet, srcTag, srcVal
        Next iMap
    Else
        WriteLog "無法取得 FieldValuePositionMap for " & reportTitle
    End If

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

Public Function ImportQueryTables(ByVal DBPath As String, _
                                  ByVal xlsht As Worksheet, _
                                  ByVal reportName As String, _
                                  ByVal dataMonth As String) As Collection
    Dim queryMap As Variant
    Dim importCols As New Collection
    Dim iMap As Long
    
    ' 1. 取配置
    queryMap = GetMapData(DBPath, reportName, "QueryTableMap")
    If Not IsArray(queryMap) Or UBound(queryMap, 1) < 0 Then
        WriteLog "未在 QueryTableMap 找到 " & reportName & " 的任何配置"
        Exit Function
    End If
    
    ' 2. 迭代每張子表
    For iMap = 0 To UBound(queryMap, 1)
        Dim tblName As String
        Dim colLetter As String
        Dim startCol As Long
        Dim dataArr As Variant
        Dim r As Long, c As Long
        
        tblName = CStr(queryMap(iMap, 0))
        colLetter = CStr(queryMap(iMap, 1))
        ' A→數字
        startCol = xlsht.Range(colLetter & "1").Column
        importCols.Add startCol
        
        ' 取資料
        dataArr = GetAccessDataAsArray(DBPath, tblName, dataMonth)
        If Not IsArray(dataArr) Or UBound(dataArr, 1) < 1 Then
            WriteLog "資料有誤: " & reportName & " | " & tblName & " 無法取得資料或只有欄位名稱"
            GoTo NextMap
        End If
        
        ' 貼到工作表
        For r = 0 To UBound(dataArr, 1)
            For c = 0 To UBound(dataArr, 2)
                xlsht.Cells(r + 1, startCol + c).Value = dataArr(r, c)
            Next c
        Next r
NextMap:
    Next iMap
    
    ' 回傳已記錄的所有起始欄位
    Set ImportQueryTables = importCols
End Function

Public Sub Process_TABLE22()
    '=== Equal Setting ===
    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("TABLE22")

    Dim reportTitle As String
    reportTitle = rpt.ReportName    

    Dim xlsht As Worksheet
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    ' xlsht.Range("A:G").ClearContents
    ' xlsht.Range("T2:T100").ClearContents

    ' xlsht.Cells.ClearContents  ' 或只清指定区
    
    ' 呼叫通用函式，並取得 importCols
    Dim importCols As Collection
    Set importCols = ImportQueryTables(gDBPath, xlsht, reportTitle, gDataMonthString)

    If importCols Is Nothing Or importCols.Count = 0 Then
        Exit Sub
    End If

    '--------------
    'Unique Setting
    '--------------
            
    Dim Bank_FaceValue As Double: Bank_FaceValue = 0
    Dim BillSecurity_FaceValue As Double: BillSecurity_FaceValue = 0
    Dim Other_FaceValue As Double: Other_FaceValue = 0

    Dim lastRow1 As Long, lastRow2 As Long
    Dim rngs1 As Range, rngs2 As Range, rng As Range

    ' --- 【修改1.2】不迴圈，直接取用 importCols(1) 和 importCols(2) ---
    If importCols.Count >= 1 Then
        lastRow1 = xlsht.Cells(xlsht.Rows.Count, importCols(1)).End(xlUp).Row
        Set rngs1 = xlsht.Range(xlsht.Cells(2, importCols(1)), xlsht.Cells(lastRow1, importCols(1)))
        
        For Each rng In rngs1
            Select Case CStr(rng.Value)
                Case "銀行"
                    Bank_FaceValue = Bank_FaceValue + rng.Offset(0, 1).Value
                Case "票券"
                    BillSecurity_FaceValue = BillSecurity_FaceValue + rng.Offset(0, 1).Value
                Case "其他"
                    Other_FaceValue = Other_FaceValue + rng.Offset(0, 1).Value                    
            End Select
        Next rng
    End If

    ' HANDLE方式

    Table22_001銀行_融資性商業本票
    Table22_002票券金融公司_融資性商業本票
    Table22_006民營企業_融資性商業本票

    ' END HANDLE
    
    Bank_FaceValue = Round(Bank_FaceValue / 1000000, 0)
    BillSecurity_FaceValue = Round(BillSecurity_FaceValue / 1000000, 0)
    Other_FaceValue = Round(Other_FaceValue / 1000000, 0)
    
    xlsht.Range("Table22_001銀行_融資性商業本票").Value = Bank_FaceValue
    xlsht.Range("Table22_002票券金融公司_融資性商業本票").Value = BillSecurity_FaceValue
    xlsht.Range("Table22_006民營企業_融資性商業本票").Value = Other_FaceValue
    
    ' 資料一開始要先清掉

    ' --- 【修改2】用 FieldValuePositionMap 批次填 rpt ---
    Dim SetFieldMap As Variant
    Dim tgtSheet As String, srcTag As String, srcVal As Variant
    
    SetFieldMap = GetMapData(gDBPath, reportTitle, "FieldValuePositionMap")
    If Not IsNull(SetFieldMap) And IsArray(SetFieldMap) Then
        For iMap = 0 To UBound(SetFieldMap, 1)
            tgtSheet = CStr(SetFieldMap(iMap, 0))
            srcTag    = CStr(SetFieldMap(iMap, 1))
            
            On Error Resume Next
            srcVal = ThisWorkbook.Sheets(tgtSheet).Range(srcTag).Value
            On Error GoTo 0
            
            rpt.SetField tgtSheet, srcTag, srcVal
        Next iMap
    Else
        WriteLog "無法取得 FieldValuePositionMap for " & reportTitle
    End If

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

' ****這張表要問一下怎麼做    
Public Sub Process_TABLE23()
    '=== Equal Setting ===
    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("TABLE23")

    Dim reportTitle As String
    reportTitle = rpt.ReportName    

    Dim xlsht As Worksheet
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    ' xlsht.Range("A:G").ClearContents
    ' xlsht.Range("T2:T100").ClearContents

    ' xlsht.Cells.ClearContents  ' 或只清指定区
    
    ' 呼叫通用函式，並取得 importCols
    Dim importCols As Collection
    Set importCols = ImportQueryTables(gDBPath, xlsht, reportTitle, gDataMonthString)

    If importCols Is Nothing Or importCols.Count = 0 Then
        Exit Sub
    End If

    '--------------
    'Unique Setting
    '--------------

    Dim Days30_Interest As Double: Days30_Interest = 0
    Dim Days90_Interest As Double: Days90_Interest = 0
    Dim Days180_Interest As Double: Days180_Interest = 0
    Dim Days270_Interest As Double: Days270_Interest = 0
    Dim Days365_Interest As Double: Days365_Interest = 0

    Dim lastRow1 As Long, lastRow2 As Long
    Dim rngs1 As Range, rngs2 As Range, rng As Range

    ' --- 【修改1.2】不迴圈，直接取用 importCols(1) 和 importCols(2) ---
    If importCols.Count >= 1 Then
        lastRow1 = xlsht.Cells(xlsht.Rows.Count, importCols(1)).End(xlUp).Row
        Set rngs1 = xlsht.Range(xlsht.Cells(2, importCols(1)), xlsht.Cells(lastRow1, importCols(1)))
        
        For Each rng In rngs1
            Select Case CStr(rng.Value)
                Case "0-30天"
                    Days30_Interest = (rng.Offset(0, 1).Value / rng.Offset(0, 2).Value)
                Case "31-90天"
                    Days90_Interest = (rng.Offset(0, 1).Value / rng.Offset(0, 2).Value)
                Case "91-180天"
                    Days180_Interest = (rng.Offset(0, 1).Value / rng.Offset(0, 2).Value)
                Case "181-270天"
                    Days270_Interest = (rng.Offset(0, 1).Value / rng.Offset(0, 2).Value)
                Case "271-365天"
                    Days365_Interest = (rng.Offset(0, 1).Value / rng.Offset(0, 2).Value)                
            End Select
        Next rng
    End If    

    Days30_Interest = 0
    Days90_Interest = 0
    Days180_Interest = 0
    Days270_Interest = 0
    Days365_Interest = 0
  
    xlsht.Range("Table23_001_加權買入利率_融資性商業本票_30天").Value = Days30_Interest
    xlsht.Range("Table23_002_加權買入利率_融資性商業本票_90天").Value = Days90_Interest
    xlsht.Range("Table23_003_加權買入利率_融資性商業本票_180天").Value = Days180_Interest
    xlsht.Range("Table23_004_加權買入利率_融資性商業本票_270天").Value = Days270_Interest
    xlsht.Range("Table23_005_加權買入利率_融資性商業本票_365天").Value = Days365_Interest
    
    ' 資料一開始要先清掉

    ' --- 【修改2】用 FieldValuePositionMap 批次填 rpt ---
    Dim SetFieldMap As Variant
    Dim tgtSheet As String, srcTag As String, srcVal As Variant
    
    SetFieldMap = GetMapData(gDBPath, reportTitle, "FieldValuePositionMap")
    If Not IsNull(SetFieldMap) And IsArray(SetFieldMap) Then
        For iMap = 0 To UBound(SetFieldMap, 1)
            tgtSheet = CStr(SetFieldMap(iMap, 0))
            srcTag    = CStr(SetFieldMap(iMap, 1))
            
            On Error Resume Next
            srcVal = ThisWorkbook.Sheets(tgtSheet).Range(srcTag).Value
            On Error GoTo 0
            
            rpt.SetField tgtSheet, srcTag, srcVal
        Next iMap
    Else
        WriteLog "無法取得 FieldValuePositionMap for " & reportTitle
    End If

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

' ******修改到這邊
' ******修改到這邊
' ******修改到這邊
' ******修改到這邊
' ******修改到這邊

' 詢問怎麼判斷
Public Sub Process_TABLE24()
    '=== Equal Setting ===
    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("TABLE24")

    Dim reportTitle As String
    reportTitle = rpt.ReportName

    Dim xlsht As Worksheet
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    ' xlsht.Range("A:G").ClearContents
    ' xlsht.Range("T2:T100").ClearContents

    ' xlsht.Cells.ClearContents  ' 或只清指定区

    ' 呼叫通用函式，並取得 importCols
    Dim importCols As Collection
    Set importCols = ImportQueryTables(gDBPath, xlsht, reportTitle, gDataMonthString)

    If importCols Is Nothing Or importCols.Count = 0 Then
        Exit Sub
    End If
    
    '--------------
    'Unique Setting
    '--------------

    Dim RP_GovBond_Cost As Double: RP_GovBond_Cost = 0


    Dim lastRow1 As Long, lastRow2 As Long
    Dim rngs1 As Range, rngs2 As Range, rng As Range

    ' --- 【修改1.2】不迴圈，直接取用 importCols(1) 和 importCols(2) ---
    If importCols.Count >= 1 Then
        lastRow1 = xlsht.Cells(xlsht.Rows.Count, importCols(1)).End(xlUp).Row
        Set rngs1 = xlsht.Range(xlsht.Cells(2, importCols(1)), xlsht.Cells(lastRow1, importCols(1)))
        
        For Each rng In rngs1
            Select Case CStr(rng.Value)
                Case "RP_GovBond_Cost"
                    RP_GovBond_Cost = RP_GovBond_Cost + rng.Offset(0, 1).Value
                Case "AC_CompanyBond_Domestic_ImpairmentLoss"
                    RP_CompanyBond_Cost = RP_CompanyBond_Cost + rng.Offset(0, 1).Value
            End Select
        Next rng
    End If
    
    
    RP_GovBond_Cost = Round(RP_GovBond_Cost / 1000, 0)

    
    xlsht.Range("Table20_0200_二公債_民營企業_其他到期日").Value = RP_GovBond_Cost
    
    ' 資料一開始要先清掉

    ' --- 【修改2】用 FieldValuePositionMap 批次填 rpt ---
    Dim SetFieldMap As Variant
    Dim tgtSheet As String, srcTag As String, srcVal As Variant
    
    SetFieldMap = GetMapData(gDBPath, reportTitle, "FieldValuePositionMap")
    If Not IsNull(SetFieldMap) And IsArray(SetFieldMap) Then
        For iMap = 0 To UBound(SetFieldMap, 1)
            tgtSheet = CStr(SetFieldMap(iMap, 0))
            srcTag    = CStr(SetFieldMap(iMap, 1))
            
            On Error Resume Next
            srcVal = ThisWorkbook.Sheets(tgtSheet).Range(srcTag).Value
            On Error GoTo 0
            
            rpt.SetField tgtSheet, srcTag, srcVal
        Next iMap
    Else
        WriteLog "無法取得 FieldValuePositionMap for " & reportTitle
    End If

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

Public Sub Process_TABLE27()
    '=== Equal Setting ===
    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("TABLE27")

    Dim reportTitle As String
    reportTitle = rpt.ReportName

    Dim xlsht As Worksheet
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    ' xlsht.Range("A:G").ClearContents
    ' xlsht.Range("T2:T100").ClearContents

    ' xlsht.Cells.ClearContents  ' 或只清指定区

    ' 呼叫通用函式，並取得 importCols
    Dim importCols As Collection
    Set importCols = ImportQueryTables(gDBPath, xlsht, reportTitle, gDataMonthString)

    If importCols Is Nothing Or importCols.Count = 0 Then
        Exit Sub
    End If
    
    '--------------
    'Unique Setting
    '--------------

    Dim RP_GovBond_Cost As Double: RP_GovBond_Cost = 0

    Dim lastRow1 As Long, lastRow2 As Long
    Dim rngs1 As Range, rngs2 As Range, rng As Range

    ' --- 【修改1.2】不迴圈，直接取用 importCols(1) 和 importCols(2) ---
    If importCols.Count >= 1 Then
        lastRow1 = xlsht.Cells(xlsht.Rows.Count, importCols(1)).End(xlUp).Row
        Set rngs1 = xlsht.Range(xlsht.Cells(2, importCols(1)), xlsht.Cells(lastRow1, importCols(1)))
        
        For Each rng In rngs1
            Select Case CStr(rng.Value)
                Case "RP_GovBond_Cost"
                    RP_GovBond_Cost = RP_GovBond_Cost + rng.Offset(0, 1).Value
                Case "AC_CompanyBond_Domestic_ImpairmentLoss"
                    RP_CompanyBond_Cost = RP_CompanyBond_Cost + rng.Offset(0, 1).Value
            End Select
        Next rng
    End If
    
    RP_GovBond_Cost = Round(RP_GovBond_Cost / 1000, 0)
    RP_CompanyBond_Cost = Round(RP_CompanyBond_Cost / 1000, 0)
    RP_CP_Cost = Round(RP_CP_Cost / 1000, 0)
    
    xlsht.Range("Table20_0200_二公債_民營企業_其他到期日").Value = RP_GovBond_Cost
    xlsht.Range("Table20_0300_三公司債_民營企業_其他到期日").Value = RP_CompanyBond_Cost
    xlsht.Range("Table20_0400_四商業本票_民營企業_其他到期日").Value = RP_CP_Cost
    
    ' 資料一開始要先清掉

    ' --- 【修改2】用 FieldValuePositionMap 批次填 rpt ---
    Dim SetFieldMap As Variant
    Dim tgtSheet As String, srcTag As String, srcVal As Variant
    
    SetFieldMap = GetMapData(gDBPath, reportTitle, "FieldValuePositionMap")
    If Not IsNull(SetFieldMap) And IsArray(SetFieldMap) Then
        For iMap = 0 To UBound(SetFieldMap, 1)
            tgtSheet = CStr(SetFieldMap(iMap, 0))
            srcTag    = CStr(SetFieldMap(iMap, 1))
            
            On Error Resume Next
            srcVal = ThisWorkbook.Sheets(tgtSheet).Range(srcTag).Value
            On Error GoTo 0
            
            rpt.SetField tgtSheet, srcTag, srcVal
        Next iMap
    Else
        WriteLog "無法取得 FieldValuePositionMap for " & reportTitle
    End If

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

Public Sub Process_TABLE36()
    '=== Equal Setting ===
    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("TABLE36")

    Dim reportTitle As String
    reportTitle = rpt.ReportName

    Dim xlsht As Worksheet
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    ' xlsht.Range("A:G").ClearContents
    ' xlsht.Range("T2:T100").ClearContents

    ' xlsht.Cells.ClearContents  ' 或只清指定区

    ' 呼叫通用函式，並取得 importCols
    Dim importCols As Collection
    Set importCols = ImportQueryTables(gDBPath, xlsht, reportTitle, gDataMonthString)

    If importCols Is Nothing Or importCols.Count = 0 Then
        Exit Sub
    End If
    
    '--------------
    'Unique Setting
    '--------------

    Dim RsByBank_Amount As Double: RsByBank_Amount = 0
    Dim RsByBillInsti_Amount As Double: RsByBillInsti_Amount = 0

    Dim BillByBank_Amount As Double: BillByBank_Amount = 0
    Dim BillByBillInsti_Amount As Double: BillByBillInsti_Amount = 0

    Dim lastRow1 As Long, lastRow2 As Long
    Dim rngs1 As Range, rngs2 As Range, rng As Range

    ' --- 【修改1.2】不迴圈，直接取用 importCols(1) 和 importCols(2) ---
    If importCols.Count >= 1 Then
        lastRow1 = xlsht.Cells(xlsht.Rows.Count, importCols(1)).End(xlUp).Row
        Set rngs1 = xlsht.Range(xlsht.Cells(2, importCols(1)), xlsht.Cells(lastRow1, importCols(1)))
        
        For Each rng In rngs1
            Select Case CStr(rng.Value)
                Case "銀行"
                    RsByBank_Amount = RsByBank_Amount + rng.Offset(0, 1).Value
                Case "票券"
                    RsByBillInsti_Amount = RsByBillInsti_Amount + rng.Offset(0, 1).Value
            End Select
        Next rng
    End If
    
    If importCols.Count >= 2 Then
        lastRow2 = xlsht.Cells(xlsht.Rows.Count, importCols(2)).End(xlUp).Row
        Set rngs2 = xlsht.Range(xlsht.Cells(2, importCols(2)), xlsht.Cells(lastRow2, importCols(2)))
        
        For Each rng In rngs2
            ' 如果第二筆表也有需要累計的 tag，可以在這裡加
            Select Case CStr(rng.Value)
                Case "銀行"
                    BillByBank_Amount = BillByBank_Amount + rng.Offset(0, 1).Value
                Case "票券"
                    BillByBillInsti_Amount = BillByBillInsti_Amount + rng.Offset(0, 1).Value
            End Select
        Next rng
    End If

    RsByBank_Amount = Round(RsByBank_Amount / 1000, 0)
    RsByBillInsti_Amount = Round(RsByBillInsti_Amount / 1000, 0)
    BillByBank_Amount = Round(BillByBank_Amount / 1000, 0)
    BillByBillInsti_Amount = Round(BillByBillInsti_Amount / 1000, 0)
    
    xlsht.Range("Table36_0200_公債_民營企業").Value = RsByBillInsti_Amount
    xlsht.Range("Table36_0200_公債_貨幣機構").Value = RsByBank_Amount
    xlsht.Range("Table36_0400_商業本票_民營企業").Value = BillByBillInsti_Amount
    xlsht.Range("Table36_0400_商業本票_貨幣機構").Value = BillByBank_Amount
    
    ' 資料一開始要先清掉

    ' --- 【修改2】用 FieldValuePositionMap 批次填 rpt ---
    Dim SetFieldMap As Variant
    Dim tgtSheet As String, srcTag As String, srcVal As Variant
    
    SetFieldMap = GetMapData(gDBPath, reportTitle, "FieldValuePositionMap")
    If Not IsNull(SetFieldMap) And IsArray(SetFieldMap) Then
        For iMap = 0 To UBound(SetFieldMap, 1)
            tgtSheet = CStr(SetFieldMap(iMap, 0))
            srcTag    = CStr(SetFieldMap(iMap, 1))
            
            On Error Resume Next
            srcVal = ThisWorkbook.Sheets(tgtSheet).Range(srcTag).Value
            On Error GoTo 0
            
            rpt.SetField tgtSheet, srcTag, srcVal
        Next iMap
    Else
        WriteLog "無法取得 FieldValuePositionMap for " & reportTitle
    End If

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


- 透過損益按公允價值衡量A1		
    - 投資成本		
        - 政府公債A		
120050101   強制FVPL金融資產-公債-中央政府(我國)	
120050103   強制FVPL金融資產-公債-地方政府(我國)
FVPL_GovBond_Domestic_Cost +
        - 公司債B		
120050121	強制FVPL金融資產-普通公司債(公營)	
120050123   強制FVPL金融資產-普通公司債(民營)                 		
120057307   強制FVPL金融資產-衍生性ＳＷＡＰ-ASS 

FVPL_CompanyBond_Domestic_Cost + FVPL_SWAP_Cost

        - 金融債券C市                              		
        - 權益證券投資D		
120050301   強制FVPL金融資產-普通股-上市公司                  		
120050302   強制FVPL金融資產-普通股-上櫃公司                  		
120050311   強制FVPL金融資產-特別股-上  

FVPL_Stock_CommonStock_Listed_Cost + FVPL_Stock_CommonStock_OTC_Cost + FVPL_Stock_PreferredStock_Listed_Cost

        - 結構型商品E		
        - 證券化商品F		
        - 其它投資部位G		
120050501	強制FVPL金融資產-受益憑證                         
120050903	強制FVPL金融資產-商業本票                         

FVPL_AssetCertificate_Cost + FVPL_CP_Cost

      		
        - 央行可轉讓定期存單H		
    - 帳面價值 = 投資成本 + 以下		
        - 政府公債A		
120070101   強制FVPL金融資產評價調整-公債-中央(我國)          		
120070103   強制FVPL金融資產評價調整-公債-地方(我國)                 		
FVPL_GovBond_Domestic_Cost + FVPL_GovBond_Domestic_ValuationAdjust

        - 公司債B		
120070121   強制FVPL金融資產評價調整-普通公司債(公營)         		
120070123   強制FVPL金融資產評價調整-普通公司債(民營)         		
120077307   強制FVPL金融資產評價調整-ＳＷＡＰ-ASS             		
120079031   強制FVPL金融資產評價調整-CVA-SWAP-ASS             	

FVPL_CompanyBond_Domestic_Cost + FVPL_SWAP_Cost + 
FVPL_CompanyBond_Domestic_ValuationAdjust + FVPL_SWAP_ValuationAdjust + FVPL_CVASWAP_ValuationAdjust


        - 金融債券C		
        - 權益證券投資D		
120070301   強制FVPL金融資產評價調整-上市公司                 		
120070302   強制FVPL金融資產評價調整-上櫃公司                 		
120070311   強制FVPL金融資產評價調整-特別股-上市                      		

FVPL_Stock_CommonStock_Listed_Cost + FVPL_Stock_CommonStock_OTC_Cost + FVPL_Stock_PreferredStock_Listed_Cost + 
FVPL_Stock_CommonStock_Listed_ValuationAdjust + FVPL_Stock_CommonStock_OTC_ValuationAdjust + FVPL_Stock_CommonStock_Emergin_ValuationAdjust + FVPL_Stock_PreferredStock_Listed_ValuationAdjust


        - 結構型商品E		
        - 證券化商品F		
        - 其它投資部位G		
120070501   強制FVPL金融資產評價調整-受益憑證                 		
120070903   強制FVPL金融資產評價調整-商業本票                 		

FVPL_AssetCertificate_ValuationAdjust + FVPL_CP_ValuationAdjust          

        - 央行可轉讓定期存單H		           		
		
    		
- 透過其他綜合損益按公允價值衡量A6		
    - 投資成本		
        - 政府公債A		
            121110101   FVOCI債務工具-公債-中央政府(我國)                 		
            121110103   FVOCI債務工具-公債-地方政府(我國)
FVOCI_GovBond_Domestic_Cost +

                		
        - 公司債B		
            121110121   FVOCI債務工具-普通公司債（公營）                  		
            121110123   FVOCI債務工具-普通公司債（民營） 
FVOCI_CompanyBond_Domestic_Cost +
                 		
        - 金融債券C		
        - 權益證券投資D		
            121010301   FVOCI權益工具-普通股-上市公司                     		
            121010302   FVOCI權益工具-普通股-上櫃公司                     		
            121019901   FVOCI權益工具-其他                               

FVOCI_Stock_CommonStock_Listed_Cost + FVOCI_Stock_CommonStock_OTC_Cost + FVOCI_Equity_Other_Cost
 		
        - 結構型商品E		
        - 證券化商品F    		
        - 其它投資部位G		
        - 央行可轉讓定期存單H		
            121110911   FVOCI債務工具-央行NCD  
FVOCI_NCD_CentralBank_Cost +                          
  		
    - 帳面價值		
        - 政府公債A		
            121130101   FVOCI債務工具評價調整-公債-中央政府               		
            121130103   FVOCI債務工具評價調整-公債-地方政府                       		
            325350103   FVOCI債務備抵損失-公債-地方政府(我國)           

FVOCI_GovBond_Domestic_Cost + FVOCI_GovBond_Domestic_ValuationAdjust + FVOCI_GovBond_Domestic_ImpairmentAllowance
  		
        - 公司債B		
            121130121   FVOCI債務工具評價調整-普通公司債（公營)           		
            121130123   FVOCI債務工具評價調整-普通公司債（民營)                       		
            325350121   FVOCI債務備抵損失-普通公司債(公營)                		
            325350123   FVOCI債務備抵損失-普通公司債(民營)  

FVOCI_CompanyBond_Domestic_Cost + FVOCI_CompanyBond_Domestic_ValuationAdjust + FVOCI_CompanyBond_Domestic_ImpairmentAllowance
             		
        - 金融債券C		
        - 權益證券投資D		
            121030301   FVOCI權益工具評價調整-普通股-上市                 		
            121030302   FVOCI權益工具評價調整-普通股-上櫃                 		
            121039901   FVOCI權益工具評價調整-其他                         

FVOCI_Stock_CommonStock_Listed_Cost + FVOCI_Stock_CommonStock_OTC_Cost + FVOCI_Equity_Other_Cost +
FVOCI_Stock_CommonStock_Listed_ValuationAdjust + FVOCI_Stock_CommonStock_OTC_ValuationAdjust + FVOCI_Equity_Other_ValuationAdjust
       		
        - 結構型商品E		
        - 證券化商品F    		
        - 其它投資部位G		
        - 央行可轉讓定期存單H		
            121130911   FVOCI債務工具評價調整-央行NCD                     		
            325350911   FVOCI債務備抵損失-央行NCD                         		
FVOCI_NCD_CentralBank_Cost + FVOCI_NCD_CentralBank_ValuationAdjust + FVOCI_NCD_CentralBank_ImpairmentAllowance
		
    		
- 按攤銷後成本衡量A7		
    - 投資成本		
        - 政府公債A		
            122010101   AC債務工具投資-公債-中央政府(我國)                		
            122010103   AC債務工具投資-公債-地方政府(我國)                 

AC_GovBond_Domestic_Cost + 
               		
        - 公司債B		
            122010121   AC債務工具投資-普通公司債(公營)                   		
            122010123   AC債務工具投資-普通公司債(民營)                     

AC_CompanyBond_Domestic_Cost + 
           		
        - 金融債券C		
        - 權益證券投資D		
        - 結構型商品E		
        - 證券化商品F		
        - 其它投資部位G		
        - 央行可轉讓定期存單H        		
            122010911   AC債務工具投資-央行NCD  
AC_NCD_CentralBank_Cost +

     		
    - 帳面價值		
        - 政府公債A		
            122030101   累積減損-AC債務工具投資-公債-中央                 		
            122030103   累積減損-AC債務工具投資-公債-地方 

AC_GovBond_Domestic_Cost + AC_GovBond_Domestic_ImpairmentLoss
                        		
        - 公司債B		
            122030121   累積減損-累積減損-AC債務工具投資-普通公司(公營)   		
            122030123   累積減損-AC債務工具投資-普通公司(民營)            		
AC_CompanyBond_Domestic_Cost + AC_CompanyBond_Domestic_ImpairmentLoss

        - 金融債券C		
        - 權益證券投資D		
        - 結構型商品E		
        - 證券化商品F    		
        - 其它投資部位G		
        - 央行可轉讓定期存單H		
            122030911   累積減損-AC債務工具投資-央行NCD                   		
AC_NCD_CentralBank_Cost + AC_NCD_CentralBank_ImpairmentLoss


' ===========================無使用到的變數


AFS_FinancialBond_Domestic_Cost
AFS_FinancialBond_Domestic_ValuationAdjust

EquityMethod_Other_Cost



' FVOCI_Stock_特別股_Cost
FVOCI_Stock_PreferredStock_Cost
' FVOCI_Stock_特別股_上市_Cost
FVOCI_Stock_PreferredStock_Listed_Cost
' FVOCI_Stock_特別股_上市_ValuationAdjust
FVOCI_Stock_PreferredStock_Listed_ValuationAdjust


' FVOCI_Stock_普通股_興櫃_Cost
FVOCI_Stock_CommonStock_Emergin_Cost
' FVOCI_Stock_普通股_興櫃_ValuationAdjust
FVOCI_Stock_CommonStock_Emergin_ValuationAdjust

FVPL_Debt_FRA_ValuationAdjust
FVPL_Debt_Future_ValuationAdjust
FVPL_Debt_NDF_ValuationAdjust
FVPL_Future_ValuationAdjust


' FVPL_Stock_普通股_興櫃_Cost
FVPL_Stock_CommonStock_Emergin_Cost

' ===========================無使用到的變數



'尚無有交易紀錄
Public Sub Process_AI233()
    '=== Equal Setting ===
    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("AI233")

    Dim reportTitle As String
    reportTitle = rpt.ReportName

    Dim xlsht As Worksheet
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    ' xlsht.Range("A:G").ClearContents
    ' xlsht.Range("T2:T100").ClearContents

    ' xlsht.Cells.ClearContents  ' 或只清指定区

    ' 呼叫通用函式，並取得 importCols
    Dim importCols As Collection
    Set importCols = ImportQueryTables(gDBPath, xlsht, reportTitle, gDataMonthString)

    If importCols Is Nothing Or importCols.Count = 0 Then
        Exit Sub
    End If
    
    '--------------
    'Unique Setting
    '--------------

    Dim RP_GovBond_Cost As Double: RP_GovBond_Cost = 0
    Dim RP_CompanyBond_Cost As Double: RP_CompanyBond_Cost = 0
    Dim RP_CP_Cost As Double: RP_CP_Cost = 0

    ' ==============

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
    FVOCI_CompanyBond_Domestic_ImpairmentAllowance
    Dim FVOCI_CompanyBond_Domestic_ValuationAdjust As Double


    Dim FVOCI_GovBond_Domestic_Cost As Double
    FVOCI_GovBond_Domestic_ImpairmentAllowance
    Dim FVOCI_GovBond_Domestic_ValuationAdjust As Double

    Dim FVOCI_NCD_CentralBank_Cost As Double
    FVOCI_NCD_CentralBank_ImpairmentAllowance
    Dim FVOCI_NCD_CentralBank_ValuationAdjust As Double

    ' FVOCI_Stock_特別股_Cost
    Dim FVOCI_Stock_PreferredStock_Cost As Double
    ' FVOCI_Stock_特別股_上市_Cost
    Dim FVOCI_Stock_PreferredStock_Listed_Cost As Double
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

    FVPL_CVASWAP_ValuationAdjust
    FVPL_Debt_FRA_ValuationAdjust
    FVPL_Debt_Future_ValuationAdjust
    FVPL_Debt_NDF_ValuationAdjust
    FVPL_Future_ValuationAdjust

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

    FVPL_SWAP_Cost
    FVPL_SWAP_ValuationAdjust






    FVPL_AssetCertificate_Cost
    FVPL_AssetCertificate_ValuationAdjust
    
    FVPL_CP_Cost
    FVPL_CP_ValuationAdjust  

    Dim lastRow1 As Long, lastRow2 As Long
    Dim rngs1 As Range, rngs2 As Range, rng As Range

    ' --- 【修改1.2】不迴圈，直接取用 importCols(1) 和 importCols(2) ---
    If importCols.Count >= 1 Then
        lastRow1 = xlsht.Cells(xlsht.Rows.Count, importCols(1)).End(xlUp).Row
        Set rngs1 = xlsht.Range(xlsht.Cells(2, importCols(1)), xlsht.Cells(lastRow1, importCols(1)))
        
        For Each rng In rngs1
            Select Case CStr(rng.Value)
                Case "AC_CompanyBond_Domestic_Cost"
                    AC_CompanyBond_Domestic_Cost = AC_CompanyBond_Domestic_Cost + rng.Offset(0, 1).Value
                Case "AC_CompanyBond_Domestic_ImpairmentLoss"
                    AC_CompanyBond_Domestic_ImpairmentLoss = AC_CompanyBond_Domestic_ImpairmentLoss + rng.Offset(0, 1).Value
                Case "AC_GovBond_Domestic_Cost"
                    AC_GovBond_Domestic_Cost = AC_GovBond_Domestic_Cost + rng.Offset(0, 1).Value
                Case "AC_GovBond_Domestic_ImpairmentLoss"
                    AC_GovBond_Domestic_ImpairmentLoss = AC_GovBond_Domestic_ImpairmentLoss + rng.Offset(0, 1).Value
                Case "AC_NCD_CentralBank_Cost"
                    AC_NCD_CentralBank_Cost = AC_NCD_CentralBank_Cost + rng.Offset(0, 1).Value
                Case "AC_NCD_CentralBank_ImpairmentLoss"
                    AC_NCD_CentralBank_ImpairmentLoss = AC_NCD_CentralBank_ImpairmentLoss + rng.Offset(0, 1).Value
                Case "AFS_FinancialBond_Domestic_Cost"
                    AFS_FinancialBond_Domestic_Cost = AFS_FinancialBond_Domestic_Cost + rng.Offset(0, 1).Value
                Case "AFS_FinancialBond_Domestic_ValuationAdjust"
                    AFS_FinancialBond_Domestic_ValuationAdjust = AFS_FinancialBond_Domestic_ValuationAdjust + rng.Offset(0, 1).Value
                Case "EquityMethod_Cost"
                    EquityMethod_Cost = EquityMethod_Cost + rng.Offset(0, 1).Value
                Case "EquityMethod_ValuationAdjust"
                    EquityMethod_ValuationAdjust = EquityMethod_ValuationAdjust + rng.Offset(0, 1).Value            
                Case "EquityMethod_Other_Cost"
                    EquityMethod_Other_Cost = EquityMethod_Other_Cost + rng.Offset(0, 1).Value
                Case "FVOCI_CompanyBond_Domestic_Cost"
                    FVOCI_CompanyBond_Domestic_Cost = FVOCI_CompanyBond_Domestic_Cost + rng.Offset(0, 1).Value
                Case "FVOCI_CompanyBond_Domestic_ValuationAdjust"
                    FVOCI_CompanyBond_Domestic_ValuationAdjust = FVOCI_CompanyBond_Domestic_ValuationAdjust + rng.Offset(0, 1).Value
                Case "FVOCI_GovBond_Domestic_Cost"
                    FVOCI_GovBond_Domestic_Cost = FVOCI_GovBond_Domestic_Cost + rng.Offset(0, 1).Value
                Case "FVOCI_GovBond_Domestic_ValuationAdjust"
                    FVOCI_GovBond_Domestic_ValuationAdjust = FVOCI_GovBond_Domestic_ValuationAdjust + rng.Offset(0, 1).Value
                Case "FVOCI_NCD_CentralBank_Cost"
                    FVOCI_NCD_CentralBank_Cost = FVOCI_NCD_CentralBank_Cost + rng.Offset(0, 1).Value
                Case "FVOCI_NCD_CentralBank_ValuationAdjust"
                    FVOCI_NCD_CentralBank_ValuationAdjust = FVOCI_NCD_CentralBank_ValuationAdjust + rng.Offset(0, 1).Value
                Case "FVOCI_Stock_特別股_Cost"
                    FVOCI_Stock_PreferredStock_Cost = FVOCI_Stock_PreferredStock_Cost + rng.Offset(0, 1).Value
                Case "FVOCI_Stock_特別股_上市_Cost"
                    FVOCI_Stock_PreferredStock_Listed_Cost = FVOCI_Stock_PreferredStock_Listed_Cost + rng.Offset(0, 1).Value                        
                Case "FVOCI_Stock_特別股_上市_ValuationAdjust"
                    FVOCI_Stock_PreferredStock_Listed_ValuationAdjust = FVOCI_Stock_PreferredStock_Listed_ValuationAdjust + rng.Offset(0, 1).Value
                Case "FVOCI_Stock_普通股_上市_Cost"
                    FVOCI_Stock_CommonStock_Listed_Cost = FVOCI_Stock_CommonStock_Listed_Cost + rng.Offset(0, 1).Value
                Case "FVOCI_Stock_普通股_上市_ValuationAdjust"
                    FVOCI_Stock_CommonStock_Listed_ValuationAdjust = FVOCI_Stock_CommonStock_Listed_ValuationAdjust + rng.Offset(0, 1).Value
                Case "FVOCI_Stock_普通股_上櫃_Cost"
                    FVOCI_Stock_CommonStock_OTC_Cost = FVOCI_Stock_CommonStock_OTC_Cost + rng.Offset(0, 1).Value
                Case "FVOCI_Stock_普通股_上櫃_ValuationAdjust"
                    FVOCI_Stock_CommonStock_OTC_ValuationAdjust = FVOCI_Stock_CommonStock_OTC_ValuationAdjust + rng.Offset(0, 1).Value
                Case "FVOCI_Stock_普通股_興櫃_Cost"
                    FVOCI_Stock_CommonStock_Emergin_Cost = FVOCI_Stock_CommonStock_Emergin_Cost + rng.Offset(0, 1).Value
                Case "FVOCI_Stock_普通股_興櫃_ValuationAdjust"
                    FVOCI_Stock_CommonStock_Emergin_ValuationAdjust = FVOCI_Stock_CommonStock_Emergin_ValuationAdjust + rng.Offset(0, 1).Value
                Case "FVOCI_Equity_Other_Cost"
                    FVOCI_Equity_Other_Cost = FVOCI_Equity_Other_Cost + rng.Offset(0, 1).Value
                Case "FVOCI_Equity_Other_ValuationAdjust"
                    FVOCI_Equity_Other_ValuationAdjust = FVOCI_Equity_Other_ValuationAdjust + rng.Offset(0, 1).Value
                Case "FVPL_AssetCertificate_Cost"
                    FVPL_AssetCertificate_Cost = FVPL_AssetCertificate_Cost + rng.Offset(0, 1).Value
                Case "FVPL_AssetCertificate_ValuationAdjust"
                    FVPL_AssetCertificate_ValuationAdjust = FVPL_AssetCertificate_ValuationAdjust + rng.Offset(0, 1).Value
                Case "FVPL_CompanyBond_Domestic_Cost"
                    FVPL_CompanyBond_Domestic_Cost = FVPL_CompanyBond_Domestic_Cost + rng.Offset(0, 1).Value
                Case "FVPL_CompanyBond_Domestic_ValuationAdjust"
                    FVPL_CompanyBond_Domestic_ValuationAdjust = FVPL_CompanyBond_Domestic_ValuationAdjust + rng.Offset(0, 1).Value
                Case "FVPL_CP_Cost"
                    FVPL_CP_Cost = FVPL_CP_Cost + rng.Offset(0, 1).Value
                Case "FVPL_CP_ValuationAdjust"
                    FVPL_CP_ValuationAdjust = FVPL_CP_ValuationAdjust + rng.Offset(0, 1).Value
                Case "FVPL_GovBond_Domestic_Cost"
                    FVPL_GovBond_Domestic_Cost = FVPL_GovBond_Domestic_Cost + rng.Offset(0, 1).Value
                Case "FVPL_GovBond_Domestic_ValuationAdjust"
                    FVPL_GovBond_Domestic_ValuationAdjust = FVPL_GovBond_Domestic_ValuationAdjust + rng.Offset(0, 1).Value
                Case "FVPL_Stock_特別股_上市_Cost"
                    FVPL_Stock_PreferredStock_Listed_Cost = FVPL_Stock_PreferredStock_Listed_Cost + rng.Offset(0, 1).Value
                Case "FVPL_Stock_特別股_上市_ValuationAdjust"
                    FVPL_Stock_PreferredStock_Listed_ValuationAdjust = FVPL_Stock_PreferredStock_Listed_ValuationAdjust + rng.Offset(0, 1).Value
                Case "FVPL_Stock_普通股_上市_Cost"
                    FVPL_Stock_CommonStock_Listed_Cost = FVPL_Stock_CommonStock_Listed_Cost + rng.Offset(0, 1).Value
                Case "FVPL_Stock_普通股_上市_ValuationAdjust"
                    FVPL_Stock_CommonStock_Listed_ValuationAdjust = FVPL_Stock_CommonStock_Listed_ValuationAdjust + rng.Offset(0, 1).Value
                Case "FVPL_Stock_普通股_上櫃_Cost"
                    FVPL_Stock_CommonStock_OTC_Cost = FVPL_Stock_CommonStock_OTC_Cost + rng.Offset(0, 1).Value
                Case "FVPL_Stock_普通股_上櫃_ValuationAdjust"
                    FVPL_Stock_CommonStock_OTC_ValuationAdjust = FVPL_Stock_CommonStock_OTC_ValuationAdjust + rng.Offset(0, 1).Value
                Case "FVPL_Stock_普通股_興櫃_Cost"
                    FVPL_Stock_CommonStock_Emergin_Cost = FVPL_Stock_CommonStock_Emergin_Cost + rng.Offset(0, 1).Value
                Case "FVPL_Stock_普通股_興櫃_ValuationAdjust"
                    FVPL_Stock_CommonStock_Emergin_ValuationAdjust = FVPL_Stock_CommonStock_Emergin_ValuationAdjust + rng.Offset(0, 1).Value

                Case "FVPL_AssetCertificate_Cost"
                    FVPL_AssetCertificate_Cost = FVPL_AssetCertificate_Cost + rng.Offset(0, 1).Value
                Case "FVPL_AssetCertificate_ValuationAdjust"
                    FVPL_AssetCertificate_ValuationAdjust = FVPL_AssetCertificate_ValuationAdjust + rng.Offset(0, 1).Value
                Case "FVPL_CP_Cost"
                    FVPL_CP_Cost = FVPL_CP_Cost + rng.Offset(0, 1).Value
                Case "FVPL_CP_ValuationAdjust "
                    FVPL_CP_ValuationAdjust  = FVPL_CP_ValuationAdjust  + rng.Offset(0, 1).Value
            End Select
        Next rng
    End If
    

    ' 單位：新臺幣千元


    
    RP_GovBond_Cost = Round(RP_GovBond_Cost / 1000, 0)
    RP_CompanyBond_Cost = Round(RP_CompanyBond_Cost / 1000, 0)
    RP_CP_Cost = Round(RP_CP_Cost / 1000, 0)
    
    xlsht.Range("Table20_0200_二公債_民營企業_其他到期日").Value = RP_GovBond_Cost
    xlsht.Range("Table20_0300_三公司債_民營企業_其他到期日").Value = RP_CompanyBond_Cost
    xlsht.Range("Table20_0400_四商業本票_民營企業_其他到期日").Value = RP_CP_Cost
    
    ' 資料一開始要先清掉

    ' --- 【修改2】用 FieldValuePositionMap 批次填 rpt ---
    Dim SetFieldMap As Variant
    Dim tgtSheet As String, srcTag As String, srcVal As Variant
    
    SetFieldMap = GetMapData(gDBPath, reportTitle, "FieldValuePositionMap")
    If Not IsNull(SetFieldMap) And IsArray(SetFieldMap) Then
        For iMap = 0 To UBound(SetFieldMap, 1)
            tgtSheet = CStr(SetFieldMap(iMap, 0))
            srcTag    = CStr(SetFieldMap(iMap, 1))
            
            On Error Resume Next
            srcVal = ThisWorkbook.Sheets(tgtSheet).Range(srcTag).Value
            On Error GoTo 0
            
            rpt.SetField tgtSheet, srcTag, srcVal
        Next iMap
    Else
        WriteLog "無法取得 FieldValuePositionMap for " & reportTitle
    End If

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

Public Sub Process_AI345()
    '=== Equal Setting ===
    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("AI345")

    Dim reportTitle As String
    reportTitle = rpt.ReportName

    Dim xlsht As Worksheet
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    ' xlsht.Range("A:G").ClearContents
    ' xlsht.Range("T2:T100").ClearContents

    ' xlsht.Cells.ClearContents  ' 或只清指定区

    ' 呼叫通用函式，並取得 importCols
    Dim importCols As Collection
    Set importCols = ImportQueryTables(gDBPath, xlsht, reportTitle, gDataMonthString)

    If importCols Is Nothing Or importCols.Count = 0 Then
        Exit Sub
    End If
    
    '--------------
    'Unique Setting
    '--------------

    Dim RP_GovBond_Cost As Double: RP_GovBond_Cost = 0
    Dim RP_CompanyBond_Cost As Double: RP_CompanyBond_Cost = 0
    Dim RP_CP_Cost As Double: RP_CP_Cost = 0

    Dim lastRow1 As Long, lastRow2 As Long
    Dim rngs1 As Range, rngs2 As Range, rng As Range

    ' --- 【修改1.2】不迴圈，直接取用 importCols(1) 和 importCols(2) ---
    If importCols.Count >= 1 Then
        lastRow1 = xlsht.Cells(xlsht.Rows.Count, importCols(1)).End(xlUp).Row
        Set rngs1 = xlsht.Range(xlsht.Cells(2, importCols(1)), xlsht.Cells(lastRow1, importCols(1)))
        
        For Each rng In rngs1
            Select Case CStr(rng.Value)
                Case "RP_GovBond_Cost"
                    RP_GovBond_Cost = RP_GovBond_Cost + rng.Offset(0, 1).Value
                Case "AC_CompanyBond_Domestic_ImpairmentLoss"
                    RP_CompanyBond_Cost = RP_CompanyBond_Cost + rng.Offset(0, 1).Value
            End Select
        Next rng
    End If
    
    ' If importCols.Count >= 2 Then
    '     lastRow2 = xlsht.Cells(xlsht.Rows.Count, importCols(2)).End(xlUp).Row
    '     Set rngs2 = xlsht.Range(xlsht.Cells(2, importCols(2)), xlsht.Cells(lastRow2, importCols(2)))
        
    '     For Each rng In rngs2
    '         ' 如果第二筆表也有需要累計的 tag，可以在這裡加
    '         Select Case CStr(rng.Value)
    '             Case "RP_GovBond_Cost"
    '                 RP_GovBond_Cost = RP_GovBond_Cost + rng.Offset(0, 1).Value
    '             Case "AC_CompanyBond_Domestic_ImpairmentLoss"
    '                 RP_CompanyBond_Cost = RP_CompanyBond_Cost + rng.Offset(0, 1).Value
    '         End Select
    '     Next rng
    ' End If

    ' HANDLE方式
    - 民營企業		
        - 其他到期日		
            - 公債		
    225010101		附買回票券及債券負債-公債  
    RP_GovBond_Cost +                       
            - 公司債		
    225010105		附買回票券及債券負債-公司債    
    RP_CompanyBond_Cost +                   
            - 商業本票		
    225010303		#N/A
    找不到這個ACCOUNT CODE

    Table20_0200_二公債_民營企業_其他到期日
    Table20_0300_三公司債_民營企業_其他到期日
    Table20_0400_四商業本票_民營企業_其他到期日

    ' END HANDLE
    
    RP_GovBond_Cost = Round(RP_GovBond_Cost / 1000, 0)
    RP_CompanyBond_Cost = Round(RP_CompanyBond_Cost / 1000, 0)
    RP_CP_Cost = Round(RP_CP_Cost / 1000, 0)
    
    xlsht.Range("Table20_0200_二公債_民營企業_其他到期日").Value = RP_GovBond_Cost
    xlsht.Range("Table20_0300_三公司債_民營企業_其他到期日").Value = RP_CompanyBond_Cost
    xlsht.Range("Table20_0400_四商業本票_民營企業_其他到期日").Value = RP_CP_Cost
    
    ' 資料一開始要先清掉

    ' --- 【修改2】用 FieldValuePositionMap 批次填 rpt ---
    Dim SetFieldMap As Variant
    Dim tgtSheet As String, srcTag As String, srcVal As Variant
    
    SetFieldMap = GetMapData(gDBPath, reportTitle, "FieldValuePositionMap")
    If Not IsNull(SetFieldMap) And IsArray(SetFieldMap) Then
        For iMap = 0 To UBound(SetFieldMap, 1)
            tgtSheet = CStr(SetFieldMap(iMap, 0))
            srcTag    = CStr(SetFieldMap(iMap, 1))
            
            On Error Resume Next
            srcVal = ThisWorkbook.Sheets(tgtSheet).Range(srcTag).Value
            On Error GoTo 0
            
            rpt.SetField tgtSheet, srcTag, srcVal
        Next iMap
    Else
        WriteLog "無法取得 FieldValuePositionMap for " & reportTitle
    End If

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



ai405		
    
    
- 持有餘額		
    - 政府債券		
120050101		強制FVPL金融資產-公債-中央政府(我國)              
120050103		強制FVPL金融資產-公債-地方政府(我國)              
125010101		附賣回票券及債券投資-公債                         
121110101		FVOCI債務工具-公債-中央政府(我國)                 
121110103		FVOCI債務工具-公債-地方政府(我國)                 
122010101		AC債務工具投資-公債-中央政府(我國)                
122010103		AC債務工具投資-公債-地方政府(我國)                
    - 金融債券		
122010147		AC債務工具投資-金融債券-海外                      
121110147		FVOCI債務工具-金融債券-海外                       
120050147		#N/A
        
122010147		AC債務工具投資-金融債券-海外                      
121110147		FVOCI債務工具-金融債券-海外                       
120050147		#N/A
        
    - 公司債		
120057307		強制FVPL金融資產-衍生性ＳＷＡＰ-ASS               
120050121		強制FVPL金融資產-普通公司債(公營)                 
120050123		強制FVPL金融資產-普通公司債(民營)                 
121110121		FVOCI債務工具-普通公司債（公營）                  
121110123		FVOCI債務工具-普通公司債（民營）                  
122010121		AC債務工具投資-普通公司債(公營)                   
122010123		AC債務工具投資-普通公司債(民營)                   
        
121110127		FVOCI債務工具-普通公司債(民營)(外國)              
121110127		FVOCI債務工具-普通公司債(民營)(外國)              
    








Public Sub Process_AI405()
    '=== Equal Setting ===
    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("AI405")

    Dim reportTitle As String
    reportTitle = rpt.ReportName

    Dim xlsht As Worksheet
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    ' xlsht.Range("A:G").ClearContents
    ' xlsht.Range("T2:T100").ClearContents

    ' xlsht.Cells.ClearContents  ' 或只清指定区

    ' 呼叫通用函式，並取得 importCols
    Dim importCols As Collection
    Set importCols = ImportQueryTables(gDBPath, xlsht, reportTitle, gDataMonthString)

    If importCols Is Nothing Or importCols.Count = 0 Then
        Exit Sub
    End If
    
    '--------------
    'Unique Setting
    '--------------

    Dim RP_GovBond_Cost As Double: RP_GovBond_Cost = 0
    Dim RP_CompanyBond_Cost As Double: RP_CompanyBond_Cost = 0
    Dim RP_CP_Cost As Double: RP_CP_Cost = 0

    Dim lastRow1 As Long, lastRow2 As Long
    Dim rngs1 As Range, rngs2 As Range, rng As Range

    ' --- 【修改1.2】不迴圈，直接取用 importCols(1) 和 importCols(2) ---
    If importCols.Count >= 1 Then
        lastRow1 = xlsht.Cells(xlsht.Rows.Count, importCols(1)).End(xlUp).Row
        Set rngs1 = xlsht.Range(xlsht.Cells(2, importCols(1)), xlsht.Cells(lastRow1, importCols(1)))
        
        For Each rng In rngs1
            Select Case CStr(rng.Value)
                Case "RP_GovBond_Cost"
                    RP_GovBond_Cost = RP_GovBond_Cost + rng.Offset(0, 1).Value
                Case "AC_CompanyBond_Domestic_ImpairmentLoss"
                    RP_CompanyBond_Cost = RP_CompanyBond_Cost + rng.Offset(0, 1).Value
            End Select
        Next rng
    End If
    
    ' If importCols.Count >= 2 Then
    '     lastRow2 = xlsht.Cells(xlsht.Rows.Count, importCols(2)).End(xlUp).Row
    '     Set rngs2 = xlsht.Range(xlsht.Cells(2, importCols(2)), xlsht.Cells(lastRow2, importCols(2)))
        
    '     For Each rng In rngs2
    '         ' 如果第二筆表也有需要累計的 tag，可以在這裡加
    '         Select Case CStr(rng.Value)
    '             Case "RP_GovBond_Cost"
    '                 RP_GovBond_Cost = RP_GovBond_Cost + rng.Offset(0, 1).Value
    '             Case "AC_CompanyBond_Domestic_ImpairmentLoss"
    '                 RP_CompanyBond_Cost = RP_CompanyBond_Cost + rng.Offset(0, 1).Value
    '         End Select
    '     Next rng
    ' End If

    ' END HANDLE
    
    RP_GovBond_Cost = Round(RP_GovBond_Cost / 1000, 0)
    RP_CompanyBond_Cost = Round(RP_CompanyBond_Cost / 1000, 0)
    RP_CP_Cost = Round(RP_CP_Cost / 1000, 0)
    
    xlsht.Range("Table20_0200_二公債_民營企業_其他到期日").Value = RP_GovBond_Cost
    xlsht.Range("Table20_0300_三公司債_民營企業_其他到期日").Value = RP_CompanyBond_Cost
    xlsht.Range("Table20_0400_四商業本票_民營企業_其他到期日").Value = RP_CP_Cost
    
    ' 資料一開始要先清掉

    ' --- 【修改2】用 FieldValuePositionMap 批次填 rpt ---
    Dim SetFieldMap As Variant
    Dim tgtSheet As String, srcTag As String, srcVal As Variant
    
    SetFieldMap = GetMapData(gDBPath, reportTitle, "FieldValuePositionMap")
    If Not IsNull(SetFieldMap) And IsArray(SetFieldMap) Then
        For iMap = 0 To UBound(SetFieldMap, 1)
            tgtSheet = CStr(SetFieldMap(iMap, 0))
            srcTag    = CStr(SetFieldMap(iMap, 1))
            
            On Error Resume Next
            srcVal = ThisWorkbook.Sheets(tgtSheet).Range(srcTag).Value
            On Error GoTo 0
            
            rpt.SetField tgtSheet, srcTag, srcVal
        Next iMap
    Else
        WriteLog "無法取得 FieldValuePositionMap for " & reportTitle
    End If

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

Public Sub Process_AI410()
    '=== Equal Setting ===
    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("AI410")

    Dim reportTitle As String
    reportTitle = rpt.ReportName

    Dim xlsht As Worksheet
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    ' xlsht.Range("A:G").ClearContents
    ' xlsht.Range("T2:T100").ClearContents

    ' xlsht.Cells.ClearContents  ' 或只清指定区

    ' 呼叫通用函式，並取得 importCols
    Dim importCols As Collection
    Set importCols = ImportQueryTables(gDBPath, xlsht, reportTitle, gDataMonthString)

    If importCols Is Nothing Or importCols.Count = 0 Then
        Exit Sub
    End If
    
    '--------------
    'Unique Setting
    '--------------

    Dim RP_GovBond_Cost As Double: RP_GovBond_Cost = 0
    Dim RP_CompanyBond_Cost As Double: RP_CompanyBond_Cost = 0
    Dim RP_CP_Cost As Double: RP_CP_Cost = 0

    Dim lastRow1 As Long, lastRow2 As Long
    Dim rngs1 As Range, rngs2 As Range, rng As Range

    ' --- 【修改1.2】不迴圈，直接取用 importCols(1) 和 importCols(2) ---
    If importCols.Count >= 1 Then
        lastRow1 = xlsht.Cells(xlsht.Rows.Count, importCols(1)).End(xlUp).Row
        Set rngs1 = xlsht.Range(xlsht.Cells(2, importCols(1)), xlsht.Cells(lastRow1, importCols(1)))
        
        For Each rng In rngs1
            Select Case CStr(rng.Value)
                Case "RP_GovBond_Cost"
                    RP_GovBond_Cost = RP_GovBond_Cost + rng.Offset(0, 1).Value
                Case "AC_CompanyBond_Domestic_ImpairmentLoss"
                    RP_CompanyBond_Cost = RP_CompanyBond_Cost + rng.Offset(0, 1).Value
            End Select
        Next rng
    End If
    
    ' If importCols.Count >= 2 Then
    '     lastRow2 = xlsht.Cells(xlsht.Rows.Count, importCols(2)).End(xlUp).Row
    '     Set rngs2 = xlsht.Range(xlsht.Cells(2, importCols(2)), xlsht.Cells(lastRow2, importCols(2)))
        
    '     For Each rng In rngs2
    '         ' 如果第二筆表也有需要累計的 tag，可以在這裡加
    '         Select Case CStr(rng.Value)
    '             Case "RP_GovBond_Cost"
    '                 RP_GovBond_Cost = RP_GovBond_Cost + rng.Offset(0, 1).Value
    '             Case "AC_CompanyBond_Domestic_ImpairmentLoss"
    '                 RP_CompanyBond_Cost = RP_CompanyBond_Cost + rng.Offset(0, 1).Value
    '         End Select
    '     Next rng
    ' End If

    ' HANDLE方式
    - 民營企業		
        - 其他到期日		
            - 公債		
    225010101		附買回票券及債券負債-公債  
    RP_GovBond_Cost +                       
            - 公司債		
    225010105		附買回票券及債券負債-公司債    
    RP_CompanyBond_Cost +                   
            - 商業本票		
    225010303		#N/A
    找不到這個ACCOUNT CODE

    Table20_0200_二公債_民營企業_其他到期日
    Table20_0300_三公司債_民營企業_其他到期日
    Table20_0400_四商業本票_民營企業_其他到期日

    ' END HANDLE
    
    RP_GovBond_Cost = Round(RP_GovBond_Cost / 1000, 0)
    RP_CompanyBond_Cost = Round(RP_CompanyBond_Cost / 1000, 0)
    RP_CP_Cost = Round(RP_CP_Cost / 1000, 0)
    
    xlsht.Range("Table20_0200_二公債_民營企業_其他到期日").Value = RP_GovBond_Cost
    xlsht.Range("Table20_0300_三公司債_民營企業_其他到期日").Value = RP_CompanyBond_Cost
    xlsht.Range("Table20_0400_四商業本票_民營企業_其他到期日").Value = RP_CP_Cost
    
    ' 資料一開始要先清掉

    ' --- 【修改2】用 FieldValuePositionMap 批次填 rpt ---
    Dim SetFieldMap As Variant
    Dim tgtSheet As String, srcTag As String, srcVal As Variant
    
    SetFieldMap = GetMapData(gDBPath, reportTitle, "FieldValuePositionMap")
    If Not IsNull(SetFieldMap) And IsArray(SetFieldMap) Then
        For iMap = 0 To UBound(SetFieldMap, 1)
            tgtSheet = CStr(SetFieldMap(iMap, 0))
            srcTag    = CStr(SetFieldMap(iMap, 1))
            
            On Error Resume Next
            srcVal = ThisWorkbook.Sheets(tgtSheet).Range(srcTag).Value
            On Error GoTo 0
            
            rpt.SetField tgtSheet, srcTag, srcVal
        Next iMap
    Else
        WriteLog "無法取得 FieldValuePositionMap for " & reportTitle
    End If

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

Public Sub Process_AI430()
    '=== Equal Setting ===
    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("AI430")

    Dim reportTitle As String
    reportTitle = rpt.ReportName

    Dim xlsht As Worksheet
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    ' xlsht.Range("A:G").ClearContents
    ' xlsht.Range("T2:T100").ClearContents

    ' xlsht.Cells.ClearContents  ' 或只清指定区

    ' 呼叫通用函式，並取得 importCols
    Dim importCols As Collection
    Set importCols = ImportQueryTables(gDBPath, xlsht, reportTitle, gDataMonthString)

    If importCols Is Nothing Or importCols.Count = 0 Then
        Exit Sub
    End If
    
    '--------------
    'Unique Setting
    '--------------

    Dim RP_GovBond_Cost As Double: RP_GovBond_Cost = 0
    Dim RP_CompanyBond_Cost As Double: RP_CompanyBond_Cost = 0
    Dim RP_CP_Cost As Double: RP_CP_Cost = 0

    Dim lastRow1 As Long, lastRow2 As Long
    Dim rngs1 As Range, rngs2 As Range, rng As Range

    ' --- 【修改1.2】不迴圈，直接取用 importCols(1) 和 importCols(2) ---
    If importCols.Count >= 1 Then
        lastRow1 = xlsht.Cells(xlsht.Rows.Count, importCols(1)).End(xlUp).Row
        Set rngs1 = xlsht.Range(xlsht.Cells(2, importCols(1)), xlsht.Cells(lastRow1, importCols(1)))
        
        For Each rng In rngs1
            Select Case CStr(rng.Value)
                Case "RP_GovBond_Cost"
                    RP_GovBond_Cost = RP_GovBond_Cost + rng.Offset(0, 1).Value
                Case "AC_CompanyBond_Domestic_ImpairmentLoss"
                    RP_CompanyBond_Cost = RP_CompanyBond_Cost + rng.Offset(0, 1).Value
            End Select
        Next rng
    End If
    
    ' If importCols.Count >= 2 Then
    '     lastRow2 = xlsht.Cells(xlsht.Rows.Count, importCols(2)).End(xlUp).Row
    '     Set rngs2 = xlsht.Range(xlsht.Cells(2, importCols(2)), xlsht.Cells(lastRow2, importCols(2)))
        
    '     For Each rng In rngs2
    '         ' 如果第二筆表也有需要累計的 tag，可以在這裡加
    '         Select Case CStr(rng.Value)
    '             Case "RP_GovBond_Cost"
    '                 RP_GovBond_Cost = RP_GovBond_Cost + rng.Offset(0, 1).Value
    '             Case "AC_CompanyBond_Domestic_ImpairmentLoss"
    '                 RP_CompanyBond_Cost = RP_CompanyBond_Cost + rng.Offset(0, 1).Value
    '         End Select
    '     Next rng
    ' End If

    ' HANDLE方式
    - 民營企業		
        - 其他到期日		
            - 公債		
    225010101		附買回票券及債券負債-公債  
    RP_GovBond_Cost +                       
            - 公司債		
    225010105		附買回票券及債券負債-公司債    
    RP_CompanyBond_Cost +                   
            - 商業本票		
    225010303		#N/A
    找不到這個ACCOUNT CODE

    Table20_0200_二公債_民營企業_其他到期日
    Table20_0300_三公司債_民營企業_其他到期日
    Table20_0400_四商業本票_民營企業_其他到期日

    ' END HANDLE
    
    RP_GovBond_Cost = Round(RP_GovBond_Cost / 1000, 0)
    RP_CompanyBond_Cost = Round(RP_CompanyBond_Cost / 1000, 0)
    RP_CP_Cost = Round(RP_CP_Cost / 1000, 0)
    
    xlsht.Range("Table20_0200_二公債_民營企業_其他到期日").Value = RP_GovBond_Cost
    xlsht.Range("Table20_0300_三公司債_民營企業_其他到期日").Value = RP_CompanyBond_Cost
    xlsht.Range("Table20_0400_四商業本票_民營企業_其他到期日").Value = RP_CP_Cost
    
    ' 資料一開始要先清掉

    ' --- 【修改2】用 FieldValuePositionMap 批次填 rpt ---
    Dim SetFieldMap As Variant
    Dim tgtSheet As String, srcTag As String, srcVal As Variant
    
    SetFieldMap = GetMapData(gDBPath, reportTitle, "FieldValuePositionMap")
    If Not IsNull(SetFieldMap) And IsArray(SetFieldMap) Then
        For iMap = 0 To UBound(SetFieldMap, 1)
            tgtSheet = CStr(SetFieldMap(iMap, 0))
            srcTag    = CStr(SetFieldMap(iMap, 1))
            
            On Error Resume Next
            srcVal = ThisWorkbook.Sheets(tgtSheet).Range(srcTag).Value
            On Error GoTo 0
            
            rpt.SetField tgtSheet, srcTag, srcVal
        Next iMap
    Else
        WriteLog "無法取得 FieldValuePositionMap for " & reportTitle
    End If

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

Public Sub Process_AI601()
    '=== Equal Setting ===
    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("AI601")

    Dim reportTitle As String
    reportTitle = rpt.ReportName

    Dim xlsht As Worksheet
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    ' xlsht.Range("A:G").ClearContents
    ' xlsht.Range("T2:T100").ClearContents

    ' xlsht.Cells.ClearContents  ' 或只清指定区

    ' 呼叫通用函式，並取得 importCols
    Dim importCols As Collection
    Set importCols = ImportQueryTables(gDBPath, xlsht, reportTitle, gDataMonthString)

    If importCols Is Nothing Or importCols.Count = 0 Then
        Exit Sub
    End If
    
    '--------------
    'Unique Setting
    '--------------

    Dim RP_GovBond_Cost As Double: RP_GovBond_Cost = 0
    Dim RP_CompanyBond_Cost As Double: RP_CompanyBond_Cost = 0
    Dim RP_CP_Cost As Double: RP_CP_Cost = 0

    Dim lastRow1 As Long, lastRow2 As Long
    Dim rngs1 As Range, rngs2 As Range, rng As Range

    ' --- 【修改1.2】不迴圈，直接取用 importCols(1) 和 importCols(2) ---
    If importCols.Count >= 1 Then
        lastRow1 = xlsht.Cells(xlsht.Rows.Count, importCols(1)).End(xlUp).Row
        Set rngs1 = xlsht.Range(xlsht.Cells(2, importCols(1)), xlsht.Cells(lastRow1, importCols(1)))
        
        For Each rng In rngs1
            Select Case CStr(rng.Value)
                Case "RP_GovBond_Cost"
                    RP_GovBond_Cost = RP_GovBond_Cost + rng.Offset(0, 1).Value
                Case "AC_CompanyBond_Domestic_ImpairmentLoss"
                    RP_CompanyBond_Cost = RP_CompanyBond_Cost + rng.Offset(0, 1).Value
            End Select
        Next rng
    End If
    
    ' If importCols.Count >= 2 Then
    '     lastRow2 = xlsht.Cells(xlsht.Rows.Count, importCols(2)).End(xlUp).Row
    '     Set rngs2 = xlsht.Range(xlsht.Cells(2, importCols(2)), xlsht.Cells(lastRow2, importCols(2)))
        
    '     For Each rng In rngs2
    '         ' 如果第二筆表也有需要累計的 tag，可以在這裡加
    '         Select Case CStr(rng.Value)
    '             Case "RP_GovBond_Cost"
    '                 RP_GovBond_Cost = RP_GovBond_Cost + rng.Offset(0, 1).Value
    '             Case "AC_CompanyBond_Domestic_ImpairmentLoss"
    '                 RP_CompanyBond_Cost = RP_CompanyBond_Cost + rng.Offset(0, 1).Value
    '         End Select
    '     Next rng
    ' End If

    ' HANDLE方式
    - 民營企業		
        - 其他到期日		
            - 公債		
    225010101		附買回票券及債券負債-公債  
    RP_GovBond_Cost +                       
            - 公司債		
    225010105		附買回票券及債券負債-公司債    
    RP_CompanyBond_Cost +                   
            - 商業本票		
    225010303		#N/A
    找不到這個ACCOUNT CODE

    Table20_0200_二公債_民營企業_其他到期日
    Table20_0300_三公司債_民營企業_其他到期日
    Table20_0400_四商業本票_民營企業_其他到期日

    ' END HANDLE
    
    RP_GovBond_Cost = Round(RP_GovBond_Cost / 1000, 0)
    RP_CompanyBond_Cost = Round(RP_CompanyBond_Cost / 1000, 0)
    RP_CP_Cost = Round(RP_CP_Cost / 1000, 0)
    
    xlsht.Range("Table20_0200_二公債_民營企業_其他到期日").Value = RP_GovBond_Cost
    xlsht.Range("Table20_0300_三公司債_民營企業_其他到期日").Value = RP_CompanyBond_Cost
    xlsht.Range("Table20_0400_四商業本票_民營企業_其他到期日").Value = RP_CP_Cost
    
    ' 資料一開始要先清掉

    ' --- 【修改2】用 FieldValuePositionMap 批次填 rpt ---
    Dim SetFieldMap As Variant
    Dim tgtSheet As String, srcTag As String, srcVal As Variant
    
    SetFieldMap = GetMapData(gDBPath, reportTitle, "FieldValuePositionMap")
    If Not IsNull(SetFieldMap) And IsArray(SetFieldMap) Then
        For iMap = 0 To UBound(SetFieldMap, 1)
            tgtSheet = CStr(SetFieldMap(iMap, 0))
            srcTag    = CStr(SetFieldMap(iMap, 1))
            
            On Error Resume Next
            srcVal = ThisWorkbook.Sheets(tgtSheet).Range(srcTag).Value
            On Error GoTo 0
            
            rpt.SetField tgtSheet, srcTag, srcVal
        Next iMap
    Else
        WriteLog "無法取得 FieldValuePositionMap for " & reportTitle
    End If

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

Public Sub Process_AI605()
    '=== Equal Setting ===
    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("AI605")

    Dim reportTitle As String
    reportTitle = rpt.ReportName

    Dim xlsht As Worksheet
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    ' xlsht.Range("A:G").ClearContents
    ' xlsht.Range("T2:T100").ClearContents

    ' xlsht.Cells.ClearContents  ' 或只清指定区

    ' 呼叫通用函式，並取得 importCols
    Dim importCols As Collection
    Set importCols = ImportQueryTables(gDBPath, xlsht, reportTitle, gDataMonthString)

    If importCols Is Nothing Or importCols.Count = 0 Then
        Exit Sub
    End If
    
    '--------------
    'Unique Setting
    '--------------

    Dim RP_GovBond_Cost As Double: RP_GovBond_Cost = 0
    Dim RP_CompanyBond_Cost As Double: RP_CompanyBond_Cost = 0
    Dim RP_CP_Cost As Double: RP_CP_Cost = 0

    Dim lastRow1 As Long, lastRow2 As Long
    Dim rngs1 As Range, rngs2 As Range, rng As Range

    ' --- 【修改1.2】不迴圈，直接取用 importCols(1) 和 importCols(2) ---
    If importCols.Count >= 1 Then
        lastRow1 = xlsht.Cells(xlsht.Rows.Count, importCols(1)).End(xlUp).Row
        Set rngs1 = xlsht.Range(xlsht.Cells(2, importCols(1)), xlsht.Cells(lastRow1, importCols(1)))
        
        For Each rng In rngs1
            Select Case CStr(rng.Value)
                Case "RP_GovBond_Cost"
                    RP_GovBond_Cost = RP_GovBond_Cost + rng.Offset(0, 1).Value
                Case "AC_CompanyBond_Domestic_ImpairmentLoss"
                    RP_CompanyBond_Cost = RP_CompanyBond_Cost + rng.Offset(0, 1).Value
            End Select
        Next rng
    End If
    
    ' If importCols.Count >= 2 Then
    '     lastRow2 = xlsht.Cells(xlsht.Rows.Count, importCols(2)).End(xlUp).Row
    '     Set rngs2 = xlsht.Range(xlsht.Cells(2, importCols(2)), xlsht.Cells(lastRow2, importCols(2)))
        
    '     For Each rng In rngs2
    '         ' 如果第二筆表也有需要累計的 tag，可以在這裡加
    '         Select Case CStr(rng.Value)
    '             Case "RP_GovBond_Cost"
    '                 RP_GovBond_Cost = RP_GovBond_Cost + rng.Offset(0, 1).Value
    '             Case "AC_CompanyBond_Domestic_ImpairmentLoss"
    '                 RP_CompanyBond_Cost = RP_CompanyBond_Cost + rng.Offset(0, 1).Value
    '         End Select
    '     Next rng
    ' End If

    ' HANDLE方式
    - 民營企業		
        - 其他到期日		
            - 公債		
    225010101		附買回票券及債券負債-公債  
    RP_GovBond_Cost +                       
            - 公司債		
    225010105		附買回票券及債券負債-公司債    
    RP_CompanyBond_Cost +                   
            - 商業本票		
    225010303		#N/A
    找不到這個ACCOUNT CODE

    Table20_0200_二公債_民營企業_其他到期日
    Table20_0300_三公司債_民營企業_其他到期日
    Table20_0400_四商業本票_民營企業_其他到期日

    ' END HANDLE
    
    RP_GovBond_Cost = Round(RP_GovBond_Cost / 1000, 0)
    RP_CompanyBond_Cost = Round(RP_CompanyBond_Cost / 1000, 0)
    RP_CP_Cost = Round(RP_CP_Cost / 1000, 0)
    
    xlsht.Range("Table20_0200_二公債_民營企業_其他到期日").Value = RP_GovBond_Cost
    xlsht.Range("Table20_0300_三公司債_民營企業_其他到期日").Value = RP_CompanyBond_Cost
    xlsht.Range("Table20_0400_四商業本票_民營企業_其他到期日").Value = RP_CP_Cost
    
    ' 資料一開始要先清掉

    ' --- 【修改2】用 FieldValuePositionMap 批次填 rpt ---
    Dim SetFieldMap As Variant
    Dim tgtSheet As String, srcTag As String, srcVal As Variant
    
    SetFieldMap = GetMapData(gDBPath, reportTitle, "FieldValuePositionMap")
    If Not IsNull(SetFieldMap) And IsArray(SetFieldMap) Then
        For iMap = 0 To UBound(SetFieldMap, 1)
            tgtSheet = CStr(SetFieldMap(iMap, 0))
            srcTag    = CStr(SetFieldMap(iMap, 1))
            
            On Error Resume Next
            srcVal = ThisWorkbook.Sheets(tgtSheet).Range(srcTag).Value
            On Error GoTo 0
            
            rpt.SetField tgtSheet, srcTag, srcVal
        Next iMap
    Else
        WriteLog "無法取得 FieldValuePositionMap for " & reportTitle
    End If

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
