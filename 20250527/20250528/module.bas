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
    Dim rngs As Range
    Dim rng As Range

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
        ElseIf CStr(rng.Value) = "EquityMethod_Cost" Then
            EquityMethod_Cost = EquityMethod_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "EquityMethod_ValuationAdjust" Then
            EquityMethod_ValuationAdjust = EquityMethod_ValuationAdjust + rng.Offset(0, 1).Value            
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
        ElseIf CStr(rng.Value) = "FVOCI_Stock_特別股_Cost" Then
            FVOCI_Stock_PreferredStock_Cost = FVOCI_Stock_PreferredStock_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVOCI_Stock_特別股_上市_Cost" Then
            FVOCI_Stock_PreferredStock_Listed_Cost = FVOCI_Stock_PreferredStock_Listed_Cost + rng.Offset(0, 1).Value                        
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




    ' HANDLE方式
    
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

    FVPL_CompanyBond_Public_Domestic_Cost + FVOCI_CompanyBond_Public_Domestic_Cost + AC_CompanyBond_Public_Domestic_Cost

            
        透過損益按公允價值衡量之金融資產2 A		
    120050121		強制FVPL金融資產-普通公司債(公營)                 
    120070121		強制FVPL金融資產評價調整-普通公司債(公營)   
    FVPL_CompanyBond_Public_Domestic_Cost + FVPL_CompanyBond_Public_Domestic_ValuationAdjust
    
            
        透過其他綜合損益按公允價值衡量之金融資產2 B		
    121110121		FVOCI債務工具-普通公司債（公營）                  
            
    FVOCI_CompanyBond_Public_Domestic_Cost +


        按攤銷後成本衡量之債務工具投資2 C		
    122010121		AC債務工具投資-普通公司債(公營)                   
            

    AC_CompanyBond_Public_Domestic_Cost +


    2.2.民營企業-國內公司債		
        原始取得成本1		
    120050123		強制FVPL金融資產-普通公司債(民營)                 
    121110123		FVOCI債務工具-普通公司債（民營）                  
    122010123		AC債務工具投資-普通公司債(民營)                   
    
    FVPL_CompanyBond_Private_Domestic_Cost + FVOCI_CompanyBond_Private_Domestic_Cost + AC_CompanyBond_Private_Domestic_Cost


        透過損益按公允價值衡量之金融資產2 A		
    120050123		強制FVPL金融資產-普通公司債(民營)                 
    120070123		強制FVPL金融資產評價調整-普通公司債(民營)         

    FVPL_CompanyBond_Private_Domestic_Cost + FVPL_CompanyBond_Private_Domestic_ValuationAdjust

        透過其他綜合損益按公允價值衡量之金融資產2 B		
    121110123		FVOCI債務工具-普通公司債（民營）                  

    FVOCI_CompanyBond_Private_Domestic_Cost +

        按攤銷後成本衡量之債務工具投資2 C		
    122010123		AC債務工具投資-普通公司債(民營)                   
    AC_CompanyBond_Private_Domestic_Cost +
    

    3.股票及股權投資-民營企業		
        原始取得成本1		    
    15503
    ' * 原來公式寫 15503，實際上這是 15003
    EquityMethod_ValuationAdjust +



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
    1200503		強制FVPL金融資產-股票    
    FVPL_Stock_PreferredStock_Listed_Cost + FVPL_Stock_CommonStock_Listed_Cost + FVPL_Stock_CommonStock_OTC_Cost + FVPL_Stock_CommonStock_Emergin_Cost                         
            
        透過其他綜合損益按公允價值衡量之金融資產2 B		
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
    15001		採用權益法之投資成本 
    EquityMethod_Other_Cost +                              
    15003		加（減）：採用權益法認列之投資權益調整            
    EquityMethod_ValuationAdjust +            
    4.受益憑證-其他		
            
        原始取得成本1		
    1200505		強制FVPL金融資產-受益憑證              
    
    FVPL_AssetCertificate_Cost +
            
        透過損益按公允價值衡量之金融資產2 A		
    1200505		強制FVPL金融資產-受益憑證                         
    1200705		強制FVPL金融資產評價調整-受益憑證                 
    FVPL_AssetCertificate_Cost + FVPL_AssetCertificate_ValuationAdjust

        透過其他綜合損益按公允價值衡量之金融資產2 B		
            
        按攤銷後成本衡量之債務工具投資2 C		
            
    5.新台幣可轉讓定期存單-中央銀行發行		
            
        原始取得成本1		
    121110911		FVOCI債務工具-央行NCD                             
    122010911		AC債務工具投資-央行NCD   
    FVOCI_NCD_CentralBank_Cost + AC_NCD_CentralBank_Cost
            
        透過損益按公允價值衡量之金融資產2 A		
            
        透過其他綜合損益按公允價值衡量之金融資產2 B		
    121110911		FVOCI債務工具-央行NCD                             
    121130911		FVOCI債務工具評價調整-央行NCD                     
    FVOCI_NCD_CentralBank_Cost + FVOCI_NCD_CentralBank_ValuationAdjust
            
        按攤銷後成本衡量之債務工具投資2 C		
    122010911		AC債務工具投資-央行NCD                            
    122030911		累積減損-AC債務工具投資-央行NCD                   

    AC_NCD_CentralBank_Cost + AC_NCD_CentralBank_ImpairmentLoss
            
    6.商業本票-民營企業		
            
        原始取得成本1		
    120050903		強制FVPL金融資產-商業本票                         
    FVPL_CP_Cost + 
    
            
        透過損益按公允價值衡量之金融資產2 A		
    120050903		強制FVPL金融資產-商業本票                         
    120070903		強制FVPL金融資產評價調整-商業本票                 
    FVPL_CP_Cost + FVPL_CP_ValuationAdjust
            
        透過其他綜合損益按公允價值衡量之金融資產2 B		
            
        按攤銷後成本衡量之債務工具投資2 C		
            
    7.國外機構發行-在國外發行-長期債票券6		
            
        原始取得成本1		
    140010147		備供出售-金融債券-海外                  
    AFS_FinancialBond_Domestic_Cost +
    
            
        透過損益按公允價值衡量之金融資產2 A		
            
        透過其他綜合損益按公允價值衡量之金融資產2 B		
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
    Dim rngs As Range
    Dim rng As Range
    
    lastRow = xlsht.Cells(xlsht.Rows.Count, "A").End(xlUp).Row
    Set rngs = xlsht.Range("A2:A" & lastRow)
    
    domesticInterestRevenue = Round(domesticInterestRevenue / 1000, 0)
    
    xlsht.Range("FM11_一利息股息收入_利息_其他").Value = foreignInterestRevenue

    xlsht.Range("FM11_三證券投資處分利益_一年期以上之債權證券").Value = gainOnSecuritiesDisposal

    xlsht.Range("FM11_三證券投資處分損失_一年期以上之債權證券").Value = lossOnSecuritiesDisposal

    xlsht.Range("FM11_五證券投資評價及減損迴轉利益_一年期以上之債權證券").Value = reversalImpairmentPL

    xlsht.Range("FM11_五證券投資評價及減損損失_一年期以上之債權證券").Value = valuationImpairmentLoss

    xlsht.Range("FM11_一利息收入_自中華民國境內其他客戶").Value = domesticInterestRevenue
    
    xlsht.Range("T2:T100").NumberFormat = "#,##,##"

    

    rpt.SetField "FOA", "FM11_一利息股息收入_利息_其他", CStr(foreignInterestRevenue)

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
    Dim rngs As Range
    Dim rng As Range
    
    lastRow = xlsht.Cells(xlsht.Rows.Count, "A").End(xlUp).Row
    Set rngs = xlsht.Range("A2:A" & lastRow)
    
    domesticInterestRevenue = Round(domesticInterestRevenue / 1000, 0)
    
    xlsht.Range("FM11_一利息股息收入_利息_其他").Value = foreignInterestRevenue

    xlsht.Range("FM11_三證券投資處分利益_一年期以上之債權證券").Value = gainOnSecuritiesDisposal

    xlsht.Range("FM11_三證券投資處分損失_一年期以上之債權證券").Value = lossOnSecuritiesDisposal

    xlsht.Range("FM11_五證券投資評價及減損迴轉利益_一年期以上之債權證券").Value = reversalImpairmentPL

    xlsht.Range("FM11_五證券投資評價及減損損失_一年期以上之債權證券").Value = valuationImpairmentLoss

    xlsht.Range("FM11_一利息收入_自中華民國境內其他客戶").Value = domesticInterestRevenue
    
    xlsht.Range("T2:T100").NumberFormat = "#,##,##"

    

    rpt.SetField "FOA", "FM11_一利息股息收入_利息_其他", CStr(foreignInterestRevenue)

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
    Dim rngs As Range
    Dim rng As Range
    
    domesticInterestRevenue = Round(domesticInterestRevenue / 1000, 0)
    
    xlsht.Range("FM11_一利息股息收入_利息_其他").Value = foreignInterestRevenue

    xlsht.Range("FM11_三證券投資處分利益_一年期以上之債權證券").Value = gainOnSecuritiesDisposal

    xlsht.Range("FM11_三證券投資處分損失_一年期以上之債權證券").Value = lossOnSecuritiesDisposal

    xlsht.Range("FM11_五證券投資評價及減損迴轉利益_一年期以上之債權證券").Value = reversalImpairmentPL

    xlsht.Range("FM11_五證券投資評價及減損損失_一年期以上之債權證券").Value = valuationImpairmentLoss

    xlsht.Range("FM11_一利息收入_自中華民國境內其他客戶").Value = domesticInterestRevenue
    
    xlsht.Range("T2:T100").NumberFormat = "#,##,##"

    

    rpt.SetField "FOA", "FM11_一利息股息收入_利息_其他", CStr(foreignInterestRevenue)

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


'尚無有交易紀錄
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
    Dim rngs As Range
    Dim rng As Range

    Dim RP_GovBond_Cost As Double
    Dim RP_CompanyBond_Cost As Double
    Dim RP_CP_Cost As Double

    RP_GovBond_Cost = 0
    RP_CompanyBond_Cost = 0
    RP_CP_Cost = 0

    lastRow = xlsht.Cells(xlsht.Rows.Count, "A").End(xlUp).Row
    Set rngs = xlsht.Range("A2:A" & lastRow)

    For Each rng In rngs
        If CStr(rng.Value) = "RP_GovBond_Cost" Then
            RP_GovBond_Cost = RP_GovBond_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "AC_CompanyBond_Domestic_ImpairmentLoss" Then
            RP_CompanyBond_Cost = RP_CompanyBond_Cost + rng.Offset(0, 1).Value
        End If
    Next rng



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

    rpt.SetField "FOA", "FM11_一利息股息收入_利息_其他", CStr(foreignInterestRevenue)

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
    Dim rngs As Range
    Dim rng As Range
            
    Dim Bank_FaceValue As Double
    Dim BillSecurity_FaceValue As Double
    Dim Other_FaceValue As Double

    Bank_FaceValue = 0
    BillSecurity_FaceValue = 0
    Other_FaceValue = 0

    lastRow = xlsht.Cells(xlsht.Rows.Count, "A").End(xlUp).Row
    Set rngs = xlsht.Range("A2:A" & lastRow)

    For Each rng In rngs
        If CStr(rng.Value) = "銀行" Then
            Bank_FaceValue = Bank_FaceValue + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "票券" Then
            BillSecurity_FaceValue = BillSecurity_FaceValue + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "其他" Then
            Other_FaceValue = Other_FaceValue + rng.Offset(0, 1).Value            
        End If
    Next rng

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

    rpt.SetField "FOA", "FM11_一利息股息收入_利息_其他", CStr(foreignInterestRevenue)

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
    Dim rngs As Range
    Dim rng As Range
            
PARAMETERS DataMonthParam TEXT;
SELECT 
    IIf([BillTransactionByTradeDate.Days] >= 0 AND [BillTransactionByTradeDate.Days] <= 30, '0-30天',
    IIf([BillTransactionByTradeDate.Days] > 30 AND [BillTransactionByTradeDate.Days] <= 90, '31-90天',
    IIf([BillTransactionByTradeDate.Days] > 90 AND [BillTransactionByTradeDate.Days] <= 180, '91-180天',
    IIf([BillTransactionByTradeDate.Days] > 180 AND [BillTransactionByTradeDate.Days] <= 270, '181-270天',
    IIf([BillTransactionByTradeDate.Days] > 270 AND [BillTransactionByTradeDate.Days] <= 365, '271-365天', '其他'))))) AS DayPeriod,
    SUM(([BillTransactionByTradeDate.FaceValue] * [BillTransactionByTradeDate.TradeYield])/[BillTransactionByTradeDate.FaceValue]) AS 'FaceValue*TradeYield'
FROM 
    BillTransactionByTradeDate
WHERE
    BillTransactionByTradeDate.BillType NOT IN ('央行NCD', '一年以上央行NCD')
    AND BillTransactionByTradeDate.TransactionType NOT IN ('兌償/到期還本', '攤提', '附買回履約', '附買回解約', '附賣回履約', '附賣回解約')
    AND BillTransactionByTradeDate.DataMonthString = [DataMonthParam]
GROUP BY 
    IIf([BillTransactionByTradeDate.Days] >= 0 AND [BillTransactionByTradeDate.Days] <= 30, '0-30天',
    IIf([BillTransactionByTradeDate.Days] > 30 AND [BillTransactionByTradeDate.Days] <= 90, '31-90天',
    IIf([BillTransactionByTradeDate.Days] > 90 AND [BillTransactionByTradeDate.Days] <= 180, '91-180天',
    IIf([BillTransactionByTradeDate.Days] > 180 AND [BillTransactionByTradeDate.Days] <= 270, '181-270天',
    IIf([BillTransactionByTradeDate.Days] > 270 AND [BillTransactionByTradeDate.Days] <= 365, '271-365天', '其他')))));
    
    Dim Days30_Interest As Double
    Dim Days90_Interest As Double
    Dim Days180_Interest As Double
    Dim Days270_Interest As Double
    Dim Days365_Interest As Double

    Days30_Interest = 0
    Days90_Interest = 0
    Days180_Interest = 0
    Days270_Interest = 0
    Days365_Interest = 0

    lastRow = xlsht.Cells(xlsht.Rows.Count, "A").End(xlUp).Row
    Set rngs = xlsht.Range("A2:A" & lastRow)

    For Each rng In rngs
        If CStr(rng.Value) = "0-30天" Then
            Days30_Interest = (rng.Offset(0, 1).Value / rng.Offset(0, 2).Value)
        ElseIf CStr(rng.Value) = "31-90天" Then
            Days90_Interest = (rng.Offset(0, 1).Value / rng.Offset(0, 2).Value)
        ElseIf CStr(rng.Value) = "91-180天" Then
            Days180_Interest = (rng.Offset(0, 1).Value / rng.Offset(0, 2).Value)
        ElseIf CStr(rng.Value) = "181-270天" Then
            Days270_Interest = (rng.Offset(0, 1).Value / rng.Offset(0, 2).Value)
        ElseIf CStr(rng.Value) = "271-365天" Then
            Days365_Interest = (rng.Offset(0, 1).Value / rng.Offset(0, 2).Value)
        End If
    Next rng

    ' ****這張表要問一下怎麼做

    ' HANDLE方式

    ' END HANDLE
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

    rpt.SetField "FOA", "FM11_一利息股息收入_利息_其他", CStr(foreignInterestRevenue)

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

Public Sub Process_TABLE24()
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
    Set rpt = gReports("AI821")
    
    reportTitle = rpt.ReportName
    queryTable_1 = "AI821_OBU_MM4901B_LIST"
    queryTable_2 = "AI821_OBU_MM4901B_SUM"

    ' dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1, gDataMonthString)
    ' dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2, gDataMonthString)
    dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1, gDataMonthString)
    dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2, gDataMonthString)

    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    xlsht.Range("A:K").ClearContents
    xlsht.Range("T2:T100").ClearContents

    '=== Paste Queyr Table into Excel ===
    If Err.Number <> 0 Or LBound(dataArr_1) > UBound(dataArr_1) Then
        ' MsgBox reportTitle & ": " & queryTable_1 & " 資料表無資料"
        WriteLog reportTitle & ": " & queryTable_1 & " 資料表無資料"
    Else
        For j = 0 To UBound(dataArr_1, 2)
            For i = 0 To UBound(dataArr_1, 1)
                xlsht.Cells(i + 1, j + 1).Value = dataArr_1(i, j)
            Next i
        Next j
    End If

    If Err.Number <> 0 Or LBound(dataArr_2) > UBound(dataArr_2) Then
        ' MsgBox reportTitle & ": " & queryTable_2 & " 資料表無資料"
        WriteLog reportTitle & ": " & queryTable_2 & " 資料表無資料"
    Else
        For j = 0 To UBound(dataArr_2, 2)
            For i = 0 To UBound(dataArr_2, 1)
                xlsht.Cells(i + 1, j + 9).Value = dataArr_2(i, j)
            Next i
        Next j
    End If

    '--------------
    'Unique Setting
    '--------------
    Dim rngs As Range
    Dim rng As Range

    Dim domesticBank As Double
    Dim chinaBranchBank As Double
    Dim foreignBranchBank As Double
    Dim chinaBank As Double
    Dim others As Double

    domesticBank = 0
    chinaBranchBank = 0
    foreignBranchBank = 0
    chinaBank = 0
    others = 0

    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    If (lastRow > 1) Then       
        Set rngs = xlsht.Range("I2:I" & lastRow)

        For Each rng In rngs
            If CStr(rng.Value) = "本國銀行" Then
                domesticBank = domesticBank + rng.Offset(0, 2).Value
            ElseIf CStr(rng.Value) = "陸銀在臺分行" Then
                chinaBranchBank = chinaBranchBank + rng.Offset(0, 2).Value
            ElseIf CStr(rng.Value) = "外商銀行在臺分行" Then
                foreignBranchBank = foreignBranchBank + rng.Offset(0, 2).Value
            ElseIf CStr(rng.Value) = "大陸地區銀行" Then
                chinaBank = chinaBank + rng.Offset(0, 2).Value
            ElseIf CStr(rng.Value) = "其他" Then
                others = others + rng.Offset(0, 2).Value
            End If
        Next rng

        domesticBank = Round(domesticBank, 0)
        chinaBranchBank = Round(chinaBranchBank, 0)
        foreignBranchBank = Round(foreignBranchBank, 0)
        chinaBank = Round(chinaBank, 0)
        others = Round(others, 0)
    End If
    
    xlsht.Range("AI821_本國銀行").Value = domesticBank
    rpt.SetField "Table1", "AI821_本國銀行", CStr(domesticBank)

    xlsht.Range("AI821_陸銀在臺分行").Value = chinaBranchBank
    rpt.SetField "Table1", "AI821_陸銀在臺分行", CStr(chinaBranchBank)

    xlsht.Range("AI821_外商銀行在臺分行").Value = foreignBranchBank
    rpt.SetField "Table1", "AI821_外商銀行在臺分行", CStr(foreignBranchBank)

    xlsht.Range("AI821_大陸地區銀行").Value = chinaBank
    rpt.SetField "Table1", "AI821_大陸地區銀行", CStr(chinaBank)

    xlsht.Range("AI821_其他").Value = others
    rpt.SetField "Table1", "AI821_其他", CStr(others)
    
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

Public Sub Process_TABLE27()
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
    Set rpt = gReports("TABLE2")
    
    reportTitle = rpt.ReportName
    queryTable_1 = "表2_DBU_AC5602_TWD"
    queryTable_2 = "表2_CloseRate_USDTWD"


    ' dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1, gDataMonthString)
    ' dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2, gDataMonthString)
    dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1, gDataMonthString)
    dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2, gDataMonthString)

    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    xlsht.Range("A:I").ClearContents
    xlsht.Range("T2:T100").ClearContents

    '=== Paste Queyr Table into Excel ===
    If Err.Number <> 0 Or LBound(dataArr_1) > UBound(dataArr_1) Then
        ' MsgBox reportTitle & ": " & queryTable_1 & "資料表無資料"
        WriteLog reportTitle & ": " & queryTable_1 & "資料表無資料"
    Else
        For j = 0 To UBound(dataArr_1, 2)
            For i = 0 To UBound(dataArr_1, 1)
                xlsht.Cells(i + 1, j + 1).Value = dataArr_1(i, j)
            Next i
        Next j
    End If

    If Err.Number <> 0 Or LBound(dataArr_2) > UBound(dataArr_2) Then
        ' MsgBox reportTitle & ": " & queryTable_2 & "資料表無資料"
        WriteLog reportTitle & ": " & queryTable_2 & "資料表無資料"
    Else
        For j = 0 To UBound(dataArr_2, 2)
            For i = 0 To UBound(dataArr_2, 1)
                xlsht.Cells(i + 1, j + 7).Value = dataArr_2(i, j)
            Next i
        Next j
    End If

    '--------------
    'Unique Setting
    '--------------
    Dim rngs As Range
    Dim rng As Range

    Dim marginDeposit_TWD As Double
    Dim marginDeposit_USD As Double
    Dim rateUSDtoTWD As Double

    marginDeposit_TWD = 0
    marginDeposit_USD = 0
    rateUSDtoTWD = 0

    rateUSDtoTWD = xlsht.Range("I2").Value
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row

    If (lastRow > 1) Then
        Set rngs = xlsht.Range("C2:C" & lastRow)
        
        For Each rng In rngs
            If CStr(rng.Value) = "196017703" Then
                marginDeposit_TWD = marginDeposit_TWD + rng.Offset(0, 3).Value
            End If
        Next rng

        marginDeposit_USD = Round((marginDeposit_TWD / rateUSDtoTWD) / 1000, 0)
        marginDeposit_TWD = Round(marginDeposit_TWD / 1000, 0)

    End If
    
    xlsht.Range("Table2_A_1011100_其他").Value = marginDeposit_TWD
    rpt.SetField "FOA", "Table2_A_1011100_其他", CStr(marginDeposit_TWD)

    xlsht.Range("Table2_A_1010000_合計").Value = marginDeposit_TWD
    rpt.SetField "FOA", "Table2_A_1010000_合計", CStr(marginDeposit_TWD)

    xlsht.Range("Table2_B_01_F1_原幣國外資產").Value = marginDeposit_USD
    rpt.SetField "FOA", "Table2_B_01_F1_原幣國外資產", CStr(marginDeposit_USD)

    xlsht.Range("Table2_B_01_F3_折合率").Value = rateUSDtoTWD
    rpt.SetField "FOA", "Table2_B_01_F3_折合率", CStr(rateUSDtoTWD)

    xlsht.Range("Table2_B_01_F4_折合新台幣國外資產").Value = marginDeposit_TWD
    rpt.SetField "FOA", "Table2_B_01_F4_折合新台幣國外資產", CStr(marginDeposit_TWD)

    xlsht.Range("Table2_B_01_F4_合計").Value = marginDeposit_TWD
    rpt.SetField "FOA", "Table2_B_01_F4_合計", CStr(marginDeposit_TWD)
    
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

Public Sub Process_TABLE36()
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
    Set rpt = gReports("FB5")
    
    reportTitle = rpt.ReportName
    queryTable = "FB5_DL6320"

    ' dataArr = GetAccessDataAsArray(gDBPath, queryTable, gDataMonthString)
    dataArr = GetAccessDataAsArray(gDBPath, queryTable, gDataMonthString)
    
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    xlsht.Range("A:G").ClearContents
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

    Dim SpotToDBU_CNY As Double

    SpotToDBU_CNY = 0

    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    Set rngs = xlsht.Range("D1:D" & lastRow)

    For Each rng In rngs
        If CStr(rng.Value) = "CNY" Then
            SpotToDBU_CNY = SpotToDBU_CNY + rng.Offset(0, 1).Value
        End If
    Next rng

    SpotToDBU_CNY = Round(SpotToDBU_CNY / 1000, 0)
    
    xlsht.Range("FB5_外匯交易_即期外匯_DBU").Value = SpotToDBU_CNY
    rpt.SetField "FOA", "FB5_外匯交易_即期外匯_DBU", CStr(SpotToDBU_CNY) 
    
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


'尚無有交易紀錄
Public Sub Process_AI233()
    '=== Equal Setting ===
    'Fetch Query Access DB table
    Dim dataArr_1 As Variant
    Dim dataArr_2 As Variant

    'Declare worksheet and handle data
    Dim xlsht As Worksheet

    Dim i As Integer, j As Integer

    Dim reportTitle As String
    Dim queryTable_1 As String
    Dim queryTable_2 As String

    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("FB5A")
    
    reportTitle = rpt.ReportName
    queryTable_1 = "FB5A_OBU_FC7700B_LIST"
    queryTable_2 = "FB5A_OBU_CF6320_LIST"

    ' dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1, gDataMonthString)
    ' dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2, gDataMonthString)
    dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1, gDataMonthString)
    dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2, gDataMonthString)
    
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    xlsht.Range("A:G").ClearContents

    '=== Paste Queyr Table into Excel ===
    If Err.Number <> 0 Or LBound(dataArr_1) > UBound(dataArr_1) Then
        ' MsgBox reportTitle & ": " & queryTable_1 & " 資料表無資料"
        WriteLog reportTitle & ": " & queryTable_1 & " 資料表無資料"
    Else
        For j = 0 To UBound(dataArr_1, 2)
            For i = 0 To UBound(dataArr_1, 1)
                xlsht.Cells(i + 1, j + 1).Value = dataArr_1(i, j)
            Next i
        Next j
        MsgBox reportTitle & ": " & queryTable_1 & " 資料表有資料，此表單尚無有資料紀錄，尚請確認。"
        WriteLog reportTitle & ": " & queryTable_1 & " 資料表有資料，此表單尚無有資料紀錄，尚請確認。"
    End If

    If Err.Number <> 0 Or LBound(dataArr_2) > UBound(dataArr_2) Then
        ' MsgBox reportTitle & ": " & queryTable_2 & " 資料表無資料"
        WriteLog reportTitle & ": " & queryTable_2 & " 資料表無資料"
    Else
        For j = 0 To UBound(dataArr_2, 2)
            For i = 0 To UBound(dataArr_2, 1)
                xlsht.Cells(i + 1, j + 1).Value = dataArr_2(i, j)
            Next i
        Next j
        MsgBox reportTitle & ": " & queryTable_2 & " 資料表有資料，此表單尚無有資料紀錄，尚請確認。"
        WriteLog reportTitle & ": " & queryTable_2 & " 資料表有資料，此表單尚無有資料紀錄，尚請確認。"
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

Public Sub Process_AI345()
    '=== Equal Setting ===
    'Fetch Query Access DB table
    Dim dataArr_1 As Variant
    Dim dataArr_2 As Variant
    Dim dataArr_3 As Variant

    'Declare worksheet and handle data
    Dim xlsht As Worksheet

    Dim i As Integer, j As Integer
    Dim lastRow As Integer

    Dim reportTitle As String
    Dim queryTable_1 As String
    Dim queryTable_2 As String
    Dim queryTable_3 As String

    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("FM2")
    
    reportTitle = rpt.ReportName
    queryTable_1 = "FM2_OBU_MM4901B_LIST"
    queryTable_2 = "FM2_OBU_MM4901B_Subtotal"
    queryTable_3 = "FM2_OBU_MM4901B_Subtotal_BankCode"

    ' dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1, gDataMonthString)
    ' dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2, gDataMonthString)
    ' dataArr_3 = GetAccessDataAsArray(gDBPath, queryTable_3, gDataMonthString)
    dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1, gDataMonthString)
    dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2, gDataMonthString)
    dataArr_3 = GetAccessDataAsArray(gDBPath, queryTable_3, gDataMonthString)

    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    xlsht.Range("A:P").ClearContents
    xlsht.Range("Q2:W200").ClearContents

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


    If Err.Number <> 0 Or LBound(dataArr_3) > UBound(dataArr_3) Then
        MsgBox "資料有誤: " & reportTitle & "| " & queryTable_3 & " 資料表無資料"
        WriteLog "資料有誤: " & reportTitle & "| " & queryTable_3 & " 資料表無資料"
    Else
        For j = 0 To UBound(dataArr_3, 2)
            For i = 0 To UBound(dataArr_3, 1)
                xlsht.Cells(i + 1, j + 13).Value = dataArr_3(i, j)
            Next i
        Next j
    End If

    '--------------
    'Unique Setting
    '--------------
    Dim BankCode As Variant
    Dim CounterParty As String, Category As String
    Dim Amount As Double

    Dim pasteRow As Long
    Dim targetRow As Long
    Dim targetCol As String
    
    ' 設定第二部份記錄的起始列（Row 10）
    pasteRow = 2
    targetRow = 10
    lastRow = xlsht.Cells(xlsht.Rows.Count, "M").End(xlUp).Row
    
    ' 逐列處理原始資料（從第二列開始）
    For i = 2 To lastRow
        ' 讀取原始資料欄位值（依照題目定義的欄位順序）
        ' 原始資料欄位：
        ' A: DataID
        ' B: DataMonthString
        ' C: DealDate
        ' D: DealID
        ' E: CounterParty
        ' F: MaturityDate
        ' G: CurrencyType
        ' H: Amount
        ' I: Category
        ' J: BankCode
        

        If (Not IsEmpty(xlsht.Cells(i, "P").Value)) Then
            '銀行代碼
            BankCode = xlsht.Cells(i, "P").Value        
            'CounterParty
            CounterParty = xlsht.Cells(i, "M").Value
            ' 金額
            Amount = Round(xlsht.Cells(i, "O").value / 1000, 0)
            ' 類別 
            Category = xlsht.Cells(i, "N").Value               
            'TWTP_MP / OBU_MP / TWTP_MT / OBU_MT
            
            ' K：BankCode
            xlsht.Cells(pasteRow, "Q").Value = BankCode
            ' L：CounterParty
            xlsht.Cells(pasteRow, "R").Value = CounterParty

            ' 根據 Category 將金額填入對應分類欄位
            Select Case Category
                Case "DBU_MP"
                    ' M：TWTP_MP
                    xlsht.Cells(pasteRow, "S").Value = Amount
                Case "OBU_MP"
                    ' N：OBU_MP
                    xlsht.Cells(pasteRow, "T").Value = Amount
                Case "DBU_MT"
                    ' O：TWTP_MT
                    xlsht.Cells(pasteRow, "U").Value = Amount
                Case "OBU_MT"
                    ' P：OBU_MT
                    xlsht.Cells(pasteRow, "V").Value = Amount
            End Select
        

            ' 二、記錄儲存格位置和數值（輸出位置由 Row 10 開始）
            ' 這邊假設：BankCode 記錄在 C 欄；金額根據 Category 記錄在 E (TWTP_MP) / F (OBU_MP) / G (TWTP_MT) / H (OBU_MT)
            Select Case Category
                Case "DBU_MP"
                    targetCol = "E"
                Case "OBU_MP"
                    targetCol = "F"
                Case "DBU_MT"
                    targetCol = "G"
                Case "OBU_MT"
                    targetCol = "H"
            End Select
            
            xlsht.Cells(pasteRow, "W").Value =  targetCol & CStr(targetRow)
            ' rpt.SetField "FOA", "FM2_BankCode", "C" & CStr(targetRow), BankCode
            ' rpt.SetField "FOA", "FM2_Amount", targetCol & CStr(targetRow), Amount

            rpt.AddDynamicField "FOA", "FM2_BankCode_" & Format(BankCode, "0000"), "C" & CStr(targetRow), CStr(Format(BankCode, "0000"))
            rpt.AddDynamicField "FOA", "FM2_Amount_" & Format(BankCode, "0000"), targetCol & CStr(targetRow), CStr(Amount) 
            
            InsertIntoTable gDBPath, "MonthlyDeclarationReport", gDataMonthString, "FM2", "FOA|FM2_BankCode_" & Format(BankCode, "0000"), CStr(Format(BankCode, "0000")), "C" & CStr(targetRow)
            InsertIntoTable gDBPath, "MonthlyDeclarationReport", gDataMonthString, "FM2", "FOA|FM2_Amount_" & Format(BankCode, "0000"), CStr(Amount), targetCol & CStr(targetRow)

            UpdateRecord gDBPath, gDataMonthString, rpt.ReportName, "FOA|FM2_申報時間", "C2", gDataMonthStringROC
            
            pasteRow = pasteRow + 1
            targetRow = targetRow + 1
        End If
    Next i

    xlsht.Range("S2:S100").NumberFormat = "#,##,##"
    xlsht.Range("T2:T100").NumberFormat = "#,##,##"
    xlsht.Range("U2:U100").NumberFormat = "#,##,##"
    xlsht.Range("V2:V100").NumberFormat = "#,##,##"
    ' 更改分頁顏色為黃色(6)
    xlsht.Tab.ColorIndex = 6
End Sub

Public Sub Process_AI405()
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
    Set rpt = gReports("FM10")
    
    reportTitle = rpt.ReportName
    queryTable_1 = "FM10_OBU_AC4603_LIST"
    queryTable_2 = "FM10_OBU_AC4603_Subtotal"

    ' dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1, gDataMonthString)
    ' dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2, gDataMonthString)
    dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1, gDataMonthString)
    dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2, gDataMonthString)

    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    xlsht.Range("A:H").ClearContents
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
                xlsht.Cells(i + 1, j + 7).Value = dataArr_2(i, j)
            Next i
        Next j
    End If

    '--------------
    'Unique Setting
    '--------------
    Dim rngs As Range
    Dim rng As Range

    ' FVPL
    Dim FVPL_VALUE As Double
    Dim FVPL_ADJUSTMENT As Double
    Dim FVPL_NET_VALUE As Double
    ' FVOCI
    Dim FVOCI_VALUE As Double
    Dim FVOCI_ADJUSTMENT As Double
    Dim FVOCI_NET_VALUE As Double
    ' AC
    Dim AC_VALUE As Double
    Dim AC_ADJUSTMENT As Double
    Dim AC_NET_VALUE As Double
    Dim otherFinancialAssets As Double

    ' FVPL
    FVPL_VALUE = 0
    FVPL_ADJUSTMENT = 0
    FVPL_NET_VALUE = 0
    ' FVOCI
    FVOCI_VALUE = 0
    FVOCI_ADJUSTMENT = 0
    FVOCI_NET_VALUE = 0
    ' AC
    AC_VALUE = 0
    AC_ADJUSTMENT = 0
    AC_NET_VALUE = 0
    ' Other
    otherFinancialAssets = 0

    lastRow = xlsht.Cells(xlsht.Rows.Count, "G").End(xlUp).Row
    Set rngs = xlsht.Range("G2:G" & lastRow)

    For Each rng In rngs
        If CStr(rng.Value) = "FVPL_Cost" Then
            FVPL_VALUE = FVPL_VALUE + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVPL_ValuationAdjust" Then
            FVPL_ADJUSTMENT = FVPL_ADJUSTMENT + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVOCI_Cost" Then
            FVOCI_VALUE = FVOCI_VALUE + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVOCI_ValuationAdjust" Then
            FVOCI_ADJUSTMENT = FVOCI_ADJUSTMENT + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "AC_Cost" Then
            AC_VALUE = AC_VALUE + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "AC_ImpairmentLoss" Then
            AC_ADJUSTMENT = AC_ADJUSTMENT + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "拆放證券公司_OSU" Then
            otherFinancialAssets = otherFinancialAssets + rng.Offset(0, 1).Value
        End If
    Next rng

    FVPL_NET_VALUE = FVPL_VALUE + FVPL_ADJUSTMENT
    FVOCI_NET_VALUE = FVOCI_VALUE + FVOCI_ADJUSTMENT
    AC_NET_VALUE = AC_VALUE + AC_ADJUSTMENT

    FVPL_VALUE = Round(FVPL_VALUE / 1000, 0)
    FVPL_NET_VALUE = Round(FVPL_NET_VALUE / 1000, 0)
    FVOCI_VALUE = Round(FVOCI_VALUE / 1000, 0)
    FVOCI_NET_VALUE = Round(FVOCI_NET_VALUE / 1000, 0)
    AC_VALUE = Round(AC_VALUE / 1000, 0)
    AC_NET_VALUE = Round(AC_NET_VALUE / 1000, 0)
    otherFinancialAssets = Round(otherFinancialAssets / 1000, 0)
 
    
    xlsht.Range("FM10_FVPL_總額A").Value = FVPL_VALUE
    rpt.SetField "FOA", "FM10_FVPL_總額A", CStr(FVPL_VALUE)

    xlsht.Range("FM10_FVPL_淨額B").Value = FVPL_NET_VALUE
    rpt.SetField "FOA", "FM10_FVPL_淨額B", CStr(FVPL_NET_VALUE)
    
    xlsht.Range("FM10_FVOCI_總額C").Value = FVOCI_VALUE
    rpt.SetField "FOA", "FM10_FVOCI_總額C", CStr(FVOCI_VALUE)

    xlsht.Range("FM10_FVOCI_淨額D").Value = FVOCI_NET_VALUE
    rpt.SetField "FOA", "FM10_FVOCI_淨額D", CStr(FVOCI_NET_VALUE)

    xlsht.Range("FM10_AC_總額E").Value = AC_VALUE
    rpt.SetField "FOA", "FM10_AC_總額E", CStr(AC_VALUE)

    xlsht.Range("FM10_AC_淨額F").Value = AC_NET_VALUE
    rpt.SetField "FOA", "FM10_AC_淨額F", CStr(AC_NET_VALUE)

    xlsht.Range("FM10_四其他_境內_總額H").Value = otherFinancialAssets
    rpt.SetField "FOA", "FM10_四其他_境內_總額H", CStr(otherFinancialAssets)

    xlsht.Range("FM10_四其他_境內_淨額I").Value = otherFinancialAssets
    rpt.SetField "FOA", "FM10_四其他_境內_淨額I", CStr(otherFinancialAssets)
    
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

Public Sub Process_AI410()
    '=== Equal Setting ===
    'Fetch Query Access DB table
    ' F1
    Dim dataArr_1 As Variant
    Dim dataArr_2 As Variant
    Dim dataArr_3 As Variant
    Dim dataArr_4 As Variant
    Dim dataArr_5 As Variant
    Dim dataArr_6 As Variant
    ' F2
    Dim dataArr_7 As Variant
    Dim dataArr_8 As Variant
    Dim dataArr_9 As Variant
    Dim dataArr_10 As Variant

    'Declare worksheet and handle data
    Dim xlsht As Worksheet

    Dim i As Integer, j As Integer
    Dim lastRow As Long

    Dim reportTitle As String
    ' F1
    Dim queryTable_1 As String
    Dim queryTable_2 As String
    Dim queryTable_3 As String
    Dim queryTable_4 As String
    Dim queryTable_5 As String
    Dim queryTable_6 As String
    ' F2
    Dim queryTable_7 As String
    Dim queryTable_8 As String
    Dim queryTable_9 As String
    Dim queryTable_10 As String

    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("F1_F2")

    reportTitle = rpt.ReportName
    ' F1
    queryTable_1 = "F1_Foreign_DL6850_FS"
    queryTable_2 = "F1_Foreign_DL6850_SS"
    queryTable_3 = "F1_Domestic_DL6850_FS"
    queryTable_4 = "F1_Domestic_DL6850_SS"
    queryTable_5 = "F1_CM2810_LIST"
    queryTable_6 = "F1_CM2810_Subtotal"
    ' F2
    queryTable_7 = "F2_Foreign_DL6850_FS"
    queryTable_8 = "F2_Foreign_DL6850_SS"
    queryTable_9 = "F2_Domestic_DL6850_FS"
    queryTable_10 = "F2_Domestic_DL6850_SS"

    ' dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1, gDataMonthString)
    ' dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2, gDataMonthString)
    ' dataArr_3 = GetAccessDataAsArray(gDBPath, queryTable_3, gDataMonthString)
    ' dataArr_4 = GetAccessDataAsArray(gDBPath, queryTable_4, gDataMonthString)
    ' dataArr_5 = GetAccessDataAsArray(gDBPath, queryTable_5, gDataMonthString)
    ' dataArr_5 = GetAccessDataAsArray(gDBPath, queryTable_5, gDataMonthString)
    ' F1
    dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1, gDataMonthString)
    dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2, gDataMonthString)
    dataArr_3 = GetAccessDataAsArray(gDBPath, queryTable_3, gDataMonthString)
    dataArr_4 = GetAccessDataAsArray(gDBPath, queryTable_4, gDataMonthString)
    dataArr_5 = GetAccessDataAsArray(gDBPath, queryTable_5, gDataMonthString)
    dataArr_6 = GetAccessDataAsArray(gDBPath, queryTable_6, gDataMonthString)
    ' F2
    dataArr_7 = GetAccessDataAsArray(gDBPath, queryTable_7, gDataMonthString)
    dataArr_8 = GetAccessDataAsArray(gDBPath, queryTable_8, gDataMonthString)
    dataArr_9 = GetAccessDataAsArray(gDBPath, queryTable_9, gDataMonthString)
    dataArr_10 = GetAccessDataAsArray(gDBPath, queryTable_10, gDataMonthString)

    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    xlsht.Range("A:AA").ClearContents
    ' F1
    xlsht.Range("AC2:AC100").ClearContents
    xlsht.Range("AG2:AG100").ClearContents
    xlsht.Range("AK2:AK100").ClearContents
    xlsht.Range("AO2:AO100").ClearContents
    xlsht.Range("AS2:AS100").ClearContents
    ' F2
    xlsht.Range("AW2:AW100").ClearContents
    xlsht.Range("BA2:BA100").ClearContents
    xlsht.Range("BE2:BE100").ClearContents
    xlsht.Range("BI2:BI100").ClearContents

    
    '=== Paste Queyr Table into Excel ===
    ' F1
    If Err.Number <> 0 Or LBound(dataArr_1) > UBound(dataArr_1) Then
        MsgBox "資料有誤: " & reportTitle & "| " & queryTable_1 & "資料表無資料"
        WriteLog "資料有誤: " & reportTitle & "| " & queryTable_1 & "資料表無資料"
    Else
        For j = 0 To UBound(dataArr_1, 2)
            For i = 0 To UBound(dataArr_1, 1)
                xlsht.Cells(i + 1, j + 1).Value = dataArr_1(i, j)
            Next i
        Next j
    End If

    If Err.Number <> 0 Or LBound(dataArr_2) > UBound(dataArr_2) Then
        MsgBox "資料有誤: " & reportTitle & "| " & queryTable_2 & "資料表無資料"
        WriteLog "資料有誤: " & reportTitle & "| " & queryTable_2 & "資料表無資料"
    Else
        For j = 0 To UBound(dataArr_2, 2)
            For i = 0 To UBound(dataArr_2, 1)
                xlsht.Cells(i + 1, j + 3).Value = dataArr_2(i, j)
            Next i
        Next j
    End If

    If Err.Number <> 0 Or LBound(dataArr_3) > UBound(dataArr_3) Then
        MsgBox "資料有誤: " & reportTitle & "| " & queryTable_3 & "資料表無資料"
        WriteLog "資料有誤: " & reportTitle & "| " & queryTable_3 & "資料表無資料"
    Else
        For j = 0 To UBound(dataArr_3, 2)
            For i = 0 To UBound(dataArr_3, 1)
                xlsht.Cells(i + 1, j + 5).Value = dataArr_3(i, j)
            Next i
        Next j
    End If

    If Err.Number <> 0 Or LBound(dataArr_4) > UBound(dataArr_4) Then
        ' MsgBox reportTitle & ": " & queryTable_4 & "資料表無資料"
        WriteLog reportTitle & ": " & queryTable_4 & "資料表無資料"
    Else
        For j = 0 To UBound(dataArr_4, 2)
            For i = 0 To UBound(dataArr_4, 1)
                xlsht.Cells(i + 1, j + 7).Value = dataArr_4(i, j)
            Next i
        Next j
    End If

    If Err.Number <> 0 Or LBound(dataArr_5) > UBound(dataArr_5) Then
        MsgBox "資料有誤: " & reportTitle & "| " & queryTable_5 & "資料表無資料"
        WriteLog "資料有誤: " & reportTitle & "| " & queryTable_5 & "資料表無資料"
    Else
        For j = 0 To UBound(dataArr_5, 2)
            For i = 0 To UBound(dataArr_5, 1)
                xlsht.Cells(i + 1, j + 9).Value = dataArr_5(i, j)
            Next i
        Next j
    End If

    If Err.Number <> 0 Or LBound(dataArr_6) > UBound(dataArr_6) Then
        MsgBox "資料有誤: " & reportTitle & "| " & queryTable_6 & "資料表無資料"
        WriteLog "資料有誤: " & reportTitle & "| " & queryTable_6 & "資料表無資料"
    Else
        For j = 0 To UBound(dataArr_6, 2)
            For i = 0 To UBound(dataArr_6, 1)
                xlsht.Cells(i + 1, j + 17).Value = dataArr_6(i, j)
            Next i
        Next j
    End If

    ' F2
    If Err.Number <> 0 Or LBound(dataArr_7) > UBound(dataArr_7) Then
        MsgBox "資料有誤: " & reportTitle & "| " & queryTable_7 & "資料表無資料"
        WriteLog "資料有誤: " & reportTitle & "| " & queryTable_7 & "資料表無資料"
    Else
        For j = 0 To UBound(dataArr_7, 2)
            For i = 0 To UBound(dataArr_7, 1)
                xlsht.Cells(i + 1, j + 20).Value = dataArr_7(i, j)
            Next i
        Next j
    End If

    If Err.Number <> 0 Or LBound(dataArr_8) > UBound(dataArr_8) Then
        MsgBox "資料有誤: " & reportTitle & "| " & queryTable_8 & "資料表無資料"
        WriteLog "資料有誤: " & reportTitle & "| " & queryTable_8 & "資料表無資料"
    Else
        For j = 0 To UBound(dataArr_8, 2)
            For i = 0 To UBound(dataArr_8, 1)
                xlsht.Cells(i + 1, j + 22).Value = dataArr_8(i, j)
            Next i
        Next j
    End If

    If Err.Number <> 0 Or LBound(dataArr_9) > UBound(dataArr_9) Then
        MsgBox "資料有誤: " & reportTitle & "| " & queryTable_9 & "資料表無資料"
        WriteLog "資料有誤: " & reportTitle & "| " & queryTable_9 & "資料表無資料"
    Else
        For j = 0 To UBound(dataArr_9, 2)
            For i = 0 To UBound(dataArr_9, 1)
                xlsht.Cells(i + 1, j + 24).Value = dataArr_9(i, j)
            Next i
        Next j
    End If

    If Err.Number <> 0 Or LBound(dataArr_10) > UBound(dataArr_10) Then
        ' MsgBox reportTitle & ": " & queryTable_10 & "資料表無資料"
        WriteLog reportTitle & ": " & queryTable_10 & "資料表無資料"
    Else
        For j = 0 To UBound(dataArr_10, 2)
            For i = 0 To UBound(dataArr_10, 1)
                xlsht.Cells(i + 1, j + 26).Value = dataArr_10(i, j)
            Next i
        Next j
    End If

    Dim currencies_F1 As Variant
    Dim currencies_F2 As Variant

    currencies_F1 = Array("JPY", "GBP", "CHF", "CAD", "AUD", "NZD", "SGD", "HKD", "ZAR", "SEK", "THB", "RM", "EUR", "CNY", "OTHER")

    currencies_F2 = Array("EUR_JPY", "EUR_GBP", "EUR_CHF", "EUR_CAD", "EUR_AUD", "EUR_SGD", "EUR_HKD", "EUR_CNY", "EUR_OTHER", _
    "GBP_JPY", "GBP_CHF", "GBP_CAD", "GBP_AUD", "GBP_SGD", "GBP_HKD", "GBP_CNY", "GBP_OTHER",  _
    "JPY_CHF", "JPY_CAD", "JPY_AUD", "JPY_SGD", "JPY_HKD", "JPY_CNY", "JPY_OTHER", _
    "CNY_AUD", "CNY_SGD", "CNY_HKD", "CNY_OTHER")
    
    ' 定義交易名稱，對應到不同資料表
    Dim transactionTypes_F1 As Variant
    Dim transactionTypes_F2 As Variant

    transactionTypes_F1 = Array("F1_與國外金融機構及非金融機構間交易_SPOT", _
                            "F1_與國外金融機構及非金融機構間交易_SWAP", _
                            "F1_與國內金融機構間交易_SPOT", _
                            "F1_與國內金融機構間交易_SWAP", _
                            "F1_與國內顧客間交易_SPOT")

    transactionTypes_F2 = Array("F2_與國外金融機構及非金融機構間交易_SPOT", _
                            "F2_與國外金融機構及非金融機構間交易_SWAP", _
                            "F2_與國內金融機構間交易_SPOT", _
                            "F2_與國內金融機構間交易_SWAP")
    
    ' 對應每個交易類型在 Excel 中的欄位範圍
    Dim dataRanges_F1 As Variant
    Dim dataRanges_F2 As Variant
    ' Cur 在前一欄, Value 在後一欄
    dataRanges_F1 = Array("A:B", "C:D", "E:F", "G:H", "Q:R") 
    dataRanges_F2 = Array("T:U", "V:W", "X:Y", "Z:AA")
    
    Dim curDict As Object
    Dim currCol As Integer
    For i = LBound(transactionTypes_F1) To UBound(transactionTypes_F1)
        ' 建立字典儲存貨幣數值，並初始化為 0
        Set curDict = CreateObject("Scripting.Dictionary")
        For j = LBound(currencies_F1) To UBound(currencies_F1)
            curDict.Add currencies_F1(j), 0
        Next j
        
        ' 確定當前交易的資料範圍
        ' 取得起始欄位（Cur欄）
        currCol = xlsht.Range(dataRanges_F1(i)).Column
        lastRow = xlsht.Cells(xlsht.Rows.Count, currCol).End(xlUp).Row

        For j = 2 To lastRow ' 假設第1列是標題，從第2列開始
            Dim curCode As String, curValue As Variant
            ' 貨幣名稱
            curCode = xlsht.Cells(j, currCol).Value 
            ' 貨幣數值 百萬元，四捨五入小數第一位
            curValue = Round(xlsht.Cells(j, currCol + 1).Value / 1000000, 1) 
            
            ' 確保 Value 為數字，且 Cur 是已定義的貨幣
            If IsNumeric(curValue) And curDict.Exists(curCode) Then
                ' 若累加改成 curDict(curCode) = curDict(curCode) + curValue
                curDict(curCode) = curValue 
            End If
        Next j
        
        ' 依照固定貨幣順序填入 Excel 和報表
        For j = LBound(currencies_F1) To UBound(currencies_F1)
            Dim fieldName As String, valueNum As Variant
            ' 產生field名稱
            fieldName = transactionTypes_F1(i) & "_" & currencies_F1(j) 
            valueNum = curDict(currencies_F1(j))
        
            ' 設定 Excel 的 Range 值
            xlsht.Range(fieldName).Value = valueNum
            
            ' 設定報表欄位
            rpt.SetField "f1", fieldName, CStr(valueNum)
        Next j
    Next i


    For i = LBound(transactionTypes_F2) To UBound(transactionTypes_F2)
        ' 建立字典儲存貨幣數值，並初始化為 0
        Set curDict = CreateObject("Scripting.Dictionary")
        For j = LBound(currencies_F2) To UBound(currencies_F2)
            curDict.Add currencies_F2(j), 0
        Next j
        
        ' 確定當前交易的資料範圍
        ' 取得起始欄位（Cur欄）
        currCol = xlsht.Range(dataRanges_F2(i)).Column
        lastRow = xlsht.Cells(xlsht.Rows.Count, currCol).End(xlUp).Row

        For j = 2 To lastRow ' 假設第1列是標題，從第2列開始
            ' 貨幣名稱
            curCode = xlsht.Cells(j, currCol).Value 
            ' 貨幣數值 百萬元，四捨五入小數第一位
            curValue = Round(xlsht.Cells(j, currCol + 1).Value / 1000000, 1) 
            
            ' 確保 Value 為數字，且 Cur 是已定義的貨幣
            If IsNumeric(curValue) And curDict.Exists(curCode) Then
                ' 若累加改成 curDict(curCode) = curDict(curCode) + curValue
                curDict(curCode) = curValue 
            End If
        Next j
        
        ' 依照固定貨幣順序填入 Excel 和報表
        For j = LBound(currencies_F2) To UBound(currencies_F2)
            ' 產生field名稱
            fieldName = transactionTypes_F2(i) & "_" & currencies_F2(j) 
            valueNum = curDict(currencies_F2(j))
        
            ' 設定 Excel 的 Range 值
            xlsht.Range(fieldName).Value = valueNum
            
            ' 設定報表欄位
            rpt.SetField "f2", fieldName, CStr(valueNum)
        Next j
    Next i
    
    ' F1
    xlsht.Range("AC2:AC100").NumberFormat = "#,##,##.0"
    xlsht.Range("AG2:AG100").NumberFormat = "#,##,##.0"
    xlsht.Range("AK2:AK100").NumberFormat = "#,##,##.0"
    xlsht.Range("AO2:AO100").NumberFormat = "#,##,##.0"
    xlsht.Range("AS2:AS100").NumberFormat = "#,##,##.0"
    
    ' F2
    xlsht.Range("AW2:AW100").NumberFormat = "#,##,##.0"
    xlsht.Range("BA2:BA100").NumberFormat = "#,##,##.0"
    xlsht.Range("BE2:BE100").NumberFormat = "#,##,##.0"
    xlsht.Range("BI2:BI100").NumberFormat = "#,##,##.0"

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
    Set rpt = gReports("TABLE41")
    
    reportTitle = rpt.ReportName
    queryTable_1 = "表41_DBU_DL9360_LIST"
    queryTable_2 = "表41_DBU_DL9360_Subtotal"

    ' dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1, gDataMonthString)
    ' dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2, gDataMonthString)
    dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1, gDataMonthString)
    dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2, gDataMonthString)

    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    xlsht.Range("A:E").ClearContents
    xlsht.Range("T2:T3").ClearContents

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
                xlsht.Cells(i + 1, j + 8).Value = dataArr_2(i, j)
            Next i
        Next j
    End If

    '--------------
    'Unique Setting
    '--------------
    Dim derivativeGain As Double
    Dim derivativeLoss As Double

    derivativeGain = 0
    derivativeLoss = 0
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row

    If xlsht.Cells(1, "L").Value = "SumProfit_USD" Then
        If NOT IsEmpty(xlsht.Cells(2, "L").Value) Then
            derivativeGain = xlsht.Cells(2, "L").Value
        Else
            MsgBox "Error: No Data for Derivative Profit"
            WriteLog "Error: No Data for Derivative Profit"
        End If
    Else
        MsgBox "Error: No Data for Derivative Profit/Loss"
        WriteLog "Error: No Data for Derivative Profit/Loss"
    End If

    If xlsht.Cells(1, "M").Value = "SumLoss_USD" Then
        If NOT IsEmpty(xlsht.Cells(2, "M").Value) Then
            derivativeLoss = xlsht.Cells(2, "M").Value
        Else
            MsgBox "Error: No Data for Derivative Loss"
            WriteLog "Error: No Data for Derivative Loss"
        End If
    Else
        MsgBox "Error: No Data for Derivative Profit/Loss"
        WriteLog "Error: No Data for Derivative Profit/Loss"
    End If

    derivativeGain = Round(derivativeGain / 1000, 0)
    derivativeLoss = ABs(Round(derivativeLoss / 1000, 0))
    
    xlsht.Range("Table41_四衍生工具處分利益").Value = derivativeGain
    rpt.SetField "FOA", "Table41_四衍生工具處分利益", CStr(derivativeGain)

    xlsht.Range("Table41_四衍生工具處分損失").Value = derivativeLoss
    rpt.SetField "FOA", "Table41_四衍生工具處分損失", CStr(derivativeLoss)

    rpt.SetField "FOA", "Table41_一利息收入", CStr(xlsht.Range("Table41_一利息收入").Value)
    rpt.SetField "FOA", "Table41_一利息收入_利息", CStr(xlsht.Range("Table41_一利息收入_利息").Value)
    rpt.SetField "FOA", "Table41_一利息收入_利息_存放銀行同業", CStr(xlsht.Range("Table41_一利息收入_利息_存放銀行同業").Value)
    rpt.SetField "FOA", "Table41_二金融服務收入", CStr(xlsht.Range("Table41_二金融服務收入").Value)
    rpt.SetField "FOA", "Table41_一利息支出", CStr(xlsht.Range("Table41_一利息支出").Value)
    rpt.SetField "FOA", "Table41_一利息支出_利息", CStr(xlsht.Range("Table41_一利息支出_利息").Value)
    rpt.SetField "FOA", "Table41_一利息支出_利息_外國人新台幣存款", CStr(xlsht.Range("Table41_一利息支出_利息_外國人新台幣存款").Value)
    rpt.SetField "FOA", "Table41_一利息支出_利息_外國人外匯存款", CStr(xlsht.Range("Table41_一利息支出_利息_外國人外匯存款").Value)
    rpt.SetField "FOA", "Table41_二金融服務支出", CStr(xlsht.Range("Table41_二金融服務支出").Value)
    
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

Public Sub Process_AI601()
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
    Set rpt = gReports("AI602")
    
    reportTitle = rpt.ReportName
    queryTable_1 = "AI602_SumIpUSD"
    queryTable_2 = "AI602_Subtotal"

    ' dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1, gDataMonthString)
    ' dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2, gDataMonthString)
    ' dataArr_3 = GetAccessDataAsArray(gDBPath, queryTable_3, gDataMonthString)
    dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1, gDataMonthString)
    dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2, gDataMonthString)

    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    xlsht.Range("A:D").ClearContents
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
                xlsht.Cells(i + 1, j + 3).Value = dataArr_2(i, j)
            Next i
        Next j
    End If

    '--------------
    'Unique Setting
    '--------------
    Dim rngs As Range
    Dim rng As Range

    Dim FVPL_GovDebt_Cost As Double
    Dim FVPL_GovDebt_Adjustment As Double
    Dim FVPL_GovDebt_Impairment As Double
    Dim FVPL_GovDebt_BookValue As Double

    Dim FVPL_CompanyDebt_Cost As Double
    Dim FVPL_CompanyDebt_Adjustment As Double
    Dim FVPL_CompanyDebt_Impairment As Double
    Dim FVPL_CompanyDebt_BookValue As Double

    Dim FVPL_FinanceDebt_Cost As Double
    Dim FVPL_FinanceDebt_Adjustment As Double
    Dim FVPL_FinanceDebt_Impairment As Double
    Dim FVPL_FinanceDebt_BookValue As Double

    Dim FVOCI_GovDebt_Cost As Double
    Dim FVOCI_GovDebt_Adjustment As Double
    Dim FVOCI_GovDebt_Impairment As Double
    Dim FVOCI_GovDebt_BookValue As Double
    
    Dim FVOCI_CompanyDebt_Cost As Double
    Dim FVOCI_CompanyDebt_Adjustment As Double
    Dim FVOCI_CompanyDebt_Impairment As Double
    Dim FVOCI_CompanyDebt_BookValue As Double
    
    Dim FVOCI_FinanceDebt_Cost As Double
    Dim FVOCI_FinanceDebt_Adjustment As Double
    Dim FVOCI_FinanceDebt_Impairment As Double
    Dim FVOCI_FinanceDebt_BookValue As Double

    Dim AC_GovDebt_Cost As Double
    Dim AC_GovDebt_Impairment As Double
    Dim AC_GovDebt_BookValue As Double
    
    Dim AC_CompanyDebt_Cost As Double
    Dim AC_CompanyDebt_Impairment As Double
    Dim AC_CompanyDebt_BookValue As Double
    
    Dim AC_FinanceDebt_Cost As Double
    Dim AC_FinanceDebt_Impairment As Double
    Dim AC_FinanceDebt_BookValue As Double

    FVPL_GovDebt_Cost = 0
    FVPL_GovDebt_Adjustment = 0
    FVPL_GovDebt_Impairment = 0
    FVPL_GovDebt_BookValue = 0

    FVPL_CompanyDebt_Cost = 0
    FVPL_CompanyDebt_Adjustment = 0
    FVPL_CompanyDebt_Impairment = 0
    FVPL_CompanyDebt_BookValue = 0

    FVPL_FinanceDebt_Cost = 0
    FVPL_FinanceDebt_Adjustment = 0
    FVPL_FinanceDebt_Impairment = 0
    FVPL_FinanceDebt_BookValue = 0

    FVOCI_GovDebt_Cost = 0
    FVOCI_GovDebt_Adjustment = 0
    FVOCI_GovDebt_Impairment = 0
    FVOCI_GovDebt_BookValue = 0

    FVOCI_CompanyDebt_Cost = 0
    FVOCI_CompanyDebt_Adjustment = 0
    FVOCI_CompanyDebt_Impairment = 0
    FVOCI_CompanyDebt_BookValue = 0

    FVOCI_FinanceDebt_Cost = 0
    FVOCI_FinanceDebt_Adjustment = 0
    FVOCI_FinanceDebt_Impairment = 0
    FVOCI_FinanceDebt_BookValue = 0

    AC_GovDebt_Cost = 0
    AC_GovDebt_Impairment = 0
    AC_GovDebt_BookValue = 0

    AC_CompanyDebt_Cost = 0
    AC_CompanyDebt_Impairment = 0
    AC_CompanyDebt_BookValue = 0

    AC_FinanceDebt_Cost = 0
    AC_FinanceDebt_Impairment = 0
    AC_FinanceDebt_BookValue = 0

    lastRow = xlsht.Cells(xlsht.Rows.Count, "C").End(xlUp).Row
    Set rngs = xlsht.Range("C2:C" & lastRow)

    For Each rng In rngs
        ' FVPL 政府公債
        If CStr(rng.Value) = "FVPL_GovBond_Foreign_Cost" Then
            FVPL_GovDebt_Cost = FVPL_GovDebt_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVPL_GovBond_Foreign_ValuationAdjust" Then
            FVPL_GovDebt_Adjustment = FVPL_GovDebt_Adjustment + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVPL_GovBond_Foreign_減損" Then
            FVPL_GovDebt_Impairment = FVPL_GovDebt_Impairment + rng.Offset(0, 1).Value
        ' FVOCI 政府公債
        ElseIf CStr(rng.Value) = "FVOCI_GovBond_Foreign_Cost" Then
            FVOCI_GovDebt_Cost = FVOCI_GovDebt_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVOCI_GovBond_Foreign_ValuationAdjust" Then
            FVOCI_GovDebt_Adjustment = FVOCI_GovDebt_Adjustment + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVOCI_GovBond_Foreign_減損" Then
            FVOCI_GovDebt_Impairment = FVOCI_GovDebt_Impairment + rng.Offset(0, 1).Value
        ' AC 政府公債
        ElseIf CStr(rng.Value) = "AC_GovBond_Foreign_Cost" Then
            AC_GovDebt_Cost = AC_GovDebt_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "AC_GovBond_Foreign_減損" Then
            AC_GovDebt_Impairment = AC_GovDebt_Impairment + rng.Offset(0, 1).Value
        ' FVPL 公司債
        ElseIf CStr(rng.Value) = "FVPL_CompanyBond_Foreign_Cost" Then
            FVPL_CompanyDebt_Cost = FVPL_CompanyDebt_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVPL_CompanyBond_Foreign_ValuationAdjust" Then
            FVPL_CompanyDebt_Adjustment = FVPL_CompanyDebt_Adjustment + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVPL_CompanyBond_Foreign_減損" Then
            FVPL_CompanyDebt_Impairment = FVPL_CompanyDebt_Impairment + rng.Offset(0, 1).Value
        ' FVOCI 公司債
        ElseIf CStr(rng.Value) = "FVOCI_CompanyBond_Foreign_Cost" Then
            FVOCI_CompanyDebt_Cost = FVOCI_CompanyDebt_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVOCI_CompanyBond_Foreign_ValuationAdjust" Then
            FVOCI_CompanyDebt_Adjustment = FVOCI_CompanyDebt_Adjustment + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVOCI_CompanyBond_Foreign_減損" Then
            FVOCI_CompanyDebt_Impairment = FVOCI_CompanyDebt_Impairment + rng.Offset(0, 1).Value
        ' AC 公司債
        ElseIf CStr(rng.Value) = "AC_CompanyBond_Foreign_Cost" Then
            AC_CompanyDebt_Cost = AC_CompanyDebt_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "AC_CompanyBond_Foreign_減損" Then
            AC_CompanyDebt_Impairment = AC_CompanyDebt_Impairment + rng.Offset(0, 1).Value
        ' FVPL 金融債
        ElseIf CStr(rng.Value) = "FVPL_FinancialBond_Foreign_Cost" Then
            FVPL_FinanceDebt_Cost = FVPL_FinanceDebt_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVPL_FinancialBond_Foreign_ValuationAdjust" Then
            FVPL_FinanceDebt_Adjustment = FVPL_FinanceDebt_Adjustment + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVPL_FinancialBond_Foreign_減損" Then
            FVPL_FinanceDebt_Impairment = FVPL_FinanceDebt_Impairment + rng.Offset(0, 1).Value
        ' FVOCI 金融債
        ElseIf CStr(rng.Value) = "FVOCI_FinancialBond_Foreign_Cost" Then
            FVOCI_FinanceDebt_Cost = FVOCI_FinanceDebt_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVOCI_FinancialBond_Foreign_ValuationAdjust" Then
            FVOCI_FinanceDebt_Adjustment = FVOCI_FinanceDebt_Adjustment + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVOCI_FinancialBond_Foreign_減損" Then
            FVOCI_FinanceDebt_Impairment = FVOCI_FinanceDebt_Impairment + rng.Offset(0, 1).Value
        ' AC 金融債
        ElseIf CStr(rng.Value) = "AC_FinancialBond_Foreign_Cost" Then
            AC_FinanceDebt_Cost = AC_FinanceDebt_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "AC_FinancialBond_Foreign_減損" Then
            AC_FinanceDebt_Impairment = AC_FinanceDebt_Impairment + rng.Offset(0, 1).Value
        End If
    Next rng

    'FVOCI減損數為正數，需要扣除
    FVPL_GovDebt_BookValue = FVPL_GovDebt_Cost + FVPL_GovDebt_Adjustment - FVPL_GovDebt_Impairment 
    FVPL_CompanyDebt_BookValue = FVPL_CompanyDebt_Cost + FVPL_CompanyDebt_Adjustment - FVPL_CompanyDebt_Impairment
    FVPL_FinanceDebt_BookValue = FVPL_FinanceDebt_Cost + FVPL_FinanceDebt_Adjustment - FVPL_FinanceDebt_Impairment

    'FVOCI減損數為正數，需要扣除
    FVOCI_GovDebt_BookValue = FVOCI_GovDebt_Cost + FVOCI_GovDebt_Adjustment - FVOCI_GovDebt_Impairment 
    FVOCI_CompanyDebt_BookValue = FVOCI_CompanyDebt_Cost + FVOCI_CompanyDebt_Adjustment - FVOCI_CompanyDebt_Impairment
    FVOCI_FinanceDebt_BookValue = FVOCI_FinanceDebt_Cost + FVOCI_FinanceDebt_Adjustment - FVOCI_FinanceDebt_Impairment
    
    'AC減損數為負數，相加即可
    AC_GovDebt_BookValue = AC_GovDebt_Cost - AC_GovDebt_Impairment
    AC_CompanyDebt_BookValue = AC_CompanyDebt_Cost - AC_CompanyDebt_Impairment
    AC_FinanceDebt_BookValue = AC_FinanceDebt_Cost - AC_FinanceDebt_Impairment

    Dim sum_GovDebt_Cost As Double
    Dim sum_GovDebt_BookValue As Double
    sum_GovDebt_Cost = 0
    sum_GovDebt_BookValue = 0

    Dim sum_CompanyDebt_Cost As Double
    Dim sum_CompanyDebt_BookValue As Double
    sum_CompanyDebt_Cost = 0
    sum_CompanyDebt_BookValue = 0

    Dim sum_FinanceDebt_Cost As Double
    Dim sum_FinanceDebt_BookValue As Double
    sum_FinanceDebt_Cost = 0
    sum_FinanceDebt_BookValue = 0

    FVPL_GovDebt_Cost = Round(FVPL_GovDebt_Cost / 1000, 0)
    FVPL_GovDebt_BookValue = Round(FVPL_GovDebt_BookValue / 1000, 0)
    FVOCI_GovDebt_Cost = Round(FVOCI_GovDebt_Cost / 1000, 0)
    FVOCI_GovDebt_BookValue = Round(FVOCI_GovDebt_BookValue / 1000, 0)
    AC_GovDebt_Cost = Round(AC_GovDebt_Cost / 1000, 0)
    AC_GovDebt_BookValue = Round(AC_GovDebt_BookValue / 1000, 0)
    sum_GovDebt_Cost = FVPL_GovDebt_Cost + FVOCI_GovDebt_Cost + AC_GovDebt_Cost
    sum_GovDebt_BookValue = FVPL_GovDebt_BookValue + FVOCI_GovDebt_BookValue + AC_GovDebt_BookValue

    FVPL_CompanyDebt_Cost = Round(FVPL_CompanyDebt_Cost / 1000, 0)
    FVPL_CompanyDebt_BookValue = Round(FVPL_CompanyDebt_BookValue / 1000, 0)
    FVOCI_CompanyDebt_Cost = Round(FVOCI_CompanyDebt_Cost / 1000, 0)
    FVOCI_CompanyDebt_BookValue = Round(FVOCI_CompanyDebt_BookValue / 1000, 0)
    AC_CompanyDebt_Cost = Round(AC_CompanyDebt_Cost / 1000, 0)
    AC_CompanyDebt_BookValue = Round(AC_CompanyDebt_BookValue / 1000, 0)
    sum_CompanyDebt_Cost = FVPL_CompanyDebt_Cost + FVOCI_CompanyDebt_Cost + AC_CompanyDebt_Cost
    sum_CompanyDebt_BookValue = FVPL_CompanyDebt_BookValue + FVOCI_CompanyDebt_BookValue + AC_CompanyDebt_BookValue

    FVPL_FinanceDebt_Cost = Round(FVPL_FinanceDebt_Cost / 1000, 0)
    FVPL_FinanceDebt_BookValue = Round(FVPL_FinanceDebt_BookValue / 1000, 0)    
    FVOCI_FinanceDebt_Cost = Round(FVOCI_FinanceDebt_Cost / 1000, 0)
    FVOCI_FinanceDebt_BookValue = Round(FVOCI_FinanceDebt_BookValue / 1000, 0)
    AC_FinanceDebt_Cost = Round(AC_FinanceDebt_Cost / 1000, 0)
    AC_FinanceDebt_BookValue = Round(AC_FinanceDebt_BookValue / 1000, 0)
    sum_FinanceDebt_Cost = FVPL_FinanceDebt_Cost + FVOCI_FinanceDebt_Cost + AC_FinanceDebt_Cost
    sum_FinanceDebt_BookValue = FVPL_FinanceDebt_BookValue + FVOCI_FinanceDebt_BookValue + AC_FinanceDebt_BookValue

    xlsht.Range("AI602_政府公債_投資成本_FVPL_F1").Value = FVPL_GovDebt_Cost
    rpt.SetField "Table1", "AI602_政府公債_投資成本_FVPL_F1", CStr(FVPL_GovDebt_Cost)

    xlsht.Range("AI602_政府公債_帳面價值_FVPL_F1").Value = FVPL_GovDebt_BookValue
    rpt.SetField "Table1", "AI602_政府公債_帳面價值_FVPL_F1", CStr(FVPL_GovDebt_BookValue)

    xlsht.Range("AI602_政府公債_投資成本_FVOCI_F2").Value = FVOCI_GovDebt_Cost
    rpt.SetField "Table1", "AI602_政府公債_投資成本_FVOCI_F2", CStr(FVOCI_GovDebt_Cost)

    xlsht.Range("AI602_政府公債_帳面價值_FVOCI_F2").Value = FVOCI_GovDebt_BookValue
    rpt.SetField "Table1", "AI602_政府公債_帳面價值_FVOCI_F2", CStr(FVOCI_GovDebt_BookValue)

    xlsht.Range("AI602_政府公債_投資成本_AC_F3").Value = AC_GovDebt_Cost
    rpt.SetField "Table1", "AI602_政府公債_投資成本_AC_F3", CStr(AC_GovDebt_Cost)

    xlsht.Range("AI602_政府公債_帳面價值_AC_F3").Value = AC_GovDebt_BookValue
    rpt.SetField "Table1", "AI602_政府公債_帳面價值_AC_F3", CStr(AC_GovDebt_BookValue)

    xlsht.Range("AI602_政府公債_投資成本_合計_F5").Value = sum_GovDebt_Cost
    rpt.SetField "Table1", "AI602_政府公債_投資成本_合計_F5", CStr(sum_GovDebt_Cost)

    xlsht.Range("AI602_政府公債_帳面價值_合計_F5").Value = sum_GovDebt_BookValue
    rpt.SetField "Table1", "AI602_政府公債_帳面價值_合計_F5", CStr(sum_GovDebt_BookValue)

    xlsht.Range("AI602_公司債_投資成本_FVPL_F6").Value = FVPL_CompanyDebt_Cost
    rpt.SetField "Table1", "AI602_公司債_投資成本_FVPL_F6", CStr(FVPL_CompanyDebt_Cost)

    xlsht.Range("AI602_公司債_帳面價值_FVPL_F6").Value = FVPL_CompanyDebt_BookValue
    rpt.SetField "Table1", "AI602_公司債_帳面價值_FVPL_F6", CStr(FVPL_CompanyDebt_BookValue)
    
    xlsht.Range("AI602_公司債_投資成本_FVOCI_F7").Value = FVOCI_CompanyDebt_Cost
    rpt.SetField "Table1", "AI602_公司債_投資成本_FVOCI_F7", CStr(FVOCI_CompanyDebt_Cost)

    xlsht.Range("AI602_公司債_帳面價值_FVOCI_F7").Value = FVOCI_CompanyDebt_BookValue
    rpt.SetField "Table1", "AI602_公司債_帳面價值_FVOCI_F7", CStr(FVOCI_CompanyDebt_BookValue)

    xlsht.Range("AI602_公司債_投資成本_AC_F8").Value = AC_CompanyDebt_Cost
    rpt.SetField "Table1", "AI602_公司債_投資成本_AC_F8", CStr(AC_CompanyDebt_Cost)

    xlsht.Range("AI602_公司債_帳面價值_AC_F8").Value = AC_CompanyDebt_BookValue
    rpt.SetField "Table1", "AI602_公司債_帳面價值_AC_F8", CStr(AC_CompanyDebt_BookValue)

    xlsht.Range("AI602_公司債_投資成本_合計_F10").Value = sum_CompanyDebt_Cost
    rpt.SetField "Table1", "AI602_公司債_投資成本_合計_F10", CStr(sum_CompanyDebt_Cost)

    xlsht.Range("AI602_公司債_帳面價值_合計_F10").Value = sum_CompanyDebt_BookValue
    rpt.SetField "Table1", "AI602_公司債_帳面價值_合計_F10", CStr(sum_CompanyDebt_BookValue)

    xlsht.Range("AI602_金融債_投資成本_FVPL_F1").Value = FVPL_FinanceDebt_Cost
    rpt.SetField "Table2", "AI602_金融債_投資成本_FVPL_F1", CStr(FVPL_FinanceDebt_Cost)

    xlsht.Range("AI602_金融債_帳面價值_FVPL_F1").Value = FVPL_FinanceDebt_BookValue
    rpt.SetField "Table2", "AI602_金融債_帳面價值_FVPL_F1", CStr(FVPL_FinanceDebt_BookValue)

    xlsht.Range("AI602_金融債_投資成本_FVOCI_F2").Value = FVOCI_FinanceDebt_Cost
    rpt.SetField "Table2", "AI602_金融債_投資成本_FVOCI_F2", CStr(FVOCI_FinanceDebt_Cost)

    xlsht.Range("AI602_金融債_帳面價值_FVOCI_F2").Value = FVOCI_FinanceDebt_BookValue
    rpt.SetField "Table2", "AI602_金融債_帳面價值_FVOCI_F2", CStr(FVOCI_FinanceDebt_BookValue)

    xlsht.Range("AI602_金融債_投資成本_AC_F3").Value = AC_FinanceDebt_Cost
    rpt.SetField "Table2", "AI602_金融債_投資成本_AC_F3", CStr(AC_FinanceDebt_Cost)

    xlsht.Range("AI602_金融債_帳面價值_AC_F3").Value = AC_FinanceDebt_BookValue
    rpt.SetField "Table2", "AI602_金融債_帳面價值_AC_F3", CStr(AC_FinanceDebt_BookValue)

    xlsht.Range("AI602_金融債_投資成本_合計_F5").Value = sum_FinanceDebt_Cost
    rpt.SetField "Table2", "AI602_金融債_投資成本_合計_F5", CStr(sum_FinanceDebt_Cost)

    xlsht.Range("AI602_金融債_帳面價值_合計_F5").Value = sum_FinanceDebt_BookValue
    rpt.SetField "Table2", "AI602_金融債_帳面價值_合計_F5", CStr(sum_FinanceDebt_BookValue)
    
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

Public Sub Process_AI605()
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
    Set rpt = gReports("AI240")
    
    reportTitle = rpt.ReportName
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
