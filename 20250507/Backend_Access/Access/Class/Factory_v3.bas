Option Compare Database

Public Function GetImporter(ByVal cleaningType As String) As IImporter
    Set GetImporter = New ImporterStandard
End Function

Public Function GetCleaner(ByVal cleaningType As String) As ICleaner
    Select Case cleaningType
        'CleanerRowsDelete
        Case "DBU_MM4901B", "OBU_MM4901B"
            Set GetCleaner = New CleanerRowsDelete
            GetCleaner.Initialize 1, 1, 2, 3, Array("USD", "AUD", "CAD", "CHF", "CNY", "EUR", "GBP", "HKD", "JPY", "NZD", "SGD", "THB", "ZAR", "TWD", "總計:", "主管")
        
        'CleanerUnitACCurr
        Case "DBU_AC5602", "OBU_AC5602"
            Set GetCleaner = New CleanerUnitCurr
            GetCleaner.Initialize 1, 3, 2, 3, Array("合計", "總資產"), Array("K", "G", "E", "D", "C", "A")
        Case "OBU_AC4603"
            Set GetCleaner = New CleanerUnitCurr
            GetCleaner.Initialize 1, 3, 2, 4, Array("合計", "總資產:"), Array("K", "H", "F", "D", "C", "A")

        'CleanerRowsColsDelete
        Case "OBU_AC5411B"
            Set GetCleaner = New CleanerRowsColsDelete
            GetCleaner.Initialize 1, 1, , , Array("小計", "總收入", "總支出", "純益"), Array("J", "H", "F", "E", "D", "A")
        Case "DBU_CM2810"
            Set GetCleaner = New CleanerRowsColsDelete
            GetCleaner.Initialize 1, 1, 2, 3, Array("主管", "TWD", "總計"), Array("Q", "O")
        Case "DBU_DL9360"
            Set GetCleaner = New CleanerRowsColsDelete
            GetCleaner.Initialize 1, 2, 2, 3, Array("交易"), Array("K")
            '*******要測試為什麼isempty和其他方法刪除不掉空格row，因為lastRow沒有那一行
        Case "OBU_DL6320"
            Set GetCleaner = New CleanerRowsColsDelete
            GetCleaner.Initialize 1, 1, 2, 2, Array("總計", "襄理"), Array("R", "Q", "O", "N", "M")
        Case "DBU_DL6850"
            Set GetCleaner = New CleanerRowsColsDelete
            GetCleaner.Initialize 1, 1, 2, 2, Array("小計", "總計", "主管"), Array("L")
        
        'CleanerAC_PluralCurr
        Case "DBU_AC5601", "OBU_AC5601", "OBU_AC4620B"
            Set GetCleaner = New CleanerPluralCurr
        'CleanerIsEmpty
        Case "OBU_CF6320", "DBU_CF6850", "OBU_FC7700B", "OBU_FC9450B"
            Set GetCleaner = New CleanerIsEmpty
        
        'Special
        Case "FXDebtEvaluation"
            Set GetCleaner = New CleanerFXDebtEvaluation
        Case "AssetsImpairment"
            Set GetCleaner = New CleanerCleanAssetsImpairment
            
        ' Case "DBU_FC7810B"
        '     Set GetCleaner = New CleanerAC5601

        ' ============
        ' NT Position
        ' ============

        Case "AccountBalance"
            Set GetCleaner = New CleanerAccountBalance

        Case "BillTransactionDetails"
            ' 統一編號要補零（F 欄）
            Set GetCleaner = New CleanerOMReport
            GetCleaner.Initialize , , , , , , Array("RemoveSpaceCols", 4, "RemovePercentCols", 19)

        Case "BillRiskPositionBalance"
            Set GetCleaner = New CleanerOMReport
            GetCleaner.Initialize , , , , , , Array("FormatDateCols", 2, "FormatDateCols", 11, "FormatDateCols", 12, "FormatDateCols", 13)

        Case "BillHoldingDetails"
            Set GetCleaner = New CleanerOMReport
            GetCleaner.Initialize , , , , , , Array("RemoveSpaceCols", 7)

        Case "BillValuation"
            ' 移除空白銀行名稱
            Set GetCleaner = New CleanerOMReport
            GetCleaner.Initialize , , , , , , Array("FormatDateCols", 1, "FormatDateCols", 8, "FormatDateCols", 9, "FormatDateCols", 10)

        Case "BondValuation"
            Set GetCleaner = New CleanerOMReport
            GetCleaner.Initialize , , , , , , Array("FormatDateCols", 2, "FormatDateCols", 3, "FormatDateCols", 8, "FormatDateCols", 9)
            
        Case "BondTransactionDetails", "BondRiskPositionBalance"
            Set GetCleaner = New CleanerOMReport
            GetCleaner.Initialize , , , , , , Array()
        
        Case Else
            MsgBox "未知的清理類型: " & cleaningType
            WriteLog "未知的清理類型: " & cleaningType
            Set GetCleaner = Nothing
    End Select
End Function

Public Sub ProcessFile(ByVal filePath As String, _
                       ByVal cleaningType As String, _
                       ByVal ReportDataDate As Date, _
                       ByVal ReportMonth As Date, _
                       ByVal ReportMonthString As String, _
                       ByVal xlApp As Excel.Application)

    Dim cleaner As ICleaner
    Dim additionalCleaner As ICleaner
    Dim importer As IImporter
    Dim accessDBPath As String

    ' 取得相應的清理物件
    Set cleaner = GetCleaner(cleaningType)

    If cleaner Is Nothing Then
        MsgBox "找不到對應的清理程序: " & cleaningType
        WriteLog "找不到對應的清理程序: " & cleaningType
        Exit Sub
    End If

    ' 執行清理作業 (假設清理直接修改原檔或產生新分頁)
    cleaner.CleanReport filePath, cleaningType, xlApp

    If cleaner.HasFile Then
        If cleaner.HasData Then
            If cleaningType = "OBU_CF6320" Or _
            cleaningType = "DBU_CF6850" Or _
            cleaningType = "OBU_FC7700B" Or _
            cleaningType = "OBU_FC9450B" Then
            MsgBox "請檢查報表" & cleaningType & "，資料有異常", vbExclamation
            WriteLog "請檢查報表" & cleaningType & "，資料有異常"
            '    Debug.Print "請檢查報表" & cleaningType & "，資料有異常", vbExclamation
            Else
                'Handle additional Cleaner
                Set additionalCleaner = New CleanerAddColumns
                additionalCleaner.additionalClean filePath, cleaningType, ReportDataDate, ReportMonth, ReportMonthString, xlApp

                ' 取得相應的匯入物件
                Set importer = GetImporter(cleaningType)
                If importer Is Nothing Then
                    MsgBox "找不到對應的匯入程序: " & cleaningType
                    WriteLog "找不到對應的匯入程序: " & cleaningType
                    ' Debug.Print "找不到對應的匯入程序: " & cleaningType
                    Exit Sub
                End If

                ' 設定 Access 資料庫路徑及目標資料表
                accessDBPath = CurrentProject.FullName

                ' 執行 Access 匯入
                importer.ImportData filePath, accessDBPath, cleaningType, xlApp
            End If
        Else
            MsgBox "報表" & cleaningType & "沒有資料，不需匯入Access資料庫"
            WriteLog "報表" & cleaningType & "沒有資料，不需匯入Access資料庫"
        End If
    Else
        MsgBox "報表" & cleaningType & "檔案路徑存取錯誤，請確認檔案或路徑是否正確。"
        WriteLog "報表" & cleaningType & "檔案路徑存取錯誤，請確認檔案或路徑是否正確。"
    End If

    ' 處理完該檔案後，清除任何剪貼模式
    ' Call ClearCutCopyMode(xlApp)
End Sub

'---------------------------
'Call Return Object with Dictionary
'---------------------------
Public Sub ProcessAllReports()
    Dim ReportDataDate As Date
    Dim ReportMonth As Date
    Dim ReportMonthString As String

    'Handle clean data and import data
    Dim configDict As Object
    Dim filePathDict As Object
    Dim reportType As Variant
    Dim filePath As String

    'Custom Paths
    Dim customPathArray As Variant
    Dim customDate As Date

    'Share Excel Application
    Dim xlApp As Excel.Application
    'Set CustomPath
    customPathArray = Array()

    '取得設定
    Set configDict = GetConfigsReturnDict(customPathArray, customDate)

    '檢查是否成功取得設定
    If configDict.count = 0 Then
        MsgBox "Error: 無法取得設定資料", vbCritical
        WriteLog "Error: 無法取得設定資料"
        Exit Sub
    End If

    ' 設定報表時間資訊
    ReportDataDate = configDict("ReportDataDate")
    ReportMonth = configDict("ReportMonth")
    ReportMonthString = configDict("ReportMonthString")

    ' 設定 FilePaths
    Set filePathDict = configDict("FilePaths")

    ' Set Unit Excel Application
    Set xlApp = New Excel.Application
    Call ConfigureExcelApp(xlApp)

    ' 遍歷 FilePaths
    For Each reportType In filePathDict.Keys
        
        filePath = filePathDict(reportType)
        ProcessFile filePath, reportType, ReportDataDate, ReportMonth, ReportMonthString, xlApp
    Next reportType

    ' 所有檔案處理完畢後，還原 Excel 屬性
    Call RestoreExcelAppSettings(xlApp)
    ' 最後再呼叫 ClearCutCopyMode，作為保險備援
    ' Call ClearCutCopyMode(xlApp)

    ' Close Excel Application
    xlApp.Quit
    Set xlApp = Nothing
    MsgBox "***完成所有原始報表匯入作業***"
    WriteLog "***完成所有原始報表匯入作業***"
    ' Debug.Print "***完成所有原始報表匯入作業***"
End Sub

'---------------------------
'Call Object with Collection
'---------------------------
' Public Sub ProcessAllReports()
'     Dim reportDataDate As Date
'     Dim reportMonth As Date
'     Dim reportMonthString As String
    
'     'Handle clean data and import data
'     Dim configCollection As Object
'     Dim filePathDict As Object
'     Dim ReportType As Variant
'     Dim filePath As String

'     'Custom Paths
'     Dim customPathArray As Variant
'     Dim customDate As Date

'     'Set CustomPath
'     customPathArray = Array()

'     Set configCollection = GetConfigsReturnCollection(customPathArray, customDate)
'     If configCollection.Count < 4 Then
'         MsgBox "Error: 無法取得設定資料", vbCritical
'         Exit Sub
'     End If

'     'Set reportDataDAte/reportMonth/reportMonthString
'     reportDataDate = configCollection(1)
'     reportMonth = configCollection(2)
'     reportMonthString = configCollection(3)

'     'Set FilePaths
'     Set filePathDict = configCollection(4)
'     For Each ReportType in filePathDict.Keys
'         filePath = filePathDict(ReportType)
'         ProcessFile filePath, ReportType, reportDataDate
'     Next ReportType
' End Sub
