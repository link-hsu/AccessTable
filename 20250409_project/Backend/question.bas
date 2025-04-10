我說明一下我的架構，請幫我修改
我有一個interface ICleaner
Option Compare Database
Public Sub Initialize(Optional ByVal sheetName As Variant = 1, _
                      Optional ByVal loopColumn As Integer = 1, _
                      Optional ByVal leftToDelete As Integer = 2, _
                      Optional ByVal rightToDelete As Integer = 3, _
                      Optional ByVal rowsToDelete As Variant, _
                      Optional ByVal colsToDelete As Variant)
    'implement operations here
End Sub

Public Sub CleanReport(ByVal fullFilePath As String, _
                       ByVal cleaningType As String)
    'implement operations here
End Sub

Public Sub additionalClean(ByVal fullFilePath As String, _
                           ByVal cleaningType As String, _
                           ByVal dataDate As Date, _
                           ByVal dataMonth As Date, _
                           ByVal dataMonthString As String)
    'implement operations here
End Sub

這是第一個實作範例，主要是用在CleanReport
' CleanRowsColsDelete
Option Compare Database

Implements ICleaner

Private clsSheetName As Variant
Private clsLoopColumn As Integer
Private clsLeftToDelete As Integer
Private clsRightToDelete As Integer
Private clsRowsToDelete As Variant
Private clsColsToDelete As Variant

Public Sub ICleaner_Initialize(Optional ByVal sheetName As Variant = 1, _
                               Optional ByVal loopColumn As Integer = 1, _
                               Optional ByVal leftToDelete As Integer = 2, _
                               Optional ByVal rightToDelete As Integer = 3, _
                               Optional ByVal rowsToDelete As Variant, _
                               Optional ByVal colsToDelete As Variant)
    clsSheetName = sheetName
    clsLoopColumn = loopColumn
    clsLeftToDelete = leftToDelete
    clsRightToDelete = rightToDelete
    clsRowsToDelete = rowsToDelete
    clsColsToDelete = colsToDelete
End Sub

Public Sub ICleaner_CleanReport(ByVal fullFilePath As String, _
                                ByVal cleaningType As String)
    Dim xlApp As Excel.Application
    Dim xlbk As Excel.Workbook
    Dim xlsht As Excel.Worksheet
    Dim i As Long, lastRow As Long
    Dim col As Variant

    Dim isRowDelete As Boolean
    Dim rowToCheck As Variant
    Dim tableColumns As Variant

    If IsEmpty(clsRowsToDelete) Then clsRowsToDelete = Array()
    If IsEmpty(clsColsToDelete) Then clsColsToDelete = Array()

    ' 檢查檔案是否存在
    If Dir(fullFilePath) = "" Then
        MsgBox "File does not exist in path: " & fullFilePath
        Exit Sub
    End If

    ' 啟動 Excel 並開啟工作簿
    Set xlApp = Excel.Application
    Set xlbk = xlApp.Workbooks.Open(fullFilePath)
    Set xlsht = xlbk.Sheets(clsSheetName)

    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    For i = lastRow To 2 Step -1
        isRowDelete = False
        For Each rowToCheck In clsRowsToDelete
            If xlsht.Cells(i, clsLoopColumn).Value = rowToCheck Or _
               IsEmpty(xlsht.Cells(i, clsLoopColumn).Value) Or _
               Trim(xlsht.Cells(i, clsLoopColumn).Value) = "" Or _
               Left(xlsht.Cells(i, clsLoopColumn).Value, clsLeftToDelete) = rowToCheck Or _
               Right(xlsht.Cells(i, clsLoopColumn).Value, clsRightToDelete) = rowToCheck Then
                isRowDelete = True
                Exit For
            End If
        Next rowToCheck

        ' 刪除該列
        If isRowDelete Then xlsht.Rows(i).Delete
    Next i

    For i = lastRow To 2 Step -1
        If Left(xlsht.Cells(i, 1).Value, 2) = "主管" Or _
           Left(xlsht.Cells(i, 3).Value, 2) = "主管" Then
            xlsht.Rows.Delete
        End If
    Next i

    ' 刪除指定的欄位
    For Each col In clsColsToDelete
        xlsht.Columns(col).Delete
    Next col
    xlbk.Save
    xlbk.Close False
    Set xlsht = Nothing
    Set xlbk = Nothing
    xlApp.Quit
    Set xlApp = Nothing

    MsgBox "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
    Debug.Print "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
End Sub

Public Sub ICleaner_additionalClean(ByVal fullFilePath As String, _
                                    ByVal cleaningType As String, _
                                    ByVal dataDate As Date, _
                                    ByVal dataMonth As Date, _
                                    ByVal dataMonthString As String)
    'implement operations here
End Sub

這是第二個實作範例，主要是用在CleanReport
Option Compare Database

Implements ICleaner

Public Sub ICleaner_Initialize(Optional ByVal sheetName As Variant = 1, _
                               Optional ByVal loopColumn As Integer = 1, _
                               Optional ByVal leftToDelete As Integer = 2, _
                               Optional ByVal rightToDelete As Integer = 3, _
                               Optional ByVal rowsToDelete As Variant, _
                               Optional ByVal colsToDelete As Variant)
End Sub

Public Sub ICleaner_CleanReport(ByVal fullFilePath As String, _
                                ByVal cleaningType As String)
    Dim xlApp As Excel.Application
    Dim xlbk As Excel.Workbook
    Dim xlsht As Excel.Worksheet    

    ' Dim xlbk As Workbook
    ' Dim xlsht As Worksheet

    Dim lastRow As Long
    Dim i As Integer
    Dim colArray As Variant
    Dim valueType As Variant
    Dim eachType As Variant
    'collection save different type of row
    Dim securityIndex As Collection
    Dim securityName As Collection

    Dim startRow As Integer
    Dim endRow As Integer
    Dim innerLastRow As Integer
    Dim sheetName As String

    Dim toDelete As Boolean

    'fullFilePath = "D:\DavidHsu\testFile\vba\test\金融資產減損 1140123.xlsx"
    If Dir(fullFilePath) = "" Then
        MsgBox "File does not exist in path: " & fullFilePath
    End If
    
    Set xlApp = Excel.Application
    xlApp.ScreenUpdating = False
    xlApp.DisplayAlerts = False
    xlApp.AskToUpdateLinks = False
    Set xlbk = xlApp.Workbooks.Open(fullFilePath)
    ' Set xlbk = Workbooks.Open(fullFilePath)
    Set xlsht = xlbk.Sheets("減損")
        
    colArray = Array("Security_id", _
                     "issuer", _
                     "成本", _
                     "應收利息", _
                     "信評", _
                     "PD", _
                     "LGD", _
                     "上期減損數(成本)", _
                     "本期減損數(成本)", _
                     "上期減損數(利息)", _
                     "本期減損數(利息)")
                     
    '分頁名稱不能夠有?，會報錯誤
    valueType = Array("強制FVPL金融資產-公債-中央政府", "強制FVPL金融資產-公債-地方政府(我國)", _
                        "強制FVPL金融資產-普通公司債(公", "強制FVPL金融資產-普通公司債(民", _
                        "強制FVPL金融資產-商業本票", "FVOCI債務工具-央行NCD", _
                        "FVOCI債務工具-公債-中央政府(我", _
                        "FVOCI債務工具-公債-地方政府(我國)", _
                        "FVOCI債務工具-普通公司債（公營", _
                        "FVOCI債務工具-普通公司債（民營", _
                        "AC債務工具-央行NCD", _
                        "AC債務工具投資-公債-中央政府(?", _
                        "AC債務工具投資-公債-地方政府(?", _
                        "AC債務工具投資-普通公司債(公營", _
                        "AC債務工具投資-普通公司債(民營", _
                        "強制FVPL金融資產-公債-中央政府(外國)", _
                        "強制FVPL金融資產-普通公司債(公營)-海外", _
                        "強制FVPL金融資產-普通公司債(民營)-海外", _
                        "FVOCI債務工具-公債-中央政府(外國)", _
                        "FVOCI債務工具-普通公司債(公營)-海外", _
                        "FVOCI債務工具-普通公司債(民營)-海外", _
                        "FVOCI債務工具-金融債券-海外", _
                        "AC債務工具投資-公債-中央政府(外國)", _
                        "AC債務工具投資-普通公司債(公營)-海外", _
                        "AC債務工具投資-普通公司債(民營)-海外", _
                        "AC債務工具投資-金融債券-海外")
                        
    lastRow = xlsht.Cells(xlsht.Rows.Count, 3).End(xlUp).Row
    
    '開始建立分頁和處理資料
    Set securityName = New Collection
    Set securityIndex = New Collection
    tableColumns = GetTableColumns(cleaningType)

    For i = lastRow To 1 Step -1
        If IsEmpty(xlsht.Cells(i, 3).Value) Then
            xlsht.Rows(i).Delete
        End If
        If Left(xlsht.Cells(i, "I").Value, 5) = "利息備抵數" Then
            xlsht.Rows(i & ":" & lastRow).Delete
        End If
    Next i
    
    lastRow = xlsht.Cells(xlsht.Rows.Count, 3).End(xlUp).Row
    
    For i = 1 To lastRow
        For Each eachType In valueType
            If Trim(xlsht.Cells(i, 3).Value) = eachType Then
                securityIndex.Add i
                securityName.Add eachType
            End If
        Next eachType
    Next i
    
    For i = 1 To securityIndex.Count
        If i + 1 <= securityIndex.Count Then
            If securityIndex(i) + 1 = securityIndex(i + 1) Then
                GoTo ContinueLoop
            Else
                startRow = securityIndex(i) + 1
                endRow = securityIndex(i + 1) - 1
            End If
        Else
            startRow = securityIndex(i) + 1
            endRow = lastRow
        End If

        If InStr(securityName(i), "?") > 0 Then
            sheetName = Replace(securityName(i), "?", "")
        Else
            sheetName = securityName(i)
        End If

        xlbk.Sheets.Add After:=xlbk.Sheets(xlbk.Sheets.count)

        With ActiveSheet
            .Name = sheetName
            xlsht.Range(xlsht.Cells(startRow, "C"), xlsht.Cells(endRow, "M")).Copy
            .Range("A2").PasteSpecial Paste:=xlPasteValues
            innerLastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            .Range("L2:L" & innerLastRow).Value = securityName(i)
        End With
ContinueLoop:
    Next i
    
    For i = xlbk.Sheets.Count To 1 Step -1
        'Default cancel
        toDelete = True
        'Check wehther in valueType or not
        For Each eachType In valueType
            If xlbk.Sheets(i).Name = eachType Then
                toDelete = False
                Exit For
            End If
        Next eachType
        'Delete if not in list
        If toDelete Then
            xlbk.Sheets(i).Delete
        End If
    Next i

    xlApp.CutCopyMode = False
    xlApp.ScreenUpdating = True
    xlApp.DisplayAlerts = True
    xlApp.AskToUpdateLinks = True
    
    Set securityIndex = Nothing
    Set securityName = Nothing

    xlbk.Save
    xlbk.Close False
    Set xlsht = Nothing
    Set xlbk = Nothing
    xlApp.Quit
    Set xlApp = Nothing

    MsgBox "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
    Debug.Print "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
End Sub

Public Sub ICleaner_additionalClean(ByVal fullFilePath As String, _
                                    ByVal cleaningType As String, _
                                    ByVal dataDate As Date, _
                                    ByVal dataMonth As Date, _
                                    ByVal dataMonthString As String)
    'implement operations here
End Sub

這是針對第三個針對 additionalClean method 的實作

'CleanerAddColumns
Option Compare Database

Implements ICleaner

Public Sub ICleaner_Initialize(Optional ByVal sheetName As Variant = 1, _
                               Optional ByVal loopColumn As Integer = 1, _
                               Optional ByVal leftToDelete As Integer = 2, _
                               Optional ByVal rightToDelete As Integer = 3, _
                               Optional ByVal rowsToDelete As Variant, _
                               Optional ByVal colsToDelete As Variant)
End Sub

Public Sub ICleaner_CleanReport(ByVal fullFilePath As String, _
                                ByVal cleaningType As String)    
    'implement operations here
End Sub

Public Sub ICleaner_additionalClean(ByVal fullFilePath As String, _
                                    ByVal cleaningType As String, _
                                    ByVal dataDate As Date, _
                                    ByVal dataMonth As Date, _
                                    ByVal dataMonthString As String)
    Dim xlApp As Excel.Application
    Dim xlbk As Excel.Workbook
    Dim xlsht As Excel.Worksheet
    Dim i As Long, lastRow As Long
    Dim tableColumns As Variant

    ' 檢查檔案是否存在
    If Dir(fullFilePath) = "" Then
        MsgBox "File does not exist in path: " & fullFilePath
        Exit Sub
    End If

    ' 啟動 Excel 並開啟工作簿
    Set xlApp = Excel.Application
    xlApp.Visible = False '不顯示Excel
    Set xlbk = xlApp.Workbooks.Open(fullFilePath)
    
    'Set Table Columns
    tableColumns = GetTableColumns(cleaningType)

    For Each xlsht In xlbk.sheets
        lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
        xlsht.Columns("A:D").Insert shift:=xlToRight
        ' 反向遍歷以刪除符合條件的列
        For i = 2 To lastRow
            xlsht.Cells(i, "B").value = dataDate
            xlsht.Cells(i, "C").value = dataMonth
            xlsht.Cells(i, "D").value = dataMonthString
            If cleaningType = "DBU_AC5602" Or _
               cleaningType = "OBU_AC5602" Then
                xlsht.Cells(i, "J").Value = "TWD"
            ElseIf cleaningType = "OBU_AC4603" Then
                xlsht.Cells(i, "J").Value = "USD"
            ElseIf cleaningType = "OBU_AC5411B" Then
                xlsht.Cells(i, "I").Value = "USD"
            End If
        Next i
        xlsht.Range("A1").Resize(1, UBound(tableColumns) - LBound(tableColumns) + 1).Value = tableColumns
        xlsht.Columns("A").Delete
    Next xlsht

    ' 儲存並關閉檔案    
    xlbk.Save
    xlbk.Close False
    Set xlsht = Nothing
    Set xlbk = Nothing
    xlApp.Quit
    Set xlApp = Nothing

    MsgBox "完成Addition Clean，路徑為: " &  fullFilePath
    Debug.Print "完成Addition Clean，路徑為: " &  fullFilePath
End Sub



以下是我的主執行Sub

Option Compare Database

Public Function GetImporter(ByVal cleaningType As String) As IImporter
    ' 目前以標準匯入方式處理所有報表
    ' 若未來有特殊匯入需求，可根據 cleaningType 擴充其他 Importer 類別
    Set GetImporter = New ImporterStandard
End Function

Public Function GetCleaner(ByVal cleaningType As String) As ICleaner
    Select Case cleaningType
        'CleanerRowsDelete
        Case "DBU_MM4901B"
            Set GetCleaner = New CleanerRowsDelete
            GetCleaner.Initialize 1, 1, 2, 3, Array("USD", "AUD", "CAD", "CHF", "CNY", "EUR", "GBP", "HKD", "JPY", "NZD", "SGD", "THB", "ZAR", "TWD", "總計:", "主管")
        Case "OBU_MM4901B"
            Set GetCleaner = New CleanerRowsDelete
            GetCleaner.Initialize 1, 1, 2, 3, Array("USD", "AUD", "CAD", "CHF", "CNY", "EUR", "GBP", "HKD", "JPY", "NZD", "SGD", "THB", "ZAR", "TWD", "總計:", "主管")
        
        'CleanerUnitACCurr
        Case "DBU_AC5602"
            Set GetCleaner = New CleanerUnitCurr
            GetCleaner.Initialize 1, 3, 2, 3, Array("合計", "總資產"), Array("K", "G", "E", "D", "C", "A")
        Case "OBU_AC5602"
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
        Case "DBU_AC5601"
            Set GetCleaner = New CleanerPluralCurr
        Case "OBU_AC5601"
            Set GetCleaner = New CleanerPluralCurr
        Case "OBU_AC4620B"
            Set GetCleaner = New CleanerPluralCurr

        'CleanerIsEmpty
        Case "OBU_CF6320"
            Set GetCleaner = New CleanerIsEmpty
        Case "DBU_CF6850"
            Set GetCleaner = New CleanerIsEmpty
        Case "OBU_FC7700B"
            Set GetCleaner = New CleanerIsEmpty
        Case "OBU_FC9450B"
            Set GetCleaner = New CleanerIsEmpty
        
        'Special
        Case "FXDebtEvaluation"
            Set GetCleaner = New CleanerFXDebtEvaluation
        Case "AssetsImpairment"
            Set GetCleaner = New CleanerCleanAssetsImpairment
            
        ' Case "DBU_FC7810B"
        '     Set GetCleaner = New CleanerAC5601
        Case Else
            MsgBox "未知的清理類型: " & cleaningType
            Set GetCleaner = Nothing
    End Select
End Function



Public Sub ProcessFile(ByVal filePath As String, _
                       ByVal cleaningType As String, _
                       ByVal ReportDataDate As Date, _
                       ByVal ReportMonth As Date, _
                       ByVal ReportMonthString As String)

    Dim cleaner As ICleaner
    Dim additionalCleaner As ICleaner
    Dim importer As IImporter
    Dim accessDBPath As String

    ' 取得相應的清理物件
    Set cleaner = GetCleaner(cleaningType)

    If cleaner Is Nothing Then
        MsgBox "找不到對應的清理程序: " & cleaningType
        Exit Sub
    End If

    ' 執行清理作業 (假設清理直接修改原檔或產生新分頁)
    cleaner.CleanReport filePath, cleaningType

    'Handle additional Cleaner
     Set additionalCleaner = New CleanerAddColumns
     additionalCleaner.additionalClean filePath, cleaningType, ReportDataDate, ReportMonth, ReportMonthString

    ' 取得相應的匯入物件
     Set importer = GetImporter(cleaningType)
     If importer Is Nothing Then
         MsgBox "找不到對應的匯入程序: " & cleaningType
         Exit Sub
     End If

    ' ' 設定 Access 資料庫路徑及目標資料表
     accessDBPath = "D:\DavidHsu\ReportCreator\DB_MonthlyReport.accdb"

    ' ' 執行 Access 匯入
     importer.ImportData filePath, accessDBPath, cleaningType
End Sub

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

    'Set CustomPath
    customPathArray = Array()

    '取得設定
    Set configDict = GetConfigsReturnDict(customPathArray, customDate)

    '檢查是否成功取得設定
    If configDict.count = 0 Then
        MsgBox "Error: 無法取得設定資料", vbCritical
        Exit Sub
    End If

    ' 設定報表時間資訊
    ReportDataDate = configDict("ReportDataDate")
    ReportMonth = configDict("ReportMonth")
    ReportMonthString = configDict("ReportMonthString")

    ' 設定 FilePaths
    Set filePathDict = configDict("FilePaths")

    ' 遍歷 FilePaths
    For Each reportType In filePathDict.Keys
        filePath = filePathDict(reportType)
        ProcessFile filePath, reportType, ReportDataDate, ReportMonth, ReportMonthString
    Next reportType
End Sub

以上是我的程式碼架構，請延續之前的話題，幫我統一使用單一 Excel 執行個體來修改，給我完整內容

















Option Compare Database

Public Sub ImportData(ByVal fullFilePath As String, ByVal DBsPath As String, ByVal tableName As String)
    'implement operations here
End Sub


Option Compare Database

Implements IImporter

Public Sub IImporter_ImportData(ByVal fullFilePath As String, _
                                ByVal accessDBPath As String, _
                                ByVal tableName As String)
    Dim cn As Object
    Dim xlApp As Object
    Dim xlbk As Object
    Dim sqlString As String
    Dim sheetName As String
    Dim i As Integer, j As Integer

    Dim tableColumns As Variant

    Dim fieldList As String
    Dim selectList As String

    ' 使用 ADODB 連接 Access 資料庫
    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & accessDBPath & ";"

    ' 使用 Excel 來開啟檔案，取得所有分頁名稱
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False ' Excel 不顯示
    Set xlbk = xlApp.Workbooks.Open(fullFilePath)

    ' 取得 Access 資料表的欄位名稱
    tableColumns = GetTableColumns(tableName) ' 假設這個函式回傳欄位名稱陣列

    ' 確保 tableColumns 至少有 2 個欄位（避免 Primary Key 之外沒有欄位）
    If UBound(tableColumns) < 1 Then
    MsgBox "資料表 " & tableName & " 至少需要 2 個欄位（Primary Key + 其他欄位）。", vbCritical
    Exit Sub
    End If

    ' 動態構建 `INSERT INTO` 及 `SELECT` 語法（忽略 Primary Key）
    fieldList = ""
    selectList = ""

    For i = 1 To UBound(tableColumns) ' 從 tableColumns(1) 開始，略過 Primary Key
        fieldList = fieldList & "[" & tableColumns(i) & "],"
        selectList = selectList & "[" & tableColumns(i) & "],"
    Next i

    ' clear last comma
    fieldList = Left(fieldList, Len(fieldList) - 1)
    selectList = Left(selectList, Len(selectList) - 1)

    For i = 1 To xlbk.Sheets.count
        sheetName = xlbk.Sheets(i).Name
        ' Dynamic structure SQL Query language，skip Primary Key
        sqlString = "INSERT INTO " & tableName & " (" & fieldList & ") " & _
        "SELECT " & selectList & " FROM [Excel 12.0 Xml;HDR=YES;Database=" & fullFilePath & "].[" & sheetName & "$]"
        ' 執行 SQL
        cn.Execute sqlString
    Next i

    ' 關閉 Excel 檔案和釋放物件
    xlbk.Close False
    Set xlbk = Nothing
    xlApp.Quit
    Set xlApp = Nothing

    ' 關閉 ADODB 連接
    cn.Close
    Set cn = Nothing

    MsgBox "完成 " & tableName & " 資料表匯入作業"
    Debug.Print "完成 " & tableName & " 資料表匯入作業"

End Sub

