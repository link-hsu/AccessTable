
Answer1.




以下提供修改後的完整範例，主要改動包括：

- **修改 ICleaner 介面**：新增 xlApp 參數，讓外部呼叫者將已建立好的 Excel.Application 傳入。
- **修改各 ICleaner 實作**：移除原本在各自內部建立 Excel 物件與 Quit 的程式碼，改為使用傳入的 xlApp。
- **修改主程式**：在最外層（例如 ProcessAllReports）建立單一的 Excel.Application 物件，並將該物件傳遞給各個 Cleaner 與 Additional Cleaner。

下面分別說明並附上完整程式碼範例：

---

## 1. ICleaner 介面 (ICleaner.bas)

將原本介面中的方法，修改為接受 Excel.Application 參數

```vba
Option Compare Database

' 介面定義 (ICleaner)
Public Sub Initialize(Optional ByVal sheetName As Variant = 1, _
                      Optional ByVal loopColumn As Integer = 1, _
                      Optional ByVal leftToDelete As Integer = 2, _
                      Optional ByVal rightToDelete As Integer = 3, _
                      Optional ByVal rowsToDelete As Variant, _
                      Optional ByVal colsToDelete As Variant)
    ' implement operations here
End Sub

' 新增 xlApp 參數，讓呼叫者傳入同一個 Excel.Application 物件
Public Sub CleanReport(ByVal fullFilePath As String, _
                       ByVal cleaningType As String, _
                       ByVal xlApp As Excel.Application)
    ' implement operations here
End Sub

Public Sub additionalClean(ByVal fullFilePath As String, _
                           ByVal cleaningType As String, _
                           ByVal dataDate As Date, _
                           ByVal dataMonth As Date, _
                           ByVal dataMonthString As String, _
                           ByVal xlApp As Excel.Application)
    ' implement operations here
End Sub
```

---

## 2. 第一個 ICleaner 實作 (例如 CleanerRowsColsDelete)

修改說明：
- 移除裡面建立 xlApp（`Set xlApp = Excel.Application`）與 Quit 的程式碼，直接使用傳進來的 `xlApp` 。
- 釋放 Workbook 及 Worksheet 物件後，不再呼叫 xlApp.Quit

```vba
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

' 修改介面方法，加入 xlApp 參數
Public Sub ICleaner_CleanReport(ByVal fullFilePath As String, _
                                ByVal cleaningType As String, _
                                ByVal xlApp As Excel.Application)
    Dim xlbk As Excel.Workbook
    Dim xlsht As Excel.Worksheet
    Dim i As Long, lastRow As Long
    Dim col As Variant
    Dim isRowDelete As Boolean
    Dim rowToCheck As Variant

    If IsEmpty(clsRowsToDelete) Then clsRowsToDelete = Array()
    If IsEmpty(clsColsToDelete) Then clsColsToDelete = Array()

    ' 檢查檔案是否存在
    If Dir(fullFilePath) = "" Then
        MsgBox "File does not exist in path: " & fullFilePath
        Exit Sub
    End If

    ' 直接使用傳入的 xlApp 開啟工作簿
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
        If isRowDelete Then xlsht.Rows(i).Delete
    Next i

    For i = lastRow To 2 Step -1
        If Left(xlsht.Cells(i, 1).Value, 2) = "主管" Or _
           Left(xlsht.Cells(i, 3).Value, 2) = "主管" Then
            xlsht.Rows(i).Delete
        End If
    Next i

    For Each col In clsColsToDelete
        xlsht.Columns(col).Delete
    Next col

    xlbk.Save
    xlbk.Close False
    Set xlsht = Nothing
    Set xlbk = Nothing

    MsgBox "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
    Debug.Print "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
End Sub

' 如果有額外清理需求，也請加入 xlApp 參數 (此處僅留空作為範例)
Public Sub ICleaner_additionalClean(ByVal fullFilePath As String, _
                                    ByVal cleaningType As String, _
                                    ByVal dataDate As Date, _
                                    ByVal dataMonth As Date, _
                                    ByVal dataMonthString As String, _
                                    ByVal xlApp As Excel.Application)
    ' implement operations here
End Sub
```

---

## 3. 第二個 ICleaner 實作 (另一範例)

同樣調整成使用傳入的 xlApp，移除內部自行產生的 Excel.Application 與 Quit 的部份

```vba
Option Compare Database
Implements ICleaner

Public Sub ICleaner_Initialize(Optional ByVal sheetName As Variant = 1, _
                               Optional ByVal loopColumn As Integer = 1, _
                               Optional ByVal leftToDelete As Integer = 2, _
                               Optional ByVal rightToDelete As Integer = 3, _
                               Optional ByVal rowsToDelete As Variant, _
                               Optional ByVal colsToDelete As Variant)
    ' 如有需要，可於此初始化欄位值
End Sub

Public Sub ICleaner_CleanReport(ByVal fullFilePath As String, _
                                ByVal cleaningType As String, _
                                ByVal xlApp As Excel.Application)
    Dim xlbk As Excel.Workbook
    Dim xlsht As Excel.Worksheet    
    Dim lastRow As Long, i As Integer
    Dim colArray As Variant
    Dim valueType As Variant
    Dim eachType As Variant
    Dim securityIndex As Collection
    Dim securityName As Collection
    Dim startRow As Integer, endRow As Integer, innerLastRow As Integer
    Dim sheetName As String
    Dim toDelete As Boolean

    If Dir(fullFilePath) = "" Then
        MsgBox "File does not exist in path: " & fullFilePath
        Exit Sub
    End If
    
    xlApp.ScreenUpdating = False
    xlApp.DisplayAlerts = False
    xlApp.AskToUpdateLinks = False
    Set xlbk = xlApp.Workbooks.Open(fullFilePath)
    Set xlsht = xlbk.Sheets("減損")
        
    colArray = Array("Security_id", "issuer", "成本", "應收利息", _
                     "信評", "PD", "LGD", "上期減損數(成本)", _
                     "本期減損數(成本)", "上期減損數(利息)", "本期減損數(利息)")
                     
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
    
    Set securityName = New Collection
    Set securityIndex = New Collection

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

        xlbk.Sheets.Add After:=xlbk.Sheets(xlbk.Sheets.Count)
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
        toDelete = True
        For Each eachType In valueType
            If xlbk.Sheets(i).Name = eachType Then
                toDelete = False
                Exit For
            End If
        Next eachType
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

    MsgBox "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
    Debug.Print "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
End Sub

Public Sub ICleaner_additionalClean(ByVal fullFilePath As String, _
                                    ByVal cleaningType As String, _
                                    ByVal dataDate As Date, _
                                    ByVal dataMonth As Date, _
                                    ByVal dataMonthString As String, _
                                    ByVal xlApp As Excel.Application)
    ' implement operations here
End Sub
```

---

## 4. 主程式 (ProcessFile 與 ProcessAllReports)

修改重點在於：
- 在 ProcessAllReports 建立一個 Excel.Application 物件，並在每次呼叫 ProcessFile 時將同一個 xlApp 傳入。
- ProcessFile 內部呼叫 Cleaner 及 Additional Cleaner 時，使用新版的介面方法。

```vba
Option Compare Database

Public Function GetImporter(ByVal cleaningType As String) As IImporter
    ' 目前以標準匯入方式處理所有報表
    Set GetImporter = New ImporterStandard
End Function

Public Function GetCleaner(ByVal cleaningType As String) As ICleaner
    Select Case cleaningType
        ' 以下依據清理類型建立對應的 Cleaner 實體
        Case "DBU_MM4901B", "OBU_MM4901B"
            Set GetCleaner = New CleanerRowsDelete
            GetCleaner.Initialize 1, 1, 2, 3, Array("USD", "AUD", "CAD", "CHF", "CNY", "EUR", "GBP", "HKD", "JPY", "NZD", "SGD", "THB", "ZAR", "TWD", "總計:", "主管")
        Case "DBU_AC5602", "OBU_AC5602"
            Set GetCleaner = New CleanerUnitCurr
            GetCleaner.Initialize 1, 3, 2, 3, Array("合計", "總資產"), Array("K", "G", "E", "D", "C", "A")
        Case "OBU_AC4603"
            Set GetCleaner = New CleanerUnitCurr
            GetCleaner.Initialize 1, 3, 2, 4, Array("合計", "總資產:"), Array("K", "H", "F", "D", "C", "A")
        Case "OBU_AC5411B"
            Set GetCleaner = New CleanerRowsColsDelete
            GetCleaner.Initialize 1, 1, , , Array("小計", "總收入", "總支出", "純益"), Array("J", "H", "F", "E", "D", "A")
        Case "DBU_CM2810"
            Set GetCleaner = New CleanerRowsColsDelete
            GetCleaner.Initialize 1, 1, 2, 3, Array("主管", "TWD", "總計"), Array("Q", "O")
        Case "DBU_DL9360"
            Set GetCleaner = New CleanerRowsColsDelete
            GetCleaner.Initialize 1, 2, 2, 3, Array("交易"), Array("K")
        Case "OBU_DL6320"
            Set GetCleaner = New CleanerRowsColsDelete
            GetCleaner.Initialize 1, 1, 2, 2, Array("總計", "襄理"), Array("R", "Q", "O", "N", "M")
        Case "DBU_DL6850"
            Set GetCleaner = New CleanerRowsColsDelete
            GetCleaner.Initialize 1, 1, 2, 2, Array("小計", "總計", "主管"), Array("L")
        Case "DBU_AC5601", "OBU_AC5601", "OBU_AC4620B"
            Set GetCleaner = New CleanerPluralCurr
        Case "OBU_CF6320", "DBU_CF6850", "OBU_FC7700B", "OBU_FC9450B"
            Set GetCleaner = New CleanerIsEmpty
        Case "FXDebtEvaluation"
            Set GetCleaner = New CleanerFXDebtEvaluation
        Case "AssetsImpairment"
            Set GetCleaner = New CleanerCleanAssetsImpairment
        Case Else
            MsgBox "未知的清理類型: " & cleaningType
            Set GetCleaner = Nothing
    End Select
End Function

' 修改 ProcessFile，增加 xlApp 參數
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

    Set cleaner = GetCleaner(cleaningType)
    If cleaner Is Nothing Then
        MsgBox "找不到對應的清理程序: " & cleaningType
        Exit Sub
    End If

    ' 執行清理作業，傳入共用的 xlApp
    cleaner.CleanReport filePath, cleaningType, xlApp

    ' 如有額外的清理需求
    Set additionalCleaner = New CleanerAddColumns
    additionalCleaner.additionalClean filePath, cleaningType, ReportDataDate, ReportMonth, ReportMonthString, xlApp

    Set importer = GetImporter(cleaningType)
    If importer Is Nothing Then
        MsgBox "找不到對應的匯入程序: " & cleaningType
        Exit Sub
    End If

    accessDBPath = "D:\DavidHsu\ReportCreator\DB_MonthlyReport.accdb"
    importer.ImportData filePath, accessDBPath, cleaningType
End Sub

' 修改 ProcessAllReports，在此建立單一 Excel.Application 物件
Public Sub ProcessAllReports()
    Dim ReportDataDate As Date
    Dim ReportMonth As Date
    Dim ReportMonthString As String
    Dim configDict As Object
    Dim filePathDict As Object
    Dim reportType As Variant
    Dim filePath As String
    Dim customPathArray As Variant
    Dim customDate As Date
    Dim xlApp As Excel.Application

    customPathArray = Array()
    Set configDict = GetConfigsReturnDict(customPathArray, customDate)
    If configDict.Count = 0 Then
        MsgBox "Error: 無法取得設定資料", vbCritical
        Exit Sub
    End If

    ReportDataDate = configDict("ReportDataDate")
    ReportMonth = configDict("ReportMonth")
    ReportMonthString = configDict("ReportMonthString")
    Set filePathDict = configDict("FilePaths")

    ' 建立單一 Excel.Application 物件
    Set xlApp = New Excel.Application
    xlApp.ScreenUpdating = False
    xlApp.DisplayAlerts = False
    xlApp.AskToUpdateLinks = False

    For Each reportType In filePathDict.Keys
        filePath = filePathDict(reportType)
        ProcessFile filePath, reportType, ReportDataDate, ReportMonth, ReportMonthString, xlApp
    Next reportType

    ' 全部工作處理完後，統一關閉 Excel
    xlApp.Quit
    Set xlApp = Nothing
End Sub
```

---

## 小結

1. 在所有與 Excel 操作有關的方法中，統一改為接收一個外部建立的 **單一 Excel.Application 物件**（即 xlApp）。
2. 在主程式的 ProcessAllReports 中，先建立 xlApp，再在每個檔案的處理中傳入 xlApp，確保只使用單一 Excel 執行個體。
3. 各 Cleaner 與 Additional Cleaner 不再自行建立或關閉 Excel 物件，減少殘留進程與資源浪費。

使用上述修改後，你就可以在整個批次作業中只用一個 Excel 執行個體，有效避免多重開啟 Excel 所導致的錯誤與資源耗盡問題。如果需要進一步調整或整合其他元件，可依此模式延伸修改。



Answer2.

以下提供修改後的完整範例，主要改動重點如下：

- **修改 ICleaner 介面**：在 CleanReport 與 additionalClean 方法中，多加入一個參數 xlApp As Excel.Application，這樣所有實作都能利用由外部傳入的同一個 Excel 執行個體。  
- **各實作模組變更**：移除各自內部建立 Excel.Application 的程式碼，改由使用傳入之 xlApp，同時也不要在各模組中呼叫 xlApp.Quit。  
- **主程式改寫**：在 ProcessAllReports 中建立唯一的 Excel 執行個體，並在每次呼叫 ProcessFile 時傳入這個物件；所有檔案處理完畢後，再統一關閉 Excel 執行個體。

下面提供完整修改範例（以下以部分範例說明，其他相似模組可依此模式修改）：

---

### 1. 修改介面 ICleaner

假設原本的 ICleaner 介面內容如下，現調整方法簽章，多一個 xlApp 參數：

```vba
' ICleaner 介面 (ICleaner.bas)
Option Compare Database

' 初始化
Public Sub Initialize(Optional ByVal sheetName As Variant = 1, _
                      Optional ByVal loopColumn As Integer = 1, _
                      Optional ByVal leftToDelete As Integer = 2, _
                      Optional ByVal rightToDelete As Integer = 3, _
                      Optional ByVal rowsToDelete As Variant, _
                      Optional ByVal colsToDelete As Variant)
    ' implement operations here
End Sub

' 改良過的清理方法，傳入 Excel 執行個體
Public Sub CleanReport(ByVal fullFilePath As String, _
                       ByVal cleaningType As String, _
                       ByVal xlApp As Excel.Application)
    ' implement operations here
End Sub

' 改良過的 Additional Clean 方法，傳入 Excel 執行個體
Public Sub additionalClean(ByVal fullFilePath As String, _
                           ByVal cleaningType As String, _
                           ByVal dataDate As Date, _
                           ByVal dataMonth As Date, _
                           ByVal dataMonthString As String, _
                           ByVal xlApp As Excel.Application)
    ' implement operations here
End Sub
```

---

### 2. 修改第一個實作模組 – CleanerRowsColsDelete

以下為修改後的範例，注意重點是移除自己建立 Excel.Application 的動作，改為使用傳入的 xlApp，且結束前不呼叫 xlApp.Quit：

```vba
' CleanerRowsColsDelete 模組 (CleanerRowsColsDelete.bas)
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

' 修改後使用傳入的 xlApp
Public Sub ICleaner_CleanReport(ByVal fullFilePath As String, _
                                ByVal cleaningType As String, _
                                ByVal xlApp As Excel.Application)
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

    ' 直接使用傳入的 xlApp 開啟工作簿
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
            xlsht.Rows(i).Delete
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

    MsgBox "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
    Debug.Print "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
End Sub

Public Sub ICleaner_additionalClean(ByVal fullFilePath As String, _
                                    ByVal cleaningType As String, _
                                    ByVal dataDate As Date, _
                                    ByVal dataMonth As Date, _
                                    ByVal dataMonthString As String, _
                                    ByVal xlApp As Excel.Application)
    ' 此模組若沒有額外處理，可直接留空或進行適當實作
    ' implement operations here
End Sub
```

---

### 3. 修改第二個實作模組（例如 CleanerUnitCurr 或其他 CleanReport 實作）

以下示範第二個 CleanReport 實作的部分改法，主要同樣將內部建立 xlApp 的部分拿掉，改用傳入參數；同時所有結束前僅關閉 Workbook，不呼叫 xlApp.Quit：

```vba
' 例如：CleanerUnitCurr 模組 (示意內容)
Option Compare Database
Implements ICleaner

Public Sub ICleaner_Initialize(Optional ByVal sheetName As Variant = 1, _
                               Optional ByVal loopColumn As Integer = 1, _
                               Optional ByVal leftToDelete As Integer = 2, _
                               Optional ByVal rightToDelete As Integer = 3, _
                               Optional ByVal rowsToDelete As Variant, _
                               Optional ByVal colsToDelete As Variant)
    ' 若有需要初始化的參數，可在此加入
End Sub

Public Sub ICleaner_CleanReport(ByVal fullFilePath As String, _
                                ByVal cleaningType As String, _
                                ByVal xlApp As Excel.Application)
    Dim xlbk As Excel.Workbook
    Dim xlsht As Excel.Worksheet
    Dim lastRow As Long
    Dim i As Integer
    ' ...其他變數宣告...

    If Dir(fullFilePath) = "" Then
        MsgBox "File does not exist in path: " & fullFilePath
        Exit Sub
    End If

    xlApp.ScreenUpdating = False
    xlApp.DisplayAlerts = False
    xlApp.AskToUpdateLinks = False
    Set xlbk = xlApp.Workbooks.Open(fullFilePath)
    Set xlsht = xlbk.Sheets("減損")
    
    ' 輸入你現有的清理邏輯
    lastRow = xlsht.Cells(xlsht.Rows.Count, 3).End(xlUp).Row
    
    ' 略過部分處理，請依原邏輯置入
    ' …(處理程式碼)…
    
    xlApp.CutCopyMode = False
    xlApp.ScreenUpdating = True
    xlApp.DisplayAlerts = True
    xlApp.AskToUpdateLinks = True
    
    xlbk.Save
    xlbk.Close False
    Set xlsht = Nothing
    Set xlbk = Nothing

    MsgBox "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
    Debug.Print "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
End Sub

Public Sub ICleaner_additionalClean(ByVal fullFilePath As String, _
                                    ByVal cleaningType As String, _
                                    ByVal dataDate As Date, _
                                    ByVal dataMonth As Date, _
                                    ByVal dataMonthString As String, _
                                    ByVal xlApp As Excel.Application)
    ' implement operations here，同理改用傳入之 xlApp
End Sub
```

---

### 4. 修改 CleanerAddColumns 實作（針對 additionalClean 方法）

同樣的，將內部建立 Excel.Application 的部分拿掉，使用傳入的 xlApp：

```vba
' CleanerAddColumns 模組 (CleanerAddColumns.bas)
Option Compare Database
Implements ICleaner

Public Sub ICleaner_Initialize(Optional ByVal sheetName As Variant = 1, _
                               Optional ByVal loopColumn As Integer = 1, _
                               Optional ByVal leftToDelete As Integer = 2, _
                               Optional ByVal rightToDelete As Integer = 3, _
                               Optional ByVal rowsToDelete As Variant, _
                               Optional ByVal colsToDelete As Variant)
    ' 如有初始化參數則實作
End Sub

Public Sub ICleaner_CleanReport(ByVal fullFilePath As String, _
                                ByVal cleaningType As String, _
                                ByVal xlApp As Excel.Application)
    ' 本模組主要針對 addition clean，所以此處可留空或簡單實作
    ' implement operations here (如不需要此方法，可不做處理)
End Sub

Public Sub ICleaner_additionalClean(ByVal fullFilePath As String, _
                                    ByVal cleaningType As String, _
                                    ByVal dataDate As Date, _
                                    ByVal dataMonth As Date, _
                                    ByVal dataMonthString As String, _
                                    ByVal xlApp As Excel.Application)
    Dim xlbk As Excel.Workbook
    Dim xlsht As Excel.Worksheet
    Dim i As Long, lastRow As Long
    Dim tableColumns As Variant

    If Dir(fullFilePath) = "" Then
        MsgBox "File does not exist in path: " & fullFilePath
        Exit Sub
    End If

    xlApp.Visible = False
    Set xlbk = xlApp.Workbooks.Open(fullFilePath)
    
    tableColumns = GetTableColumns(cleaningType)

    For Each xlsht In xlbk.Sheets
        lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
        xlsht.Columns("A:D").Insert Shift:=xlToRight
        For i = 2 To lastRow
            xlsht.Cells(i, "B").Value = dataDate
            xlsht.Cells(i, "C").Value = dataMonth
            xlsht.Cells(i, "D").Value = dataMonthString
            If cleaningType = "DBU_AC5602" Or cleaningType = "OBU_AC5602" Then
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

    xlbk.Save
    xlbk.Close False
    Set xlsht = Nothing
    Set xlbk = Nothing

    MsgBox "完成Addition Clean，路徑為: " & fullFilePath
    Debug.Print "完成Addition Clean，路徑為: " & fullFilePath
End Sub
```

---

### 5. 修改主程式模組

在主程式中，我們先建立單一的 Excel 執行個體，並將此物件傳遞至各個檔案處理程序，處理完所有檔案後再關閉此 Excel 物件。以下是一個完整的範例：

```vba
' 主程式模組 (MainModule.bas)
Option Compare Database

Public Function GetImporter(ByVal cleaningType As String) As IImporter
    ' 目前以標準匯入方式處理所有報表
    Set GetImporter = New ImporterStandard
End Function

Public Function GetCleaner(ByVal cleaningType As String) As ICleaner
    Select Case cleaningType
        ' CleanerRowsDelete
        Case "DBU_MM4901B", "OBU_MM4901B"
            Set GetCleaner = New CleanerRowsDelete
            GetCleaner.Initialize 1, 1, 2, 3, Array("USD", "AUD", "CAD", "CHF", "CNY", "EUR", "GBP", _
                                                     "HKD", "JPY", "NZD", "SGD", "THB", "ZAR", "TWD", "總計:", "主管")
        ' CleanerUnitACCurr
        Case "DBU_AC5602", "OBU_AC5602"
            Set GetCleaner = New CleanerUnitCurr
            GetCleaner.Initialize 1, 3, 2, 3, Array("合計", "總資產"), Array("K", "G", "E", "D", "C", "A")
        Case "OBU_AC4603"
            Set GetCleaner = New CleanerUnitCurr
            GetCleaner.Initialize 1, 3, 2, 4, Array("合計", "總資產:"), Array("K", "H", "F", "D", "C", "A")
        ' CleanerRowsColsDelete
        Case "OBU_AC5411B"
            Set GetCleaner = New CleanerRowsColsDelete
            GetCleaner.Initialize 1, 1, , , Array("小計", "總收入", "總支出", "純益"), Array("J", "H", "F", "E", "D", "A")
        Case "DBU_CM2810"
            Set GetCleaner = New CleanerRowsColsDelete
            GetCleaner.Initialize 1, 1, 2, 3, Array("主管", "TWD", "總計"), Array("Q", "O")
        Case "DBU_DL9360"
            Set GetCleaner = New CleanerRowsColsDelete
            GetCleaner.Initialize 1, 2, 2, 3, Array("交易"), Array("K")
        Case "OBU_DL6320"
            Set GetCleaner = New CleanerRowsColsDelete
            GetCleaner.Initialize 1, 1, 2, 2, Array("總計", "襄理"), Array("R", "Q", "O", "N", "M")
        Case "DBU_DL6850"
            Set GetCleaner = New CleanerRowsColsDelete
            GetCleaner.Initialize 1, 1, 2, 2, Array("小計", "總計", "主管"), Array("L")
        ' CleanerAC_PluralCurr
        Case "DBU_AC5601", "OBU_AC5601", "OBU_AC4620B"
            Set GetCleaner = New CleanerPluralCurr
        ' CleanerIsEmpty
        Case "OBU_CF6320", "DBU_CF6850", "OBU_FC7700B", "OBU_FC9450B"
            Set GetCleaner = New CleanerIsEmpty
        ' Special
        Case "FXDebtEvaluation"
            Set GetCleaner = New CleanerFXDebtEvaluation
        Case "AssetsImpairment"
            Set GetCleaner = New CleanerCleanAssetsImpairment
        Case Else
            MsgBox "未知的清理類型: " & cleaningType
            Set GetCleaner = Nothing
    End Select
End Function

' 修改 ProcessFile，增加 xlApp 參數
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
        Exit Sub
    End If

    ' 執行清理作業 (傳入單一的 xlApp)
    cleaner.CleanReport filePath, cleaningType, xlApp

    ' 處理 Additional Clean
    Set additionalCleaner = New CleanerAddColumns
    additionalCleaner.additionalClean filePath, cleaningType, ReportDataDate, ReportMonth, ReportMonthString, xlApp

    ' 取得相應的匯入物件
    Set importer = GetImporter(cleaningType)
    If importer Is Nothing Then
        MsgBox "找不到對應的匯入程序: " & cleaningType
        Exit Sub
    End If

    accessDBPath = "D:\DavidHsu\ReportCreator\DB_MonthlyReport.accdb"
    importer.ImportData filePath, accessDBPath, cleaningType
End Sub

' 主程序，建立單一 Excel 執行個體供所有檔案使用
Public Sub ProcessAllReports()
    Dim ReportDataDate As Date, ReportMonth As Date
    Dim ReportMonthString As String
    Dim configDict As Object, filePathDict As Object
    Dim reportType As Variant, filePath As String
    Dim customPathArray As Variant, customDate As Date
    Dim xlApp As Excel.Application

    customPathArray = Array()
    Set configDict = GetConfigsReturnDict(customPathArray, customDate)
    If configDict.Count = 0 Then
        MsgBox "Error: 無法取得設定資料", vbCritical
        Exit Sub
    End If

    ReportDataDate = configDict("ReportDataDate")
    ReportMonth = configDict("ReportMonth")
    ReportMonthString = configDict("ReportMonthString")
    Set filePathDict = configDict("FilePaths")

    ' 建立一個 Excel 執行個體（僅建立一次）
    Set xlApp = New Excel.Application
    xlApp.Visible = False
    xlApp.ScreenUpdating = False
    xlApp.DisplayAlerts = False
    xlApp.AskToUpdateLinks = False

    ' 遍歷 FilePaths 處理各報表
    For Each reportType In filePathDict.Keys
        filePath = filePathDict(reportType)
        ProcessFile filePath, reportType, ReportDataDate, ReportMonth, ReportMonthString, xlApp
    Next reportType

    ' 處理完畢後關閉 Excel 應用程式
    xlApp.Quit
    Set xlApp = Nothing
End Sub
```

---

### 小結

1. **單一 Excel 執行個體**：在 ProcessAllReports 中建立一次 Excel.Application，並傳遞給所有 CleanReport 與 additionalClean 方法，避免每次重複建立與關閉 Excel。  
2. **物件釋放**：各個模組在關閉 Workbook 與釋放 Worksheet 後，**不再呼叫 xlApp.Quit**，以免中斷共用的 Excel 實例。  
3. **統一管理**：最後在 ProcessAllReports 裡完成所有處理後，再統一呼叫 xlApp.Quit。

依照此架構修改後，便能有效避免因多重 Excel 實例導致系統資源不足或 "Excel 正在忙碌中" 的問題。



Answer3.

延續上面問題，請問有關
    xlApp.Visible
    xlApp.ScreenUpdating
    xlApp.DisplayAlerts
    xlApp.AskToUpdateLinks
    xlApp.CutCopyMode
寫的好像有點混亂，有沒有建議是否可以一起管理？請給我完整版本



以下提供一個建議做法，把 Excel 的設定封裝在獨立的副程式中，用來設定與還原 Excel 的屬性，讓主程式以及各個 Cleaner 模組都能統一管理這些設定。以下是一個完整版本的範例：

---

### 1. 建立 Excel 屬性管理模組

```vba
' 模組：modExcelSettings
Option Compare Database

' 設定 Excel 屬性（在開始工作前呼叫）
Public Sub ConfigureExcelApp(ByRef xlApp As Excel.Application)
    With xlApp
        .Visible = False             ' 不顯示 Excel 畫面
        .ScreenUpdating = False      ' 關閉畫面更新，提高效能
        .DisplayAlerts = False       ' 關閉警告訊息，自動接受預設
        .AskToUpdateLinks = False    ' 關閉連結更新詢問
    End With
End Sub

' 還原 Excel 屬性（在工作完成後呼叫）
Public Sub RestoreExcelAppSettings(ByRef xlApp As Excel.Application)
    With xlApp
        .ScreenUpdating = True
        .DisplayAlerts = True
        .AskToUpdateLinks = True
    End With
End Sub

' 清除剪貼簿狀態
Public Sub ClearCutCopyMode(ByRef xlApp As Excel.Application)
    xlApp.CutCopyMode = False
End Sub
```

---

### 2. 調整 ICleaner 介面

這裡僅修改方法簽章，增加 xlApp 參數（其他內容與之前一致）：

```vba
' ICleaner 介面 (ICleaner.bas)
Option Compare Database

Public Sub Initialize(Optional ByVal sheetName As Variant = 1, _
                      Optional ByVal loopColumn As Integer = 1, _
                      Optional ByVal leftToDelete As Integer = 2, _
                      Optional ByVal rightToDelete As Integer = 3, _
                      Optional ByVal rowsToDelete As Variant, _
                      Optional ByVal colsToDelete As Variant)

End Sub

' 修改過的清理方法，增加 xlApp 參數
Public Sub CleanReport(ByVal fullFilePath As String, _
                       ByVal cleaningType As String, _
                       ByVal xlApp As Excel.Application)

End Sub

' 修改過的 additionalClean 方法，增加 xlApp 參數
Public Sub additionalClean(ByVal fullFilePath As String, _
                           ByVal cleaningType As String, _
                           ByVal dataDate As Date, _
                           ByVal dataMonth As Date, _
                           ByVal dataMonthString As String, _
                           ByVal xlApp As Excel.Application)

End Sub
```

---

### 3. 修改各個 ICleaner 實作模組

以「CleanerRowsColsDelete」為例，其修改後的內容如下，重點是移除內部建立 xlApp 的部分，使用傳入的物件，同時不重複設定或還原 Excel 屬性，交由主程式統一管理：

```vba
' CleanerRowsColsDelete 模組 (CleanerRowsColsDelete.bas)
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
                                ByVal cleaningType As String, _
                                ByVal xlApp As Excel.Application)
    Dim xlbk As Excel.Workbook
    Dim xlsht As Excel.Worksheet
    Dim i As Long, lastRow As Long
    Dim col As Variant
    Dim isRowDelete As Boolean
    Dim rowToCheck As Variant

    If IsEmpty(clsRowsToDelete) Then clsRowsToDelete = Array()
    If IsEmpty(clsColsToDelete) Then clsColsToDelete = Array()

    ' 檢查檔案是否存在
    If Dir(fullFilePath) = "" Then
        MsgBox "File does not exist in path: " & fullFilePath
        Exit Sub
    End If

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

        If isRowDelete Then xlsht.Rows(i).Delete
    Next i

    For i = lastRow To 2 Step -1
        If Left(xlsht.Cells(i, 1).Value, 2) = "主管" Or _
           Left(xlsht.Cells(i, 3).Value, 2) = "主管" Then
            xlsht.Rows(i).Delete
        End If
    Next i

    For Each col In clsColsToDelete
        xlsht.Columns(col).Delete
    Next col

    xlbk.Save
    xlbk.Close False
    Set xlsht = Nothing
    Set xlbk = Nothing

    MsgBox "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
    Debug.Print "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
End Sub

Public Sub ICleaner_additionalClean(ByVal fullFilePath As String, _
                                    ByVal cleaningType As String, _
                                    ByVal dataDate As Date, _
                                    ByVal dataMonth As Date, _
                                    ByVal dataMonthString As String, _
                                    ByVal xlApp As Excel.Application)
    ' 此模組若不需要額外處理，可留空或依需求實作
End Sub
```

其他實作模組（例如 CleanerUnitCurr、CleanerAddColumns 等）皆採用相同方式：  
*移除內部建立 xlApp 與關閉 Quit 的動作，直接使用傳入的 xlApp 物件。*

---

### 4. 修改主程式模組

在 ProcessAllReports 中先建立單一的 Excel 執行個體，透過剛才所建立的模組進行屬性設定與還原。每次呼叫 ProcessFile 時皆傳入該 xlApp 物件，所有檔案處理完畢後再一次關閉 Excel 執行個體。以下為完整範例：

```vba
' 主程式模組 (MainModule.bas)
Option Compare Database

Public Function GetImporter(ByVal cleaningType As String) As IImporter
    ' 目前以標準匯入方式處理所有報表
    Set GetImporter = New ImporterStandard
End Function

Public Function GetCleaner(ByVal cleaningType As String) As ICleaner
    Select Case cleaningType
        ' CleanerRowsDelete
        Case "DBU_MM4901B", "OBU_MM4901B"
            Set GetCleaner = New CleanerRowsDelete
            GetCleaner.Initialize 1, 1, 2, 3, _
                Array("USD", "AUD", "CAD", "CHF", "CNY", "EUR", "GBP", "HKD", "JPY", "NZD", "SGD", "THB", "ZAR", "TWD", "總計:", "主管")
        ' CleanerUnitACCurr
        Case "DBU_AC5602", "OBU_AC5602"
            Set GetCleaner = New CleanerUnitCurr
            GetCleaner.Initialize 1, 3, 2, 3, Array("合計", "總資產"), Array("K", "G", "E", "D", "C", "A")
        Case "OBU_AC4603"
            Set GetCleaner = New CleanerUnitCurr
            GetCleaner.Initialize 1, 3, 2, 4, Array("合計", "總資產:"), Array("K", "H", "F", "D", "C", "A")
        ' CleanerRowsColsDelete
        Case "OBU_AC5411B"
            Set GetCleaner = New CleanerRowsColsDelete
            GetCleaner.Initialize 1, 1, , , Array("小計", "總收入", "總支出", "純益"), Array("J", "H", "F", "E", "D", "A")
        Case "DBU_CM2810"
            Set GetCleaner = New CleanerRowsColsDelete
            GetCleaner.Initialize 1, 1, 2, 3, Array("主管", "TWD", "總計"), Array("Q", "O")
        Case "DBU_DL9360"
            Set GetCleaner = New CleanerRowsColsDelete
            GetCleaner.Initialize 1, 2, 2, 3, Array("交易"), Array("K")
        Case "OBU_DL6320"
            Set GetCleaner = New CleanerRowsColsDelete
            GetCleaner.Initialize 1, 1, 2, 2, Array("總計", "襄理"), Array("R", "Q", "O", "N", "M")
        Case "DBU_DL6850"
            Set GetCleaner = New CleanerRowsColsDelete
            GetCleaner.Initialize 1, 1, 2, 2, Array("小計", "總計", "主管"), Array("L")
        ' CleanerAC_PluralCurr
        Case "DBU_AC5601", "OBU_AC5601", "OBU_AC4620B"
            Set GetCleaner = New CleanerPluralCurr
        ' CleanerIsEmpty
        Case "OBU_CF6320", "DBU_CF6850", "OBU_FC7700B", "OBU_FC9450B"
            Set GetCleaner = New CleanerIsEmpty
        ' Special
        Case "FXDebtEvaluation"
            Set GetCleaner = New CleanerFXDebtEvaluation
        Case "AssetsImpairment"
            Set GetCleaner = New CleanerCleanAssetsImpairment
        Case Else
            MsgBox "未知的清理類型: " & cleaningType
            Set GetCleaner = Nothing
    End Select
End Function

' 修改 ProcessFile，增加 xlApp 參數
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

    Set cleaner = GetCleaner(cleaningType)
    If cleaner Is Nothing Then
        MsgBox "找不到對應的清理程序: " & cleaningType
        Exit Sub
    End If

    ' 執行 CleanReport，統一傳入單一 xlApp 物件
    cleaner.CleanReport filePath, cleaningType, xlApp

    ' 執行 additionalClean 處理
    Set additionalCleaner = New CleanerAddColumns
    additionalCleaner.additionalClean filePath, cleaningType, ReportDataDate, ReportMonth, ReportMonthString, xlApp

    Set importer = GetImporter(cleaningType)
    If importer Is Nothing Then
        MsgBox "找不到對應的匯入程序: " & cleaningType
        Exit Sub
    End If

    accessDBPath = "D:\DavidHsu\ReportCreator\DB_MonthlyReport.accdb"
    importer.ImportData filePath, accessDBPath, cleaningType
End Sub

' 主程序，建立單一 Excel 執行個體並統一管理屬性
Public Sub ProcessAllReports()
    Dim ReportDataDate As Date, ReportMonth As Date
    Dim ReportMonthString As String
    Dim configDict As Object, filePathDict As Object
    Dim reportType As Variant, filePath As String
    Dim customPathArray As Variant, customDate As Date
    Dim xlApp As Excel.Application

    customPathArray = Array()
    Set configDict = GetConfigsReturnDict(customPathArray, customDate)
    If configDict.Count = 0 Then
        MsgBox "Error: 無法取得設定資料", vbCritical
        Exit Sub
    End If

    ReportDataDate = configDict("ReportDataDate")
    ReportMonth = configDict("ReportMonth")
    ReportMonthString = configDict("ReportMonthString")
    Set filePathDict = configDict("FilePaths")

    ' 建立單一的 Excel 執行個體
    Set xlApp = New Excel.Application
    Call ConfigureExcelApp(xlApp)

    For Each reportType In filePathDict.Keys
        filePath = filePathDict(reportType)
        ProcessFile filePath, reportType, ReportDataDate, ReportMonth, ReportMonthString, xlApp
    Next reportType

    ' 處理結束前還原屬性，並清除剪貼簿狀態
    Call RestoreExcelAppSettings(xlApp)
    Call ClearCutCopyMode(xlApp)

    xlApp.Quit
    Set xlApp = Nothing
End Sub
```

---

### 小結

1. **Excel 屬性管理**：透過獨立模組 modExcelSettings 將 xlApp 的 Visible、ScreenUpdating、DisplayAlerts、AskToUpdateLinks 與 CutCopyMode 統一管理，方便在主程式中一次設定/還原。  
2. **傳入單一 Excel 執行個體**：各個 Cleaner 模組的 CleanReport 及 additionalClean 方法改為傳入 xlApp，移除了重複建立與關閉 Excel 的動作。  
3. **主程式統一處理**：ProcessAllReports 建立單一 Excel.Application，呼叫管理屬性的副程式，處理完所有報表後再還原屬性、清除剪貼簿狀態與關閉 Excel。

這樣整體的管理會更有系統且容易維護，同時也能避免因多重 Excel 實例而產生的問題。



Answer 4.
延續上面問題，我看到你在
ProcessAllReports
最開始和結束分別使用
   Call ConfigureExcelApp(xlApp)
和
    Call RestoreExcelAppSettings(xlApp)
和
   Call ClearCutCopyMode(xlApp)
但是因為for each迴圈中，每一個就是處理一個excel檔案，
CutCopyMode這個部分，如果我在for迴圈裡面不少程序都會進行copy的話，在最後使用會不會有問題？
會建議在for迴圈裡面操作還是for迴圈外面操作？
哪個方式比較不容易有問題，
請給我所有需要修改的內容


以下提供兩種建議的做法，並說明哪一種較不易產生問題：

---

### 建議與說明

**問題來源：**  
因為在你的作業中，多數程序都使用了 copy/paste 操作，且 Excel 的 CutCopyMode 狀態有時會遺留（例如使用 PasteSpecial 後可能還留在「剪貼模式」），若等到 For Each 迴圈外面統一做 ClearCutCopyMode，有可能上一個檔案還沒完全「清理」好後，就直接開啟下一個檔案，這樣可能導致後續的複製貼上行為出現不正常狀況。  

**建議：**  
比較穩妥的作法是**在每次處理完單一檔案的所有操作後**，都立即清除該檔案的剪貼狀態，也就是在 ProcessFile 裡面處理完該檔案前就呼叫 ClearCutCopyMode(xlApp)。這樣可以避免上一個檔案操作遺留的剪貼模式狀態影響到後續檔案的處理。  

換句話說，建議在 For Each 迴圈中，每處理完一個檔案後立即清除剪貼模式，而非等到所有檔案全部處理完才一次清除。

---

### 所需修改內容

以下我將以先前的完整版本為基礎，修改 ProcessFile 程序，使其在結束時呼叫 ClearCutCopyMode，再示範主程式 ProcessAllReports 的部分也可能不需要在外部再次呼叫，但可以保留還原 Excel 屬性與關閉前的統一呼叫。

#### 1. 修改 modExcelSettings 模組（無需變動，僅供參考）

```vba
' modExcelSettings 模組
Option Compare Database

Public Sub ConfigureExcelApp(ByRef xlApp As Excel.Application)
    With xlApp
        .Visible = False             ' 不顯示 Excel 畫面
        .ScreenUpdating = False      ' 關閉畫面更新，提高效能
        .DisplayAlerts = False       ' 關閉警告訊息，自動接受預設
        .AskToUpdateLinks = False    ' 關閉連結更新詢問
    End With
End Sub

Public Sub RestoreExcelAppSettings(ByRef xlApp As Excel.Application)
    With xlApp
        .ScreenUpdating = True
        .DisplayAlerts = True
        .AskToUpdateLinks = True
    End With
End Sub

Public Sub ClearCutCopyMode(ByRef xlApp As Excel.Application)
    xlApp.CutCopyMode = False
End Sub
```

#### 2. 修改 ProcessFile 程序

將 ProcessFile 程序修改為在所有清理與匯入操作完成後，立即呼叫 ClearCutCopyMode(xlApp)（操作這個檔案完畢時呼叫）。

```vba
' 主程式模組中 ProcessFile 部分修改後內容：
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

    Set cleaner = GetCleaner(cleaningType)
    If cleaner Is Nothing Then
        MsgBox "找不到對應的清理程序: " & cleaningType
        Exit Sub
    End If

    ' 呼叫 CleanReport 時傳入單一 xlApp
    cleaner.CleanReport filePath, cleaningType, xlApp

    ' 呼叫 additionalClean 處理
    Set additionalCleaner = New CleanerAddColumns
    additionalCleaner.additionalClean filePath, cleaningType, ReportDataDate, ReportMonth, ReportMonthString, xlApp

    Set importer = GetImporter(cleaningType)
    If importer Is Nothing Then
        MsgBox "找不到對應的匯入程序: " & cleaningType
        Exit Sub
    End If

    accessDBPath = "D:\DavidHsu\ReportCreator\DB_MonthlyReport.accdb"
    importer.ImportData filePath, accessDBPath, cleaningType

    ' 處理完該檔案後，清除任何剪貼模式
    Call ClearCutCopyMode(xlApp)
End Sub
```

#### 3. 修改 ProcessAllReports 程序

在 ProcessAllReports 中，先建立單一 Excel 執行個體並設定屬性，依據每個檔案處理時已在 ProcessFile 中呼叫 ClearCutCopyMode，這裡只需要在所有檔案迴圈完成後還原其屬性與關閉 Excel。

```vba
Public Sub ProcessAllReports()
    Dim ReportDataDate As Date, ReportMonth As Date
    Dim ReportMonthString As String
    Dim configDict As Object, filePathDict As Object
    Dim reportType As Variant, filePath As String
    Dim customPathArray As Variant, customDate As Date
    Dim xlApp As Excel.Application

    customPathArray = Array()
    Set configDict = GetConfigsReturnDict(customPathArray, customDate)
    If configDict.Count = 0 Then
        MsgBox "Error: 無法取得設定資料", vbCritical
        Exit Sub
    End If

    ReportDataDate = configDict("ReportDataDate")
    ReportMonth = configDict("ReportMonth")
    ReportMonthString = configDict("ReportMonthString")
    Set filePathDict = configDict("FilePaths")

    ' 建立單一的 Excel 執行個體
    Set xlApp = New Excel.Application
    Call ConfigureExcelApp(xlApp)

    For Each reportType In filePathDict.Keys
        filePath = filePathDict(reportType)
        ProcessFile filePath, reportType, ReportDataDate, ReportMonth, ReportMonthString, xlApp
    Next reportType

    ' 所有檔案處理完畢後，還原 Excel 屬性
    Call RestoreExcelAppSettings(xlApp)
    ' 最後再呼叫 ClearCutCopyMode，作為保險備援
    Call ClearCutCopyMode(xlApp)
    
    xlApp.Quit
    Set xlApp = Nothing
End Sub
```

---

### 總結

- **在 For Each 迴圈中，每個檔案處理完後**，在 ProcessFile 裡面呼叫 ClearCutCopyMode(xlApp)，這樣能確保每個檔案的 copy 狀態清除，不會影響到下一個檔案處理。  
- 主程式最終還原屬性與呼叫 ClearCutCopyMode 作保險，但主要是在每個檔案後進行清理，較不易出現剪貼模式未清的情況。  

這樣修改後，你的 Excel 操作設定會更集中管理，也可以降低因剪貼狀態遺留而導致的問題。



Answer 5.

延續上面的架構，我給你我的 IImporter_ImportData 實作如上，
但是實際匯入microsoft access之後，在資料表中最下面會有很多空白的Data匯入access資料表中，因此我想要在下面實作中增加檢查是否在符合Table上傳格式的Excel表格中，將最下面沒有資料的row刪除，我要增加在下面的CleanerAddColumns裡面

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


回答
以下提供一個完整修改後的 CleanerAddColumns 模組範例，重點在於在 additionalClean 程式中，針對每一個工作表在完成其他處理後，再進行「刪除最下面空白列」的動作。這裡採用傳入單一的 xlApp 物件，並以一個關鍵欄（例如已填入 dataDate 的欄位，或是最後處理後的表頭所在欄）來判斷最後有資料的列，然後刪除底下的空白列。你可以根據實際上傳格式調整判斷依據。以下是完整範例：

---

```vba
' CleanerAddColumns 模組 (CleanerAddColumns.bas)
Option Compare Database

Implements ICleaner

' 此模組不需要初始化參數，如有需求可自行加入
Public Sub ICleaner_Initialize(Optional ByVal sheetName As Variant = 1, _
                               Optional ByVal loopColumn As Integer = 1, _
                               Optional ByVal leftToDelete As Integer = 2, _
                               Optional ByVal rightToDelete As Integer = 3, _
                               Optional ByVal rowsToDelete As Variant, _
                               Optional ByVal colsToDelete As Variant)
    ' 初始化動作(如果需要)
End Sub

' 這裡的 CleanReport 不做處理，因為此模組主要針對 additionalClean
Public Sub ICleaner_CleanReport(ByVal fullFilePath As String, _
                                ByVal cleaningType As String, _
                                ByVal xlApp As Excel.Application)
    ' implement operations here, 如無需求可留空
End Sub

' 修改後的 additionalClean，使用傳入的 xlApp 物件
Public Sub ICleaner_additionalClean(ByVal fullFilePath As String, _
                                    ByVal cleaningType As String, _
                                    ByVal dataDate As Date, _
                                    ByVal dataMonth As Date, _
                                    ByVal dataMonthString As String, _
                                    ByVal xlApp As Excel.Application)
    Dim xlbk As Excel.Workbook
    Dim xlsht As Excel.Worksheet
    Dim i As Long, lastRow As Long
    Dim tableColumns As Variant

    ' 檢查檔案是否存在
    If Dir(fullFilePath) = "" Then
        MsgBox "File does not exist in path: " & fullFilePath
        Exit Sub
    End If

    ' 開啟檔案 (使用傳入的 xlApp，不建立新的 Excel 實例)
    Set xlbk = xlApp.Workbooks.Open(fullFilePath)
    
    ' 取得 Table 的標題欄位（你必須自行實作 GetTableColumns 函數）
    tableColumns = GetTableColumns(cleaningType)

    ' 逐一處理工作表
    For Each xlsht In xlbk.Sheets
        ' --- 插入處理動作 ---
        ' 在現有資料之前插入四個欄位，用來儲存額外資料
        xlsht.Columns("A:D").Insert Shift:=xlToRight
        
        ' 以第一個欄位（原本的資料）來判斷資料列數
        lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
        
        ' 逐列填入額外資料及幣別設定
        For i = 2 To lastRow
            xlsht.Cells(i, "B").Value = dataDate
            xlsht.Cells(i, "C").Value = dataMonth
            xlsht.Cells(i, "D").Value = dataMonthString
            If cleaningType = "DBU_AC5602" Or cleaningType = "OBU_AC5602" Then
                xlsht.Cells(i, "J").Value = "TWD"
            ElseIf cleaningType = "OBU_AC4603" Then
                xlsht.Cells(i, "J").Value = "USD"
            ElseIf cleaningType = "OBU_AC5411B" Then
                xlsht.Cells(i, "I").Value = "USD"
            End If
        Next i
        
        ' 設定標題列，假設標題列位於第1列，並依照 Table 格式填入標題
        xlsht.Range("A1").Resize(1, UBound(tableColumns) - LBound(tableColumns) + 1).Value = tableColumns
        
        ' 刪除剛剛多餘的第一欄（調整欄位結構，使標題正確）
        xlsht.Columns("A").Delete

        ' === 以下為新增的動作：刪除工作表中最下面空白列 ===
        ' 以關鍵欄判斷：假設第一資料欄 (現在的 A 欄) 保證有資料，且應該在每一列都有值
        Dim lastDataRow As Long, currentLastRow As Long
        lastDataRow = xlsht.Cells(xlsht.Rows.Count, "A").End(xlUp).Row
        ' 如果 UsedRange 的列數大於 lastDataRow，表示底部可能有多餘的空白列
        currentLastRow = xlsht.UsedRange.Rows.Count
        If lastDataRow < currentLastRow Then
            xlsht.Rows(lastDataRow + 1 & ":" & currentLastRow).Delete
        End If
        ' === 刪除動作結束 ===

    Next xlsht

    ' 儲存並關閉檔案
    xlbk.Save
    xlbk.Close False
    Set xlsht = Nothing
    Set xlbk = Nothing

    MsgBox "完成Addition Clean，路徑為: " & fullFilePath
    Debug.Print "完成Addition Clean，路徑為: " & fullFilePath
End Sub
```

---

### 說明重點

1. **統一使用傳入的 Excel 執行個體**  
   – 不再於本模組中建立 xlApp，而是直接用 ProcessAllReports（或 ProcessFile）傳入的同一個 Excel 實例。  
2. **操作流程**  
   – 開啟檔案後，針對每一個工作表：  
  a. 插入額外欄位、填入額外資料（dataDate、dataMonth、dataMonthString）及依據 cleaningType 設定幣別  
  b. 設定標題列（tableColumns）  
  c. 刪除多餘的欄位以調整表格結構  
3. **刪除底部空白列**  
   – 使用工作表中關鍵欄（本範例取 A 欄）的最後一個資料列作為最後有效列  
   – 與 xlsht.UsedRange.Rows.Count 作比較，若 UsedRange 定義的列數大於最後有資料的列數，則刪除該區間中的空白列。  
4. **關閉及釋放**  
   – 儲存並關閉 Workbook 後，清除物件參考

---

依照此版本進行修改後，於每個檔案處理完成前，都能在該檔案中刪除底部空白列，確保上傳至 Access 的資料表不會多出空白資料行。
