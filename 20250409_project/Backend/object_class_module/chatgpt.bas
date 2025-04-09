=== Question ===

Option Compare Database

Public Function GetImporter(ByVal cleaningType As String) As IImporter
    ' 目前以標準匯入方式處理所有報表
    ' 若未來有特殊匯入需求，可根據 cleaningType 擴充其他 Importer 類別
    Set GetImporter = New ImporterStandard
End Function

Public Function GetCleaner(ByVal cleaningType As String) As ICleaner
    Select Case cleaningType
        'CleanerRowsDelete
        Case "OBU_MM4901B"
            Set GetCleaner = New CleanerRowsDelete
            GetCleaner.Initialize 1, 1, 2, 3, Array("USD", "AUD", "CAD", "CHF", "CNY", "EUR", "GBP", "HKD", "JPY", "NZD", "SGD", "THB", "ZAR", "TWD", "總計:", "主管")
        
        'CleanerUnitACCurr
        Case "OBU_AC4603"
            Set GetCleaner = New CleanerUnitCurr
            GetCleaner.Initialize 1, 3, 2, 4, Array("合計", "總資產:"), Array("K", "H", "F", "D", "C", "A")

        'CleanerRowsColsDelete
        Case "DBU_DL6850"
            Set GetCleaner = New CleanerRowsColsDelete
            GetCleaner.Initialize 1, 1, 2, 2, Array("小計", "總計", "主管"), Array("L")
        
        'CleanerAC_PluralCurr
        Case "OBU_AC4620B"
            Set GetCleaner = New CleanerPluralCurr

        'CleanerIsEmpty
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


以上是我的主程序，另外我有建立下面的物件模組 ICleaner interface
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

以及implementA和Ｂ(這邊只舉兩個例子)


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


    ' 找出最後一列
    '**注意這邊是使用A欄位最後儲存格當作Row Number, 為了DL9360修改
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    '**原來
    ' lastRow = xlsht.Cells(xlsht.Rows.Count, clsLoopColumn).End(xlUp).Row

    ' 反向遍歷以刪除符合條件的列
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

    'Set Table Columns
    ' tableColumns = GetTableColumns(cleaningType)
    ' xlsht.Range("A1").Resize(1, UBound(tableColumns) - LBound(tableColumns) + 1).Value = tableColumns

    ' 儲存並關閉檔案

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

    Dim i As Integer, j As Integer
    Dim lastRow As Integer, startRow As Integer, endRow As Integer, newLastRow As Integer

    Dim assetRows As Collection

    Dim currArray() As String
    Dim tableColumns As Variant


    Redim currArray(0)
    ' Dim fullFilePath As String
    ' fullFilePath = "D:\DavidHsu\testFile\vba\test\OBU_AC5601_33_AC_E_20241231_r.xls"
    If Dir(fullFilePath) = "" Then
        MsgBox "File does not exist in path: " & fullFilePath
    End If

    If Dir(Replace(fullFilePath, "xls", "txt")) = "" Then
        MsgBox "File does not exist in path: " & Replace(fullFilePath, "xls", "txt")
    End If

    Set xlApp = Excel.Application
    xlApp.ScreenUpdating = False
    xlApp.DisplayAlerts = False

    Set xlbk = xlApp.Workbooks.Open(fullFilePath)
    ' Set xlbk = Workbooks.Open(fullFilePath)
    Set xlsht = xlbk.Worksheets(1)
    
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row

    xlsht.Columns("B").Insert shift:=xlToRight

    For i = lastRow To 2 Step -1
        If IsEmpty(xlsht.Cells(i, 1).Value) Or _
           IsNumeric(xlsht.Cells(i, 1).Value) Or _
           xlsht.Cells(i, 1).Value = "放款類" Or _
           xlsht.Cells(i, 1).Value = "存款類" Or _
           xlsht.Cells(i, 1).Value = "負債類" Or _
           xlsht.Cells(i, 1).Value = "損益類 - 收入" Or _
           xlsht.Cells(i, 1).Value = "損益類 - 費用" Or _
           xlsht.Cells(i, 1).Value = "業主權益類" Or _
           Left(xlsht.Cells(i, 1).Value, 2) = "或有" Or _
           Left(xlsht.Cells(i, 1).Value, 2) = "主管" Then
            xlsht.Rows(i).Delete
        End If
        xlsht.Cells(i, "B") = xlsht.Cells(i, "C").Value & xlsht.Cells(i, "E").Value
    Next i

    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row

    xlsht.Columns("K").Delete
    xlsht.Columns("G").Delete
    xlsht.Columns("C:E").Delete

    currArray = ReadCurrencyFromTxt(Replace(fullFilePath, "xls", "txt"))
    'tableColumns = GetTableColumns(cleaningType)
    
    Set assetRows = New Collection

    ' 找出所有 "資產類" 出現的位置
    For i = 1 To lastRow
        If xlsht.Cells(i, 1).Value Like "*資產類*" Then
            assetRows.Add i
        End If
    Next i

    ' 依照奇數次到偶數次作為區間分割
    For i = 1 To assetRows.Count Step 2
        If i + 1 < assetRows.Count Then
            startRow = assetRows(i) + 1
            endRow = assetRows(i + 2) - 1
        Else
            startRow = assetRows(i) + 1
            endRow = lastRow
        End If
        
        ' 建立新的工作表
        xlbk.Sheets.Add After:=xlbk.Sheets(xlbk.Sheets.count)
        With ActiveSheet
            .Name = currArray(i \ 2)
            ' 複製資料 A:AA 欄
            xlsht.Range(xlsht.Cells(startRow, 1), xlsht.Cells(endRow, 27)).Copy
            .Range("A2").PasteSpecial Paste:=xlPasteValues
 
            newLastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            For j = newLastRow To 2 Step -1
                If .Cells(j, 1).Value = "資產類" Then .Rows(j).Delete
                .Cells(j, "F").Value = currArray(i \ 2)
            Next j
            .Columns("A").Delete
            'Set Table Columns
            ' .Range("A1").Resize(1, UBound(tableColumns) - LBound(tableColumns) + 1).Value = tableColumns
        End With
    Next i
    
    xlsht.Delete
    
    ' 清除剪貼簿
    xlApp.CutCopyMode = False
    xlApp.ScreenUpdating = True
    xlApp.DisplayAlerts = True

    Set assetRows = Nothing

    xlbk.Save
    xlbk.Close False
    Set xlsht = Nothing
    Set xlbk = Nothing
    xlApp.Quit
    Set xlApp = Nothing

    ' On Error Resume Next
    ' Set xlsht = Nothing
    ' Set xlbk = Nothing
    ' Set assetRows = Nothing
    ' On Error GoTo 0
    
    MsgBox "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
    Debug.Print "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
End Sub

'---------------------------------------------------------------
'function call and execute work for dictionary

Public Function ReadCurrencyFromTxt(ByVal filePath As String) As String()
    Dim fileNum As Integer
    Dim lineContent As String
    Dim arrCurr() As String
    Dim curr As String
    Dim i As Integer
    Dim dict As Object ' 使用 Dictionary 來存放不重複的貨幣

    ' 建立 Dictionary 來去除重複值
    Set dict = CreateObject("Scripting.Dictionary")

    ' 取得可用的檔案編號
    fileNum = FreeFile()

    ' 開啟檔案進行讀取
    Open filePath For Input As #fileNum
    
    ' 逐行讀取檔案內容
    Do Until EOF(fileNum)
        Line Input #fileNum, lineContent ' 讀取一行內容
    
        ' 如果該行包含 "幣? ??別"，則提取後面的值
        If Left(lineContent, 6) = "幣    別" Then
    
            ' 取出 "幣? ??別" 後的貨幣代碼
            curr = Trim(Mid(lineContent, 12, 3))
            
            ' 確保 curr 不是空值，且不重複加入
            If curr <> "" And Not dict.Exists(curr) Then
                dict.Add curr, Nothing
            End If
        End If
    Loop
    
    ' 關閉檔案
    Close #fileNum
    
    ' 將 Dictionary 轉換為陣列
    If dict.count > 0 Then
        ReDim arrCurr(dict.count - 1)
        i = 0
        Dim key As Variant
        For Each key In dict.Keys
            arrCurr(i) = key
            i = i + 1
        Next key
    End If  
    ' 回傳結果
    ReadCurrencyFromTxt = arrCurr  
End Function
    
Public Sub ICleaner_additionalClean(ByVal fullFilePath As String, _
                                    ByVal cleaningType As String, _
                                    ByVal dataDate As Date, _
                                    ByVal dataMonth As Date, _
                                    ByVal dataMonthString As String)
    'implement operations here
End Sub


請問我實際在執行主程序時，為什麼在cleaner階段中開啟excel進行清理，但是都沒辦法完整關掉由microsoft access操作Excel的使用，因為持續打開excel執行相關內容，導致開啟過多excel，變成我不能夠一次執行太多個cleaner，請問為什麼會這樣，要怎麼改善







' ===================


===Ans 1===

這個問題通常是因為 Excel 的 COM 物件在使用後並沒有被正確釋放，導致 Excel 程序持續在背景中執行而無法完全關閉。以下提供詳細的說明以及改進建議：

---

### 為什麼 Excel 會持續開啟？

1. **COM 物件釋放不完整**  
   當你透過 Microsoft Access 自動化 Excel 時，每一個 Excel 的物件（例如 Application、Workbook、Worksheet 等）都必須在使用完畢後明確地釋放。儘管程式碼中有呼叫 `xlApp.Quit` 以及設置物件為 `Nothing`，但只要其中的任何子物件（如 xlsht、xlbk，甚至是隱含建立的 Range 物件）仍然存在參考，整個 Excel 實例都不會正確關閉。  
   
2. **隱藏的物件引用**  
   在 VBA 中，有時候隱式建立的物件（例如用過程中呼叫的 `Cells`、`Range`、或 `ActiveSheet` 等）可能沒有明確指向變數而產生隱藏引用，這樣也會阻止 Excel 的完全關閉。

3. **錯誤處理不足**  
   如果執行過程中發生錯誤，導致沒有進入到釋放物件的程式區段，Excel 實例可能會持續存在。缺乏適當的錯誤處理機制也會加劇這個問題。

---

### 如何改善？

要確保 Excel 使用後能夠完整關閉，可以考慮以下幾個改善建議：

1. **建立獨立的 Excel 實例**  
   使用 `CreateObject("Excel.Application")` 而非直接用 `Excel.Application` 來取得目前的應用程式物件，這樣可以確保創建一個全新的 Excel 實例，較不容易受到其他潛在執行個體的影響。例如：  
   ```vba
   Dim xlApp As Excel.Application
   Set xlApp = CreateObject("Excel.Application")
   ```

2. **明確釋放所有 COM 物件**  
   結束時，務必按照從子物件到父物件的順序釋放引用，並在每個物件不再使用時，設置為 `Nothing`。例如：  
   ```vba
   xlbk.Close False
   Set xlsht = Nothing
   Set xlbk = Nothing
   xlApp.Quit
   Set xlApp = Nothing
   ```
   如果你有使用到其他中介物件（例如 Range、Collection 等），也需確實釋放。

3. **加入錯誤處理**  
   為了避免錯誤造成物件無法釋放，可以加入 `On Error` 機制，確保在發生錯誤時仍會正確釋放資源。例如：
   ```vba
   On Error GoTo CleanUp
   ' [程式核心區段]

   CleanUp:
       If Not xlbk Is Nothing Then xlbk.Close False
       If Not xlApp Is Nothing Then xlApp.Quit
       Set xlsht = Nothing
       Set xlbk = Nothing
       Set xlApp = Nothing
   ```

4. **避免不必要的隱式物件參考**  
   確保所有對 Excel 物件的引用都是以變數方式存取，避免直接使用 `ActiveSheet` 或 `Cells` 而沒有事先宣告。例如，取用範圍可以使用：
   ```vba
   Dim rng As Excel.Range
   Set rng = xlsht.Cells(i, 1)
   ' 使用完 rng 後，再設置 rng = Nothing
   Set rng = Nothing
   ```

5. **檢查是否有全局參考遺漏**  
   如果 CleanReport 方法中有使用到其他全域或隱式創建的物件，請檢查這些物件是否也在結束時被釋放。使用 `Debug.Print` 或 Windows 任務管理員檢查 Excel 進程，可以幫助確認改善成果。

---

### 總結

問題的根本在於在清理階段中由 Access 操控 Excel 的自動化操作，沒有完整釋放所有 COM 物件，導致 Excel 實例在背景中累積。改進方案包括：  
- 使用 `CreateObject` 來創建新的 Excel 實例。  
- 明確、依序釋放所有物件，避免任何殘留的引用。  
- 在程式中加上錯誤處理，確保即使發生例外，也能正確釋放資源。  
- 避免隱式的物件引用，所有操作都使用事先宣告好的變數。

按照這些步驟修改後，你應該可以解決 Excel 持續開啟的問題，並且可以一次執行多個 cleaner 而不會因為 Excel 殘留進程而出現資源耗盡的狀況。

===Ans 2===

下面提供兩個主要方向修改建議，並附上調整範例，供你參考與套用到你的程式中：

---

## 1. 使用 CreateObject 建立全新的 Excel 實例

在你的程式中，原本使用的是：

```vba
Set xlApp = Excel.Application
```

這行會嘗試連結現有的 Excel 物件（若有），可能導致無法完全控制或釋放。建議修改為使用 CreateObject，以確保建立全新執行個體，類似如下：

```vba
Set xlApp = CreateObject("Excel.Application")
```

這樣在結束時使用 `xlApp.Quit` 後，就能順利關閉這個獨立的實例。

---

## 2. 加入全面性的錯誤處理與明確釋放所有物件

每個涉及 Excel 物件的函數（例如 `ICleaner_CleanReport` 或其他清理模組）都應該在程式結尾處加入錯誤處理區段，確保即使發生錯誤也能釋放所有物件。以下以 `ICleaner_CleanReport` 的範例說明：

### 調整前（部分摘錄）：
```vba
Set xlApp = Excel.Application
Set xlbk = xlApp.Workbooks.Open(fullFilePath)
Set xlsht = xlbk.Sheets(clsSheetName)
' ...
' 程式操作後
xlbk.Save
xlbk.Close False
Set xlsht = Nothing
Set xlbk = Nothing
xlApp.Quit
Set xlApp = Nothing
```

### 修改後的範例：
```vba
Dim xlApp As Excel.Application
Dim xlbk As Excel.Workbook
Dim xlsht As Excel.Worksheet

On Error GoTo CleanUp  ' 加入錯誤處理

' 建立新的 Excel 實例
Set xlApp = CreateObject("Excel.Application")
xlApp.ScreenUpdating = False
xlApp.DisplayAlerts = False

' 開啟工作簿與取得工作表
Set xlbk = xlApp.Workbooks.Open(fullFilePath)
Set xlsht = xlbk.Sheets(clsSheetName)

' ==============================
' 這裡放置你的清理程式碼邏輯
' ==============================

' 儲存並關閉檔案
xlbk.Save

CleanUp:
    On Error Resume Next
    If Not xlsht Is Nothing Then Set xlsht = Nothing
    If Not xlbk Is Nothing Then xlbk.Close False: Set xlbk = Nothing
    If Not xlApp Is Nothing Then
        xlApp.Quit
        Set xlApp = Nothing
    End If

If Err.Number = 0 Then
    MsgBox "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
Else
    MsgBox "清理時發生錯誤：" & Err.Description
End If
```

### 要點說明
- **錯誤處理區**：在程式開始使用 `On Error GoTo CleanUp`，這樣若中間發生任何錯誤，就會直接跳到 CleanUp 區段，保證釋放物件。
- **依序釋放子物件到父物件**：先將工作表 (xlsht) 設為 Nothing，再關閉並釋放工作簿 (xlbk)，最後執行 `xlApp.Quit` 並將 `xlApp` 設為 Nothing。
- **使用 On Error Resume Next**：在 CleanUp 區段內部，確保釋放動作不因其他錯誤中斷。

---

## 3. 檢查其他隱式引用

請確認所有對 Excel 物件的存取都用變數儲存，避免直接使用像 `ActiveSheet`、`Cells` 等隱式方法而導致 COM 物件沒有明確指派。若必須使用，可在作業結束後也將變數設置為 Nothing。

---

## 總結

為解決 Excel 未能完全關閉的問題，你需要：
- 透過 `CreateObject("Excel.Application")` 建立全新 Excel 實例。
- 每個涉及 Excel 操作的程式段落，加上錯誤處理與明確釋放各層物件（WorkSheet、Workbook、Application）。
- 檢查並避免隱式物件引用。

請依照上面調整範例在你的各個 `ICleaner` 模組中進行修改，這樣可以確保在完成清理後 Excel 都能正確關閉，避免背景中累積大量 Excel 進程。

===Ans 3===

下列提供一個修改範例，能夠在錯誤發生時自動重試該 cleaner 作業，直到成功為止（或者達到一個你預設的重試次數上限）。這個解法主要透過下述重試機制來確保每個 cleaner 都有執行到：

---

### 重試機制概念

1. **重試迴圈**  
   使用一個重試迴圈（例如 Do While 或 For … Next）來執行清理作業。當作業執行發生錯誤時，先將錯誤記錄或顯示，然後進入 CleanUp 區段釋放所有 Excel 相關物件，最後再重設狀態後重試。

2. **錯誤計數上限**  
   設定一個最大重試次數（例如 5 次），以避免在某些例外案例下進入無限迴圈。若達到上限，則可選擇顯示錯誤訊息給使用者。

3. **完全釋放物件後重試**  
   每次重試前，必須將前一次操作所建立的 Excel 物件（Application、Workbook、Worksheet 等）全部釋放，並重新建立新的 Excel 實例，以確保狀態良好。

---

### 範例修改

下面是一個將重試機制納入 `ICleaner_CleanReport` 方法的範例修改。請根據你各模組的結構，將這段程式碼調整適配到你的專案中。

```vba
Public Sub ICleaner_CleanReport(ByVal fullFilePath As String, _
                                ByVal cleaningType As String)
    Dim xlApp As Excel.Application
    Dim xlbk As Excel.Workbook
    Dim xlsht As Excel.Worksheet
    Dim retryCount As Integer
    Dim maxRetries As Integer
    Dim success As Boolean

    maxRetries = 5  '設定最大重試次數
    retryCount = 0
    success = False

RetryLoop:
    On Error GoTo CleanUp
    ' 建立新的 Excel 實例
    Set xlApp = CreateObject("Excel.Application")
    xlApp.ScreenUpdating = False
    xlApp.DisplayAlerts = False

    ' 開啟工作簿與取得工作表
    Set xlbk = xlApp.Workbooks.Open(fullFilePath)
    Set xlsht = xlbk.Sheets(clsSheetName)  ' 假設已定義 clsSheetName

    ' ==============================
    ' 這裡開始執行你的清理邏輯
    ' 例如，刪除不需要的列、欄位
    Dim i As Long, lastRow As Long, col As Variant
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row

    For i = lastRow To 2 Step -1
        ' 此處為範例判斷，可以根據實際需求調整條件
        If IsEmpty(xlsht.Cells(i, 1).Value) Or _
           Trim(xlsht.Cells(i, 1).Value) = "" Or _
           Left(xlsht.Cells(i, 1).Value, 2) = "主管" Then
            xlsht.Rows(i).Delete
        End If
    Next i

    ' ==============================
    ' 儲存並關閉檔案
    xlbk.Save
    success = True  ' 如果能走到這邊，則視為成功
CleanUp:
    On Error Resume Next   ' 釋放資源時忽略錯誤
    If Not xlsht Is Nothing Then Set xlsht = Nothing
    If Not xlbk Is Nothing Then
        xlbk.Close False
        Set xlbk = Nothing
    End If
    If Not xlApp Is Nothing Then
        xlApp.Quit
        Set xlApp = Nothing
    End If

    ' 若發生錯誤，則 Err.Number 非 0，重試直到達到上限
    If Err.Number <> 0 Then
        retryCount = retryCount + 1
        If retryCount < maxRetries Then
            Debug.Print "清理 " & cleaningType & " 發生錯誤 (" & Err.Description & ")，進行第 " & retryCount & " 次重試..."
            Err.Clear  ' 清除錯誤
            GoTo RetryLoop
        Else
            MsgBox "清理 " & cleaningType & " 時發生錯誤超過 " & maxRetries & " 次。" & vbCrLf & "錯誤訊息: " & Err.Description, vbCritical
            Exit Sub
        End If
    End If

    If success Then
        MsgBox "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
        Debug.Print "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
    End If
End Sub
```

---

### 關鍵說明

1. **重試迴圈 (RetryLoop 標籤)**  
   程式進入重試迴圈前設置 `retryCount`，當發生錯誤（也就是 Err.Number 不為 0）時，就會跳轉回 `RetryLoop`。這樣在錯誤發生時，程式會重新建立 Excel 實例，並再次執行清理作業。

2. **清理區段 (CleanUp 標籤)**  
   在每次作業結束後，即使遇到錯誤，都會執行 CleanUp 區段。此區段負責釋放所有 Excel COM 物件，然後檢查錯誤狀態。如果重試次數尚未達上限，程式會清除錯誤訊息並回到 RetryLoop。

3. **最大重試次數**  
   設定了 `maxRetries` 來避免無限重試。你可以根據實際需求調整這個數值。

4. **成功標示**  
   當操作成功時將 `success` 設為 True，並在退出 CleanUp 後顯示成功訊息。

---

這個範例展示了一種在發生錯誤後重試的方式，確保每個 cleaner 只要遇到暫時性問題就能自動重試，直到成功執行。如果在某個 cleaner 中遇到永久性錯誤，也不會造成其他 cleaner 被略過，而只會中斷該次重試並通知使用者。你可以根據這個範例，將相似的重試機制應用到所有需使用 Excel 自動化的 cleaner 模組中。

===Ans 4===
下面提供兩個模組的完整修正範例，我在主要的 ICleaner_CleanReport 內加入了以下調整：

- **使用 CreateObject 建立全新 Excel 實例**：改用  
  ```vba
  Set xlApp = CreateObject("Excel.Application")
  ```  
  取代原本的 Set xlApp = Excel.Application  
- **加入錯誤處理與重試機制**：當發生錯誤時，程式會進入 CleanUp 區段，釋放所有 Excel 物件，並利用重試迴圈（最高重試次數設為 5 次）重新執行，直到操作成功或超出重試次數上限。

以下分別是兩個模組的完整修正後內容，並在程式碼後說明修正重點與位置。

---

## 範例 1 完整修正內容

```vba
Option Compare Database
Implements ICleaner

Public Sub ICleaner_Initialize(Optional ByVal sheetName As Variant = 1, _
                               Optional ByVal loopColumn As Integer = 1, _
                               Optional ByVal leftToDelete As Integer = 2, _
                               Optional ByVal rightToDelete As Integer = 3, _
                               Optional ByVal rowsToDelete As Variant, _
                               Optional ByVal colsToDelete As Variant)
    ' 這裡不需要進行物件初始化
End Sub

Public Sub ICleaner_CleanReport(ByVal fullFilePath As String, _
                                ByVal cleaningType As String)
    Dim xlApp As Excel.Application
    Dim xlbk As Excel.Workbook
    Dim xlsht As Excel.Worksheet
    
    Dim lastRow As Long
    Dim i As Integer
    Dim colArray As Variant
    Dim valueType As Variant
    Dim eachType As Variant
    
    Dim securityIndex As Collection
    Dim securityName As Collection
    
    Dim startRow As Integer
    Dim endRow As Integer
    Dim innerLastRow As Integer
    Dim sheetNameStr As String
    
    Dim toDelete As Boolean
    
    ' 加入重試機制用參數
    Dim retryCount As Integer, maxRetries As Integer
    Dim success As Boolean
    maxRetries = 5
    retryCount = 0
    success = False

RetryLoop:
    On Error GoTo CleanUp
    ' 檢查檔案是否存在
    If Dir(fullFilePath) = "" Then
        MsgBox "File does not exist in path: " & fullFilePath
        Exit Sub
    End If
    
    ' 建立全新 Excel 實例
    Set xlApp = CreateObject("Excel.Application")
    xlApp.ScreenUpdating = False
    xlApp.DisplayAlerts = False
    xlApp.AskToUpdateLinks = False
    
    Set xlbk = xlApp.Workbooks.Open(fullFilePath)
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
    
    ' 開始建立分頁和處理資料
    Set securityName = New Collection
    Set securityIndex = New Collection
    ' 假設 GetTableColumns 為其他已定義函數
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
            sheetNameStr = Replace(securityName(i), "?", "")
        Else
            sheetNameStr = securityName(i)
        End If

        xlbk.Sheets.Add After:=xlbk.Sheets(xlbk.Sheets.count)

        With ActiveSheet
            .Name = sheetNameStr
            xlsht.Range(xlsht.Cells(startRow, "C"), xlsht.Cells(endRow, "M")).Copy
            .Range("A2").PasteSpecial Paste:=xlPasteValues
            innerLastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            .Range("L2:L" & innerLastRow).Value = securityName(i)
            'Set Table Columns，如有需要可開啟此行
            ' .Range("A1").Resize(1, UBound(tableColumns) - LBound(tableColumns) + 1).Value = tableColumns
        End With
ContinueLoop:
    Next i
    
    For i = xlbk.Sheets.Count To 1 Step -1
        'Default cancel
        toDelete = True
        'Check whether in valueType or not
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

    success = True  ' 成功執行
CleanUp:
    On Error Resume Next
    ' 釋放所有物件
    If Not xlsht Is Nothing Then Set xlsht = Nothing
    If Not xlbk Is Nothing Then
        xlbk.Close False
        Set xlbk = Nothing
    End If
    If Not xlApp Is Nothing Then
        xlApp.Quit
        Set xlApp = Nothing
    End If

    If Err.Number <> 0 Then
        retryCount = retryCount + 1
        If retryCount < maxRetries Then
            Debug.Print "清理 " & cleaningType & " 發生錯誤 (" & Err.Description & ")，第 " & retryCount & " 次重試..."
            Err.Clear
            GoTo RetryLoop
        Else
            MsgBox "清理 " & cleaningType & " 時發生錯誤超過 " & maxRetries & " 次。" & vbCrLf & "錯誤訊息: " & Err.Description, vbCritical
            Exit Sub
        End If
    End If

    If success Then
        MsgBox "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
        Debug.Print "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
    End If
End Sub

Public Sub ICleaner_additionalClean(ByVal fullFilePath As String, _
                                    ByVal cleaningType As String, _
                                    ByVal dataDate As Date, _
                                    ByVal dataMonth As Date, _
                                    ByVal dataMonthString As String)
    ' implement operations here
End Sub
```

### 修正重點（範例 1）
1. **Excel 實例建立**：  
   原本使用 `Set xlApp = Excel.Application` 改成 `Set xlApp = CreateObject("Excel.Application")`（第 30 行左右）。

2. **錯誤處理與重試機制**：  
   在程式開始前定義重試變數，並以 `RetryLoop` 標籤包住主要作業。在 `CleanUp` 區段內依錯誤次數進行重試，若重試次數小於 5 則重新執行（CleanUp 區段從第 139 行開始）。

3. **物件釋放順序**：  
   在 CleanUp 區段內，依序釋放 xlsht、xlbk、xlApp，避免隱藏引用導致 Excel 無法完全關閉。

---

## 範例 2 完整修正內容

```vba
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

    Dim retryCount As Integer, maxRetries As Integer
    Dim success As Boolean
    
    maxRetries = 5
    retryCount = 0
    success = False

RetryLoop:
    On Error GoTo CleanUp
    ' 檢查檔案是否存在
    If Dir(fullFilePath) = "" Then
        MsgBox "File does not exist in path: " & fullFilePath
        Exit Sub
    End If

    ' 使用 CreateObject 建立全新 Excel 實例
    Set xlApp = CreateObject("Excel.Application")
    Set xlbk = xlApp.Workbooks.Open(fullFilePath)
    Set xlsht = xlbk.Sheets(clsSheetName)
    
    ' 若沒有指定刪除列與欄的條件，先初始化為空陣列
    If IsEmpty(clsRowsToDelete) Then clsRowsToDelete = Array()
    If IsEmpty(clsColsToDelete) Then clsColsToDelete = Array()

    ' 找出最後一列（使用 A 欄）
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row

    ' 反向遍歷以刪除符合條件的列
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

    ' 針對主管列進行刪除
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

    ' 以下若有設定 Table Columns 可依需要呼叫
    ' tableColumns = GetTableColumns(cleaningType)
    ' xlsht.Range("A1").Resize(1, UBound(tableColumns) - LBound(tableColumns) + 1).Value = tableColumns

    ' 儲存並關閉檔案
    xlbk.Save
    xlbk.Close False

    Set xlsht = Nothing
    Set xlbk = Nothing
    xlApp.Quit
    Set xlApp = Nothing

    success = True
CleanUp:
    On Error Resume Next
    If Not xlsht Is Nothing Then Set xlsht = Nothing
    If Not xlbk Is Nothing Then
        xlbk.Close False
        Set xlbk = Nothing
    End If
    If Not xlApp Is Nothing Then
        xlApp.Quit
        Set xlApp = Nothing
    End If
    
    If Err.Number <> 0 Then
        retryCount = retryCount + 1
        If retryCount < maxRetries Then
            Debug.Print "清理 " & cleaningType & " 發生錯誤 (" & Err.Description & ")，第 " & retryCount & " 次重試..."
            Err.Clear
            GoTo RetryLoop
        Else
            MsgBox "清理 " & cleaningType & " 時發生錯誤超過 " & maxRetries & " 次。" & vbCrLf & "錯誤訊息: " & Err.Description, vbCritical
            Exit Sub
        End If
    End If

    If success Then
        MsgBox "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
        Debug.Print "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
    End If
End Sub

Public Sub ICleaner_additionalClean(ByVal fullFilePath As String, _
                                    ByVal cleaningType As String, _
                                    ByVal dataDate As Date, _
                                    ByVal dataMonth As Date, _
                                    ByVal dataMonthString As String)
    ' implement operations here
End Sub
```

### 修正重點（範例 2）
1. **建立 Excel 實例**：  
   原本的 `Set xlApp = Excel.Application` 修改為  
   ```vba
   Set xlApp = CreateObject("Excel.Application")
   ```  
   （位於 RetryLoop 區段內）。

2. **錯誤處理與重試機制**：  
   同範例 1，在主要區段前定義重試變數，並透過 RetryLoop 與 CleanUp 區段進行錯誤處理及重試（重試上限同為 5 次）。

3. **物件釋放順序**：  
   在 CleanUp 區段中依序釋放 xlsht、xlbk、xlApp 以確保 Excel 完全退出。

---

這兩個修正版本即在原有程式的基礎上，修改了以下重點：  
- 將 Excel 實例由直接取得改成使用 CreateObject。  
- 加入錯誤處理與重試機制，並在重試前確保所有 COM 物件被正確釋放。  

你可以將以上完整程式碼直接複製到各自的模組中使用。