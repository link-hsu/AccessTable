你可以利用一個 `Scripting.Dictionary` 來把每一個 fvArray 的字串對應到你要的代碼，然後在輸出陣列的最後一欄填入這個對照值。以下示範只貼出你要新增／修改的區段，請把它插入到你原本程式的適當位置。

---

1. **在宣告區**（`Dim fvArray …` 之後）新增一個 Dictionary 並把對照表填入：  
```vb
    ' 在 fvArray 宣告之後，新增這段
    Dim mapDict As Object
    Set mapDict = CreateObject("Scripting.Dictionary")
    With mapDict
        .Add "FVPL-公司債(公營)",   "FVPL_CompanyBond_Foreign"
        .Add "FVPL-公司債(民營)",   "FVPL_CompanyBond_Foreign"
        .Add "FVPL-公債",         "FVPL_GovBond_Foreign"
        .Add "FVPL-金融債",       "FVPL_FinancialBond_Foreign"
        .Add "FVOCI-公債",        "FVOCI_GovBond_Foreign"
        .Add "FVOCI-公司債(公營)",  "FVOCI_CompanyBond_Foreign"
        .Add "FVOCI-公司債(民營)",  "FVOCI_CompanyBond_Foreign"
        .Add "FVOCI-金融債",      "FVOCI_FinancialBond_Foreign"
        .Add "AC-公債",           "AC_GovBond_Foreign"
        .Add "AC-公司債(公營)",    "AC_CompanyBond_Foreign"
        .Add "AC-公司債(民營)",    "AC_CompanyBond_Foreign"
        .Add "AC-金融債",         "AC_FinancialBond_Foreign"
    End With
```

2. **擴充輸出陣列的欄數**（在你第一次 `ReDim outputArr(1 To lastRow, 1 To 31)` 的地方，把 31 改為 32）：  
```vb
        If i = 1 Then
            ReDim outputArr(1 To lastRow, 1 To 32)    ' ← 改成 32
        End If
```

3. **在每一筆資料處理完畢時，將對照值寫入第 32 欄**，也就是在你原本這段：
```vb
    outputArr(j, 31) = category
```
之後，緊接著加上一行：
```vb
    ' 將 category 轉為對照代碼，放在第 32 欄
    If mapDict.Exists(category) Then
        outputArr(j, 32) = mapDict(category)
    Else
        outputArr(j, 32) = ""    ' 或者填入預設值
    End If
```

4. **在表頭加入新的欄名**（你原本設定欄名的地方）：  
```vb
    ' 假設原來是這樣寫
    For i = 1 To (UBound(columnsArray) - LBound(columnsArray) + 1)
        ActiveSheet.Range("A1").Resize(1, (UBound(columnsArray) - LBound(columnsArray) + 1)).Value = columnsArray
    Next i

    ' 在這之後，加入新的欄名
    ActiveSheet.Cells(1, 32).Value = "評價類別對照"
```

---

### 完整示範（只列出新增／修改處）

```vb
    ' … 省略前面宣告 …
    fvArray = Array( _
        "FVPL-公司債(公營)", _
        … _
        "AC-金融債")

    ' ===== 新增：建立對照 Dictionary =====
    Dim mapDict As Object
    Set mapDict = CreateObject("Scripting.Dictionary")
    With mapDict
        .Add "FVPL-公司債(公營)",   "FVPL_CompanyBond_Foreign"
        .Add "FVPL-公司債(民營)",   "FVPL_CompanyBond_Foreign"
        .Add "FVPL-公債",         "FVPL_GovBond_Foreign"
        .Add "FVPL-金融債",       "FVPL_FinancialBond_Foreign"
        .Add "FVOCI-公債",        "FVOCI_GovBond_Foreign"
        .Add "FVOCI-公司債(公營)",  "FVOCI_CompanyBond_Foreign"
        .Add "FVOCI-公司債(民營)",  "FVOCI_CompanyBond_Foreign"
        .Add "FVOCI-金融債",      "FVOCI_FinancialBond_Foreign"
        .Add "AC-公債",           "AC_GovBond_Foreign"
        .Add "AC-公司債(公營)",    "AC_CompanyBond_Foreign"
        .Add "AC-公司債(民營)",    "AC_CompanyBond_Foreign"
        .Add "AC-金融債",         "AC_FinancialBond_Foreign"
    End With

    ' … 省略計算 lastRow 以及 securityRows 收集 …

    For i = 1 To securityRows.count
        ' … 省略區段判斷 …

        If i = 1 Then
            ReDim outputArr(1 To lastRow, 1 To 32)    ' ← 原本 31 改成 32
        End If

        category = xlsht.Cells(startRow - 1, 1).Value

        For j = startRow To endRow Step 2
            ' … 省略原本 1–31 欄的填值 …
            outputArr(j, 31) = category

            ' ===== 新增：將 category 轉為對照代碼，放第 32 欄 =====
            If mapDict.Exists(category) Then
                outputArr(j, 32) = mapDict(category)
            Else
                outputArr(j, 32) = ""
            End If

        Next j
ContinueLoop:
    Next i

    ' … 省略建表與刪表頭空白列 …

    ' 原本放欄名
    For i = 1 To (UBound(columnsArray) - LBound(columnsArray) + 1)
        ActiveSheet.Range("A1").Resize(1, (UBound(columnsArray) - LBound(columnsArray) + 1)).Value = columnsArray
    Next i
    ' ===== 新增：第 32 欄欄名 =====
    ActiveSheet.Cells(1, 32).Value = "評價類別對照"
```

這樣就會在原本 31 欄的「評價類別」之後，多出一欄「評價類別對照」，並依你提供的對照表填入相對應的英文代碼。





----------------------------------------------------------------------------------------------

以下是整支加入「評價類別對照」欄位的完整 VBA 巨集，已將 fvArray 的中／英對照透過 Dictionary 寫入第 32 欄，並在表頭新增該欄位名稱：

```vb
Public Sub ICleaner_CleanReport(ByVal fullFilePath As String, _
                                ByVal cleaningType As String)
    Dim xlApp As Excel.Application
    Dim xlbk As Excel.Workbook
    Dim ws As Excel.Worksheet
    Dim xlsht As Excel.Worksheet

    Dim copyRg As Range
    Dim Rngs As Range
    Dim oneRng As Range

    Dim outputArr() As Variant
    Dim fvArray As Variant

    Dim i As Integer, j As Integer, k As Integer
    Dim lastRow As Integer

    Dim securityRows As Collection
    Dim category As Variant
    Dim tableColumns As Variant

    ' 對照 Dictionary
    Dim mapDict As Object

    ' fullFilePath 範例:
    ' fullFilePath = "D:\...\20250131外幣債損益評估表(月底)對AC5100B.xlsx"

    Set xlApp = Excel.Application
    xlApp.ScreenUpdating = False
    xlApp.DisplayAlerts = False
    xlApp.AskToUpdateLinks = False

    Set xlbk = xlApp.Workbooks.Open(fullFilePath)
    Set ws = xlbk.Worksheets("評估表")
    ws.Copy After:=xlbk.Sheets(xlbk.Sheets.count)
    ActiveSheet.Name = "評估表cp"
    Set xlsht = ActiveSheet

    ' 解除公式，僅保留值
    Set copyRg = xlsht.UsedRange
    copyRg.Value = copyRg.Value

    ' 讀取欄位標題 A5:T5
    Set Rngs = copyRg.Range("A5:T5")

    ' 解析多列標題
    Dim columnsArray() As Variant
    Dim tempSplit As Variant
    Dim tempSave() As Variant
    Dim count As Long
    Dim splitCount As Long

    count = 0: splitCount = 0
    For Each oneRng In Rngs
        tempSplit = Split(oneRng, vbLf)
        ReDim Preserve columnsArray(count)
        columnsArray(count) = Trim(tempSplit(0))
        count = count + 1
        If UBound(tempSplit) >= 1 Then
            ReDim Preserve tempSave(splitCount)
            tempSave(splitCount) = Trim(tempSplit(1))
            splitCount = splitCount + 1
        End If
    Next oneRng

    ReDim Preserve tempSave(splitCount)
    tempSave(splitCount) = "評價資產類別"

    For i = LBound(tempSave) To UBound(tempSave)
        ReDim Preserve columnsArray(count)
        columnsArray(count) = tempSave(i)
        count = count + 1
    Next i

    ' fvArray 原始類別
    fvArray = Array( _
        "FVPL-公司債(公營)", _
        "FVPL-公司債(民營)", _
        "FVPL-公債", _
        "FVPL-金融債", _
        "FVOCI-公債", _
        "FVOCI-公司債(公營)", _
        "FVOCI-公司債(民營)", _
        "FVOCI-金融債", _
        "AC-公債", _
        "AC-公司債(公營)", _
        "AC-公司債(民營)", _
        "AC-金融債")

    ' 建立中英文對照 Dictionary
    Set mapDict = CreateObject("Scripting.Dictionary")
    With mapDict
        .Add "FVPL-公司債(公營)",   "FVPL_CompanyBond_Foreign"
        .Add "FVPL-公司債(民營)",   "FVPL_CompanyBond_Foreign"
        .Add "FVPL-公債",         "FVPL_GovBond_Foreign"
        .Add "FVPL-金融債",       "FVPL_FinancialBond_Foreign"
        .Add "FVOCI-公債",        "FVOCI_GovBond_Foreign"
        .Add "FVOCI-公司債(公營)",  "FVOCI_CompanyBond_Foreign"
        .Add "FVOCI-公司債(民營)",  "FVOCI_CompanyBond_Foreign"
        .Add "FVOCI-金融債",      "FVOCI_FinancialBond_Foreign"
        .Add "AC-公債",           "AC_GovBond_Foreign"
        .Add "AC-公司債(公營)",    "AC_CompanyBond_Foreign"
        .Add "AC-公司債(民營)",    "AC_CompanyBond_Foreign"
        .Add "AC-金融債",         "AC_FinancialBond_Foreign"
    End With

    ' 刪除空白與標題列
    lastRow = xlsht.Cells(xlsht.Rows.count, 1).End(xlUp).Row
    For i = lastRow To 1 Step -1
        If IsEmpty(xlsht.Cells(i, 1).Value) Or _
           xlsht.Cells(i, 1).Value = "Security_Id" Then
            xlsht.Rows(i).Delete
        ElseIf Left(Trim(xlsht.Cells(i, 1).Value), 2) = "標註" Then
            xlsht.Rows(i & ":" & lastRow).Delete
        End If
    Next i

    lastRow = xlsht.Cells(xlsht.Rows.count, 1).End(xlUp).Row

    ' 搜尋所有 category 標記列
    Set securityRows = New Collection
    For i = 1 To lastRow
        For Each category In fvArray
            If xlsht.Cells(i, 1).Value = category Then
                securityRows.Add i
            End If
        Next category
    Next i

    ' 處理每段資料，並輸出到 outputArr
    Dim startRow As Integer, endRow As Integer
    For i = 1 To securityRows.count
        If i + 1 <= securityRows.count Then
            If securityRows(i) + 1 = securityRows(i + 1) Then GoTo ContinueLoop
            startRow = securityRows(i) + 1
            endRow = securityRows(i + 1) - 1
        Else
            startRow = securityRows(i) + 1
            endRow = lastRow
        End If

        If i = 1 Then
            ReDim outputArr(1 To lastRow, 1 To 32)    ' ← 改為 32 欄
        End If

        category = xlsht.Cells(startRow - 1, 1).Value

        For j = startRow To endRow Step 2
            ' 原本 1–20 欄與各種轉置
            For k = 1 To 40
                If k >= 1 And k <= 20 Then
                    outputArr(j, k) = xlsht.Cells(j, k).Value
                    If category Like "AC-*" Then
                        outputArr(j, 20) = xlsht.Cells(j, 17).Value
                    End If
                    outputArr(j, 17) = ""
                ElseIf k = 22 Then
                    outputArr(j, 21) = xlsht.Cells(j + 1, k - 20).Value
                ElseIf k = 28 Then
                    outputArr(j, 22) = xlsht.Cells(j + 1, k - 20).Value
                ElseIf k = 29 Then
                    outputArr(j, 23) = xlsht.Cells(j + 1, k - 20).Value
                ElseIf k = 30 Then
                    outputArr(j, 24) = xlsht.Cells(j + 1, k - 20).Value
                ElseIf k = 31 Then
                    outputArr(j, 25) = xlsht.Cells(j + 1, k - 20).Value
                ElseIf k = 32 Then
                    outputArr(j, 26) = xlsht.Cells(j + 1, k - 20).Value
                ElseIf k = 33 Then
                    outputArr(j, 27) = xlsht.Cells(j + 1, k - 20).Value
                ElseIf k = 34 Then
                    outputArr(j, 28) = xlsht.Cells(j + 1, k - 20).Value
                ElseIf k = 35 Then
                    outputArr(j, 29) = xlsht.Cells(j + 1, k - 20).Value
                ElseIf k = 36 Then
                    outputArr(j, 30) = xlsht.Cells(j + 1, k - 20).Value
                ElseIf k = 37 Then
                    outputArr(j, 31) = category
                End If
            Next k

            ' 新增第 32 欄：中英對照
            If mapDict.Exists(category) Then
                outputArr(j, 32) = mapDict(category)
            Else
                outputArr(j, 32) = ""
            End If
        Next j
ContinueLoop:
    Next i

    ' 新增 OutputData 工作表，寫入 outputArr
    xlbk.Sheets.Add After:=xlbk.Sheets(xlbk.Sheets.count)
    ActiveSheet.Name = "OutputData"

    Dim numRows As Integer, numCols As Integer
    numRows = UBound(outputArr, 1)
    numCols = UBound(outputArr, 2)
    ActiveSheet.Range(ActiveSheet.Cells(2, 1), ActiveSheet.Cells(numRows + 1, numCols)).Value = outputArr

    ' 刪空白列
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.count, 1).End(xlUp).Row
    For i = lastRow To 2 Step -1
        If ActiveSheet.Cells(i, 1).Value = "" Then ActiveSheet.Rows(i).Delete
    Next i

    ' 寫入欄名
    For i = 1 To (UBound(columnsArray) - LBound(columnsArray) + 1)
        ActiveSheet.Range("A1").Resize(1, UBound(columnsArray) + 1).Value = columnsArray
    Next i
    ' 第 32 欄欄名
    ActiveSheet.Cells(1, 32).Value = "評價類別對照"

    ' 刪除其他工作表
    For Each ws In xlbk.Sheets
        If Not ws Is ActiveSheet Then ws.Delete
    Next ws

    xlApp.ScreenUpdating = True
    xlApp.DisplayAlerts = True
    xlApp.AskToUpdateLinks = True

    Set securityRows = Nothing
    xlbk.Save
    xlbk.Close False
    Set xlsht = Nothing
    Set xlbk = Nothing
    xlApp.Quit
    Set xlApp = Nothing

    MsgBox "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
    Debug.Print "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
End Sub
```



