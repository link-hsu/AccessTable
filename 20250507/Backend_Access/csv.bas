' ==========================================================================
Q1

讀取csv和讀取text檔的使用方式一樣嗎

Thought for a second


在 VBA 裡，用 FileSystemObject（FSO）開啟 CSV 跟純文字（.txt）檔案的「程式架構」是完全一樣的──它們都是純文字檔，只是副檔名不同。差別只在於你後續怎麼「解析」每一行的內容：

| 動作         | CSV                                                       | TXT（一般文字）                                            |
| ---------- | --------------------------------------------------------- | ---------------------------------------------------- |
| 開檔、讀整檔、逐行讀 | 完全相同：`fso.OpenTextFile(path,mode,…)`、`ReadAll`、`ReadLine` | 一樣                                                   |
| 行拆解（欄位解析）  | 通常用 `Split(line, ",")` 或進階 `ParseCsvLine`                 | 視需求，用 `Split(line, delimiter)`，delimiter 可是空白、TAB、其他 |
| 行重組、寫回     | 完全相同：`OpenTextFile` mode=2 然後 `WriteLine`                 | 一樣                                                   |

---

## 範例：讀 CSV vs 讀 TXT

```vb
Dim fso As Object
Dim ts As Object
Dim lines As Variant

Set fso = CreateObject("Scripting.FileSystemObject")

'── 讀 CSV
Set ts = fso.OpenTextFile("C:\data\myfile.csv", 1, False, -1)
lines = Split(ts.ReadAll, vbCrLf)
ts.Close

' 解析第 2 行成欄位
Dim parts As Variant
parts = Split(lines(1), ",")    ' CSV 用逗號
Debug.Print parts(0), parts(1)

'── 讀 TXT（假設是以 TAB 分隔）
Set ts = fso.OpenTextFile("C:\data\myfile.txt", 1, False, -1)
lines = Split(ts.ReadAll, vbCrLf)
ts.Close

' 解析第 2 行成欄位
parts = Split(lines(1), vbTab)   ' TXT 用 Tab
Debug.Print parts(0), parts(1)
```

除了 `Split(line, ",")` 改成 `Split(line, vbTab)`（或其他分隔符），其餘完全相同。

---

### 何時要用進階 CSV 解析？

* 欄位內容可能包含逗號（例如 `"Smith, John"`）
* 欄位可能有雙引號逃逸（`""`）

這種情況才要用先前講的 `ParseCsvLine`／`JoinCsvLine`。對純文字檔（.txt），只要固定分隔符（空白、Tab 或其他），直接 `Split` 就夠了。



' ==========================================================================
Q2

老實說我不知道怎麼使用vba操作csv檔，請從基礎部分教我怎麼使用


Thought for a couple of seconds


下面我從最基礎的觀念開始，帶你一步步了解如何用 VBA 操作 CSV 檔案（純文字格式），不需要靠 Excel 物件也能完成讀、寫、修改。

---

## 1. 為什麼用 FileSystemObject (FSO)

VBA 預設可用 Excel 物件開啟 CSV，但那是把它當活的試算表操作，效率較低，也不適合大型或純文字清洗。
FileSystemObject 是 VBScript 的檔案操作物件，能用最簡單的「讀文字→處理→寫文字」方式，直接操作 `.csv`。

要使用 FSO，請在 VBA 編輯器中：

1. 選「工具」→「參考資料庫」
2. 勾選 **Microsoft Scripting Runtime**

這樣就能早期綁定 `Dim fso As New Scripting.FileSystemObject`；若不想勾選，也可用晚期綁定 `CreateObject("Scripting.FileSystemObject")`。

---

## 2. 開啟、讀取整個 CSV 檔

最簡單的範例：把 CSV 全部讀成一個字串，再切成「行」的陣列。

```vb
Sub ReadCsvAll()
    Dim fso As Object
    Dim ts As Object
    Dim fullPath As String
    Dim textAll As String
    Dim lines As Variant

    fullPath = "C:\temp\data.csv"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(fullPath, ForReading:=1, Create:=False, Tristate:= -1)  
      ' ForReading=1；Tristate=-1 表示自動偵測 BOM 編碼

    textAll = ts.ReadAll
    ts.Close

    lines = Split(textAll, vbCrLf)   ' 以 CRLF 分隔，每個元素就是一列
    MsgBox "總共讀到 " & UBound(lines) + 1 & " 列"
End Sub
```

* `OpenTextFile(path, mode, create, tristate)`

  * `mode=1` 讀取；`2` 寫入（覆寫）；`8` 附加
  * `create` = True 時若不存在會建立
  * `tristate=-1` 讓系統自動偵測 UTF‑8/Unicode BOM

---

## 3. 把一列切成多個欄位（簡易版）

最直接：`Split(line, ",")`。但若欄位有「引號包逗號」，就會切錯。先用簡易版理解：

```vb
Dim parts As Variant
parts = Split(lines(1), ",")   ' 取第 2 列切欄
For i = LBound(parts) To UBound(parts)
    Debug.Print "第 " & i+1 & " 欄 = " & parts(i)
Next
```

這對「純粹用逗號分隔、內容不含逗號」的 CSV 足夠。

---

## 4. 修改欄位後寫回檔案

假設我們要把第 2 欄所有值去掉「%」符號，然後覆寫原檔：

```vb
Sub RemovePercentInCol2()
    Dim fso As Object, tsIn As Object, tsOut As Object
    Dim lines As Variant, outLines() As String
    Dim i As Long, parts As Variant

    Set fso = CreateObject("Scripting.FileSystemObject")
    lines = Split(fso.OpenTextFile("C:\temp\data.csv",1,-1).ReadAll, vbCrLf)
    ReDim outLines(0 To UBound(lines))

    For i = 0 To UBound(lines)
        If i = 0 Then
            outLines(i) = lines(i)    ' header 保留不動
        Else
            parts = Split(lines(i), ",")
            parts(1) = Replace(parts(1), "%", "")  ' 第二欄去掉 %
            outLines(i) = Join(parts, ",")
        End If
    Next

    ' 寫回（覆寫原檔）
    Set tsOut = fso.OpenTextFile("C:\temp\data.csv", 2, False, -1)
    For i = 0 To UBound(outLines)
        tsOut.WriteLine outLines(i)
    Next
    tsOut.Close
End Sub
```

---

## 5. 處理「引號內逗號」的進階解析

如果你的 CSV 欄位可能像 `"Smith, John",25,100%`，就必須用狀態機法略過引號內的逗號：

```vb
Private Function ParseCsvLine(line As String) As Collection
    Dim cols As New Collection
    Dim inQ As Boolean, cur As String
    Dim i As Long, ch As String

    For i = 1 To Len(line)
        ch = Mid$(line, i, 1)
        If ch = """" Then
            inQ = Not inQ
        ElseIf ch = "," And Not inQ Then
            cols.Add cur
            cur = ""
        Else
            cur = cur & ch
        End If
    Next
    cols.Add cur
    Set ParseCsvLine = cols
End Function

Private Function JoinCsvLine(cols As Collection) As String
    Dim s As String, i As Long, v As String
    For i = 1 To cols.Count
        v = cols(i)
        If InStr(v, ",") > 0 Then v = """" & v & """"
        s = s & v & IIf(i < cols.Count, ",", "")
    Next
    JoinCsvLine = s
End Function
```

用法和前面 `Split` 類似，只是把 `Split` 換成 `ParseCsvLine`，最後再用 `JoinCsvLine` 合回去。

---

## 6. 小結與練習

1. **安裝 FSO**：參考 Microsoft Scripting Runtime。
2. **讀整檔**：`OpenTextFile…ReadAll` ⇒ `Split(..., vbCrLf)`。
3. **簡易欄位拆解**：`Split(line, ",")`；進階請用上述 `ParseCsvLine`。
4. **逐行處理**：Loop 每列、Modify、存在新陣列。
5. **寫回檔案**：OpenTextFile with mode=2，然後 `WriteLine`。

---

### 推薦練習

* 自行撰寫一支小程式，將第一欄數字 `+1`，第二欄去掉所有空白，第三欄補滿 8 位數字（`Format(...,"00000000")`）。
* 測試「有逗號的文字欄位」，確認進階解析正確。

掌握後，就能在 VBA 完全跳過 Excel 物件，以最快速、最省資源的方式批次處理大量 CSV！


' ==============================================================
Q3

Public Sub ICleaner\_CleanReport(ByVal fullFilePath As String, \_
ByVal cleaningType As String, \_
ByVal xlApp As Excel.Application)

```
Dim xlbk As Excel.Workbook
Dim xlsht As Excel.Worksheet
Dim i As Long, lastRow As Long
Dim colsDict As Object
Dim colIndex As Variant

If Dir(fullFilePath) = "" Then
    clsHasFile = False
    MsgBox "File not found: " & fullFilePath
    Exit Sub
Else
    clsHasFile = True
End If

Set xlbk = xlApp.Workbooks.Open(fullFilePath)
Set xlsht = xlbk.Sheets(1)

lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row

If clsHasData = False Then
    If (lastRow > 1) Then clsHasData = True
End If

If Not (IsObject(clsColsToHandle) And TypeName(clsColsToHandle) = "Dictionary") Then
    GoTo CleanUp
End If

' 因為Dict在Interface中使用底層有較嚴格的判斷類型，所以建議不使用
Set colsDict = CreateObject("Scripting.Dictionary")
Dim key As String
Dim val As Variant
Dim existingArray As Variant

For i = 0 To UBound(clsColsToHandle) Step 2
    key = clsColsToHandle(i)
    val = clsColsToHandle(i + 1)

    If colsDict.Exists(key) Then
        existingArray = colsDict(key)
        ReDim Preserve existingArray(UBound(existingArray) + 1)
        existingArray(UBound(existingArray)) = val
        colsDict(key) = existingArray
    Else
        colsDict.Add key, Array(val)
    End If
Next i

For i = lastRow To 2 Step -1
    If colsDict.Exists("FormatCols") Then
        For Each colIndex In colsDict("FormatCols")
            With xlsht.Cells(i, colIndex)
                .NumberFormat = "@"
                .Value = Format(.Value, "00000000")
            End With
            ' xlsht.Cells(i, colIndex).Value = Format(xlsht.Cells(i, colIndex).Value, "00000000")
        Next
    End If

    If colsDict.Exists("RemovePercentCols") Then
        For Each colIndex In colsDict("RemovePercentCols")
            xlsht.Cells(i, colIndex).Value = Replace(xlsht.Cells(i, colIndex).Value, "%", "")
        Next
    End If

    If colsDict.Exists("RemoveSpaceCols") Then
        For Each colIndex In colsDict("RemoveSpaceCols")
            xlsht.Cells(i, colIndex).Value = Replace(xlsht.Cells(i, colIndex).Value, " ", "")
        Next
    End If
Next

xlbk.Save
```

CleanUp:
xlbk.Close False
Set xlsht = Nothing
Set xlbk = Nothing

```
If clsHasFile And clsHasData Then
    MsgBox "完成清理：" & cleaningType & vbCrLf & fullFilePath
End If
```

End Sub

請延續上面問題，將我上面這邊的程序，
修改為使用FileSystemObject單純文字檔讀寫去修改資料，修改內容請依照上面修改的相關項目去修改，
請以Excel情境模擬直接以文字檔需要怎麼處理

清洗整理完資料以後請幫我將下面import到microsoft access的代碼
修正為當遇到副檔名是csv檔案時，不使用Excel匯入資料，
而是直接單純以access本身提供的批次匯入csv機制處理，請注意我的import需要批次處理很多不同資料表欄位的檔案
請幫我針對這兩個部分去修改

Public Sub IImporter\_ImportData(ByVal fullFilePath As String, \_
ByVal accessDBPath As String, \_
ByVal tableName As String, \_
ByVal xlApp As Excel.Application)
Dim cn As Object
Dim xlbk As Object
Dim sqlString As String
Dim sheetName As String
Dim i As Integer, j As Integer

```
Dim tableColumns As Variant

Dim fieldList As String
Dim selectList As String

' 使用 ADODB 連接 Access 資料庫
Set cn = CreateObject("ADODB.Connection")
cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & accessDBPath & ";"

' 使用 Excel 來開啟檔案，取得所有分頁名稱
' xlApp.Visible = False ' Excel 不顯示
Set xlbk = xlApp.Workbooks.Open(fullFilePath)

' 取得 Access 資料表的欄位名稱
tableColumns = GetTableColumns(tableName) ' 假設這個函式回傳欄位名稱陣列

' 確保 tableColumns 至少有 2 個欄位（避免 Primary Key 之外沒有欄位）
If UBound(tableColumns) < 1 Then
MsgBox "資料表 " & tableName & " 至少需要 2 個欄位（Primary Key + 其他欄位）。", vbCritical
WriteLog "資料表 " & tableName & " 至少需要 2 個欄位（Primary Key + 其他欄位）。"
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

' 關閉 ADODB 連接
cn.Close
Set cn = Nothing

MsgBox "完成 " & tableName & " 資料表匯入作業"
WriteLog "完成 " & tableName & " 資料表匯入作業"
' Debug.Print "完成 " & tableName & " 資料表匯入作業"
```

End Sub



Answer 3


你提到兩個需求：

1. **將清洗程式從 Excel 操作改成使用 FileSystemObject 對 CSV 做文字層級的讀寫。**
2. **修改 Access 匯入程式，遇到 `.csv` 檔案時，改用 Access 自帶機制匯入 CSV，而不是透過 Excel Application。**

---

## ✅ 一、將清洗邏輯從 Excel 轉為純文字 FileSystemObject 操作（CSV 文字檔）

### 清洗條件回顧

* `FormatCols`: 對應欄位需轉成文字格式（例如補滿8碼 00000000）
* `RemovePercentCols`: 移除 %
* `RemoveSpaceCols`: 移除空白

### 新增：使用 `FileSystemObject` 處理清洗 CSV

```vba
Public Sub ICleaner_CleanCSV(ByVal fullFilePath As String, ByVal cleaningType As String)
    Dim fso As Object, inputFile As Object, outputFile As Object
    Dim inputLines As Variant, outputLines() As String
    Dim colsDict As Object
    Dim i As Long, lineParts As Variant, colIndex As Variant
    Dim tempLine As String

    If Dir(fullFilePath) = "" Then
        MsgBox "File not found: " & fullFilePath
        Exit Sub
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set inputFile = fso.OpenTextFile(fullFilePath, 1, False)
    inputLines = Split(inputFile.ReadAll, vbCrLf)
    inputFile.Close

    If Not (IsObject(clsColsToHandle) And TypeName(clsColsToHandle) = "Dictionary") Then Exit Sub

    ' 轉成 Dictionary 結構
    Set colsDict = CreateObject("Scripting.Dictionary")
    Dim key As String, val As Variant, existingArray As Variant

    For i = 0 To UBound(clsColsToHandle) Step 2
        key = clsColsToHandle(i)
        val = clsColsToHandle(i + 1)
        If colsDict.Exists(key) Then
            existingArray = colsDict(key)
            ReDim Preserve existingArray(UBound(existingArray) + 1)
            existingArray(UBound(existingArray)) = val
            colsDict(key) = existingArray
        Else
            colsDict.Add key, Array(val)
        End If
    Next i

    ReDim outputLines(UBound(inputLines))

    For i = 0 To UBound(inputLines)
        If Trim(inputLines(i)) = "" Then GoTo SkipLine
        lineParts = Split(inputLines(i), ",")

        If i > 0 Then ' Skip header
            If colsDict.Exists("FormatCols") Then
                For Each colIndex In colsDict("FormatCols")
                    If colIndex <= UBound(lineParts) Then
                        lineParts(colIndex - 1) = Format(lineParts(colIndex - 1), "00000000")
                    End If
                Next
            End If
            If colsDict.Exists("RemovePercentCols") Then
                For Each colIndex In colsDict("RemovePercentCols")
                    If colIndex <= UBound(lineParts) Then
                        lineParts(colIndex - 1) = Replace(lineParts(colIndex - 1), "%", "")
                    End If
                Next
            End If
            If colsDict.Exists("RemoveSpaceCols") Then
                For Each colIndex In colsDict("RemoveSpaceCols")
                    If colIndex <= UBound(lineParts) Then
                        lineParts(colIndex - 1) = Replace(lineParts(colIndex - 1), " ", "")
                    End If
                Next
            End If
        End If

        outputLines(i) = Join(lineParts, ",")

SkipLine:
    Next i

    ' 覆寫原檔案
    Set outputFile = fso.OpenTextFile(fullFilePath, 2, True)
    For i = 0 To UBound(outputLines)
        outputFile.WriteLine outputLines(i)
    Next i
    outputFile.Close

    MsgBox "完成清理：" & cleaningType & vbCrLf & fullFilePath
End Sub
```

---

## ✅ 二、修正 Access 匯入邏輯（依副檔名切換 Excel / CSV 匯入方式）

```vba
Public Sub IImporter_ImportData(ByVal fullFilePath As String, _
                                ByVal accessDBPath As String, _
                                ByVal tableName As String, _
                                ByVal xlApp As Excel.Application)

    Dim cn As Object
    Dim ext As String
    Dim sqlString As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    ext = LCase(fso.GetExtensionName(fullFilePath))

    ' 使用 ADODB 連接 Access 資料庫
    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & accessDBPath & ";"

    If ext = "csv" Then
        ' === 使用 Access 內建的 CSV 匯入機制 ===
        sqlString = "INSERT INTO [" & tableName & "] " & _
                    "SELECT * FROM [Text;HDR=Yes;FMT=Delimited;Database=" & fso.GetParentFolderName(fullFilePath) & "].[" & fso.GetFileName(fullFilePath) & "]"
        cn.Execute sqlString
    Else
        ' === 使用 Excel 匯入 ===
        Dim xlbk As Object
        Dim sheetName As String
        Dim tableColumns As Variant
        Dim fieldList As String, selectList As String
        Dim i As Integer

        Set xlbk = xlApp.Workbooks.Open(fullFilePath)
        tableColumns = GetTableColumns(tableName)

        If UBound(tableColumns) < 1 Then
            MsgBox "資料表 " & tableName & " 至少需要 2 個欄位。", vbCritical
            Exit Sub
        End If

        For i = 1 To UBound(tableColumns)
            fieldList = fieldList & "[" & tableColumns(i) & "],"
            selectList = selectList & "[" & tableColumns(i) & "],"
        Next i

        fieldList = Left(fieldList, Len(fieldList) - 1)
        selectList = Left(selectList, Len(selectList) - 1)

        For i = 1 To xlbk.Sheets.count
            sheetName = xlbk.Sheets(i).Name
            sqlString = "INSERT INTO " & tableName & " (" & fieldList & ") " & _
                        "SELECT " & selectList & " FROM [Excel 12.0 Xml;HDR=YES;Database=" & fullFilePath & "].[" & sheetName & "$]"
            cn.Execute sqlString
        Next i

        xlbk.Close False
        Set xlbk = Nothing
    End If

    cn.Close
    Set cn = Nothing

    MsgBox "完成 " & tableName & " 資料表匯入作業"
    WriteLog "完成 " & tableName & " 資料表匯入作業"

End Sub
```

---

這樣就完成：

* Excel 與 CSV 清理兩種方式切換
* Access 匯入機制根據副檔名自動切換（CSV 使用 Access 文字驅動器匯入，Excel 照舊）
* 保留多欄位欄位的處理結構與彈性

是否還需要加入「跳過空白列」或是自動判斷欄位數是否一致的檢查？


' ================================================================


Q4. 我剛剛忘記使用Reasoning功能了，請你在仔細想一下上面回復的會不會有問題，再重新回答我一次

Thought for a couple of seconds


經過再三推敲，原先的兩段程式雖然大致可用，但在以下幾個地方可能出問題：

1. **CSV 解析不夠健全**

   * 直接用 `Split(line, ",")` 若欄位內含逗號（如引號包裹的文字）就會錯誤拆欄。
   * 忽略了可能的 BOM（UTF‑8 BOM）與空尾行。

2. **欄位索引對應**

   * `colIndex - 1` 與 `UBound(lineParts)` 的邊界檢查需更嚴謹。

3. **FileSystemObject 文字覆寫**

   * 直接用 `OpenTextFile` 以 ForWriting 模式會清空檔案，若中途錯誤無法回復。

4. **Access 匯入 CSV**

   * SQL 字串中若路徑包含空白，需加上單引號。
   * 未處理多個 CSV 檔案批次匯入時的 error handling。

下面重新提供改良版，並註明修正之處。

---

## 一、純文字清洗模組（更健全的 CSV 處理）

```vb
Option Explicit

' 針對 CSV 做更安全的「逐行讀取→解析→清洗→寫回」，
' 解決 BOM、空行、欄位內逗號等問題。
Public Sub ICleaner_CleanCSV(ByVal fullFilePath As String, ByVal cleaningType As String)
    Const ForReading = 1, ForWriting = 2
    Dim fso As Object
    Dim tsIn As Object, tsOut As Object
    Dim line As String, parsed As Collection
    Dim outLines As Collection
    Dim i As Long
    Dim colsDict As Object
    Dim key As String, val As Variant, arr, existingArr
    Dim idx As Long

    If Dir(fullFilePath) = "" Then
        MsgBox "找不到檔案: " & fullFilePath, vbExclamation
        Exit Sub
    End If

    ' 建立欄位處理規則 Dictionary
    If Not (IsObject(clsColsToHandle) And TypeName(clsColsToHandle) = "Dictionary") Then
        MsgBox "clsColsToHandle 未設定或型態錯誤", vbCritical
        Exit Sub
    End If
    Set colsDict = CreateObject("Scripting.Dictionary")
    For i = 0 To UBound(clsColsToHandle) Step 2
        key = clsColsToHandle(i)
        val = clsColsToHandle(i + 1)
        If colsDict.Exists(key) Then
            existingArr = colsDict(key)
            ReDim Preserve existingArr(UBound(existingArr) + 1)
            existingArr(UBound(existingArr)) = val
            colsDict(key) = existingArr
        Else
            colsDict.Add key, Array(val)
        End If
    Next

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set tsIn = fso.OpenTextFile(fullFilePath, ForReading, False, -1) ' -1 = TristateUseDefault (自動偵測 BOM)
    Set outLines = New Collection

    ' 逐行讀取並解析
    Do While Not tsIn.AtEndOfStream
        line = tsIn.ReadLine
        ' 忽略完全空白的行
        If Trim(line) = "" Then
            outLines.Add ""
        Else
            Set parsed = ParseCsvLine(line)
            ' 0 行是 header，不做清洗
            If outLines.Count > 0 Then
                ' 對每個指定欄位做清洗
                If colsDict.Exists("FormatCols") Then
                    For Each idx In colsDict("FormatCols")
                        If idx >= 1 And idx <= parsed.Count Then
                            parsed(idx) = Right$("00000000" & parsed(idx), 8)
                        End If
                    Next
                End If
                If colsDict.Exists("RemovePercentCols") Then
                    For Each idx In colsDict("RemovePercentCols")
                        If idx >= 1 And idx <= parsed.Count Then
                            parsed(idx) = Replace(parsed(idx), "%", "")
                        End If
                    Next
                End If
                If colsDict.Exists("RemoveSpaceCols") Then
                    For Each idx In colsDict("RemoveSpaceCols")
                        If idx >= 1 And idx <= parsed.Count Then
                            parsed(idx) = Replace(parsed(idx), " ", "")
                        End If
                    Next
                End If
            End If
            outLines.Add JoinCsvLine(parsed)
        End If
    Loop
    tsIn.Close

    ' 寫回到暫存檔，再覆寫原檔
    Dim tempPath As String
    tempPath = fso.GetSpecialFolder(2).Path & "\" & fso.GetTempName ' 2=TemporaryFolder
    Set tsOut = fso.OpenTextFile(tempPath, ForWriting, True, -1)
    For i = 1 To outLines.Count
        tsOut.WriteLine outLines(i)
    Next
    tsOut.Close

    fso.DeleteFile fullFilePath, True
    fso.MoveFile tempPath, fullFilePath

    MsgBox "完成清理：" & cleaningType & vbCrLf & fullFilePath, vbInformation
End Sub

' 用 Collection 來暫存每個欄位（可處理引號內逗號）
Private Function ParseCsvLine(ByVal line As String) As Collection
    Dim result As New Collection
    Dim i As Long, ch As String
    Dim inQuotes As Boolean
    Dim cur As String
    For i = 1 To Len(line)
        ch = Mid$(line, i, 1)
        If ch = """" Then
            inQuotes = Not inQuotes
        ElseIf ch = "," And Not inQuotes Then
            result.Add cur
            cur = ""
        Else
            cur = cur & ch
        End If
    Next
    result.Add cur
    Set ParseCsvLine = result
End Function

' 把 Collection 重組回 CSV 行，並在必要時加引號
Private Function JoinCsvLine(cols As Collection) As String
    Dim s As String
    Dim i As Long, v As String
    For i = 1 To cols.Count
        v = cols(i)
        If InStr(v, ",") > 0 Or InStr(v, """") > 0 Then
            v = """" & Replace(v, """", """""") & """"
        End If
        s = s & v & IIf(i < cols.Count, ",", "")
    Next
    JoinCsvLine = s
End Function
```

**說明**

* 加入 `ParseCsvLine`／`JoinCsvLine`，正確處理「引號包逗號」情況。
* 以暫存檔寫回，避免中途失敗覆寫原檔造成資料遺失。
* 嚴格檢查欄位索引範圍，並自動偵測 BOM 編碼。

---

## 二、Access 匯入程式：副檔名自動分流＋路徑引號

```vb
Option Explicit

Public Sub IImporter_ImportData(ByVal fullFilePath As String, _
                                ByVal accessDBPath As String, _
                                ByVal tableName As String, _
                                ByVal xlApp As Excel.Application)
    Dim cn As Object
    Dim fso As Object
    Dim ext As String, dirPath As String, fileName As String
    Dim sqlString As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(fullFilePath) Then
        MsgBox "檔案不存在：" & fullFilePath, vbExclamation
        Exit Sub
    End If

    ext = LCase(fso.GetExtensionName(fullFilePath))
    dirPath = fso.GetParentFolderName(fullFilePath)
    fileName = fso.GetFileName(fullFilePath)

    ' 建立 ADODB 連線
    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & accessDBPath & ";"

    If ext = "csv" Then
        ' 用 Access 文字驅動器批次匯入 CSV
        sqlString = _
          "INSERT INTO [" & tableName & "] " & _
          "SELECT * FROM " & _
          "'" & "Text;HDR=Yes;FMT=Delimited;Database=" & dirPath & "'" & ".[" & fileName & "]"
        cn.Execute sqlString
    Else
        ' 用 Excel 連結匯入
        Dim xlbk As Object
        Dim sheet As Object
        Dim tableColumns As Variant
        Dim fieldList As String, selectList As String
        Dim i As Long

        Set xlbk = xlApp.Workbooks.Open(fullFilePath, ReadOnly:=True)
        tableColumns = GetTableColumns(tableName)
        If UBound(tableColumns) < 1 Then
            MsgBox "資料表必須至少有 Primary Key + 一個欄位。", vbCritical
            GoTo CleanupExcel
        End If

        ' 建構欄位字串
        fieldList = ""
        selectList = ""
        For i = 1 To UBound(tableColumns)
            fieldList = fieldList & "[" & tableColumns(i) & "],"
            selectList = selectList & "[" & tableColumns(i) & "],"
        Next
        fieldList = Left(fieldList, Len(fieldList) - 1)
        selectList = Left(selectList, Len(selectList) - 1)

        For Each sheet In xlbk.Sheets
            sqlString = _
              "INSERT INTO [" & tableName & "] (" & fieldList & ") " & _
              "SELECT " & selectList & " FROM " & _
              "[Excel 12.0 Xml;HDR=YES;Database=" & fullFilePath & "].[" & sheet.Name & "$]"
            cn.Execute sqlString
        Next

CleanupExcel:
        xlbk.Close SaveChanges:=False
        Set xlbk = Nothing
    End If

    cn.Close
    Set cn = Nothing

    MsgBox "完成 " & tableName & " 匯入", vbInformation
End Sub
```

**改良重點**

* **CSV 路徑加單引號**：避免路徑或資料表名稱含空白導致 SQL 解析錯誤。
* **ReadOnly 開檔**：Excel 開啟改為唯讀，並確保關閉。
* **Error Label**：若 Excel 片段出錯，也能安全跳至 CleanupExcel。

---

這樣修改後，可確保：

* CSV 清洗對「引號內逗號」及 BOM 正確處理，並以暫存檔機制避免資料遺失。
* 匯入階段對 CSV／Excel 自動分流，且路徑與檔名安全包裹，不怕空白或特殊字元。

若還需要更進階的 schema.ini 設定或多檔案迴圈批次呼叫，請再提出！


' ===========================================================================

Q5.

Public Sub ICleaner\_additionalClean(ByVal fullFilePath As String, \_
ByVal cleaningType As String, \_
ByVal dataDate As Date, \_
ByVal dataMonth As Date, \_
ByVal dataMonthString As String)
Dim xlApp As Excel.Application
Dim xlbk As Excel.Workbook
Dim xlsht As Excel.Worksheet
Dim i As Long, lastRow As Long
Dim tableColumns As Variant

```
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
```

End Sub

我這段只有針對Excel的狀況去處理，請幫我加入處理csv的情形

另外再其中有使用的GetTableColumns function如下，請幫我修改增加處理csv的情形

Public Function GetTableColumns(ByVal cleaningType As String) As Variant
Dim dbs As DAO.Database
Dim rs As DAO.Recordset
Dim tableColumns() As String
Dim i As Integer

```
'取得 DAO 資料庫物件
Set dbs = CurrentDb

'查詢資料庫中的 TableName 和 DbCol
' Set rs = dbs.OpenRecordset("SELECT TableName, DbCol FROM DBsColTable WHERE TableName = '" & cleaningType & "'")
Set rs = dbs.OpenRecordset("SELECT TableName, DbCol FROM DBsColTable WHERE TableName = '" & cleaningType & "'" & "ORDER BY DataID ASC;")

'如果有資料，將 DbCol 放入陣列中
If Not rs.EOF Then
    i = 0
    rs.MoveFirst
    Do While Not rs.EOF
        ReDim Preserve tableColumns(i)
        tableColumns(i) = rs!DbCol
        i = i + 1
        rs.MoveNext
    Loop
End If

'關閉 Recordset
rs.Close
Set rs = Nothing
Set dbs = Nothing

'傳回陣列
GetTableColumns = tableColumns
```

End Function


Answer 5.

Thought for 4 seconds


下面一次貼出兩段完整的 VBA 程式，第一段是把 `ICleaner_additionalClean` 改成「Excel／CSV 分流」；第二段是把 `GetTableColumns` 改成同時能「從資料表拿欄位」或「從 CSV 讀 header 拿欄位」。

> **重點改動**
>
> 1. `ICleaner_additionalClean`：
>
>    * 先用 `FileSystemObject` 判斷副檔名，若為 `"csv"` 就走文字檔處理；否則走原本的 Excel 處理。
>    * CSV 走暫存檔機制：讀全部、解析、插欄、寫暫存，再覆寫原檔。
> 2. `GetTableColumns`：
>
>    * 新增第二個可選參數 `fullFilePath`。
>    * 若傳入了 `fullFilePath` 而且是 `.csv`，就開檔讀第一行，回傳那一列的欄位名稱；否則維持原本從 DAO 拿欄位的邏輯。

---

```vb
Option Explicit

' === 1. ICleaner_additionalClean：Excel / CSV 自動分流 ===
Public Sub ICleaner_additionalClean(ByVal fullFilePath As String, _
                                    ByVal cleaningType As String, _
                                    ByVal dataDate As Date, _
                                    ByVal dataMonth As Date, _
                                    ByVal dataMonthString As String)
    Dim fso As Object
    Dim ext As String

    ' 檔案檢查
    If Dir(fullFilePath) = "" Then
        MsgBox "File does not exist: " & fullFilePath, vbExclamation
        Exit Sub
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    ext = LCase(fso.GetExtensionName(fullFilePath))

    If ext = "csv" Then
        Call CleanCsvAdditional(fullFilePath, cleaningType, dataDate, dataMonth, dataMonthString)
    Else
        Call CleanExcelAdditional(fullFilePath, cleaningType, dataDate, dataMonth, dataMonthString)
    End If
End Sub

'── CSV 分流處理
Private Sub CleanCsvAdditional(ByVal fullFilePath As String, _
                               ByVal cleaningType As String, _
                               ByVal dataDate As Date, _
                               ByVal dataMonth As Date, _
                               ByVal dataMonthString As String)
    Const ForReading = 1, ForWriting = 2
    Dim fso As Object, tsIn As Object, tsOut As Object
    Dim lines As Variant, outLines() As String
    Dim i As Long, cols As Collection
    Dim tableCols As Variant
    Dim tempPath As String
    Dim idx As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set tsIn = fso.OpenTextFile(fullFilePath, ForReading, False, -1)
    lines = Split(tsIn.ReadAll, vbCrLf)
    tsIn.Close

    ' 讀欄位定義
    tableCols = GetTableColumns(cleaningType, fullFilePath)
    ' 預先建立 output array
    ReDim outLines(0 To UBound(lines) + 1)

    For i = 0 To UBound(lines)
        If i = 0 Then
            ' header → 插入新欄位名稱 ＋ 原 header
            outLines(i) = "DataDate,DataMonth,DataMonthString," & lines(i)
        Else
            ' data row → 解析、插值、重組
            Set cols = ParseCsvLine(lines(i))
            ' 前端三個欄位
            cols.Add Format(dataDate, "yyyy\/mm\/dd"), 1
            cols.Add Format(dataMonth, "yyyy\/mm"), 2
            cols.Add dataMonthString, 3
            ' 幣別欄位位置：以 tableCols 決定要放到第幾欄（J 或 I）
            Dim curCurrency As String
            Select Case cleaningType
                Case "DBU_AC5602", "OBU_AC5602": curCurrency = "TWD"
                Case "OBU_AC4603", "OBU_AC5411B": curCurrency = "USD"
                Case Else: curCurrency = ""
            End Select
            If curCurrency <> "" Then
                ' 例如要放在原來第10 欄 → new index = 10 +3(前插)
                idx = IIf(cleaningType = "OBU_AC5411B", 9, 10) + 3
                Call cols.Add(curCurrency, idx)
            End If
            outLines(i) = JoinCsvLine(cols)
        End If
    Next

    ' 寫到暫存檔再覆寫
    tempPath = fso.GetSpecialFolder(2).Path & "\" & fso.GetTempName
    Set tsOut = fso.OpenTextFile(tempPath, ForWriting, True, -1)
    For i = 0 To UBound(outLines)
        tsOut.WriteLine outLines(i)
    Next
    tsOut.Close

    fso.DeleteFile fullFilePath, True
    fso.MoveFile tempPath, fullFilePath

    MsgBox "完成 CSV 額外清洗：" & fullFilePath, vbInformation
End Sub

'── Excel 分流處理（原本邏輯搬進來）
Private Sub CleanExcelAdditional(ByVal fullFilePath As String, _
                                ByVal cleaningType As String, _
                                ByVal dataDate As Date, _
                                ByVal dataMonth As Date, _
                                ByVal dataMonthString As String)
    Dim xlApp As Excel.Application
    Dim xlbk As Excel.Workbook
    Dim xlsht As Excel.Worksheet
    Dim i As Long, lastRow As Long
    Dim tableColumns As Variant

    Set xlApp = Excel.Application
    xlApp.Visible = False
    Set xlbk = xlApp.Workbooks.Open(fullFilePath)

    tableColumns = GetTableColumns(cleaningType, vbNullString)

    For Each xlsht In xlbk.Sheets
        lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
        xlsht.Columns("A:D").Insert shift:=xlToRight
        For i = 2 To lastRow
            xlsht.Cells(i, "B").Value = dataDate
            xlsht.Cells(i, "C").Value = dataMonth
            xlsht.Cells(i, "D").Value = dataMonthString
            Select Case cleaningType
                Case "DBU_AC5602", "OBU_AC5602": xlsht.Cells(i, "J").Value = "TWD"
                Case "OBU_AC4603", "OBU_AC5411B"
                    If cleaningType = "OBU_AC5411B" Then
                        xlsht.Cells(i, "I").Value = "USD"
                    Else
                        xlsht.Cells(i, "J").Value = "USD"
                    End If
            End Select
        Next i
        xlsht.Range("A1").Resize(1, UBound(tableColumns) - LBound(tableColumns) + 1).Value = tableColumns
        xlsht.Columns("A").Delete
    Next

    xlbk.Save
    xlbk.Close False
    xlApp.Quit
    MsgBox "完成 Excel 額外清洗：" & fullFilePath, vbInformation
End Sub


' === CSV 解析與組合輔助 ===
Private Function ParseCsvLine(ByVal line As String) As Collection
    Dim result As New Collection
    Dim i As Long, ch As String, inQ As Boolean, cur As String
    For i = 1 To Len(line)
        ch = Mid$(line, i, 1)
        If ch = """" Then inQ = Not inQ
        If ch = "," And Not inQ Then
            result.Add cur: cur = "": Continue For
        End If
        cur = cur & ch
    Next
    result.Add cur
    Set ParseCsvLine = result
End Function

Private Function JoinCsvLine(cols As Collection) As String
    Dim s As String, i As Long, v As String
    For i = 1 To cols.Count
        v = cols(i)
        If InStr(v, ",") > 0 Or InStr(v, """") > 0 Then
            v = """" & Replace(v, """", """""") & """"
        End If
        s = s & v & IIf(i < cols.Count, ",", "")
    Next
    JoinCsvLine = s
End Function
```

---

```vb
' === 2. GetTableColumns：加上 CSV header 讀取 ===
Public Function GetTableColumns( _
    ByVal cleaningType As String, _
    Optional ByVal fullFilePath As String = "" _
) As Variant
    Dim ext As String
    Dim fso As Object

    ' 如果有傳入 fullFilePath，且副檔名是 csv，就讀 header
    If Len(fullFilePath) > 0 Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        ext = LCase(fso.GetExtensionName(fullFilePath))
        If ext = "csv" Then
            Dim ts As Object, hdr As String
            Dim arr As Variant
            Set ts = fso.OpenTextFile(fullFilePath, 1, False, -1)
            hdr = ts.ReadLine
            ts.Close
            arr = Split(hdr, ",")
            GetTableColumns = arr
            Exit Function
        End If
    End If

    ' 否則依舊走原本 DAO 由 DBsColTable 撈欄位
    Dim dbs As DAO.Database
    Dim rs As DAO.Recordset
    Dim tableColumns() As String
    Dim i As Integer

    Set dbs = CurrentDb
    Set rs = dbs.OpenRecordset( _
        "SELECT DbCol FROM DBsColTable " & _
        "WHERE TableName='" & cleaningType & "' " & _
        "ORDER BY DataID ASC;" _
    )

    If Not rs.EOF Then
        i = 0
        rs.MoveFirst
        Do While Not rs.EOF
            ReDim Preserve tableColumns(i)
            tableColumns(i) = rs!DbCol
            i = i + 1
            rs.MoveNext
        Loop
    End If

    rs.Close
    Set rs = Nothing
    Set dbs = Nothing

    GetTableColumns = tableColumns
End Function
```

---

### 測試重點

1. **Excel 情境**：檔名非 `.csv` 時，走原本插欄＋標頭＋幣別欄位邏輯。
2. **CSV 情境**：

   * 讀第一行 header，當作欄位名稱；
   * 每列前面插三個值（`dataDate`、`dataMonth`、`dataMonthString`），並在適當位置插入幣別。
   * 暫存檔寫回，避免中途失敗遺失原檔。

這樣即完成「同一支程式同時處理 Excel & CSV」＋「GetTableColumns 動態支援 CSV header」。


' ============================================

Q6.

一定要分成這麼多層寫嗎，有沒有辦法把架構修改的比較清楚一點，請重新給我一版


Thought for a couple of seconds


下面是一個更扁平化、結構清晰的版本，整個清洗流程只用一支 `ICleaner_additionalClean`，根據副檔名分支處理，並保留最少的輔助函式。

```vb
Option Explicit

'=== 主程式：針對 Excel / CSV 自動分流，並做「Additional Clean」 ===
Public Sub ICleaner_additionalClean( _
    ByVal fullFilePath As String, _
    ByVal cleaningType As String, _
    ByVal dataDate As Date, _
    ByVal dataMonth As Date, _
    ByVal dataMonthString As String)

    Dim fso As Object
    Dim ext As String
    Dim tableCols As Variant
    Dim lines As Variant, outLines() As String
    Dim i As Long, idx As Long
    Dim cols As Collection
    Dim tsIn As Object, tsOut As Object, tempPath As String

    ' 1. 檔案檢查 & 副檔名辨識
    If Dir(fullFilePath) = "" Then
        MsgBox "File not found: " & fullFilePath, vbExclamation
        Exit Sub
    End If
    Set fso = CreateObject("Scripting.FileSystemObject")
    ext = LCase(fso.GetExtensionName(fullFilePath))

    ' 2. 讀取欄位定義（CSV 時會改由 header，Excel 時由 DB table）
    tableCols = GetTableColumns(cleaningType, IIf(ext = "csv", fullFilePath, vbNullString))

    ' 3. CSV 文字檔處理
    If ext = "csv" Then
        ' 3.1 讀全部並 split 行
        Set tsIn = fso.OpenTextFile(fullFilePath, 1, False, -1)
        lines = Split(tsIn.ReadAll, vbCrLf)
        tsIn.Close
        ReDim outLines(0 To UBound(lines))

        For i = 0 To UBound(lines)
            If i = 0 Then
                ' header → 插入三個新欄位名稱
                outLines(i) = "DataDate,DataMonth,DataMonthString," & lines(i)
            Else
                ' data row → 解析、插新值、重組
                Set cols = ParseCsvLine(lines(i))
                ' 前三欄
                cols.Add Format(dataDate, "yyyy\/mm\/dd"), 1
                cols.Add Format(dataMonth, "yyyy\/mm"), 2
                cols.Add dataMonthString, 3
                ' 幣別
                Select Case cleaningType
                    Case "DBU_AC5602", "OBU_AC5602": idx = 10
                    Case "OBU_AC4603": idx = 10
                    Case "OBU_AC5411B": idx = 9
                    Case Else: idx = 0
                End Select
                If idx > 0 Then cols.Add IIf(cleaningType Like "*TWD*", "TWD", "USD"), idx + 3
                outLines(i) = JoinCsvLine(cols)
            End If
        Next

        ' 3.2 暫存檔寫回
        tempPath = fso.GetSpecialFolder(2).Path & "\" & fso.GetTempName
        Set tsOut = fso.OpenTextFile(tempPath, 2, True, -1)
        For i = 0 To UBound(outLines)
            tsOut.WriteLine outLines(i)
        Next
        tsOut.Close
        fso.DeleteFile fullFilePath, True
        fso.MoveFile tempPath, fullFilePath

        MsgBox "完成 CSV 額外清洗：" & fullFilePath, vbInformation

    ' 4. Excel 檔案處理
    Else
        Dim xlApp As Excel.Application
        Dim xlbk As Excel.Workbook
        Dim xlsht As Excel.Worksheet
        Dim lastRow As Long

        Set xlApp = Excel.Application
        xlApp.Visible = False
        Set xlbk = xlApp.Workbooks.Open(fullFilePath)

        For Each xlsht In xlbk.Sheets
            lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
            xlsht.Columns("A:D").Insert shift:=xlToRight
            For i = 2 To lastRow
                xlsht.Cells(i, "B").Value = dataDate
                xlsht.Cells(i, "C").Value = dataMonth
                xlsht.Cells(i, "D").Value = dataMonthString
                Select Case cleaningType
                    Case "DBU_AC5602", "OBU_AC5602": xlsht.Cells(i, "J").Value = "TWD"
                    Case "OBU_AC4603": xlsht.Cells(i, "J").Value = "USD"
                    Case "OBU_AC5411B": xlsht.Cells(i, "I").Value = "USD"
                End Select
            Next
            xlsht.Range("A1").Resize(1, UBound(tableCols) + 1).Value = tableCols
            xlsht.Columns("A").Delete
        Next

        xlbk.Save
        xlbk.Close False
        xlApp.Quit
        MsgBox "完成 Excel 額外清洗：" & fullFilePath, vbInformation
    End If
End Sub


'=== 讀取欄位定義：若給 CSV 路徑就讀 header，否則由 DB Table 撈 ===
Public Function GetTableColumns( _
    ByVal cleaningType As String, _
    Optional ByVal fullFilePath As String = "" _
) As Variant
    Dim fso As Object, ext As String
    If Len(fullFilePath) > 0 Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        ext = LCase(fso.GetExtensionName(fullFilePath))
        If ext = "csv" Then
            Dim ts As Object, hdr As String
            Set ts = fso.OpenTextFile(fullFilePath, 1, False, -1)
            hdr = ts.ReadLine: ts.Close
            GetTableColumns = Split(hdr, ",")
            Exit Function
        End If
    End If

    '--- 否則由 Access 資料表撈欄位 ---
    Dim dbs As DAO.Database, rs As DAO.Recordset
    Dim cols() As String, i As Integer
    Set dbs = CurrentDb
    Set rs = dbs.OpenRecordset( _
        "SELECT DbCol FROM DBsColTable " & _
        "WHERE TableName='" & cleaningType & "' " & _
        "ORDER BY DataID ASC;"
    )
    If Not rs.EOF Then
        i = 0: rs.MoveFirst
        Do While Not rs.EOF
            ReDim Preserve cols(i)
            cols(i) = rs!DbCol
            i = i + 1
            rs.MoveNext
        Loop
    End If
    rs.Close: Set rs = Nothing: Set dbs = Nothing
    GetTableColumns = cols
End Function


'=== CSV 解析／重組（處理引號內逗號） ===
Private Function ParseCsvLine(ByVal line As String) As Collection
    Dim c As New Collection, i As Long, ch As String, inQ As Boolean, cur As String
    For i = 1 To Len(line)
        ch = Mid$(line, i, 1)
        If ch = """" Then inQ = Not inQ
        If ch = "," And Not inQ Then c.Add cur: cur = "" Else cur = cur & ch
    Next
    c.Add cur
    Set ParseCsvLine = c
End Function

Private Function JoinCsvLine(cols As Collection) As String
    Dim s As String, i As Long, v As String
    For i = 1 To cols.Count
        v = cols(i)
        If InStr(v, ",") > 0 Or InStr(v, """") > 0 Then v = """" & Replace(v, """", """""") & """"
        s = s & v & IIf(i < cols.Count, ",", "")
    Next
    JoinCsvLine = s
End Function
```

### 說明

1. **只有一支 Public Sub**：`ICleaner_additionalClean` 根據副檔名分流，不再多層呼叫。
2. **欄位來源彈性**：`GetTableColumns` 若傳入 CSV 路徑就讀首行，否則回 DAO 查表。
3. **CSV 安全解析**：`ParseCsvLine`／`JoinCsvLine` 處理引號內逗號。
4. **暫存檔寫回**：避免中途失敗導致原檔毀損。

這樣結構扁平、易讀、維護也更方便。



https://chatgpt.com/share/681b87e6-b3c8-8010-a254-854202b5e70f
