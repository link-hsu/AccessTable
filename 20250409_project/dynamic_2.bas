你想要把這個查詢變成 Excel 可以直接使用的自訂函數，這完全可行！以下是修改後的版本，它會回傳 2 維陣列，可在 Excel 儲存格中像這樣使用：

\=GetUSD\_FVPL\_CostValuation\_SubType("C:\你的資料庫路徑\Finance.accdb", "2024/11")

🧠 功能說明：

* 函數名稱：GetUSD\_FVPL\_CostValuation\_SubType
* 回傳：含標題列的 2 維陣列（GroupKey, TotalValue）
* 使用情境：你可在儲存格中直接輸入，資料會即時撈出顯示

👇 以下是完整程式碼：

📌 請貼入標準模組 Module 中

```vb
Function GetUSD_FVPL_CostValuation_SubType(DBPath As String, DataMonth As String) As Variant
    Dim conn As Object, rs As Object
    Dim sql As String
    Dim dataArr As Variant, headerArr() As String
    Dim i As Long, j As Long, colCount As Long, rowCount As Long
    
    On Error GoTo ErrHandler

    sql = "SELECT AccountCodeMap.AssetMeasurementSubType AS GroupKey, " & _
          "SUM(OBU_AC5601.NetBalance) AS TotalValue " & _
          "FROM AccountCodeMap INNER JOIN OBU_AC5601 " & _
          "ON AccountCodeMap.AccountCode = OBU_AC5601.AccountCode " & _
          "WHERE AccountCodeMap.Category IN ('Cost','ValuationAdjust') " & _
          "AND OBU_AC5601.CurrencyType = 'USD' " & _
          "AND OBU_AC5601.DataMonthString = '" & DataMonth & "' " & _
          "GROUP BY AccountCodeMap.AssetMeasurementSubType"

    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn

    If rs.EOF Then
        GetUSD_FVPL_CostValuation_SubType = Array()
        GoTo Cleanup
    End If

    colCount = rs.Fields.Count
    ReDim headerArr(0 To colCount - 1)
    For i = 0 To colCount - 1
        headerArr(i) = rs.Fields(i).Name
    Next i

    dataArr = rs.GetRows()
    rowCount = UBound(dataArr, 2) + 1

    Dim resultArr() As Variant
    ReDim resultArr(0 To rowCount, 0 To colCount - 1)

    ' 標題列
    For i = 0 To colCount - 1
        resultArr(0, i) = headerArr(i)
    Next i

    ' 資料列
    For i = 0 To colCount - 1
        For j = 1 To rowCount
            resultArr(j, i) = dataArr(i, j - 1)
        Next j
    Next i

    GetUSD_FVPL_CostValuation_SubType = resultArr
    GoTo Cleanup

ErrHandler:
    MsgBox "錯誤：" & Err.Description, vbCritical
    GetUSD_FVPL_CostValuation_SubType = Array()

Cleanup:
    If Not rs Is Nothing Then rs.Close
    If Not conn Is Nothing Then conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Function
```

📥 在 Excel 中使用：

假設你 Access 檔案路徑為：

C:\Reports\MyData.accdb

並想查詢 2024/11 月資料，在某個儲存格輸入：

\=GetUSD\_FVPL\_CostValuation\_SubType("C:\Reports\MyData.accdb", "2024/11")

Excel 會自動顯示：

| GroupKey               | TotalValue |
| ---------------------- | ---------- |
| FVPL\_GovBond\_Foreign | 12345678   |
| FVPL\_CompanyBond\_... | ...        |

🔧 額外提示：
如果你要將這個函數拖曳使用，記得開啟 VBA 巨集信任與「允許自訂函數」功能。

需要我幫你做一個通用版（可傳入其他表格與欄位參數）嗎？



' ============================================================
' ============================================================
' ============================================================
' Answer2



**Designing UDF for Excel**

The user wants a flexible function that can handle several parameters like DBPath, DataMonth, CurrencyType, CategoryList, and others, returning a 2D array to use in Excel. It’s important this function works as a user-defined function (UDF) without showing MsgBox during recalculations. To handle CategoryList, I'll allow it to be passed as a Variant, accepting either a range or a comma-separated string. The function will gracefully handle errors and build SQL queries based on the provided inputs. If no records are found, it'll return an empty array.


Thought for 4 seconds


```vb
'===============================================================
' 通用 UDF：從 Access 依 AccountCodeMap 串接任意資料表，回傳二維陣列給 Excel
'---------------------------------------------------------------
' 您可以在 Excel 儲存格直接寫：
' =GetMappedData(
'     "C:\Your\Path\Finance.accdb",    ← Access 檔案完整路徑
'     "2024/11",                        ← DataMonthString
'     "USD",                            ← CurrencyType
'     A1:A2,                            ← 存放 Category 的儲存格範圍 (或文字 "Cost,ValuationAdjust")
'     "OBU_AC5601",                     ← 來源表
'     "NetBalance",                     ← 值欄位
'     "AssetMeasurementSubType"         ← 群組欄位；留空串 "" 則顯示明細
' )
'
' 會自動貼出第一列為欄位名稱，接著每列為查詢結果。
' 要改其他表、其他 Category、其他群組，只要改參數即可。
'===============================================================
Public Function GetMappedData( _
    ByVal DBPath As String, _
    ByVal DataMonth As String, _
    ByVal CurrencyType As String, _
    ByVal CategoryList As Variant, _
    ByVal SourceTable As String, _
    ByVal ValueField As String, _
    ByVal GroupByField As String _
) As Variant

    Dim cats()       As String
    Dim i As Long, j As Long
    Dim sql          As String
    Dim conn As Object, rs As Object
    Dim dataArr      As Variant
    Dim headerArr()  As String
    Dim resultArr    As Variant
    Dim colCount As Long, rowCount As Long
    Dim catStr       As String

    On Error GoTo ErrHandler

    '─── 處理 CategoryList 參數（可為 Range、1D/2D 陣列、逗號字串）
    If TypeOf CategoryList Is Excel.Range Then
        Dim v As Variant, tmp() As String, idx As Long
        v = CategoryList.Value
        If IsArray(v) Then
            ReDim tmp(0 To UBound(v, 1) * UBound(v, 2) - 1)
            idx = 0
            For i = LBound(v, 1) To UBound(v, 1)
                For j = LBound(v, 2) To UBound(v, 2)
                    If Len(CStr(v(i, j))) > 0 Then
                        tmp(idx) = CStr(v(i, j))
                        idx = idx + 1
                    End If
                Next
            Next
            ReDim Preserve tmp(0 To idx - 1)
            cats = tmp
        Else
            cats = Array(CStr(v))
        End If
    ElseIf IsArray(CategoryList) Then
        ' 直接傳入 VBA 陣列
        cats = CategoryList
    Else
        ' 文字，逗號分隔
        cats = Split(CStr(CategoryList), ",")
    End If

    '─── 組成 IN 子句
    For i = LBound(cats) To UBound(cats)
        cats(i) = "'" & Replace(cats(i), "'", "''") & "'"
    Next
    catStr = "(" & Join(cats, ",") & ")"

    '─── 組 SQL
    If Len(GroupByField) > 0 Then
        sql = "SELECT AccountCodeMap." & GroupByField & " AS GroupKey, " & _
              "SUM(" & SourceTable & "." & ValueField & ") AS TotalValue " & _
              "FROM AccountCodeMap INNER JOIN " & SourceTable & " " & _
              "ON AccountCodeMap.AccountCode=" & SourceTable & ".AccountCode " & _
              "WHERE AccountCodeMap.Category IN " & catStr & " " & _
              "AND " & SourceTable & ".CurrencyType='" & CurrencyType & "' " & _
              "AND " & SourceTable & ".DataMonthString='" & DataMonth & "' " & _
              "GROUP BY AccountCodeMap." & GroupByField
    Else
        sql = "SELECT " & SourceTable & ".DataID, " & SourceTable & ".DataMonthString, " & _
              "AccountCodeMap.AccountCode, AccountCodeMap.AccountTitle, " & _
              SourceTable & ".CurrencyType, " & SourceTable & "." & ValueField & " " & _
              "FROM AccountCodeMap INNER JOIN " & SourceTable & " " & _
              "ON AccountCodeMap.AccountCode=" & SourceTable & ".AccountCode " & _
              "WHERE AccountCodeMap.Category IN " & catStr & " " & _
              "AND " & SourceTable & ".CurrencyType='" & CurrencyType & "' " & _
              "AND " & SourceTable & ".DataMonthString='" & DataMonth & "'"
    End If

    '─── 開 ADO 連線、執行
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 1, 1    ' adOpenKeyset, adLockReadOnly

    If rs.EOF Then
        GetMappedData = Array()  ' 空結果
        GoTo CleanExit
    End If

    '─── 擷取欄位名稱
    colCount = rs.Fields.Count
    ReDim headerArr(0 To colCount - 1)
    For i = 0 To colCount - 1
        headerArr(i) = rs.Fields(i).Name
    Next

    '─── 擷取所有資料
    dataArr = rs.GetRows()      ' dataArr(field#, record#)
    rowCount = UBound(dataArr, 2) + 1

    '─── 合併成含 header 的二維陣列
    ReDim resultArr(0 To rowCount, 0 To colCount - 1)
    ' header
    For i = 0 To colCount - 1
        resultArr(0, i) = headerArr(i)
    Next
    ' data
    For i = 0 To colCount - 1
        For j = 1 To rowCount
            resultArr(j, i) = dataArr(i, j - 1)
        Next
    Next

    GetMappedData = resultArr

CleanExit:
    On Error Resume Next
    rs.Close: Set rs = Nothing
    conn.Close: Set conn = Nothing
    Exit Function

ErrHandler:
    ' UDF 中避免 MsgBox，直接回空
    GetMappedData = Array()
    Resume CleanExit
End Function
```

**使用範例**

1. 在工作表某格輸入：

   ```
   =GetMappedData(
     "C:\Reports\Finance.accdb",
     "2024/11",
     "USD",
     {"Cost","ValuationAdjust"},
     "OBU_AC5601",
     "NetBalance",
     "AssetMeasurementSubType"
   )
   ```
2. 或以儲存格範圍 A1\:A2 放兩個 Category，再寫：

   ```
   =GetMappedData(
     "C:\Reports\Finance.accdb",
     "2024/11",
     "USD",
     A1:A2,
     "OBU_AC5601",
     "NetBalance",
     "AssetMeasurementSubType"
   )
   ```

3.文字逗號分隔
=GetMappedData(
  "C:\Reports\Finance.accdb",   ← 資料庫路徑
  "2024/11",                    ← DataMonthString
  "USD",                        ← CurrencyType
  "Cost,ValuationAdjust",       ← CategoryList：逗號分隔的文字
  "OBU_AC5601",                 ← SourceTable
  "NetBalance",                 ← ValueField
  "AssetMeasurementSubType"     ← GroupByField
)

函數將自動「溢出 (spill)」出包含「GroupKey」與「TotalValue」的結果表格；若不指定 `GroupByField`（傳 ""），則會列出明細欄位。










' ======================================================================================
' ======================================================================================
' ======================================================================================
' Answer2

---

The user wants to modify the example to include a function that pushes data to an Excel table. They'll need to specify the table name and initial column parameters, so the subroutine should not only return an array but also populate an Excel ListObject starting at a designated column. This will involve modifying the UDF into a Sub called `PutMappedDataToTable` or a similar variant. The sub should take various arguments, locate the ListObject by name, clear existing rows, and fill in data starting from the specified column.

---

The user is asking for a subroutine to push mapped data to an Excel table, starting at a designated column and using specified table names. The process includes clearing existing data, adding rows if needed, and writing the data into the table's DataBodyRange. By using `ListObject.ListColumns(ColumnName).Index`, we can pinpoint column positions more easily and ensure the data fits properly. The subroutine will dynamically handle row additions and column alignment based on these specifications, ensuring flexibility and efficiency in the data export.


Thought for 8 seconds


下面範例示範如何，在「Excel 資料表（ListObject）」中，指定表名與起始欄位，將從 Access 撈回的二維資料直接填入表格。

1. **先貼入 Module1**（包含通用查 UDF）
2. **再貼入 Module2**（Push 到表格的 Sub）

---

```vb
'================ Module1: 通用查 UDF ====================
Public Function GetMappedData( _
    ByVal DBPath As String, _
    ByVal DataMonth As String, _
    ByVal CurrencyType As String, _
    ByVal CategoryList As Variant, _
    ByVal SourceTable As String, _
    ByVal ValueField As String, _
    ByVal GroupByField As String _
) As Variant

    Dim cats()       As String
    Dim i As Long, j As Long
    Dim sql          As String
    Dim conn As Object, rs As Object
    Dim dataArr      As Variant
    Dim headerArr()  As String
    Dim resultArr    As Variant
    Dim colCount As Long, rowCount As Long
    Dim catStr       As String

    On Error GoTo ErrHandler

    '── 處理 CategoryList（Range / 陣列 / 逗號字串皆可）
    If TypeOf CategoryList Is Excel.Range Then
        Dim v As Variant, tmp() As String, idx As Long
        v = CategoryList.Value
        If IsArray(v) Then
            ReDim tmp(0 To UBound(v, 1) * UBound(v, 2) - 1)
            For i = LBound(v, 1) To UBound(v, 1)
                For j = LBound(v, 2) To UBound(v, 2)
                    If Len(v(i, j) & "") > 0 Then
                        tmp(idx) = v(i, j)
                        idx = idx + 1
                    End If
                Next
            Next
            ReDim Preserve tmp(0 To idx - 1)
            cats = tmp
        Else
            cats = Array(CStr(v))
        End If
    ElseIf IsArray(CategoryList) Then
        cats = CategoryList
    Else
        cats = Split(CStr(CategoryList), ",")
    End If

    For i = LBound(cats) To UBound(cats)
        cats(i) = "'" & Replace(cats(i), "'", "''") & "'"
    Next
    catStr = "(" & Join(cats, ",") & ")"

    '── 組 SQL
    If Len(GroupByField) > 0 Then
        sql = "SELECT AccountCodeMap." & GroupByField & " AS GroupKey, " & _
              "SUM(" & SourceTable & "." & ValueField & ") AS TotalValue " & _
              "FROM AccountCodeMap INNER JOIN " & SourceTable & " " & _
              "ON AccountCodeMap.AccountCode=" & SourceTable & ".AccountCode " & _
              "WHERE AccountCodeMap.Category IN " & catStr & " " & _
              "AND " & SourceTable & ".CurrencyType='" & CurrencyType & "' " & _
              "AND " & SourceTable & ".DataMonthString='" & DataMonth & "' " & _
              "GROUP BY AccountCodeMap." & GroupByField
    Else
        sql = "SELECT " & SourceTable & ".DataID, " & SourceTable & ".DataMonthString, " & _
              "AccountCodeMap.AccountCode, AccountCodeMap.AccountTitle, " & _
              SourceTable & ".CurrencyType, " & SourceTable & "." & ValueField & " " & _
              "FROM AccountCodeMap INNER JOIN " & SourceTable & " " & _
              "ON AccountCodeMap.AccountCode=" & SourceTable & ".AccountCode " & _
              "WHERE AccountCodeMap.Category IN " & catStr & " " & _
              "AND " & SourceTable & ".CurrencyType='" & CurrencyType & "' " & _
              "AND " & SourceTable & ".DataMonthString='" & DataMonth & "'"
    End If

    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 1, 1

    If rs.EOF Then
        GetMappedData = Array()
        GoTo CleanExit
    End If

    colCount = rs.Fields.Count
    ReDim headerArr(0 To colCount - 1)
    For i = 0 To colCount - 1
        headerArr(i) = rs.Fields(i).Name
    Next

    dataArr = rs.GetRows()
    rowCount = UBound(dataArr, 2) + 1

    ReDim resultArr(0 To rowCount, 0 To colCount - 1)
    For i = 0 To colCount - 1
        resultArr(0, i) = headerArr(i)
        For j = 1 To rowCount
            resultArr(j, i) = dataArr(i, j - 1)
        Next
    Next

    GetMappedData = resultArr

CleanExit:
    On Error Resume Next
    rs.Close: conn.Close
    Set rs = Nothing: Set conn = Nothing
    Exit Function

ErrHandler:
    GetMappedData = Array()
    Resume CleanExit
End Function
'===============================================================
```

```vb
'================ Module2: Push 到 Excel 資料表 ====================
Public Sub PushMappedDataToTable( _
    ByVal DBPath As String, _
    ByVal DataMonth As String, _
    ByVal CurrencyType As String, _
    ByVal CategoryList As Variant, _
    ByVal SourceTable As String, _
    ByVal ValueField As String, _
    ByVal GroupByField As String, _
    ByVal TargetSheetName As String, _
    ByVal TableName As String, _
    ByVal StartColumnName As String _
)
    Dim arr As Variant
    arr = GetMappedData(DBPath, DataMonth, CurrencyType, CategoryList, SourceTable, ValueField, GroupByField)
    If Not IsArray(arr) Then Exit Sub
    If UBound(arr, 1) < 1 Then Exit Sub   ' 無資料 or 只有 header

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(TargetSheetName)
    Dim lo As ListObject: Set lo = ws.ListObjects(TableName)
    Dim hdrRow As Long: hdrRow = lo.HeaderRowRange.Row

    ' 找起始欄位的絕對欄號
    Dim sc As Range
    Set sc = lo.HeaderRowRange.Find(What:=StartColumnName, LookAt:=xlWhole)
    If sc Is Nothing Then
        MsgBox "找不到欄 """ & StartColumnName & """", vbCritical: Exit Sub
    End If
    Dim startCol As Long: startCol = sc.Column

    Dim nCols As Long: nCols = UBound(arr, 2) + 1
    Dim nRecs As Long: nRecs = UBound(arr, 1)

    ' 更新表頭
    Dim c As Long
    For c = 1 To nCols
        ws.Cells(hdrRow, startCol + c - 1).Value = arr(0, c - 1)
    Next

    ' 刪除原有資料列
    On Error Resume Next
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
    On Error GoTo 0

    ' 填入新資料
    Dim dataArr As Variant
    ReDim dataArr(1 To nRecs, 1 To nCols)
    Dim r As Long, cc As Long
    For r = 1 To nRecs
        For cc = 1 To nCols
            dataArr(r, cc) = arr(r, cc - 1)
        Next
    Next

    ' 將資料寫入：從 header 下一列開始
    ws.Cells(hdrRow + 1, startCol).Resize(nRecs, nCols).Value = dataArr

    ' 讓 Table 自動擴展
    lo.Resize lo.Range.Resize(nRecs + 1)

End Sub
'===============================================================
```

---

## 範例：把 USD / Cost & ValuationAdjust，依 SubType 彙總結果推到 Table1 的「GroupKey」欄

```vb
Sub Test_PushToTable()
    Dim cats As Variant
    cats = Array("Cost", "ValuationAdjust")
    Call PushMappedDataToTable( _
        "C:\Reports\Finance.accdb", _
        "2024/11", _
        "USD", _
        cats, _
        "OBU_AC5601", _
        "NetBalance", _
        "AssetMeasurementSubType", _
        "Sheet1", _
        "Table1", _
        "GroupKey" _
    )
End Sub
```

* 此範例會：

  1. 清空 Sheet1!Table1 內原有資料
  2. 先在 Table1 header 找到「GroupKey」欄
  3. 更新 Table1 「GroupKey」「TotalValue」兩欄標題
  4. 自動擴展 Table，將所有彙總後的資料貼入

你只要改 `PushMappedDataToTable` 的 10 個參數，就能將任意表、任意欄的查詢結果，推到指定 Excel 資料表的指定欄位中。





https://chatgpt.com/share/681763f6-3bbc-8010-966e-b769b49387f9
