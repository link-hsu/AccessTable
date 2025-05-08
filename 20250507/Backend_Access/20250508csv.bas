Thought for a couple of seconds


以下是整合後的完整 VBA 程式碼，已將 `ICleaner_CleanCSV` 中對 header 的處理移除，並包含所有你所需的功能：空格、百分號、兩種日期格式清理，CSV／Excel 的 additionalClean，以及 Access 匯入程序。請直接貼到標準模組 (.bas) 中。

```vb
Option Explicit

' ==============================================
' 1) CSV 清理：RemoveSpaceCols, RemovePercentCols, FormatDateCols, FormatDateColsV2
' ==============================================
Public Sub ICleaner_CleanCSV(ByVal fullFilePath As String, ByVal cleaningType As String)
    Dim fso           As Object
    Dim inputFile     As Object
    Dim outputFile    As Object
    Dim inputLines    As Variant
    Dim outputLines() As String
    Dim colsDict      As Object
    Dim i             As Long
    Dim lineParts     As Variant
    Dim colIndex      As Variant

    ' 檔案檢查
    If Dir(fullFilePath) = "" Then
        MsgBox "File not found: " & fullFilePath, vbCritical
        Exit Sub
    End If

    ' 讀入所有行
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set inputFile = fso.OpenTextFile(fullFilePath, 1, False)
    inputLines = Split(inputFile.ReadAll, vbCrLf)
    inputFile.Close

    ' 如果沒有要清理的欄位，跳出
    If Not (IsObject(clsColsToHandle) And TypeName(clsColsToHandle) = "Dictionary") Then
        MsgBox "No cleaning rules defined.", vbExclamation
        Exit Sub
    End If

    ' 把 clsColsToHandle 轉成 Dictionary：key = cleaningType, value = array of colIndex
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

    ' 逐行處理（保留 header 原樣）
    For i = 0 To UBound(inputLines)
        If Trim(inputLines(i)) = "" Then GoTo SkipLine

        lineParts = Split(inputLines(i), ",")

        If i > 0 Then   ' 跳過第一列 header
            ' RemovePercentCols
            If colsDict.Exists("RemovePercentCols") Then
                For Each colIndex In colsDict("RemovePercentCols")
                    If colIndex - 1 <= UBound(lineParts) Then
                        lineParts(colIndex - 1) = Replace(lineParts(colIndex - 1), "%", "")
                    End If
                Next
            End If
            ' RemoveSpaceCols
            If colsDict.Exists("RemoveSpaceCols") Then
                For Each colIndex In colsDict("RemoveSpaceCols")
                    If colIndex - 1 <= UBound(lineParts) Then
                        lineParts(colIndex - 1) = Replace(lineParts(colIndex - 1), " ", "")
                    End If
                Next
            End If
            ' FormatDateCols （移除 .0 並轉 yyyy-mm-dd）
            If colsDict.Exists("FormatDateCols") Then
                For Each colIndex In colsDict("FormatDateCols")
                    If colIndex - 1 <= UBound(lineParts) Then
                        lineParts(colIndex - 1) = FormatDateTimeFix(lineParts(colIndex - 1))
                    End If
                Next
            End If
            ' FormatDateColsV2 （純 yyyy-mm-dd）
            If colsDict.Exists("FormatDateColsV2") Then
                For Each colIndex In colsDict("FormatDateColsV2")
                    If colIndex - 1 <= UBound(lineParts) Then
                        lineParts(colIndex - 1) = FormatDateV2(lineParts(colIndex - 1))
                    End If
                Next
            End If
        End If

        outputLines(i) = Join(lineParts, ",")
SkipLine:
    Next i

    ' 覆寫回原檔
    Set outputFile = fso.OpenTextFile(fullFilePath, 2, True)
    For i = 0 To UBound(outputLines)
        outputFile.WriteLine outputLines(i)
    Next i
    outputFile.Close

    MsgBox "完成 CSV 清理：" & cleaningType & vbCrLf & fullFilePath, vbInformation
End Sub

' 移除尾端 .0，並轉成 yyyy-mm-dd
Private Function FormatDateTimeFix(ByVal inputStr As String) As String
    On Error GoTo ErrHandler
    Dim temp As String
    temp = Replace(inputStr, ".0", "")
    FormatDateTimeFix = Format(CDate(temp), "yyyy-mm-dd")
    Exit Function
ErrHandler:
    FormatDateTimeFix = inputStr
End Function

' 強制轉 yyyy-mm-dd（若原本就是該格式，直接回傳）
Private Function FormatDateV2(ByVal inputStr As String) As String
    On Error GoTo ErrHandler
    FormatDateV2 = Format(CDate(inputStr), "yyyy-mm-dd")
    Exit Function
ErrHandler:
    FormatDateV2 = inputStr
End Function


' ==============================================
' 2) 在 CSV / Excel 前面插入 dataDate、dataMonth、dataMonthString
' ==============================================
Public Sub ICleaner_additionalClean(ByVal fullFilePath As String, _
                                    ByVal cleaningType As String, _
                                    ByVal dataDate As Date, _
                                    ByVal dataMonth As Date, _
                                    ByVal dataMonthString As String, _
                                    ByVal xlApp As Excel.Application)
    Dim ext As String
    ext = LCase(Right(fullFilePath, 4))

    If ext = ".csv" Then
        ' CSV 處理
        Dim fso       As Object, fileIn As Object
        Dim lines     As Variant, outLines() As String
        Dim parts     As Variant
        Dim i         As Long
        Dim hdrCols   As Variant

        ' 讀檔
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set fileIn = fso.OpenTextFile(fullFilePath, 1, False)
        lines = Split(fileIn.ReadAll, vbCrLf)
        fileIn.Close

        ' header 欄位取 DB 定義（跳過 DataID）
        hdrCols = GetTableColumns(cleaningType)

        ReDim outLines(UBound(lines))
        For i = 0 To UBound(lines)
            If Trim(lines(i)) = "" Then GoTo SkipLine2
            parts = Split(lines(i), ",")
            If i = 0 Then
                outLines(i) = "dataDate,dataMonth,dataMonthString," & Join(hdrCols, ",")
            Else
                outLines(i) = Format(dataDate, "yyyy-mm-dd") & "," & _
                              Format(dataMonth, "yyyy-mm-dd") & "," & _
                              dataMonthString & "," & lines(i)
            End If
SkipLine2:
        Next i

        Set fileIn = fso.OpenTextFile(fullFilePath, 2, True)
        For i = 0 To UBound(outLines)
            fileIn.WriteLine outLines(i)
        Next i
        fileIn.Close

        MsgBox "完成 CSV Additional Clean：" & fullFilePath, vbInformation
        Exit Sub
    End If

    ' Excel 處理（原邏輯）
    Dim xlbk    As Excel.Workbook
    Dim xlsht   As Excel.Worksheet
    Dim lastRow As Long
    Dim tableColumns As Variant
    Dim j       As Long

    Set xlbk = xlApp.Workbooks.Open(fullFilePath)
    tableColumns = GetTableColumns(cleaningType)

    For Each xlsht In xlbk.Sheets
        lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
        xlsht.Columns("A:D").Insert shift:=xlToRight
        For j = 2 To lastRow
            xlsht.Cells(j, "B").Value = dataDate
            xlsht.Cells(j, "C").Value = dataMonth
            xlsht.Cells(j, "D").Value = dataMonthString
            ' 幣別示例
            If cleaningType = "DBU_AC5602" Or cleaningType = "OBU_AC5602" Then
                xlsht.Cells(j, "J").Value = "TWD"
            ElseIf cleaningType = "OBU_AC4603" Then
                xlsht.Cells(j, "J").Value = "USD"
            ElseIf cleaningType = "OBU_AC5411B" Then
                xlsht.Cells(j, "I").Value = "USD"
            End If
        Next j
        xlsht.Range("A1").Resize(1, UBound(tableColumns) + 1).Value = tableColumns
        xlsht.Columns("A").Delete

        ' 刪除多餘空白列
        Dim lastDataRow As Long, usedR As Long
        lastDataRow = xlsht.Cells(xlsht.Rows.Count, "A").End(xlUp).Row
        usedR = xlsht.UsedRange.Rows.Count
        If lastDataRow < usedR Then
            xlsht.Rows(lastDataRow + 1 & ":" & usedR).Delete
        End If
    Next xlsht

    xlbk.Save
    xlbk.Close False
    MsgBox "完成 Excel Additional Clean：" & fullFilePath, vbInformation
End Sub


' ==============================================
' 3) 匯入 Access：同時支援 .csv 與 .xlsx
' ==============================================
Public Sub IImporter_ImportData(ByVal fullFilePath As String, _
                                ByVal accessDBPath As String, _
                                ByVal tableName As String, _
                                ByVal xlApp As Excel.Application)
    Dim cn          As Object
    Dim xlbk        As Object
    Dim sqlString   As String
    Dim tableColumns As Variant
    Dim i           As Integer
    Dim fieldList   As String, selectList As String
    Dim ext         As String

    ' 開 ADODB 連線
    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & accessDBPath & ";"

    ' 取欄位定義
    tableColumns = GetTableColumns(tableName)
    If UBound(tableColumns) < 1 Then
        MsgBox "資料表 " & tableName & " 至少需要 2 個欄位。", vbCritical
        GoTo CleanUpImp
    End If

    ' 組 fieldList / selectList（跳過 Primary Key）
    For i = 1 To UBound(tableColumns)
        fieldList = fieldList & "[" & tableColumns(i) & "],"
        selectList = selectList & "[" & tableColumns(i) & "],"
    Next
    fieldList = Left(fieldList, Len(fieldList) - 1)
    selectList = Left(selectList, Len(selectList) - 1)

    ext = LCase(Right(fullFilePath, 4))
    If ext = ".csv" Then
        ' CSV via Text 驅動
        sqlString = "INSERT INTO " & tableName & " (" & fieldList & ") " & _
                    "SELECT " & selectList & " FROM " & _
                    "[Text;FMT=Delimited;HDR=YES;Database=" & GetFolderPath(fullFilePath) & "].[" & GetFileName(fullFilePath) & "]"
        cn.Execute sqlString

    Else
        ' Excel 驅動
        Set xlbk = xlApp.Workbooks.Open(fullFilePath)
        For i = 1 To xlbk.Sheets.Count
            sqlString = "INSERT INTO " & tableName & " (" & fieldList & ") " & _
                        "SELECT " & selectList & " FROM " & _
                        "[Excel 12.0 Xml;HDR=YES;Database=" & fullFilePath & "].[" & xlbk.Sheets(i).Name & "$]"
            cn.Execute sqlString
        Next
        xlbk.Close False
    End If

    MsgBox "完成 " & tableName & " 匯入作業", vbInformation

CleanUpImp:
    If Not cn Is Nothing Then
        cn.Close
        Set cn = Nothing
    End If
End Sub

' 取得資料夾路徑（不含最後的 "\"）
Private Function GetFolderPath(ByVal fullPath As String) As String
    GetFolderPath = Left(fullPath, InStrRev(fullPath, "\") - 1)
End Function

' 取得檔名（含附檔名）
Private Function GetFileName(ByVal fullPath As String) As String
    GetFileName = Mid(fullPath, InStrRev(fullPath, "\") + 1)
End Function

' ==============================================
' 取得資料表欄位對應（中→英），已由使用者提供
' ==============================================
Public Function GetTableColumns(ByVal cleaningType As String) As Variant
    Dim dbs As DAO.Database
    Dim rs  As DAO.Recordset
    Dim arr() As String
    Dim i   As Integer

    Set dbs = CurrentDb
    Set rs = dbs.OpenRecordset( _
        "SELECT DbCol FROM DBsColTable WHERE TableName='" & cleaningType & "' ORDER BY DataID;")

    If rs.EOF Then
        GetTableColumns = Array()
    Else
        i = 0
        Do While Not rs.EOF
            ReDim Preserve arr(i)
            arr(i) = rs!DbCol
            i = i + 1
            rs.MoveNext
        Loop
        GetTableColumns = arr
    End If

    rs.Close: Set rs = Nothing: Set dbs = Nothing
End Function
```

---

### 說明重點

1. **ICleaner\_CleanCSV**

   * 第一列（header）完全不動，只對第 2 列開始做清理。
   * 處理四種清理類型：移除空格、移除 `%`、去尾 `.0` 並格式化日期、純 YYYY‑MM‑DD 格式化。

2. **ICleaner\_additionalClean**

   * 若副檔名 `.csv`，直接用 FileSystemObject 讀寫，將前三欄插入並加上 header。
   * 否則執行原 Excel 版邏輯。

3. **IImporter\_ImportData**

   * 偵測 `.csv` 或 `.xlsx`，分別使用 Text 驅動或 Excel 驅動，動態 `INSERT INTO`。

4. **數字格式**

   * Access Currency 四位小數，若你的 CSV 中有超過 4 位小數，Access 會自動四捨五入。如需自行控制，可在清理階段加入 `FormatNumber(val,4)`。

貼上上述程式後，請在模組頂端啟用 Microsoft DAO 及 Microsoft ActiveX Data Objects 參考。
