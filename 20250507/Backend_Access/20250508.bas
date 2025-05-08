CleanerAddColumns_test.bas

If LCase(Right(fullFilePath, 4)) = ".csv" Then
    ' CSV 匯入
    sqlString = "INSERT INTO " & tableName & " (" & fieldList & ") " & _
                "SELECT " & selectList & " FROM [Text;HDR=Yes;FMT=Delimited;Database=" & Left(fullFilePath, InStrRev(fullFilePath, "\")) & "].[" & Mid(fullFilePath, InStrRev(fullFilePath, "\") + 1) & "]"
    cn.Execute sqlString
Else




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
        WriteLog "File does not exist in path: " & fullFilePath
        Exit Sub
    End If

    ' 啟動 Excel 並開啟工作簿
    Set xlbk = xlApp.Workbooks.Open(fullFilePath)
    
    'Set Table Columns
    tableColumns = GetTableColumns(cleaningType)

    If LCase(Right(fullFilePath, 4)) = ".csv" Then

        ' 2. 建立新的 header
        Dim fso            As Object
        Dim inputLines     As Variant, outputLines() As String
        Dim newHeader As String
        newHeader = Join(tableColumns, ",")

        Set fso = CreateObject("Scripting.FileSystemObject")
        inputLines = Split(fso.OpenTextFile(fullFilePath, 1).ReadAll, vbCrLf)
        ReDim outputLines(UBound(inputLines))

        For i = 0 To UBound(inputLines)
            If Trim(inputLines(i)) = "" Then GoTo SkipLine
    
            ' header
            If i = 0 Then
                outputLines(i) = newHeader
                GoTo SkipLine
            End If
    
            ' data row → 先拆欄，再做原本清理，再加前置三欄
            lineParts = Split(inputLines(i), ",")
    
            Dim prefix As String
            prefix = "2024/1/1,2025/2,2026/03"
            outputLines(i) = prefix & "," & Join(lineParts, ",")
    SkipLine:
        Next
    Else

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

        ' 刪除最下面的空白Row
        Dim lastDataRow As Long, currentLastRow As Long
        lastDataRow = xlsht.Cells(xlsht.Rows.Count, "A").End(xlUp).Row

        ' 如果 UsedRange 的列數大於 lastDataRow，表示底部可能有多餘的空白列
        currentLastRow = xlsht.UsedRange.Rows.Count

        If lastDataRow < currentLastRow Then
            xlsht.Rows(lastDataRow + 1 & ":" & currentLastRow).Delete
        End If
    Next xlsht

    ' 儲存並關閉檔案    
    xlbk.Save
    xlbk.Close False
    Set xlsht = Nothing
    Set xlbk = Nothing

    MsgBox "完成Addition Clean，路徑為: " &  fullFilePath
    WriteLog "完成Addition Clean，路徑為: " &  fullFilePath
    ' Debug.Print "完成Addition Clean，路徑為: " &  fullFilePath
End Sub





'==========================================================

answer2.bas



下面範例示範，如何在清 CSV 的階段：
	1.	先拿到目標 Access 欄位（取代原本 CSV header）
	2.	在最前面三欄插入固定欄位名：date,month,monthstring
	3.	每一筆資料行，都在最前面加上固定值：2024/1/1,2025/2,2026/03



Public Sub ICleaner_CleanCSV_WithPrefix(ByVal fullFilePath As String, _
                                        ByVal cleaningType As String)

    Dim fso            As Object
    Dim inputLines     As Variant, outputLines() As String
    Dim colsDict       As Object
    Dim i              As Long, lineParts As Variant, colIndex As Variant

    ' 讀檔
    If Dir(fullFilePath) = "" Then
        MsgBox "File not found: " & fullFilePath: Exit Sub
    End If
    Set fso = CreateObject("Scripting.FileSystemObject")
    inputLines = Split(fso.OpenTextFile(fullFilePath, 1).ReadAll, vbCrLf)

    ' 1. 取出 Access 欄位 (利用您先前的 GetTableColumns 函式)
    Dim tableCols As Variant
    tableCols = GetTableColumns(cleaningType)   ' ← 取得對應的 DbCol 陣列

    ' 2. 建立新的 header
    Dim newHeader As String
    newHeader = "date,month,monthstring," & Join(tableCols, ",")

    ' 3. 讀 clsColsToHandle 建 Dictionary（同您原本程式） 
    If Not (IsObject(clsColsToHandle) And TypeName(clsColsToHandle) = "Dictionary") Then
        MsgBox "沒有設定要處理的欄位": Exit Sub
    End If
    Set colsDict = CreateObject("Scripting.Dictionary")
    For i = 0 To UBound(clsColsToHandle) Step 2
        Dim key As String, val As Long, arr As Variant
        key = clsColsToHandle(i): val = clsColsToHandle(i + 1)
        If colsDict.Exists(key) Then
            arr = colsDict(key)
            ReDim Preserve arr(UBound(arr) + 1)
            arr(UBound(arr)) = val
            colsDict(key) = arr
        Else
            colsDict.Add key, Array(val)
        End If
    Next

    ' 4. 開始逐列處理
    ReDim outputLines(UBound(inputLines))
    For i = 0 To UBound(inputLines)
        If Trim(inputLines(i)) = "" Then GoTo SkipLine

        ' header
        If i = 0 Then
            outputLines(i) = newHeader
            GoTo SkipLine
        End If

        ' data row → 先拆欄，再做原本清理，再加前置三欄
        lineParts = Split(inputLines(i), ",")

        ' — 原本的 RemovePercentCols／RemoveSpaceCols／FormatDateCols… etc. —
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
        If colsDict.Exists("FormatDateCols") Then
            For Each colIndex In colsDict("FormatDateCols")
                If colIndex <= UBound(lineParts) Then
                    lineParts(colIndex - 1) = CleanDateTime(lineParts(colIndex - 1))
                End If
            Next
        End If
        If colsDict.Exists("FormatDateColsV2") Then
            For Each colIndex In colsDict("FormatDateColsV2")
                If colIndex <= UBound(lineParts) Then
                    lineParts(colIndex - 1) = CleanDateOnly(lineParts(colIndex - 1))
                End If
            Next
        End If

        ' — 在最前面加三個固定值 —
        Dim prefix As String
        prefix = "2024/1/1,2025/2,2026/03"
        outputLines(i) = prefix & "," & Join(lineParts, ",")
SkipLine:
    Next

    ' 5. 覆寫回 CSV
    Dim outF As Object, j As Long
    Set outF = fso.OpenTextFile(fullFilePath, 2, True)
    For j = 0 To UBound(outputLines)
        outF.WriteLine outputLines(j)
    Next
    outF.Close

    MsgBox "已清理並加上三欄前置值：" & vbCrLf & fullFilePath
End Sub

關鍵說明
	1.	取代原 header
我們用 GetTableColumns(cleaningType) 拿到您在 DB 中定義好的欄位名稱陣列，然後 Join(..., ",") 組成新的 CSV header。
	2.	三個固定欄位
在 header 前面硬加入 date,month,monthstring，在每筆資料前面硬加上 2024/1/1,2025/2,2026/03。
	3.	清理流程不變
您原本的各種 RemoveSpaceCols、RemovePercentCols、FormatDateCols、FormatDateColsV2 還是放在拆欄後執行。
	4.	Helper 函式

Private Function CleanDateTime(ByVal s As String) As String
    On Error Resume Next
    If InStrRev(s, ".") > 0 Then s = Left(s, InStrRev(s, ".") - 1)
    If IsDate(s) Then CleanDateTime = Format(CDate(s), "yyyy-mm-dd hh:nn:ss") Else CleanDateTime = s
End Function

Private Function CleanDateOnly(ByVal s As String) As String
    On Error Resume Next
    If IsDate(s) Then CleanDateOnly = Format(CDate(s), "yyyy-mm-dd") Else CleanDateOnly = s
End Function



之後您在 NT Position 裡面，對於每個 cleaningType 一樣建好 clsColsToHandle，再呼叫：

Call ICleaner_CleanCSV_WithPrefix("C:\path\to\your.csv", cleaningType)
Call IImporter_ImportData("C:\path\to\your.csv", "C:\path\to\db.accdb", cleaningType, xlApp)

即可完成：
	1.	以 DB 定義的欄位覆蓋 CSV header
	2.	每列前新增三個固定值欄位
	3.	清理、匯入 Access 一氣呵成。





' ====================================================
answer.bas


你的需求牽涉到以下幾個重點流程：

⸻

一、CSV 清理程序（ICleaner_CleanCSV）

根據欄位類型清理資料（移除空白、% 符號、格式化日期），並支援 Dictionary 結構設定清理邏輯。

修改目標
	•	加入 FormatDateCols 處理格式 YYYY-MM-DD 00:00:00.0 -> YYYY-MM-DD
	•	加入 FormatDateColsV2 處理格式 YYYY-MM-DD，不需修改，可略過
	•	Header 轉換（中文 -> 英文）對應欄位
	•	處理 CSV 的尾行空白
	•	日期格式是否需要清理：Access 可以接受 YYYY-MM-DD，但 YYYY-MM-DD 00:00:00.0 會造成錯誤或當成文字處理，建議清理

VBA 完整修改後程式碼：

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

    If Not (IsObject(clsColsToHandle) And TypeName(clsColsToHandle) = "Dictionary") Then GoTo CleanUp

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

            If colsDict.Exists("FormatDateCols") Then
                For Each colIndex In colsDict("FormatDateCols")
                    If colIndex <= UBound(lineParts) Then
                        Dim originalDate As String
                        originalDate = lineParts(colIndex - 1)
                        If InStr(originalDate, " ") > 0 Then
                            lineParts(colIndex - 1) = Left(originalDate, InStr(originalDate, " ") - 1)
                        End If
                    End If
                Next
            End If
            ' FormatDateColsV2 可略過，因為格式已符合 Access 日期格式
        End If
        outputLines(i) = Join(lineParts, ",")
SkipLine:
    Next i

    Set outputFile = fso.OpenTextFile(fullFilePath, 2, True)
    For i = 0 To UBound(outputLines)
        If Trim(outputLines(i)) <> "" Then
            outputFile.WriteLine outputLines(i)
        End If
    Next i

CleanUp:
    outputFile.Close
    MsgBox "完成清理：" & cleaningType & vbCrLf & fullFilePath
End Sub



⸻

二、擴充 GetTableColumns：支援 CSV header

加入 CSV 檔案的欄位名稱轉換邏輯，讓清理後欄位名稱可與 Access 欄位對應。

Public Function GetTableColumns(ByVal cleaningType As String) As Variant
    Dim dbs As DAO.Database
    Dim rs As DAO.Recordset
    Dim tableColumns() As String
    Dim i As Integer

    Set dbs = CurrentDb
    Set rs = dbs.OpenRecordset("SELECT TableName, DbCol FROM DBsColTable WHERE TableName = '" & cleaningType & "' ORDER BY DataID ASC;")

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



⸻

三、擴充 IImporter_ImportData：支援匯入 CSV（非 Excel）

你需要將 Excel 部分包裝成條件式，並新增支援 CSV 檔案的 SQL 方式。

Public Sub IImporter_ImportData(ByVal fullFilePath As String, _
                                ByVal accessDBPath As String, _
                                ByVal tableName As String, _
                                ByVal xlApp As Excel.Application)

    Dim cn As Object
    Dim sqlString As String
    Dim tableColumns As Variant
    Dim fieldList As String, selectList As String
    Dim i As Integer

    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & accessDBPath & ";"

    tableColumns = GetTableColumns(tableName)
    If UBound(tableColumns) < 1 Then
        MsgBox "資料表 " & tableName & " 至少需要 2 個欄位", vbCritical
        Exit Sub
    End If

    For i = 1 To UBound(tableColumns)
        fieldList = fieldList & "[" & tableColumns(i) & "],"
        selectList = selectList & "[" & tableColumns(i) & "],"
    Next i
    fieldList = Left(fieldList, Len(fieldList) - 1)
    selectList = Left(selectList, Len(selectList) - 1)

    If LCase(Right(fullFilePath, 4)) = ".csv" Then
        ' CSV 匯入
        sqlString = "INSERT INTO " & tableName & " (" & fieldList & ") " & _
                    "SELECT " & selectList & " FROM [Text;HDR=Yes;FMT=Delimited;Database=" & Left(fullFilePath, InStrRev(fullFilePath, "\")) & "].[" & Mid(fullFilePath, InStrRev(fullFilePath, "\") + 1) & "]"
        cn.Execute sqlString
    Else
        ' Excel 匯入
        Dim xlbk As Object, sheetName As String
        Set xlbk = xlApp.Workbooks.Open(fullFilePath)
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

    MsgBox "完成 " & tableName & " 資料匯入作業"
End Sub



⸻

四、數值處理：75000000.000000 是否能匯入 Access 的貨幣欄位？

可以，但注意以下幾點：
	•	Access 的貨幣型別支援到四位小數（會自動捨入或截斷）
	•	若匯入資料中為純文字（CSV 格式中帶引號或前有空白），Access 可能會視為文字，導致欄位型別不符
	•	建議匯入前移除多餘小數點與空格，以 FormatNumber(value, 2) 格式清理到小數點後兩位

⸻

總結建議工作流程：
	1.	透過 ICleaner_CleanCSV 依據表格類型進行清理
	2.	清理時轉換日期格式、移除 %、空白等
	3.	用 GetTableColumns 對應欄位（中文對應英文）
	4.	使用 IImporter_ImportData 將清理後 CSV 匯入 Access
	5.	Access 資料表請先設好欄位型別，確保對應（日期/數字/文字）

⸻

是否也需要我幫你整理成模組、類別格式？









