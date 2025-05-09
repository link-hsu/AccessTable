'CleanerAddColumns
Option Compare Database

Implements ICleaner

Private clsHasData As Boolean

Public Function ICleaner_HasFile() As Boolean
    ICleaner_HasFile = clsHasFile
End Function

Public Function ICleaner_HasData() As Boolean
    ICleaner_HasData = clsHasData
End Function

Public Sub ICleaner_Initialize(Optional ByVal sheetName As Variant = 1, _
                               Optional ByVal loopColumn As Integer = 1, _
                               Optional ByVal leftToDelete As Integer = 2, _
                               Optional ByVal rightToDelete As Integer = 3, _
                               Optional ByVal rowsToDelete As Variant, _
                               Optional ByVal colsToDelete As Variant, _
                               Optional ByVal colsToHandle As Variant)
End Sub

Public Sub ICleaner_CleanReport(ByVal fullFilePath As String, _
                                ByVal cleaningType As String, _
                                ByVal xlApp As Excel.Application)    
    'implement operations here
End Sub

Public Sub ICleaner_additionalClean(ByVal fullFilePath As String, _
                                    ByVal cleaningType As String, _
                                    ByVal dataDate As Date, _
                                    ByVal dataMonth As Date, _
                                    ByVal dataMonthString As String, _
                                    ByVal xlApp As Excel.Application)
    Dim fso As Object
    Dim ext As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    ext = LCase(fso.GetExtensionName(fullFilePath))

    If ext = "csv" Then
        ' CSV 處理
        Dim fileIn As Object
        Dim lines As Variant
        Dim outLines() As String
        Dim parts As Variant
        Dim i As Long
        Dim headerCols As Variant

        Set fileIn = fso.OpenTextFile(fullFilePath, 1, False)
        lines = Split(fileIn.ReadAll, vbLf)
        fileIn.Close

        headerCols = GetTableColumns(cleaningType)
        
        ReDim outLines(UBound(lines))
        For i = 0 To UBound(lines)
            If Trim(lines(i)) = "" Then GoTo SkipLine2
            parts = Split(lines(i), ",")
            If i = 0 Then
                ' outLines(i) = Join(headerCols, ",")
                outLines(i) = Join(Filter(headerCols, headerCols(0), False, vbTextCompare), ",")
            Else
                outLines(i) = Format(dataDate, "yyyy-mm-dd") & "," & _
                              Format(dataMonth, "yyyy-mm-dd") & "," & _
                              dataMonthString & "," & lines(i)
            End If
SkipLine2:
        Next i

        Set fileIn = fso.OpenTextFile(fullFilePath, 2, True)
        For i = 0 To UBound(outLines)
            ' fileIn.WriteLine outLines(i)
            fileIn.Write outLines(i) & vbLf
        Next i
        fileIn.Close

        MsgBox "完成Addition Clean，路徑為: " &  fullFilePath
        WriteLog "完成Addition Clean，路徑為: " &  fullFilePath

        Exit Sub
    End If

    Dim xlbk As Excel.Workbook
    Dim xlsht As Excel.Worksheet
    Dim lastRow As Long
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
