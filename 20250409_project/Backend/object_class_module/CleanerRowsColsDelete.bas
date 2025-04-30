' CleanRowsColsDelete
Option Compare Database

Implements ICleaner

Private clsSheetName As Variant
Private clsLoopColumn As Integer
Private clsLeftToDelete As Integer
Private clsRightToDelete As Integer
Private clsRowsToDelete As Variant
Private clsColsToDelete As Variant

Private clsHasFile As Boolean
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
    Dim tableColumns As Variant

    If IsEmpty(clsRowsToDelete) Then clsRowsToDelete = Array()
    If IsEmpty(clsColsToDelete) Then clsColsToDelete = Array()

    ' 檢查檔案是否存在
    If Dir(fullFilePath) = "" Then
        clsHasFile = False
        MsgBox "File does not exist in path: " & fullFilePath
        WriteLog "File does not exist in path: " & fullFilePath
        Exit Sub
    Else
        clsHasFile = True
    End If

    ' 啟動 Excel 並開啟工作簿
    Set xlbk = xlApp.Workbooks.Open(fullFilePath)
    Set xlsht = xlbk.Sheets(clsSheetName)


    ' 找出最後一列
    '**注意這邊是使用A欄位最後儲存格當作Row Number, 為了DL9360修改
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    '**原來
    ' lastRow = xlsht.Cells(xlsht.Rows.Count, clsLoopColumn).End(xlUp).Row

    If (lastRow > 1) Then
        clsHasData = True
    Else
        clsHasData = False
    End If

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

    If clsHasFile And clsHasData Then
        MsgBox "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
        WriteLog "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
        ' Debug.Print "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
    End If
End Sub

Public Sub ICleaner_additionalClean(ByVal fullFilePath As String, _
                                    ByVal cleaningType As String, _
                                    ByVal dataDate As Date, _
                                    ByVal dataMonth As Date, _
                                    ByVal dataMonthString As String, _
                                    ByVal xlApp As Excel.Application)
    'implement operations here
End Sub

' 範例：
' 假設 xlsht 是一個 Excel 工作表，以下範例將刪除 第 3 列：

' Dim xlsht As Worksheet
' Set xlsht = ThisWorkbook.Sheets("Sheet1")

' xlsht.Columns(3).Delete ' 刪除第 3 列
' 或刪除 B 列：

' xlsht.Columns("B").Delete ' 刪除 B 列
' 刪除多列：
' 一次刪除多列，例如刪除 B 到 D 列：

' xlsht.Columns("B:D").Delete
' 這樣就可以刪除指定的列了！ 🚀
