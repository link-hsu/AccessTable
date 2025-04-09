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
