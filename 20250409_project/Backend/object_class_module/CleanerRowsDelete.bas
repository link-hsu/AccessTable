'CleanRowsDelete
Option Compare Database

Implements ICleaner

Private clsSheetName As Variant
Private clsLoopColumn As Integer
private clsLeftToDelete As Integer
private clsRightToDelete As Integer
Private clsRowsToDelete As Variant

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
End Sub

Public Sub ICleaner_CleanReport(ByVal fullFilePath As String, _
                                ByVal cleaningType As String)
    Dim xlApp As Excel.Application
    Dim xlbk As Excel.Workbook
    Dim xlsht As Excel.Worksheet
    Dim i As Long, lastRow As Long

    Dim isRowDelete As Boolean
    Dim rowToCheck As Variant

    If IsEmpty(clsRowsToDelete) Then clsRowsToDelete = Array()
    If IsEmpty(colsToDelete) Then colsToDelete = Array()
 
    'Check File Exist or not
    If Dir(fullFilePath) = "" Then
        MsgBox "File does not exist in path: " & fullFilePath
        Exit Sub
    End If

    ' 啟動 Excel 並開啟工作簿
    Set xlApp = Excel.Application
    Set xlbk = xlApp.Workbooks.Open(fullFilePath)
    Set xlsht = xlbk.Sheets(clsSheetName)

    ' 找出最後一列
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row

    ' 反向遍歷以刪除符合條件的列
    For i = lastRow To 2 Step -1
        isRowDelete = False
        For Each rowToCheck In clsRowsToDelete
            If xlsht.Cells(i, clsLoopColumn).Value = rowToCheck Or _
               IsEmpty(xlsht.Cells(i, clsLoopColumn).Value) Or _
               Left(xlsht.Cells(i, clsLoopColumn).Value, clsLeftToDelete) = rowToCheck Or _
               Right(xlsht.Cells(i, clsLoopColumn).Value, clsRightToDelete) = rowToCheck Then
                isRowDelete = True
                Exit For
            End If
        Next rowToCheck

        ' 刪除該列
        If isRowDelete Then xlsht.Rows(i).Delete
    Next i

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


' 使用範例：
' 1. 刪除 A 欄（第 1 欄）中的 "小計"、"總收入"、"總支出"、"純益"
' vba
' Copy
' Edit
' Call CleanReport("C:\Reports\example.xlsx", 1, "小計", "總收入", "總支出", "純益")
' 2. 刪除 B 欄（第 2 欄）中的 "無效資料" 和 "錯誤值"
' vba
' Copy
' Edit
' Call CleanReport("C:\Reports\example.xlsx", 2, "無效資料", "錯誤值")
' 這樣你就可以指定不同的欄位和不同的刪除值，使函式更靈活！ 🚀




