'HandleCloseRate
Option Compare Database

Private Sub btnUploadCloseRate_Click()
    Dim xlApp As Excel.Application
    Dim xlbk As Excel.Workbook
    Dim xlsht As Excel.Worksheet
    Dim i As Long, lastRow As Long
    Dim fullFilePath As String
    Dim baseCurrency As String
    Dim isRowDelete As Boolean

    'Set Date Columns
    Dim reportDataDate As Date
    Dim reportMonth As Date
    Dim reportMonthString As String

    Dim tableColumns As Variant

    fullFilePath = "D:\DavidHsu\ReportCreator\RelationData\CopyData\1742956173629-cm2610.xls"

    ' 檢查檔案是否存在
    If Dir(fullFilePath) = "" Then
        MsgBox "File does not exist in path: " & fullFilePath
        Exit Sub
    End If

    ' 啟動 Excel 並開啟工作簿
    Set xlApp = Excel.Application
    Set xlbk = xlApp.Workbooks.Open(fullFilePath)
    Set xlsht = xlbk.Sheets("Sheet1")

    'Fetch Dates
    reportDataDate = Me.inputCloseRateDate
    reportMonth = DateSerial(Year(reportDataDate), Month(reportDataDate) + 1, 0)
    reportMonthString = Format(reportMonth, "yyyy/mm")

    'Delete Columns
    xlsht.Columns("E").Delete
    xlsht.Columns("C").Delete
    xlsht.Columns("A:D").Insert shift:=xlToRight

    lastRow = xlsht.Cells(xlsht.Rows.Count, "E").End(xlUp).Row

    ' 反向遍歷以刪除符合條件的列
    For i = lastRow To 2 Step -1
        isRowDelete = False
        xlsht.Cells(i, "B").Value = reportDataDate
        xlsht.Cells(i, "C").Value = reportMonth
        xlsht.Cells(i, "D").Value = reportMonthString
        If (IsEmpty(xlsht.Cells(i, "E").Value) And IsEmpty(xlsht.Cells(i, "F").Value) And IsEmpty(xlsht.Cells(i, "G").Value)) Or _
           Left(xlsht.Cells(i, "E").Value, 4) = "經副襄理" Then
            isRowDelete = True
        End If

        ' Delete Row
        If isRowDelete Then xlsht.Rows(i).Delete
    Next i

    'Set Table Columns
    tableColumns = GetTableColumns("CloseRate")
    xlsht.Range("A1").Resize(1, UBound(tableColumns) - LBound(tableColumns) + 1).Value = tableColumns
    xlsht.Columns("A").Delete

    lastRow = xlsht.Cells(xlsht.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow
        If xlsht.Cells(i, "D").Value <> "" Then
            baseCurrency = xlsht.Cells(i, "D").Value
        Else
            xlsht.Cells(i, "D").Value = baseCurrency
        End If
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

    Dim importer As IImporter
    Dim accessDBPath As String
    Dim fullFilePath As String

    Set importer = New ImporterStandard

    accessDBPath = "D:\DavidHsu\ReportCreator\closeRate\closeRate.accdb"
    ' accessDBPath = "D:\DavidHsu\ReportCreator\DB_MonthlyReport.accdb"
    ' fullFilePath = Me.inputCloseRateFilePath

    If importer Is Nothing Then
        MsgBox "找不到對應的匯入程序: " & cleaningType
        Exit Sub
    End If

    importer.ImportData fullFilePath, accessDBPath, "CloseRate"
End Sub


