'HandleCloseRate
Option Compare Database

Private Sub Form_Load()
    Me.inputCloseRateFilePath = ""
    Me.inputCloseRateDate = Null

    ' Check Label Notification
    Me.lblCheckInputCloseRateLabel.Caption = "請輸入CloseRate日期" & vbCrLf & "(日期格式為 YYYY/MM/DD，例如: 2025/02/27)"
    Me.lblCheckInputFilePathLabel.Caption = "請輸入檔案路徑"
End Sub

Private Sub inputCloseRateDate_AfterUpdate()
    If IsDate(Me.inputCloseRateDate) Then
        Me.lblCheckInputCloseRateLabel.Caption = ""
    Else
        Me.lblCheckInputCloseRateLabel.Caption = "請輸入有效日期" & vbCrLf & "(日期格式為 YYYY/MM/DD，例如: 2025/02/27)"
        Me.inputCloseRateDate.SetFocus
    End If
End Sub

Private Sub inputCloseRateFilePath_AfterUpdate()
    If Me.inputCloseRateFilePath = "" Then
        Me.lblCheckInputFilePathLabel.Caption = "請輸入檔案路徑"
        Me.inputCloseRateFilePath.SetFocus
    Else
        If Dir(Me.inputCloseRateFilePath) = "" Then
            Me.lblCheckInputFilePathLabel.Caption = "存取路徑資料發生問題，請確認路徑是否正確"
            Me.inputCloseRateFilePath.SetFocus
        Else
            Me.lblCheckInputFilePathLabel.Caption = ""
        End If
    End If
End Sub

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

    'Fetch Inputs
    reportDataDate = Me.inputCloseRateDate
    fullFilePath = Me.inputCloseRateFilePath

    ' 檢查檔案是否存在
    If Dir(fullFilePath) = "" Then
        MsgBox "File does ont exist in path: " & fullFilePath
        WriteLog "File does not exist in path: " & fullFilePath
        Exit Sub
    End If

    ' 啟動 Excel 並開啟工作簿
    Set xlApp = Excel.Application
    Set xlbk = xlApp.Workbooks.Open(fullFilePath)
    Set xlsht = xlbk.Sheets("Sheet1")

    reportMonth = DateSerial(Year(reportDataDate), Month(reportDataDate) + 1, 0)
    reportMonthString = Format(reportMonth, "yyyy/mm")

    'Delete Columns
    xlsht.Columns("E").Delete
    xlsht.Columns("C").Delete
    xlsht.Columns("A:D").Insert shift:=xlToRight

    lastRow = xlsht.Cells(xlsht.Rows.Count, "E").End(xlUp).Row

    ' 反向遍歷刪除列
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

    MsgBox "完成清理 CloseRate，路徑為: " & fullFilePath
    WriteLog "完成清理 CloseRate，路徑為: " & fullFilePath

    Dim importer As IImporter
    Dim accessDBPath As String

    Set importer = New ImporterStandard

    accessDBPath = CurrentProject.FullName

    ' accessDBPath = "D:\DavidHsu\ReportCreator\closeRate\closeRate.accdb"
    ' accessDBPath = "D:\DavidHsu\ReportCreator\DB_MonthlyReport.accdb"
    ' fullFilePath = Me.inputCloseRateFilePath

    If importer Is Nothing Then
        MsgBox "找不到更新 CloseRate 對應的匯入程序"
        WriteLog "找不到更新 CloseRate 對應的匯入程序"
        Exit Sub
    End If

    importer.ImportData fullFilePath, accessDBPath, "CloseRate", xlApp

    xlApp.Quit
    Set xlApp = Nothing

End Sub


