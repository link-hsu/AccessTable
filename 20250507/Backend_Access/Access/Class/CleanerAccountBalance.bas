Option Compare Database

Implements ICleaner

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
                               Optional ByVal colsToDelete As Variant, _
                               Optional ByVal colsToHandle As Variant)
End Sub

Public Sub ICleaner_CleanReport(ByVal fullFilePath As String, _
                                ByVal cleaningType As String, _
                                ByVal xlApp As Excel.Application)
    Dim xlbk As Excel.Workbook
    Dim xlsht As Excel.Worksheet
    Dim i As Integer
    
    Dim balanceSheets As Variant
    Dim sheetToDelete As Boolean

    Dim colsToDelete As Variant
    Dim col As Variant
    
    ' balanceSheets = Array("餘額A", "餘額C", "餘額D", "餘額E (2)")
    balanceSheets = Array("餘額A", "餘額C", "餘額D", "餘額E (2)")

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

    For Each xlsht In xlbk.sheets
        sheetToDelete = True
        For i = LBound(balanceSheets) To UBound(balanceSheets)
            If xlsht.Name = balanceSheets(i) Then
                sheetToDelete = False
                Exit For
            End If
        Next i
        
        If sheetToDelete Then
            xlsht.Delete
        Else
            lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row

            If clsHasData = False Then
                If (lastRow > 1) Then clsHasData = True
            End If

            If xlsht.Name = "餘額A" Or xlsht.Name = "餘額C" Or xlsht.Name = "餘額D" Then
                colsToDelete = Array("E", "D", "B")
            ElseIf xlsht.Name = "餘額E (2)" Then
                colsToDelete = Array("K", "I", "H", "E", "D", "B")
                xlsht.Range("A2:C" & lastRow).Value = xlsht.Range("A2:C" & lastRow).Value
            End if

            For Each col in colsToDelete
                xlsht.Columns(col).Delete
            Next col

            For i = lastRow to 1 Step -1
                xlsht.Cells(i, "C").Value = xlsht.Name
            Next i
        End If
    Next xlsht

    xlbk.Save
    xlbk.Close False
    Set xlsht = Nothing
    Set xlbk = Nothing

    ' On Error Resume Next
    ' Set xlsht = Nothing
    ' Set xlbk = Nothing
    ' Set securityRows = Nothing
    ' On Error GoTo 0

    If clsHasFile And clsHasData Then
        MsgBox "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
        WriteLog "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
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




' Sub AccountingDataProcessing()

'     Dim ws As Worksheet
'     Dim wsName As String
'     Dim keepSheets As Variant
'     Dim i As Long
'     Dim lastRow As Long
    
'     Application.ScreenUpdating = False
'     Application.DisplayAlerts = False
    
'     ' 要保留的分頁名稱
'     keepSheets = Array("餘額A", "餘額C", "餘額D", "餘額E")
    
'     ' 先刪除不需要的工作表
'     For Each ws In ThisWorkbook.Worksheets
'         wsName = ws.Name
'         If IsError(Application.Match(wsName, keepSheets, 0)) Then
'             ws.Delete
'             End If
'     Next ws
    
'     ' 處理保留的分頁
'     For Each ws In ThisWorkbook.Worksheets
'         ws.Activate
        
'         ' 刪除 Column D 和 Column B （要從後面開始刪，不然位置會變）
'         ws.Columns("D").Delete
'         ws.Columns("B").Delete
        
'         ' 把 Column C 移到 Column A
'         ws.Columns("C").Cut
'         ws.Columns("A").Insert Shift:=xlToRight
        
'         ' 新增三列 DataDate DataMonth DataMonthString
'         ws.Cells(1, ws.Columns.Count).End(xlToLeft).Offset(0, 1).Value = "DataDate"
'         ws.Cells(1, ws.Columns.Count).End(xlToLeft).Offset(0, 1).Value = "DataMonth"
'         ws.Cells(1, ws.Columns.Count).End(xlToLeft).Offset(0, 1).Value = "DataMonthString"
'     Next ws
    
'     ' 專門處理票券交易明細表
'     On Error Resume Next
'     Set ws = ThisWorkbook.Sheets("票券交易明細表")
'     On Error GoTo 0
    
'     If Not ws Is Nothing Then
'         ws.Activate
        
'         ' 找統一編號和票載利率欄位
'         Dim unifyCol As Long
'         Dim rateCol As Long
'         Dim headerRow As Long
        
'         headerRow = 1 ' 假設標題在第一列
'         lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
'         ' 找欄位位置
'         For i = 1 To ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
'             If ws.Cells(headerRow, i).Value = "統一編號" Then
'                 unifyCol = i
'             ElseIf ws.Cells(headerRow, i).Value = "票載利率" Then
'                 rateCol = i
'             End If
'         Next i
        
'         ' 處理統一編號補零到8位
'         If unifyCol > 0 Then
'             For i = headerRow + 1 To lastRow
'                 If Len(ws.Cells(i, unifyCol).Value) > 0 Then
'                     ws.Cells(i, unifyCol).Value = Format(ws.Cells(i, unifyCol).Value, "00000000")
'                 End If
'             Next i
'         End If
        
'         ' 處理票載利率，去掉百分比
'         If rateCol > 0 Then
'             For i = headerRow + 1 To lastRow
'                 If Len(ws.Cells(i, rateCol).Value) > 0 Then
'                     ws.Cells(i, rateCol).Value = Replace(ws.Cells(i, rateCol).Value, "%", "")
'                 End If
'             Next i
'         End If
'     Else
'         MsgBox "找不到 '票券交易明細表' 工作表！"
'     End If
    
'     Application.DisplayAlerts = True
'     Application.ScreenUpdating = True
    
'     MsgBox "處理完成！"

' End Sub
