Option Compare Database

Implements ICleaner

Private clsColsToHandle As Variant

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
    clsColsToHandle = colsToHandle
End Sub

Public Sub ICleaner_CleanReport(ByVal fullFilePath As String, _
                                ByVal cleaningType As String, _
                                ByVal xlApp As Excel.Application)

    Dim xlbk As Excel.Workbook
    Dim xlsht As Excel.Worksheet
    Dim i As Long, lastRow As Long
    Dim colsDict As Object
    Dim colIndex As Variant

    If Dir(fullFilePath) = "" Then
        clsHasFile = False
        MsgBox "File not found: " & fullFilePath
        Exit Sub
    Else
        clsHasFile = True
    End If

    Set xlbk = xlApp.Workbooks.Open(fullFilePath)
    Set xlsht = xlbk.Sheets(1)

    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    clsHasData = (lastRow > 1)

    If Not (IsObject(clsColsToHandle) And TypeName(clsColsToHandle) = "Dictionary") Then
        GoTo CleanUp
    End If

    ' 因為Dict在Interface中使用底層有較嚴格的判斷類型，所以建議不使用
    Set colsDict = CreateObject("Scripting.Dictionary")
    Dim key As String
    Dim val As Variant
    Dim existingArray As Variant

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

    For i = lastRow To 2 Step -1
        If colsDict.Exists("FormatCols") Then
            For Each colIndex In colsDict("FormatCols")
                xlsht.Cells(i, colIndex).Value = Format(xlsht.Cells(i, colIndex).Value, "00000000")
            Next
        End If

        If colsDict.Exists("RemovePercentCols") Then
            For Each colIndex In colsDict("RemovePercentCols")
                xlsht.Cells(i, colIndex).Value = Replace(xlsht.Cells(i, colIndex).Value, "%", "")
            Next
        End If

        If colsDict.Exists("RemoveSpaceCols") Then
            For Each colIndex In colsDict("RemoveSpaceCols")
                xlsht.Cells(i, colIndex).Value = Replace(xlsht.Cells(i, colIndex).Value, " ", "")
            Next
        End If
    Next

    xlbk.Save
CleanUp:
    xlbk.Close False
    Set xlsht = Nothing
    Set xlbk = Nothing

    If clsHasFile And clsHasData Then
        MsgBox "完成清理：" & cleaningType & vbCrLf & fullFilePath
    End If
End Sub

' 空的 additionalClean（可留）
Public Sub ICleaner_additionalClean(ByVal fullFilePath As String, _
                                    ByVal cleaningType As String, _
                                    ByVal dataDate As Date, _
                                    ByVal dataMonth As Date, _
                                    ByVal dataMonthString As String, _
                                    ByVal xlApp As Excel.Application)
    ' No additional actions
End Sub



Public Function GetCleaner(ByVal cleaningType As String) As ICleaner

        Case "BondTransactionDetails"
            ' 統一編號要補零（F 欄）
            Set GetCleaner = New CleanerOMReport
            GetCleaner.Initialize , , , , , , Array("FormatCols", "F")
        Case "BillTransactionDetails"
            ' 統一編號要補零（F O P 欄）
            ' 票載利率要移除百分比( T 欄)
            ' 移除空白銀行名稱
            Set GetCleaner = New CleanerOMReport
            GetCleaner.Initialize , , , , , , Array("FormatCols", "F", "FormatCols", "O", "FormatCols", "P", "RemovePercentCols", "T", "RemoveSpaceCols", "")

        Case "BillHoldingDetails"
            ' 移除空白銀行名稱
            Set GetCleaner = New CleanerOMReport
            GetCleaner.Initialize , , , , , , Array("RemoveSpaceCols", "")
End Function




