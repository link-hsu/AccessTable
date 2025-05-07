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
    Dim ws As Excel.Worksheet
    Dim xlsht As Excel.Worksheet

    ' Dim xlbk As Workbook
    ' Dim ws As Worksheet
    ' Dim xlsht As Worksheet
    
    Dim copyRg As Range
    Dim Rngs As Range
    Dim oneRng As Range
    
    Dim outputArr() As Variant
    Dim fvArray As Variant
    Dim mapGroupMeasurement As Object
    Dim groupMeasurement As Variant

    Dim i As Integer, j As Integer, k As Integer
    Dim lastRow As Integer
    
    Dim securityRows As Collection
    Dim category As Variant
    Dim tableColumns As Variant
    
    ' 檢查檔案是否存在
    If Dir(fullFilePath) = "" Then
        clsHasFile = False
        MsgBox "File does not exist in path: " & fullFilePath
        WriteLog "File does not exist in path: " & fullFilePath
        Exit Sub
    Else
        clsHasFile = True
    End If

    Set xlbk = xlApp.Workbooks.Open(fullFilePath)
    ' Set xlbk = Workbooks.Open(fullFilePath)
    Set ws = xlbk.Worksheets("評估表")
    ws.Copy After:=xlbk.Sheets(xlbk.Sheets.count)
    ActiveSheet.Name = "評估表cp"
    Set xlsht = ActiveSheet
    
    Set copyRg = xlsht.UsedRange
    copyRg.Value = copyRg.Value
        
    Set Rngs = copyRg.Range("A5:T5")
    
'--------------------------
'    If ws.UsedRange.MergeCells Then
'        ws.UsedRange.UnMerge
'    End If
'--------------------------
    
    Dim columnsArray() As Variant
    Dim tempSplit As Variant
    Dim tempSave() As Variant
    Dim count As Long
    Dim splitCount As Long
    
    count = 0
    splitCount = 0
    
    For Each oneRng In Rngs
        tempSplit = Split(oneRng, vbLf)
        ReDim Preserve columnsArray(count)
        columnsArray(count) = Trim(tempSplit(0))
        count = count + 1
        
        If UBound(tempSplit) >= 1 Then
            ReDim Preserve tempSave(splitCount)
            tempSave(splitCount) = Trim(tempSplit(1))
            splitCount = splitCount + 1
        End If
    Next oneRng
    
    ReDim Preserve tempSave(splitCount)
    tempSave(splitCount) = "評價資產類別"
    
    For i = LBound(tempSave) To UBound(tempSave)
        ReDim Preserve columnsArray(count)
        columnsArray(count) = tempSave(i)
        count = count + 1
    Next i
    
    fvArray = Array("FVPL-公債", _
                    "FVPL-公司債(公營)", _
                    "FVPL-公司債(民營)", _
                    "FVPL-金融債", _
                    "FVOCI-公債", _
                    "FVOCI-公司債(公營)", _
                    "FVOCI-公司債(民營)", _
                    "FVOCI-金融債", _
                    "AC-公債", _
                    "AC-公司債(公營)", _
                    "AC-公司債(民營)", _
                    "AC-金融債")

    groupMeasurement = Array("FVPL_GovBond_Foreign", _
                             "FVPL_CompanyBond_Foreign", _
                             "FVPL_CompanyBond_Foreign", _
                             "FVPL_FinancialBond_Foreign", _
                             "FVOCI_GovBond_Foreign", _
                             "FVOCI_CompanyBond_Foreign", _
                             "FVOCI_CompanyBond_Foreign", _
                             "FVOCI_FinancialBond_Foreign", _
                             "AC_GovBond_Foreign", _
                             "AC_CompanyBond_Foreign", _
                             "AC_CompanyBond_Foreign", _
                             "AC_FinancialBond_Foreign")
    
    Set mapGroupMeasurement = CreateObject("Scripting.Dictionary")
    For i = LBound(fvArray) To UBound(fvArray)
        mapGroupMeasurement.Add fvArray(i), groupMeasurement(i)
    Next i

    lastRow = xlsht.Cells(xlsht.Rows.count, 1).End(xlUp).Row

    If (lastRow > 1) Then clsHasData = True
    
    For i = lastRow To 1 Step -1
        If IsEmpty(xlsht.Cells(i, 1).Value) Or _
            xlsht.Cells(i, 1).Value = "Security_Id" Then
                xlsht.Rows(i).Delete
        End If
        
        If Left(Trim(xlsht.Cells(i, 1).Value), 2) = "標註" Then
            xlsht.Rows(i & ":" & lastRow).Delete
        End If
    Next i

    lastRow = xlsht.Cells(xlsht.Rows.count, 1).End(xlUp).Row
    
    Set securityRows = New Collection

    For i = 1 To lastRow
        For Each category In fvArray
            If xlsht.Cells(i, 1).Value = category Then
                securityRows.Add i
            End If
        Next category
    Next i
    
    Dim startRow As Integer
    Dim endRow As Integer
    Dim numRows As Integer
    Dim numCols As Integer

    For i = 1 To securityRows.count
        If i = 1 Then
            ReDim outputArr(1 To lastRow, 1 To 32)
        End If

        If i + 1 <= securityRows.count Then
            If securityRows(i) + 1 = securityRows(i + 1) Then
                GoTo ContinueLoop
            Else
                startRow = securityRows(i) + 1
                endRow = securityRows(i + 1) - 1
            End If
        Else
            startRow = securityRows(i) + 1
            endRow = lastRow
        End If
        
        category = xlsht.Cells(startRow - 1, 1).Value

        For j = startRow To endRow Step 2
            For k = 1 To 40
                If k >= 1 And k <= 20 Then
                    outputArr(j, k) = xlsht.Cells(j, k).Value
                    If category = "AC-公債" Or category = "AC-公司債(公營)" Or _
                       category = "AC-公司債(民營)" Or category = "AC-金融債" Then
                       outputArr(j, 20) = xlsht.Cells(j, 17).Value
                    End If
                    outputArr(j, 17) = ""
                ElseIf k = 22 Then 'Issuer
                    outputArr(j, 21) = xlsht.Cells(j + 1, k - 20).Value
                ElseIf k = 28 Then 'Avg_Txnt_Rate
                    outputArr(j, 22) = xlsht.Cells(j + 1, k - 20).Value
                ElseIf k = 29 Then 'Avg_Buy_Price
                    outputArr(j, 23) = xlsht.Cells(j + 1, k - 20).Value
                ElseIf k = 30 Then 'Tot_Nominal_Amt_USD
                    outputArr(j, 24) = xlsht.Cells(j + 1, k - 20).Value
                ElseIf k = 31 Then 'Book_Value
                    outputArr(j, 25) = xlsht.Cells(j + 1, k - 20).Value
                ElseIf k = 32 Then 'PL_Amt_USD
                    outputArr(j, 26) = xlsht.Cells(j + 1, k - 20).Value
                ElseIf k = 33 Then 'Amortize_Amt
                    outputArr(j, 27) = xlsht.Cells(j + 1, k - 20).Value
                ElseIf k = 34 Then 'DVO1_USD
                    outputArr(j, 28) = xlsht.Cells(j + 1, k - 20).Value
                ElseIf k = 35 Then 'Interest_receivable_USD
                    outputArr(j, 29) = xlsht.Cells(j + 1, k - 20).Value
                ElseIf k = 36 Then '當日評等
                    outputArr(j, 30) = xlsht.Cells(j + 1, k - 20).Value
                ElseIf k = 37 Then '評價類別
                    outputArr(j, 31) = category
                End If
            Next k

            ' Add col 32 for groupMeasurement
            If mapGroupMeasurement.Exists(category) Then
                outputArr(j, 32) = mapGroupMeasurement(category)
            Else
                outputArr(j, 32) = ""
            End If
        Next j
ContinueLoop:
    Next i
    
    xlbk.Sheets.Add After:=xlbk.Sheets(xlbk.Sheets.count)
    ActiveSheet.Name = "OutputData"

    numRows = UBound(outputArr, 1)
    numCols = UBound(outputArr, 2)
    ActiveSheet.Range(ActiveSheet.Cells(2, 1), ActiveSheet.Cells(numRows + 1, numCols)).Value = outputArr
    
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.count, 1).End(xlUp).Row
    
    For i = lastRow To 2 Step -1
        If ActiveSheet.Cells(i, 1).Value = "" Then
        ActiveSheet.Rows(i).Delete
        End If
    Next i
    
    For i = 1 To (UBound(columnsArray) - LBound(columnsArray) + 1)
        ActiveSheet.Range("A1").Resize(1, (UBound(columnsArray) - LBound(columnsArray) + 1)).Value = columnsArray
    Next i
    

    For Each ws In xlbk.Sheets
        If Not ws Is ActiveSheet Then
            ws.Delete
        End If
    Next ws

    'Set Table Columns
    ' tableColumns = GetTableColumns(cleaningType)
    ' ActiveSheet.Range("A1").Resize(1, UBound(tableColumns) - LBound(tableColumns) + 1).Value = tableColumns

    Set securityRows = Nothing

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
