Public Sub cleanCloseRate()
    Dim sheet As Worksheet
    Dim lastRow As Integer
    Dim columnArray As Variant
    Dim i As Integer
    Dim dataDate As Date, dataMonth As Date
    
    Set sheet = ThisWorkbook.Worksheets("CloseRate20230731")
    
    If VarType(sheet.Cells(1, 1).Value) = 7 Then
        dataDate = sheet.Cells(1, 1).Value
        MsgBox "dataDate is date"
    Else
        MsgBox "dataDate is not date"
    End If

    dataDate = DateSerial(2023, 7, 31)

    sheet.Columns("B:D").Insert shift:=xlToRight

    columnArray = Array("DataID", "DataDate", "DataMonth", "DataMonthString", "CurrencyType", "Rate")

    sheet.Range("A1").Resize(1, UBound(columnArray) - LBound(columnArray) + 1) = columnArray
    lastRow = sheet.Cells(sheet.Rows.Count, 1).End(xlUp).Row
    
    
    dataMonth = DateSerial(Year(dataDate), Month(dataDate) + 1, 0)
    Debug.Print dataDate
    Debug.Print dataMonth
    
    Debug.Print
    
    For i = lastRow To 2 Step -1
        sheet.Cells(i, 1).Clear
        sheet.Cells(i, 2).Value = dataDate
        sheet.Cells(i, 3).Value = dataMonth
        sheet.Cells(i, 4).Value = Format(dataMonth, "yyyy/mm")
    Next i
End Sub
