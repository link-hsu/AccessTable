Option Explicit

Public dataMonthString As String

Public Sub RunTotal()
    Dim isInputValid As Boolean

    isInputValid = False
    ' via InputBox for Users to enter yyyy/mm
    Do
        dataMonthString = InputBox("請輸入資料月份 (格式: yyyy/mm):", "輸入資料月份")

        If IsValidDataMonth(dataMonthString) Then
            isInputValid = True
        ElseIf Trim(dataMonthString) = "" Then
            MsgBox "請輸入需取得報表年度/月份" & vbCrLf & "(例如: 2024/01 )", vbExclamation, "Nul Error"
        Else
            MsgBox "輸入格式有誤，請輸入正確格式(yyyy/mm)" & vbCrLf & "(例如: 2024/01 )", vbExclamation, "Format Error"
        End If

    Loop Until isInputValid



    Call CNY1
End Sub



Public Sub CNY1()
    Dim controlSheet As Worksheet
    Dim xlsht As Worksheet

    Dim DBsPath As String
    Dim i As Integer, j As Integer
    Dim lastRow As Integer
    Dim dataArr As Variant

    Dim reportTitle As String
    Dim queryTable As String
    
    Dim rngs As Range
    Dim rng As Range

    Dim fxReceive As Double
    Dim fxPay As Double

    reportTitle = "CNY1"
    queryTable = "CNY1_DBU_AC5601"

    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")

    ' 取得 Access 資料庫路徑
    DBsPath = ThisWorkbook.Path & "\" & controlSheet.Range("DBsPathFileName")

    '***Set xlsht = ThisWorkbook.Sheets("CNY1")中的CNY1要用一個字串，當作參數傳入
    ' 設定目標 Excel 分頁
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    xlsht.Range("A:E").ClearContents

    dataArr = GetAccessDataAsArray(DBsPath, queryTable)
    
    For j = 0 To UBound(dataArr, 2)
        For i = 0 To UBound(dataArr, 1)
            xlsht.Cells(i + 1, j + 1).Value = dataArr(i, j)
        Next i
    Next j
    

    ' Compute 期收/付遠匯款-換匯遠期
    fxReceive = 0
    fxPay = 0
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    Set rngs = xlsht.Range("C3:C" & lastRow)

    For Each rng In rngs
        If rng.Value = "155930402" Then
            fxReceive = fxReceive + rng.Offset(0, 2).Value
        ElseIf rng.Value = "255930402" Then
            fxPay = fxPay + rng.Offset(0, 2).Value
        End If
    Next rng


    fxReceive = Round(fxReceive / 1000, 0)
    fxPay = Round(fxPay / 1000, 0)
    
    xlsht.Range("其他金融資產_淨額").Value = fxReceive
    xlsht.Range("其他").Value = fxReceive
    xlsht.Range("CNY1_資產總計").Value = fxReceive

    xlsht.Range("其他金融負債").Value = fxPay
    xlsht.Range("其他什項金融負債").Value = fxPay
    xlsht.Range("CNY1_負債總計").Value = fxPay
    
    'Set Number Format
    xlsht.Cells(2, "T").Resize(2, 1).NumberFormat = "#,##,00"
    
    MsgBox "Finish " & reportTitle & "分頁取得資料庫資料"
    
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他金融資產_淨額", xlsht.Range("其他金融資產_淨額").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他", xlsht.Range("其他").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "CNY1_資產總計", xlsht.Range("CNY1_資產總計").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他金融負債", xlsht.Range("其他金融負債").Value, "期付遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他什項金融負債", xlsht.Range("其他什項金融負債").Value, "期付遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "CNY1_負債總計", xlsht.Range("CNY1_負債總計").Value, "期付遠匯款-換匯遠期"
    MsgBox "Finish " & reportTitle & "申報資料寫入資料庫"


    ' Release Resources
    Set rng = Nothing
    Set rngs = Nothing
    Set xlsht = Nothing
    Set controlSheet = Nothing
    
End Sub

Public Sub FB2()
    Dim controlSheet As Worksheet
    Dim xlsht As Worksheet

    Dim DBsPath As String
    Dim i As Integer, j As Integer
    Dim lastRow As Integer
    Dim dataArr As Variant

    Dim reportTitle As String
    Dim queryTable As String
    
    Dim rngs As Range
    Dim rng As Range

    Dim fxReceive As Double
    Dim fxPay As Double

    reportTitle = "FB2"
    queryTable = "FB2_OBU_AC4620B"

    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")

    ' 取得 Access 資料庫路徑
    DBsPath = ThisWorkbook.Path & "\" & controlSheet.Range("DBsPathFileName")

    '***Set xlsht = ThisWorkbook.Sheets("CNY1")中的CNY1要用一個字串，當作參數傳入
    ' 設定目標 Excel 分頁
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    xlsht.Range("A:F").ClearContents

    dataArr = GetAccessDataAsArray(DBsPath, queryTable)
    
    For j = 0 To UBound(dataArr, 2)
        For i = 0 To UBound(dataArr, 1)
            xlsht.Cells(i + 1, j + 1).Value = dataArr(i, j)
        Next i
    Next j
    

    ' ' Compute 期收/付遠匯款-換匯遠期
    ' fxReceive = 0
    ' fxPay = 0
    ' lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    ' Set rngs = xlsht.Range("C3:C" & lastRow)

    ' For Each rng In rngs
    '     If rng.Value = "155930402" Then
    '         fxReceive = fxReceive + rng.Offset(0, 2).Value
    '     ElseIf rng.Value = "255930402" Then
    '         fxPay = fxPay + rng.Offset(0, 2).Value
    '     End If
    ' Next rng


    ' fxReceive = Round(fxReceive / 1000, 0)
    ' fxPay = Round(fxPay / 1000, 0)
    
    ' xlsht.Range("其他金融資產_淨額").Value = fxReceive
    ' xlsht.Range("其他").Value = fxReceive
    ' xlsht.Range("CNY1_資產總計").Value = fxReceive

    ' xlsht.Range("其他金融負債").Value = fxPay
    ' xlsht.Range("其他什項金融負債").Value = fxPay
    ' xlsht.Range("CNY1_負債總計").Value = fxPay
    
    ' 'Set Number Format
    ' xlsht.Cells(2, "T").Resize(2, 1).NumberFormat = "#,##,00"
    
    ' MsgBox "Finish " & reportTitle & "分頁取得資料庫資料"
    
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他金融資產_淨額", xlsht.Range("其他金融資產_淨額").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他", xlsht.Range("其他").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "CNY1_資產總計", xlsht.Range("CNY1_資產總計").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他金融負債", xlsht.Range("其他金融負債").Value, "期付遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他什項金融負債", xlsht.Range("其他什項金融負債").Value, "期付遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "CNY1_負債總計", xlsht.Range("CNY1_負債總計").Value, "期付遠匯款-換匯遠期"
    MsgBox "Finish " & reportTitle & "申報資料寫入資料庫"

    ' Release Resources
    Set rng = Nothing
    Set rngs = Nothing
    Set xlsht = Nothing
    Set controlSheet = Nothing
    
End Sub

Public Sub FB3()
    Dim controlSheet As Worksheet
    Dim xlsht As Worksheet

    Dim DBsPath As String
    Dim i As Integer, j As Integer
    Dim lastRow As Integer
    Dim dataArr_1 As Variant
    Dim dataArr_2 As Variant

    Dim reportTitle As String
    Dim queryTable_1 As String
    Dim queryTable_2 As String

    Dim rngs As Range
    Dim rng As Range

    Dim fxReceive As Double
    Dim fxPay As Double

    reportTitle = "FB3"
    queryTable_1 = "FB3_OBU_MM4901B_LIST"
    queryTable_2 = "FB3_OBU_MM4901B_SUM"

    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")

    ' 取得 Access 資料庫路徑
    DBsPath = ThisWorkbook.Path & "\" & controlSheet.Range("DBsPathFileName")

    '***Set xlsht = ThisWorkbook.Sheets("CNY1")中的CNY1要用一個字串，當作參數傳入
    ' 設定目標 Excel 分頁
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    xlsht.Range("A:K").ClearContents

    dataArr_1 = GetAccessDataAsArray(DBsPath, queryTable_1)
    dataArr_2 = GetAccessDataAsArray(DBsPath, queryTable_2)
    
    For j = 0 To UBound(dataArr_1, 2)
        For i = 0 To UBound(dataArr_1, 1)
            xlsht.Cells(i + 1, j + 1).Value = dataArr_1(i, j)
        Next i
    Next j

    For j = 0 To UBound(dataArr_2, 2)
        For i = 0 To UBound(dataArr_2, 1)
            xlsht.Cells(i + 1, j + 10).Value = dataArr_2(i, j)
        Next i
    Next j
    

    ' Compute 期收/付遠匯款-換匯遠期
    ' fxReceive = 0
    ' fxPay = 0
    ' lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    ' Set rngs = xlsht.Range("C3:C" & lastRow)

    ' For Each rng In rngs
    '     If rng.Value = "155930402" Then
    '         fxReceive = fxReceive + rng.Offset(0, 2).Value
    '     ElseIf rng.Value = "255930402" Then
    '         fxPay = fxPay + rng.Offset(0, 2).Value
    '     End If
    ' Next rng


    ' fxReceive = Round(fxReceive / 1000, 0)
    ' fxPay = Round(fxPay / 1000, 0)
    
    ' xlsht.Range("其他金融資產_淨額").Value = fxReceive
    ' xlsht.Range("其他").Value = fxReceive
    ' xlsht.Range("CNY1_資產總計").Value = fxReceive

    ' xlsht.Range("其他金融負債").Value = fxPay
    ' xlsht.Range("其他什項金融負債").Value = fxPay
    ' xlsht.Range("CNY1_負債總計").Value = fxPay
    
    ' 'Set Number Format
    ' xlsht.Cells(2, "T").Resize(2, 1).NumberFormat = "#,##,00"
    
    ' MsgBox "Finish " & reportTitle & "分頁取得資料庫資料"
    
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他金融資產_淨額", xlsht.Range("其他金融資產_淨額").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他", xlsht.Range("其他").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "CNY1_資產總計", xlsht.Range("CNY1_資產總計").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他金融負債", xlsht.Range("其他金融負債").Value, "期付遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他什項金融負債", xlsht.Range("其他什項金融負債").Value, "期付遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "CNY1_負債總計", xlsht.Range("CNY1_負債總計").Value, "期付遠匯款-換匯遠期"
    MsgBox "Finish " & reportTitle & "申報資料寫入資料庫"


    ' Release Resources
    Set rng = Nothing
    Set rngs = Nothing
    Set xlsht = Nothing
    Set controlSheet = Nothing
    
End Sub

Public Sub FB3A()
    Dim controlSheet As Worksheet
    Dim xlsht As Worksheet

    Dim DBsPath As String
    Dim i As Integer, j As Integer
    Dim lastRow As Integer
    Dim dataArr As Variant

    Dim reportTitle As String
    Dim queryTable As String
    
    Dim rngs As Range
    Dim rng As Range

    Dim fxReceive As Double
    Dim fxPay As Double

    reportTitle = "FB3A"
    queryTable = "FB3A_OBU_MM4901B"

    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")

    ' 取得 Access 資料庫路徑
    DBsPath = ThisWorkbook.Path & "\" & controlSheet.Range("DBsPathFileName")

    '***Set xlsht = ThisWorkbook.Sheets("CNY1")中的CNY1要用一個字串，當作參數傳入
    ' 設定目標 Excel 分頁
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    xlsht.Range("A:J").ClearContents

    dataArr = GetAccessDataAsArray(DBsPath, queryTable)
    
    For j = 0 To UBound(dataArr, 2)
        For i = 0 To UBound(dataArr, 1)
            xlsht.Cells(i + 1, j + 1).Value = dataArr(i, j)
        Next i
    Next j
    

    ' Compute 期收/付遠匯款-換匯遠期
    ' fxReceive = 0
    ' fxPay = 0
    ' lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    ' Set rngs = xlsht.Range("C3:C" & lastRow)

    ' For Each rng In rngs
    '     If rng.Value = "155930402" Then
    '         fxReceive = fxReceive + rng.Offset(0, 2).Value
    '     ElseIf rng.Value = "255930402" Then
    '         fxPay = fxPay + rng.Offset(0, 2).Value
    '     End If
    ' Next rng


    ' fxReceive = Round(fxReceive / 1000, 0)
    ' fxPay = Round(fxPay / 1000, 0)
    
    ' xlsht.Range("其他金融資產_淨額").Value = fxReceive
    ' xlsht.Range("其他").Value = fxReceive
    ' xlsht.Range("CNY1_資產總計").Value = fxReceive

    ' xlsht.Range("其他金融負債").Value = fxPay
    ' xlsht.Range("其他什項金融負債").Value = fxPay
    ' xlsht.Range("CNY1_負債總計").Value = fxPay
    
    ' 'Set Number Format
    ' xlsht.Cells(2, "T").Resize(2, 1).NumberFormat = "#,##,00"
    
    ' MsgBox "Finish " & reportTitle & "分頁取得資料庫資料"
    
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他金融資產_淨額", xlsht.Range("其他金融資產_淨額").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他", xlsht.Range("其他").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "CNY1_資產總計", xlsht.Range("CNY1_資產總計").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他金融負債", xlsht.Range("其他金融負債").Value, "期付遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他什項金融負債", xlsht.Range("其他什項金融負債").Value, "期付遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "CNY1_負債總計", xlsht.Range("CNY1_負債總計").Value, "期付遠匯款-換匯遠期"
    MsgBox "Finish " & reportTitle & "申報資料寫入資料庫"


    ' Release Resources
    Set rng = Nothing
    Set rngs = Nothing
    Set xlsht = Nothing
    Set controlSheet = Nothing
    
End Sub


'尚無有交易紀錄
Public Sub FM5()
    Dim controlSheet As Worksheet
    Dim xlsht As Worksheet

    Dim DBsPath As String
    Dim i As Integer, j As Integer
    Dim lastRow As Integer
    Dim dataArr As Variant

    Dim reportTitle As String
    Dim queryTable As String
    
    Dim rngs As Range
    Dim rng As Range

    Dim fxReceive As Double
    Dim fxPay As Double

    reportTitle = "FM5"
    queryTable = "FM5_OBU_FC9450B"

    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")

    ' 取得 Access 資料庫路徑
    DBsPath = ThisWorkbook.Path & "\" & controlSheet.Range("DBsPathFileName")

    '***Set xlsht = ThisWorkbook.Sheets("CNY1")中的CNY1要用一個字串，當作參數傳入
    ' 設定目標 Excel 分頁
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    xlsht.Range("A:E").ClearContents

    dataArr = GetAccessDataAsArray(DBsPath, queryTable)
    
    For j = 0 To UBound(dataArr, 2)
        For i = 0 To UBound(dataArr, 1)
            xlsht.Cells(i + 1, j + 1).Value = dataArr(i, j)
        Next i
    Next j
    

    ' Compute 期收/付遠匯款-換匯遠期
    ' fxReceive = 0
    ' fxPay = 0
    ' lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    ' Set rngs = xlsht.Range("C3:C" & lastRow)

    ' For Each rng In rngs
    '     If rng.Value = "155930402" Then
    '         fxReceive = fxReceive + rng.Offset(0, 2).Value
    '     ElseIf rng.Value = "255930402" Then
    '         fxPay = fxPay + rng.Offset(0, 2).Value
    '     End If
    ' Next rng


    ' fxReceive = Round(fxReceive / 1000, 0)
    ' fxPay = Round(fxPay / 1000, 0)
    
    ' xlsht.Range("其他金融資產_淨額").Value = fxReceive
    ' xlsht.Range("其他").Value = fxReceive
    ' xlsht.Range("CNY1_資產總計").Value = fxReceive

    ' xlsht.Range("其他金融負債").Value = fxPay
    ' xlsht.Range("其他什項金融負債").Value = fxPay
    ' xlsht.Range("CNY1_負債總計").Value = fxPay
    
    ' 'Set Number Format
    ' xlsht.Cells(2, "T").Resize(2, 1).NumberFormat = "#,##,00"
    
    ' MsgBox "Finish " & reportTitle & "分頁取得資料庫資料"
    
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他金融資產_淨額", xlsht.Range("其他金融資產_淨額").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他", xlsht.Range("其他").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "CNY1_資產總計", xlsht.Range("CNY1_資產總計").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他金融負債", xlsht.Range("其他金融負債").Value, "期付遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他什項金融負債", xlsht.Range("其他什項金融負債").Value, "期付遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "CNY1_負債總計", xlsht.Range("CNY1_負債總計").Value, "期付遠匯款-換匯遠期"
    MsgBox "Finish " & reportTitle & "申報資料寫入資料庫"


    ' Release Resources
    Set rng = Nothing
    Set rngs = Nothing
    Set xlsht = Nothing
    Set controlSheet = Nothing
    
End Sub

Public Sub FM11()
    Dim controlSheet As Worksheet
    Dim xlsht As Worksheet

    Dim DBsPath As String
    Dim i As Integer, j As Integer
    Dim lastRow As Integer
    Dim dataArr As Variant

    Dim reportTitle As String
    Dim queryTable As String
    
    Dim rngs As Range
    Dim rng As Range

    Dim fxReceive As Double
    Dim fxPay As Double

    reportTitle = "FM11"
    queryTable = "FM11_OBU_AC5411B"

    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")

    ' 取得 Access 資料庫路徑
    DBsPath = ThisWorkbook.Path & "\" & controlSheet.Range("DBsPathFileName")

    '***Set xlsht = ThisWorkbook.Sheets("CNY1")中的CNY1要用一個字串，當作參數傳入
    ' 設定目標 Excel 分頁
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    xlsht.Range("A:E").ClearContents

    dataArr = GetAccessDataAsArray(DBsPath, queryTable)
    
    For j = 0 To UBound(dataArr, 2)
        For i = 0 To UBound(dataArr, 1)
            xlsht.Cells(i + 1, j + 1).Value = dataArr(i, j)
        Next i
    Next j
    

    ' ' Compute 期收/付遠匯款-換匯遠期
    ' fxReceive = 0
    ' fxPay = 0
    ' lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    ' Set rngs = xlsht.Range("C3:C" & lastRow)

    ' For Each rng In rngs
    '     If rng.Value = "155930402" Then
    '         fxReceive = fxReceive + rng.Offset(0, 2).Value
    '     ElseIf rng.Value = "255930402" Then
    '         fxPay = fxPay + rng.Offset(0, 2).Value
    '     End If
    ' Next rng


    ' fxReceive = Round(fxReceive / 1000, 0)
    ' fxPay = Round(fxPay / 1000, 0)
    
    ' xlsht.Range("其他金融資產_淨額").Value = fxReceive
    ' xlsht.Range("其他").Value = fxReceive
    ' xlsht.Range("CNY1_資產總計").Value = fxReceive

    ' xlsht.Range("其他金融負債").Value = fxPay
    ' xlsht.Range("其他什項金融負債").Value = fxPay
    ' xlsht.Range("CNY1_負債總計").Value = fxPay
    
    ' 'Set Number Format
    ' xlsht.Cells(2, "T").Resize(2, 1).NumberFormat = "#,##,00"
    
    ' MsgBox "Finish " & reportTitle & "分頁取得資料庫資料"
    
    ' InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他金融資產_淨額", xlsht.Range("其他金融資產_淨額").Value, "期收遠匯款-換匯遠期"
    ' InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他", xlsht.Range("其他").Value, "期收遠匯款-換匯遠期"
    ' InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "CNY1_資產總計", xlsht.Range("CNY1_資產總計").Value, "期收遠匯款-換匯遠期"
    ' InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他金融負債", xlsht.Range("其他金融負債").Value, "期付遠匯款-換匯遠期"
    ' InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他什項金融負債", xlsht.Range("其他什項金融負債").Value, "期付遠匯款-換匯遠期"
    ' InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "CNY1_負債總計", xlsht.Range("CNY1_負債總計").Value, "期付遠匯款-換匯遠期"
    ' MsgBox "Finish " & reportTitle & "申報資料寫入資料庫"


    ' Release Resources
    Set rng = Nothing
    Set rngs = Nothing
    Set xlsht = Nothing
    Set controlSheet = Nothing
    
End Sub

Public Sub FM13()
    Dim controlSheet As Worksheet
    Dim xlsht As Worksheet

    Dim DBsPath As String
    Dim i As Integer, j As Integer
    Dim lastRow As Integer
    Dim dataArr_1 As Variant
    Dim dataArr_2 As Variant

    Dim reportTitle As String
    Dim queryTable_1 As String
    Dim queryTable_2 As String
    
    Dim rngs As Range
    Dim rng As Range

    Dim fxReceive As Double
    Dim fxPay As Double

    reportTitle = "FM13"
    queryTable_1 = "FM13_FXDebtEvaluation_Subtotal_FVandAdjust"
    queryTable_2 = "FM13_FXDebtEvaluation_Subtotal_Impairment"

    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")

    ' 取得 Access 資料庫路徑
    DBsPath = ThisWorkbook.Path & "\" & controlSheet.Range("DBsPathFileName")

    '***Set xlsht = ThisWorkbook.Sheets("CNY1")中的CNY1要用一個字串，當作參數傳入
    ' 設定目標 Excel 分頁
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    xlsht.Range("A:E").ClearContents

    dataArr_1 = GetAccessDataAsArray(DBsPath, queryTable_1)
    dataArr_2 = GetAccessDataAsArray(DBsPath, queryTable_2)
    
    For j = 0 To UBound(dataArr_1, 2)
        For i = 0 To UBound(dataArr_1, 1)
            xlsht.Cells(i + 1, j + 1).Value = dataArr_1(i, j)
        Next i
    Next j

    For j = 0 To UBound(dataArr_2, 2)
        For i = 0 To UBound(dataArr_2, 1)
            xlsht.Cells(i + 1, j + 4).Value = dataArr_2(i, j)
        Next i
    Next j
    
    ' Compute 期收/付遠匯款-換匯遠期
    ' fxReceive = 0
    ' fxPay = 0
    ' lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    ' Set rngs = xlsht.Range("C3:C" & lastRow)

    ' For Each rng In rngs
    '     If rng.Value = "155930402" Then
    '         fxReceive = fxReceive + rng.Offset(0, 2).Value
    '     ElseIf rng.Value = "255930402" Then
    '         fxPay = fxPay + rng.Offset(0, 2).Value
    '     End If
    ' Next rng


    ' fxReceive = Round(fxReceive / 1000, 0)
    ' fxPay = Round(fxPay / 1000, 0)
    
    ' xlsht.Range("其他金融資產_淨額").Value = fxReceive
    ' xlsht.Range("其他").Value = fxReceive
    ' xlsht.Range("CNY1_資產總計").Value = fxReceive

    ' xlsht.Range("其他金融負債").Value = fxPay
    ' xlsht.Range("其他什項金融負債").Value = fxPay
    ' xlsht.Range("CNY1_負債總計").Value = fxPay
    
    ' 'Set Number Format
    ' xlsht.Cells(2, "T").Resize(2, 1).NumberFormat = "#,##,00"
    
    ' MsgBox "Finish " & reportTitle & "分頁取得資料庫資料"
    
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他金融資產_淨額", xlsht.Range("其他金融資產_淨額").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他", xlsht.Range("其他").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "CNY1_資產總計", xlsht.Range("CNY1_資產總計").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他金融負債", xlsht.Range("其他金融負債").Value, "期付遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他什項金融負債", xlsht.Range("其他什項金融負債").Value, "期付遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "CNY1_負債總計", xlsht.Range("CNY1_負債總計").Value, "期付遠匯款-換匯遠期"
    MsgBox "Finish " & reportTitle & "申報資料寫入資料庫"


    ' Release Resources
    Set rng = Nothing
    Set rngs = Nothing
    Set xlsht = Nothing
    Set controlSheet = Nothing
    
End Sub

Public Sub AI821()
    Dim controlSheet As Worksheet
    Dim xlsht As Worksheet

    Dim DBsPath As String
    Dim i As Integer, j As Integer
    Dim lastRow As Integer
    Dim dataArr_1 As Variant
    Dim dataArr_2 As Variant

    Dim reportTitle As String
    Dim queryTable_1 As String
    Dim queryTable_2 As String
    
    Dim rngs As Range
    Dim rng As Range

    Dim fxReceive As Double
    Dim fxPay As Double

    reportTitle = "AI821"
    queryTable_1 = "AI821_OBU_MM4901B_LIST"
    queryTable_2 = "AI821_OBU_MM4901B_SUM"

    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")

    ' 取得 Access 資料庫路徑
    DBsPath = ThisWorkbook.Path & "\" & controlSheet.Range("DBsPathFileName")

    '***Set xlsht = ThisWorkbook.Sheets("CNY1")中的CNY1要用一個字串，當作參數傳入
    ' 設定目標 Excel 分頁
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    xlsht.Range("A:K").ClearContents

    dataArr_1 = GetAccessDataAsArray(DBsPath, queryTable_1)
    dataArr_2 = GetAccessDataAsArray(DBsPath, queryTable_2)
    
    For j = 0 To UBound(dataArr_1, 2)
        For i = 0 To UBound(dataArr_1, 1)
            xlsht.Cells(i + 1, j + 1).Value = dataArr_1(i, j)
        Next i
    Next j

    For j = 0 To UBound(dataArr_2, 2)
        For i = 0 To UBound(dataArr_2, 1)
            xlsht.Cells(i + 1, j + 9).Value = dataArr_2(i, j)
        Next i
    Next j
    

    ' Compute 期收/付遠匯款-換匯遠期
    ' fxReceive = 0
    ' fxPay = 0
    ' lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    ' Set rngs = xlsht.Range("C3:C" & lastRow)

    ' For Each rng In rngs
    '     If rng.Value = "155930402" Then
    '         fxReceive = fxReceive + rng.Offset(0, 2).Value
    '     ElseIf rng.Value = "255930402" Then
    '         fxPay = fxPay + rng.Offset(0, 2).Value
    '     End If
    ' Next rng


    ' fxReceive = Round(fxReceive / 1000, 0)
    ' fxPay = Round(fxPay / 1000, 0)
    
    ' xlsht.Range("其他金融資產_淨額").Value = fxReceive
    ' xlsht.Range("其他").Value = fxReceive
    ' xlsht.Range("CNY1_資產總計").Value = fxReceive

    ' xlsht.Range("其他金融負債").Value = fxPay
    ' xlsht.Range("其他什項金融負債").Value = fxPay
    ' xlsht.Range("CNY1_負債總計").Value = fxPay
    
    ' 'Set Number Format
    ' xlsht.Cells(2, "T").Resize(2, 1).NumberFormat = "#,##,00"
    
    ' MsgBox "Finish " & reportTitle & "分頁取得資料庫資料"
    
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他金融資產_淨額", xlsht.Range("其他金融資產_淨額").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他", xlsht.Range("其他").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "CNY1_資產總計", xlsht.Range("CNY1_資產總計").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他金融負債", xlsht.Range("其他金融負債").Value, "期付遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他什項金融負債", xlsht.Range("其他什項金融負債").Value, "期付遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "CNY1_負債總計", xlsht.Range("CNY1_負債總計").Value, "期付遠匯款-換匯遠期"
    MsgBox "Finish " & reportTitle & "申報資料寫入資料庫"


    ' Release Resources
    Set rng = Nothing
    Set rngs = Nothing
    Set xlsht = Nothing
    Set controlSheet = Nothing
    
End Sub

Public Sub Table2()
    Dim controlSheet As Worksheet
    Dim xlsht As Worksheet

    Dim DBsPath As String
    Dim i As Integer, j As Integer
    Dim lastRow As Integer
    Dim dataArr_1 As Variant
    Dim dataArr_2 As Variant
    
    Dim reportTitle As String
    Dim queryTable_1 As String
    Dim queryTable_2 As String
    
    Dim rngs As Range
    Dim rng As Range

    Dim fxReceive As Double
    Dim fxPay As Double

    reportTitle = "Table2"
    queryTable_1 = "表2_DBU_AC5602_TWD"
    queryTable_2 = "表2_CloseRate_USDTWD"

    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")

    ' 取得 Access 資料庫路徑
    DBsPath = ThisWorkbook.Path & "\" & controlSheet.Range("DBsPathFileName")

    '***Set xlsht = ThisWorkbook.Sheets("CNY1")中的CNY1要用一個字串，當作參數傳入
    ' 設定目標 Excel 分頁
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    xlsht.Range("A:I").ClearContents

    dataArr_1 = GetAccessDataAsArray(DBsPath, queryTable_1)
    dataArr_2 = GetAccessDataAsArray(DBsPath, queryTable_2)
    
    If Err.Number <> 0 Or LBound(dataArr_1) > UBound(dataArr_1) Then
        MsgBox queryTable_1 & "資料表無資料"
    Else
        For j = 0 To UBound(dataArr_1, 2)
            For i = 0 To UBound(dataArr_1, 1)
                xlsht.Cells(i + 1, j + 1).Value = dataArr_1(i, j)
            Next i
        Next j
    End If

    If Err.Number <> 0 Or LBound(dataArr_2) > UBound(dataArr_2) Then
        MsgBox queryTable_2 & "資料表無資料"
    Else
        For j = 0 To UBound(dataArr_2, 2)
            For i = 0 To UBound(dataArr_2, 1)
                xlsht.Cells(i + 1, j + 7).Value = dataArr_2(i, j)
            Next i
        Next j
    End If

    ' Compute 期收/付遠匯款-換匯遠期
    ' fxReceive = 0
    ' fxPay = 0
    ' lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    ' Set rngs = xlsht.Range("C3:C" & lastRow)

    ' For Each rng In rngs
    '     If rng.Value = "155930402" Then
    '         fxReceive = fxReceive + rng.Offset(0, 2).Value
    '     ElseIf rng.Value = "255930402" Then
    '         fxPay = fxPay + rng.Offset(0, 2).Value
    '     End If
    ' Next rng


    ' fxReceive = Round(fxReceive / 1000, 0)
    ' fxPay = Round(fxPay / 1000, 0)
    
    ' xlsht.Range("其他金融資產_淨額").Value = fxReceive
    ' xlsht.Range("其他").Value = fxReceive
    ' xlsht.Range("CNY1_資產總計").Value = fxReceive

    ' xlsht.Range("其他金融負債").Value = fxPay
    ' xlsht.Range("其他什項金融負債").Value = fxPay
    ' xlsht.Range("CNY1_負債總計").Value = fxPay
    
    ' 'Set Number Format
    ' xlsht.Cells(2, "T").Resize(2, 1).NumberFormat = "#,##,00"
    
    ' MsgBox "Finish " & reportTitle & "分頁取得資料庫資料"
    
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他金融資產_淨額", xlsht.Range("其他金融資產_淨額").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他", xlsht.Range("其他").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "CNY1_資產總計", xlsht.Range("CNY1_資產總計").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他金融負債", xlsht.Range("其他金融負債").Value, "期付遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他什項金融負債", xlsht.Range("其他什項金融負債").Value, "期付遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "CNY1_負債總計", xlsht.Range("CNY1_負債總計").Value, "期付遠匯款-換匯遠期"
    MsgBox "Finish " & reportTitle & "申報資料寫入資料庫"


    ' Release Resources
    Set rng = Nothing
    Set rngs = Nothing
    Set xlsht = Nothing
    Set controlSheet = Nothing
    
End Sub

Public Sub FB5_FB5A()
    Dim controlSheet As Worksheet
    Dim xlsht As Worksheet

    Dim DBsPath As String
    Dim i As Integer, j As Integer
    Dim lastRow As Integer
    Dim dataArr As Variant

    Dim reportTitle As String
    Dim queryTable As String
    
    Dim rngs As Range
    Dim rng As Range

    Dim fxReceive As Double
    Dim fxPay As Double

    reportTitle = "FB5_FB5A"
    queryTable = "FB5_FB5A_DL6320"

    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")

    ' 取得 Access 資料庫路徑
    DBsPath = ThisWorkbook.Path & "\" & controlSheet.Range("DBsPathFileName")

    '***Set xlsht = ThisWorkbook.Sheets("CNY1")中的CNY1要用一個字串，當作參數傳入
    ' 設定目標 Excel 分頁
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    xlsht.Range("A:G").ClearContents

    dataArr = GetAccessDataAsArray(DBsPath, queryTable)
    
    For j = 0 To UBound(dataArr, 2)
        For i = 0 To UBound(dataArr, 1)
            xlsht.Cells(i + 1, j + 1).Value = dataArr(i, j)
        Next i
    Next j
    

    ' Compute 期收/付遠匯款-換匯遠期
    ' fxReceive = 0
    ' fxPay = 0
    ' lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    ' Set rngs = xlsht.Range("C3:C" & lastRow)

    ' For Each rng In rngs
    '     If rng.Value = "155930402" Then
    '         fxReceive = fxReceive + rng.Offset(0, 2).Value
    '     ElseIf rng.Value = "255930402" Then
    '         fxPay = fxPay + rng.Offset(0, 2).Value
    '     End If
    ' Next rng


    ' fxReceive = Round(fxReceive / 1000, 0)
    ' fxPay = Round(fxPay / 1000, 0)
    
    ' xlsht.Range("其他金融資產_淨額").Value = fxReceive
    ' xlsht.Range("其他").Value = fxReceive
    ' xlsht.Range("CNY1_資產總計").Value = fxReceive

    ' xlsht.Range("其他金融負債").Value = fxPay
    ' xlsht.Range("其他什項金融負債").Value = fxPay
    ' xlsht.Range("CNY1_負債總計").Value = fxPay
    
    ' 'Set Number Format
    ' xlsht.Cells(2, "T").Resize(2, 1).NumberFormat = "#,##,00"
    
    ' MsgBox "Finish " & reportTitle & "分頁取得資料庫資料"
    
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他金融資產_淨額", xlsht.Range("其他金融資產_淨額").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他", xlsht.Range("其他").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "CNY1_資產總計", xlsht.Range("CNY1_資產總計").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他金融負債", xlsht.Range("其他金融負債").Value, "期付遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他什項金融負債", xlsht.Range("其他什項金融負債").Value, "期付遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "CNY1_負債總計", xlsht.Range("CNY1_負債總計").Value, "期付遠匯款-換匯遠期"
    MsgBox "Finish " & reportTitle & "申報資料寫入資料庫"


    ' Release Resources
    Set rng = Nothing
    Set rngs = Nothing
    Set xlsht = Nothing
    Set controlSheet = Nothing
    
End Sub

Public Sub FM2()
    Dim controlSheet As Worksheet
    Dim xlsht As Worksheet

    Dim DBsPath As String
    Dim i As Integer, j As Integer
    Dim lastRow As Integer
    Dim dataArr_1 As Variant
    Dim dataArr_2 As Variant
    Dim dataArr_3 As Variant

    Dim reportTitle As String
    Dim queryTable_1 As String
    Dim queryTable_2 As String
    Dim queryTable_3 As String
    
    Dim rngs As Range
    Dim rng As Range

    Dim fxReceive As Double
    Dim fxPay As Double

    reportTitle = "FM2"
    queryTable_1 = "FM2_OBU_MM4901B_LIST"
    queryTable_2 = "FM2_OBU_MM4901B_Subtotal"
    queryTable_3 = "FM2_OBU_MM4901B_Subtotal_BankCode"

    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")

    ' 取得 Access 資料庫路徑
    DBsPath = ThisWorkbook.Path & "\" & controlSheet.Range("DBsPathFileName")

    '***Set xlsht = ThisWorkbook.Sheets("CNY1")中的CNY1要用一個字串，當作參數傳入
    ' 設定目標 Excel 分頁
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    xlsht.Range("A:N").ClearContents

    dataArr_1 = GetAccessDataAsArray(DBsPath, queryTable_1)
    dataArr_2 = GetAccessDataAsArray(DBsPath, queryTable_2)
    dataArr_3 = GetAccessDataAsArray(DBsPath, queryTable_3)
    
    For j = 0 To UBound(dataArr_1, 2)
        For i = 0 To UBound(dataArr_1, 1)
            xlsht.Cells(i + 1, j + 1).Value = dataArr_1(i, j)
        Next i
    Next j

    For j = 0 To UBound(dataArr_2, 2)
        For i = 0 To UBound(dataArr_2, 1)
            xlsht.Cells(i + 1, j + 10).Value = dataArr_2(i, j)
        Next i
    Next j

    For j = 0 To UBound(dataArr_3, 2)
        For i = 0 To UBound(dataArr_3, 1)
            xlsht.Cells(i + 1, j + 12).Value = dataArr_3(i, j)
        Next i
    Next j
    

    ' Compute 期收/付遠匯款-換匯遠期
    ' fxReceive = 0
    ' fxPay = 0
    ' lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    ' Set rngs = xlsht.Range("C3:C" & lastRow)

    ' For Each rng In rngs
    '     If rng.Value = "155930402" Then
    '         fxReceive = fxReceive + rng.Offset(0, 2).Value
    '     ElseIf rng.Value = "255930402" Then
    '         fxPay = fxPay + rng.Offset(0, 2).Value
    '     End If
    ' Next rng


    ' fxReceive = Round(fxReceive / 1000, 0)
    ' fxPay = Round(fxPay / 1000, 0)
    
    ' xlsht.Range("其他金融資產_淨額").Value = fxReceive
    ' xlsht.Range("其他").Value = fxReceive
    ' xlsht.Range("CNY1_資產總計").Value = fxReceive

    ' xlsht.Range("其他金融負債").Value = fxPay
    ' xlsht.Range("其他什項金融負債").Value = fxPay
    ' xlsht.Range("CNY1_負債總計").Value = fxPay
    
    ' 'Set Number Format
    ' xlsht.Cells(2, "T").Resize(2, 1).NumberFormat = "#,##,00"
    
    ' MsgBox "Finish " & reportTitle & "分頁取得資料庫資料"
    
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他金融資產_淨額", xlsht.Range("其他金融資產_淨額").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他", xlsht.Range("其他").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "CNY1_資產總計", xlsht.Range("CNY1_資產總計").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他金融負債", xlsht.Range("其他金融負債").Value, "期付遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他什項金融負債", xlsht.Range("其他什項金融負債").Value, "期付遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "CNY1_負債總計", xlsht.Range("CNY1_負債總計").Value, "期付遠匯款-換匯遠期"
    MsgBox "Finish " & reportTitle & "申報資料寫入資料庫"


    ' Release Resources
    Set rng = Nothing
    Set rngs = Nothing
    Set xlsht = Nothing
    Set controlSheet = Nothing
    
End Sub

Public Sub FM10()
    Dim controlSheet As Worksheet
    Dim xlsht As Worksheet

    Dim DBsPath As String
    Dim i As Integer, j As Integer
    Dim lastRow As Integer
    Dim dataArr_1 As Variant
    Dim dataArr_2 As Variant

    Dim reportTitle As String
    Dim queryTable_1 As String
    Dim queryTable_2 As String
    
    Dim rngs As Range
    Dim rng As Range

    Dim fxReceive As Double
    Dim fxPay As Double

    reportTitle = "FM10"
    queryTable_1 = "FM10_OBU_AC4603_LIST"
    queryTable_2 = "FM10_OBU_AC4603_Subtotal"

    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")

    ' 取得 Access 資料庫路徑
    DBsPath = ThisWorkbook.Path & "\" & controlSheet.Range("DBsPathFileName")

    '***Set xlsht = ThisWorkbook.Sheets("CNY1")中的CNY1要用一個字串，當作參數傳入
    ' 設定目標 Excel 分頁
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    xlsht.Range("A:H").ClearContents

    dataArr_1 = GetAccessDataAsArray(DBsPath, queryTable_1)
    dataArr_2 = GetAccessDataAsArray(DBsPath, queryTable_2)
    
    For j = 0 To UBound(dataArr_1, 2)
        For i = 0 To UBound(dataArr_1, 1)
            xlsht.Cells(i + 1, j + 1).Value = dataArr_1(i, j)
        Next i
    Next j

    For j = 0 To UBound(dataArr_2, 2)
        For i = 0 To UBound(dataArr_2, 1)
            xlsht.Cells(i + 1, j + 7).Value = dataArr_2(i, j)
        Next i
    Next j
    
    ' Compute 期收/付遠匯款-換匯遠期
    ' fxReceive = 0
    ' fxPay = 0
    ' lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    ' Set rngs = xlsht.Range("C3:C" & lastRow)

    ' For Each rng In rngs
    '     If rng.Value = "155930402" Then
    '         fxReceive = fxReceive + rng.Offset(0, 2).Value
    '     ElseIf rng.Value = "255930402" Then
    '         fxPay = fxPay + rng.Offset(0, 2).Value
    '     End If
    ' Next rng


    ' fxReceive = Round(fxReceive / 1000, 0)
    ' fxPay = Round(fxPay / 1000, 0)
    
    ' xlsht.Range("其他金融資產_淨額").Value = fxReceive
    ' xlsht.Range("其他").Value = fxReceive
    ' xlsht.Range("CNY1_資產總計").Value = fxReceive

    ' xlsht.Range("其他金融負債").Value = fxPay
    ' xlsht.Range("其他什項金融負債").Value = fxPay
    ' xlsht.Range("CNY1_負債總計").Value = fxPay
    
    ' 'Set Number Format
    ' xlsht.Cells(2, "T").Resize(2, 1).NumberFormat = "#,##,00"
    
    ' MsgBox "Finish " & reportTitle & "分頁取得資料庫資料"
    
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他金融資產_淨額", xlsht.Range("其他金融資產_淨額").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他", xlsht.Range("其他").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "CNY1_資產總計", xlsht.Range("CNY1_資產總計").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他金融負債", xlsht.Range("其他金融負債").Value, "期付遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他什項金融負債", xlsht.Range("其他什項金融負債").Value, "期付遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "CNY1_負債總計", xlsht.Range("CNY1_負債總計").Value, "期付遠匯款-換匯遠期"
    MsgBox "Finish " & reportTitle & "申報資料寫入資料庫"


    ' Release Resources
    Set rng = Nothing
    Set rngs = Nothing
    Set xlsht = Nothing
    Set controlSheet = Nothing
    
End Sub

Public Sub F1_F2()
    Dim controlSheet As Worksheet
    Dim xlsht As Worksheet

    Dim DBsPath As String
    Dim i As Integer, j As Integer
    Dim lastRow As Integer
    Dim dataArr_1 As Variant
    Dim dataArr_2 As Variant
    Dim dataArr_3 As Variant
    Dim dataArr_4 As Variant
    Dim dataArr_5 As Variant
    Dim dataArr_6 As Variant

    Dim reportTitle As String
    Dim queryTable_1 As String
    Dim queryTable_2 As String
    Dim queryTable_3 As String
    Dim queryTable_4 As String
    Dim queryTable_5 As String
    Dim queryTable_6 As String
    
    Dim rngs As Range
    Dim rng As Range

    Dim fxReceive As Double
    Dim fxPay As Double

    reportTitle = "F1_F2"
    queryTable_1 = "F1_Foreign_DL6850_FS"
    queryTable_2 = "F1_Foreign_DL6850_SS"
    queryTable_3 = "F1_Domestic_DL6850_FS"
    queryTable_4 = "F1_Domestic_DL6850_SS"
    queryTable_5 = "F1_CM2810_LIST"
    queryTable_6 = "F1_CM2810_Subtotal"

    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")

    ' 取得 Access 資料庫路徑
    DBsPath = ThisWorkbook.Path & "\" & controlSheet.Range("DBsPathFileName")

    '***Set xlsht = ThisWorkbook.Sheets("CNY1")中的CNY1要用一個字串，當作參數傳入
    ' 設定目標 Excel 分頁
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    xlsht.Range("A:S").ClearContents

    dataArr_1 = GetAccessDataAsArray(DBsPath, queryTable_1)
    dataArr_2 = GetAccessDataAsArray(DBsPath, queryTable_2)
    dataArr_3 = GetAccessDataAsArray(DBsPath, queryTable_3)
    dataArr_4 = GetAccessDataAsArray(DBsPath, queryTable_4)
    dataArr_5 = GetAccessDataAsArray(DBsPath, queryTable_5)
    dataArr_6 = GetAccessDataAsArray(DBsPath, queryTable_6)
    


    If Err.Number <> 0 Or LBound(dataArr_1) > UBound(dataArr_1) Then
        MsgBox queryTable_1 & "資料表無資料"
    Else
        For j = 0 To UBound(dataArr_1, 2)
            For i = 0 To UBound(dataArr_1, 1)
                xlsht.Cells(i + 1, j + 1).Value = dataArr_1(i, j)
            Next i
        Next j
    End If

    If Err.Number <> 0 Or LBound(dataArr_2) > UBound(dataArr_2) Then
        MsgBox queryTable_2 & "資料表無資料"
    Else
        For j = 0 To UBound(dataArr_2, 2)
            For i = 0 To UBound(dataArr_2, 1)
                xlsht.Cells(i + 1, j + 3).Value = dataArr_2(i, j)
            Next i
        Next j
    End If

    If Err.Number <> 0 Or LBound(dataArr_3) > UBound(dataArr_3) Then
        MsgBox queryTable_3 & "資料表無資料"
    Else
        For j = 0 To UBound(dataArr_3, 2)
            For i = 0 To UBound(dataArr_3, 1)
                xlsht.Cells(i + 1, j + 5).Value = dataArr_3(i, j)
            Next i
        Next j
    End If

    If Err.Number <> 0 Or LBound(dataArr_4) > UBound(dataArr_4) Then
        MsgBox queryTable_4 & "資料表無資料"
    Else
        For j = 0 To UBound(dataArr_4, 2)
            For i = 0 To UBound(dataArr_4, 1)
                xlsht.Cells(i + 1, j + 7).Value = dataArr_4(i, j)
            Next i
        Next j
    End If

    If Err.Number <> 0 Or LBound(dataArr_5) > UBound(dataArr_5) Then
        MsgBox queryTable_5 & "資料表無資料"
    Else
        For j = 0 To UBound(dataArr_5, 2)
            For i = 0 To UBound(dataArr_5, 1)
                xlsht.Cells(i + 1, j + 9).Value = dataArr_5(i, j)
            Next i
        Next j
    End If

    If Err.Number <> 0 Or LBound(dataArr_6) > UBound(dataArr_6) Then
        MsgBox queryTable_6 & "資料表無資料"
    Else
        For j = 0 To UBound(dataArr_6, 2)
            For i = 0 To UBound(dataArr_6, 1)
                xlsht.Cells(i + 1, j + 17).Value = dataArr_6(i, j)
            Next i
        Next j
    End If


    ' Compute 期收/付遠匯款-換匯遠期
    ' fxReceive = 0
    ' fxPay = 0
    ' lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    ' Set rngs = xlsht.Range("C3:C" & lastRow)

    ' For Each rng In rngs
    '     If rng.Value = "155930402" Then
    '         fxReceive = fxReceive + rng.Offset(0, 2).Value
    '     ElseIf rng.Value = "255930402" Then
    '         fxPay = fxPay + rng.Offset(0, 2).Value
    '     End If
    ' Next rng


    ' fxReceive = Round(fxReceive / 1000, 0)
    ' fxPay = Round(fxPay / 1000, 0)
    
    ' xlsht.Range("其他金融資產_淨額").Value = fxReceive
    ' xlsht.Range("其他").Value = fxReceive
    ' xlsht.Range("CNY1_資產總計").Value = fxReceive

    ' xlsht.Range("其他金融負債").Value = fxPay
    ' xlsht.Range("其他什項金融負債").Value = fxPay
    ' xlsht.Range("CNY1_負債總計").Value = fxPay
    
    ' 'Set Number Format
    ' xlsht.Cells(2, "T").Resize(2, 1).NumberFormat = "#,##,00"
    
    ' MsgBox "Finish " & reportTitle & "分頁取得資料庫資料"
    
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他金融資產_淨額", xlsht.Range("其他金融資產_淨額").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他", xlsht.Range("其他").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "CNY1_資產總計", xlsht.Range("CNY1_資產總計").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他金融負債", xlsht.Range("其他金融負債").Value, "期付遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他什項金融負債", xlsht.Range("其他什項金融負債").Value, "期付遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "CNY1_負債總計", xlsht.Range("CNY1_負債總計").Value, "期付遠匯款-換匯遠期"
    MsgBox "Finish " & reportTitle & "申報資料寫入資料庫"


    ' Release Resources
    Set rng = Nothing
    Set rngs = Nothing
    Set xlsht = Nothing
    Set controlSheet = Nothing
    
End Sub

Public Sub Table41()
    Dim controlSheet As Worksheet
    Dim xlsht As Worksheet

    Dim DBsPath As String
    Dim i As Integer, j As Integer
    Dim lastRow As Integer
    Dim dataArr_1 As Variant
    Dim dataArr_2 As Variant

    Dim reportTitle As String
    Dim queryTable_1 As String
    Dim queryTable_2 As String
    
    Dim rngs As Range
    Dim rng As Range

    Dim fxReceive As Double
    Dim fxPay As Double

    reportTitle = "Table41"
    queryTable_1 = "表41_DBU_DL9360_LIST"
    queryTable_2 = "表41_DBU_DL9360_Subtotal"

    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")

    ' 取得 Access 資料庫路徑
    DBsPath = ThisWorkbook.Path & "\" & controlSheet.Range("DBsPathFileName")

    '***Set xlsht = ThisWorkbook.Sheets("CNY1")中的CNY1要用一個字串，當作參數傳入
    ' 設定目標 Excel 分頁
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    xlsht.Range("A:J").ClearContents

    dataArr_1 = GetAccessDataAsArray(DBsPath, queryTable_1)
    dataArr_2 = GetAccessDataAsArray(DBsPath, queryTable_2)
    
    For j = 0 To UBound(dataArr_1, 2)
        For i = 0 To UBound(dataArr_1, 1)
            xlsht.Cells(i + 1, j + 1).Value = dataArr_1(i, j)
        Next i
    Next j

    For j = 0 To UBound(dataArr_2, 2)
        For i = 0 To UBound(dataArr_2, 1)
            xlsht.Cells(i + 1, j + 8).Value = dataArr_2(i, j)
        Next i
    Next j
    

    ' Compute 期收/付遠匯款-換匯遠期
    ' fxReceive = 0
    ' fxPay = 0
    ' lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    ' Set rngs = xlsht.Range("C3:C" & lastRow)

    ' For Each rng In rngs
    '     If rng.Value = "155930402" Then
    '         fxReceive = fxReceive + rng.Offset(0, 2).Value
    '     ElseIf rng.Value = "255930402" Then
    '         fxPay = fxPay + rng.Offset(0, 2).Value
    '     End If
    ' Next rng


    ' fxReceive = Round(fxReceive / 1000, 0)
    ' fxPay = Round(fxPay / 1000, 0)
    
    ' xlsht.Range("其他金融資產_淨額").Value = fxReceive
    ' xlsht.Range("其他").Value = fxReceive
    ' xlsht.Range("CNY1_資產總計").Value = fxReceive

    ' xlsht.Range("其他金融負債").Value = fxPay
    ' xlsht.Range("其他什項金融負債").Value = fxPay
    ' xlsht.Range("CNY1_負債總計").Value = fxPay
    
    ' 'Set Number Format
    ' xlsht.Cells(2, "T").Resize(2, 1).NumberFormat = "#,##,00"
    
    ' MsgBox "Finish " & reportTitle & "分頁取得資料庫資料"
    
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他金融資產_淨額", xlsht.Range("其他金融資產_淨額").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他", xlsht.Range("其他").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "CNY1_資產總計", xlsht.Range("CNY1_資產總計").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他金融負債", xlsht.Range("其他金融負債").Value, "期付遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他什項金融負債", xlsht.Range("其他什項金融負債").Value, "期付遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "CNY1_負債總計", xlsht.Range("CNY1_負債總計").Value, "期付遠匯款-換匯遠期"
    MsgBox "Finish " & reportTitle & "申報資料寫入資料庫"


    ' Release Resources
    Set rng = Nothing
    Set rngs = Nothing
    Set xlsht = Nothing
    Set controlSheet = Nothing
    
End Sub

Public Sub AI602()
    Dim controlSheet As Worksheet
    Dim xlsht As Worksheet

    Dim DBsPath As String
    Dim i As Integer, j As Integer
    Dim lastRow As Integer
    Dim dataArr_1 As Variant
    Dim dataArr_2 As Variant
    Dim dataArr_3 As Variant

    Dim reportTitle As String
    Dim queryTable_1 As String
    Dim queryTable_2 As String
    Dim queryTable_3 As String
    
    Dim rngs As Range
    Dim rng As Range

    Dim fxReceive As Double
    Dim fxPay As Double

    reportTitle = "AI602"
    queryTable_1 = "AI602_Impairment_USD"
    queryTable_2 = "AI602_GroupedAC5601"
    queryTable_3 = "AI602_Subtotal"

    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")

    ' 取得 Access 資料庫路徑
    DBsPath = ThisWorkbook.Path & "\" & controlSheet.Range("DBsPathFileName")

    '***Set xlsht = ThisWorkbook.Sheets("CNY1")中的CNY1要用一個字串，當作參數傳入
    ' 設定目標 Excel 分頁
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    xlsht.Range("A:K").ClearContents

    dataArr_1 = GetAccessDataAsArray(DBsPath, queryTable_1)
    dataArr_2 = GetAccessDataAsArray(DBsPath, queryTable_2)
    dataArr_3 = GetAccessDataAsArray(DBsPath, queryTable_3)
    
    For j = 0 To UBound(dataArr_1, 2)
        For i = 0 To UBound(dataArr_1, 1)
            xlsht.Cells(i + 1, j + 1).Value = dataArr_1(i, j)
        Next i
    Next j

    For j = 0 To UBound(dataArr_2, 2)
        For i = 0 To UBound(dataArr_2, 1)
            xlsht.Cells(i + 1, j + 3).Value = dataArr_2(i, j)
        Next i
    Next j
    
    For j = 0 To UBound(dataArr_3, 2)
        For i = 0 To UBound(dataArr_3, 1)
            xlsht.Cells(i + 1, j + 10).Value = dataArr_3(i, j)
        Next i
    Next j

    ' Compute 期收/付遠匯款-換匯遠期
    ' fxReceive = 0
    ' fxPay = 0
    ' lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    ' Set rngs = xlsht.Range("C3:C" & lastRow)

    ' For Each rng In rngs
    '     If rng.Value = "155930402" Then
    '         fxReceive = fxReceive + rng.Offset(0, 2).Value
    '     ElseIf rng.Value = "255930402" Then
    '         fxPay = fxPay + rng.Offset(0, 2).Value
    '     End If
    ' Next rng


    ' fxReceive = Round(fxReceive / 1000, 0)
    ' fxPay = Round(fxPay / 1000, 0)
    
    ' xlsht.Range("其他金融資產_淨額").Value = fxReceive
    ' xlsht.Range("其他").Value = fxReceive
    ' xlsht.Range("CNY1_資產總計").Value = fxReceive

    ' xlsht.Range("其他金融負債").Value = fxPay
    ' xlsht.Range("其他什項金融負債").Value = fxPay
    ' xlsht.Range("CNY1_負債總計").Value = fxPay
    
    ' 'Set Number Format
    ' xlsht.Cells(2, "T").Resize(2, 1).NumberFormat = "#,##,00"
    
    ' MsgBox "Finish " & reportTitle & "分頁取得資料庫資料"
    
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他金融資產_淨額", xlsht.Range("其他金融資產_淨額").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他", xlsht.Range("其他").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "CNY1_資產總計", xlsht.Range("CNY1_資產總計").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他金融負債", xlsht.Range("其他金融負債").Value, "期付遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他什項金融負債", xlsht.Range("其他什項金融負債").Value, "期付遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "CNY1_負債總計", xlsht.Range("CNY1_負債總計").Value, "期付遠匯款-換匯遠期"
    MsgBox "Finish " & reportTitle & "申報資料寫入資料庫"


    ' Release Resources
    Set rng = Nothing
    Set rngs = Nothing
    Set xlsht = Nothing
    Set controlSheet = Nothing
    
End Sub

Public Sub AI240()
    Dim controlSheet As Worksheet
    Dim xlsht As Worksheet

    Dim DBsPath As String
    Dim i As Integer, j As Integer
    Dim lastRow As Integer
    Dim dataArr_1 As Variant
    Dim dataArr_2 As Variant

    Dim reportTitle As String
    Dim queryTable_1 As String
    Dim queryTable_2 As String
    
    Dim rngs As Range
    Dim rng As Range

    Dim fxReceive As Double
    Dim fxPay As Double

    reportTitle = "AI240"
    queryTable_1 = "AI240_DBU_DL6850_LIST"
    queryTable_2 = "AI240_DBU_DL6850_Subtoal"

    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")

    ' 取得 Access 資料庫路徑
    DBsPath = ThisWorkbook.Path & "\" & controlSheet.Range("DBsPathFileName")

    '***Set xlsht = ThisWorkbook.Sheets("CNY1")中的CNY1要用一個字串，當作參數傳入
    ' 設定目標 Excel 分頁
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    xlsht.Range("A:L").ClearContents

    dataArr_1 = GetAccessDataAsArray(DBsPath, queryTable_1)
    dataArr_2 = GetAccessDataAsArray(DBsPath, queryTable_2)
    
    For j = 0 To UBound(dataArr_1, 2)
        For i = 0 To UBound(dataArr_1, 1)
            xlsht.Cells(i + 1, j + 1).Value = dataArr_1(i, j)
        Next i
    Next j

    For j = 0 To UBound(dataArr_2, 2)
        For i = 0 To UBound(dataArr_2, 1)
            xlsht.Cells(i + 1, j + 10).Value = dataArr_2(i, j)
        Next i
    Next j
    

    ' Compute 期收/付遠匯款-換匯遠期
    ' fxReceive = 0
    ' fxPay = 0
    ' lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    ' Set rngs = xlsht.Range("C3:C" & lastRow)

    ' For Each rng In rngs
    '     If rng.Value = "155930402" Then
    '         fxReceive = fxReceive + rng.Offset(0, 2).Value
    '     ElseIf rng.Value = "255930402" Then
    '         fxPay = fxPay + rng.Offset(0, 2).Value
    '     End If
    ' Next rng


    ' fxReceive = Round(fxReceive / 1000, 0)
    ' fxPay = Round(fxPay / 1000, 0)
    
    ' xlsht.Range("其他金融資產_淨額").Value = fxReceive
    ' xlsht.Range("其他").Value = fxReceive
    ' xlsht.Range("CNY1_資產總計").Value = fxReceive

    ' xlsht.Range("其他金融負債").Value = fxPay
    ' xlsht.Range("其他什項金融負債").Value = fxPay
    ' xlsht.Range("CNY1_負債總計").Value = fxPay
    
    ' 'Set Number Format
    ' xlsht.Cells(2, "T").Resize(2, 1).NumberFormat = "#,##,00"
    
    ' MsgBox "Finish " & reportTitle & "分頁取得資料庫資料"
    
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他金融資產_淨額", xlsht.Range("其他金融資產_淨額").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他", xlsht.Range("其他").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "CNY1_資產總計", xlsht.Range("CNY1_資產總計").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他金融負債", xlsht.Range("其他金融負債").Value, "期付遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他什項金融負債", xlsht.Range("其他什項金融負債").Value, "期付遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "CNY1_負債總計", xlsht.Range("CNY1_負債總計").Value, "期付遠匯款-換匯遠期"
    MsgBox "Finish " & reportTitle & "申報資料寫入資料庫"


    ' Release Resources
    Set rng = Nothing
    Set rngs = Nothing
    Set xlsht = Nothing
    Set controlSheet = Nothing
    
End Sub








'以下為第二階段將資料貼入申報檔案
'--------------------------------------------
'讀取access data and save to another file
'--------------------------------------------

Option Explicit

Public Sub ExportReport()
    Dim pathSheet As Worksheet
    Dim DBsPath As String
    Dim emptyReportPath As String
    Dim reportTitle As String
    Dim dataMonthString As String
    Dim rs As Object
    Dim xlsht As Worksheet
    Dim emptyReportWb As Workbook
    Dim targetWb As Workbook
    Dim lastRow As Integer
    Dim i As Integer

    ' 取得控制面板分頁
    Set pathSheet = ThisWorkbook.Sheets("ControlPanel")

    ' 取得 Access 資料庫路徑
    DBsPath = pathSheet.Range("DBsPath") & "\" & pathSheet.Range("DBsPathFileName")

    ' 轉換資料月份格式
    dataMonthString = pathSheet.Range("DataMonthString")

    ' 取得報表路徑
    emptyReportPath = pathSheet.Range("EmptyReportPath")

    ' 取得報表標題 (假設在 ControlPanel 頁面有指定)
    reportTitle = pathSheet.Range("ReportTitle")

    ' 根據 ReportTitle 和 DataMonthString 查詢資料
    Set rs = GetReportData(DBsPath, reportTitle, dataMonthString)

    ' 如果沒有資料則離開
    If rs Is Nothing Or rs.EOF Then
        MsgBox "無法找到符合條件的資料！", vbExclamation
        Exit Sub
    End If

    ' 打開空白申報 Excel 檔案
    Set emptyReportWb = Workbooks.Open(emptyReportPath)
    Set xlsht = emptyReportWb.Sheets(1) ' 假設第一個工作表是要填寫的

    ' 將資料填入空白申報檔案
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    For i = 1 To rs.Fields.Count
        ' 假設資料從 A2 開始，依序填入對應的欄位
        xlsht.Cells(lastRow + 1, i).Value = rs.Fields(i - 1).Value
    Next i

    ' 關閉 Recordset
    rs.Close
    Set rs = Nothing

    ' 另存為新的檔案
    Dim newFileName As String
    newFileName = "C:\YourPath\NewReport_" & reportTitle & "_" & dataMonthString & ".xlsx" ' 設定新檔案名稱
    emptyReportWb.SaveAs newFileName

    ' 關閉新檔案
    emptyReportWb.Close

    MsgBox "報表已成功匯出並儲存為新檔案！", vbInformation
End Sub







' 根據 ReportTitle 和 DataMonthString 查詢報表資料
Public Function GetReportData(DBPath As String, ReportTitle As String, DataMonthString As String) As Object
    Dim conn As Object
    Dim cmd As Object
    Dim rs As Object
    Dim query As String

    On Error GoTo ErrHandler

    ' 建立 ADODB 連線
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath

    ' 構建 SQL 查詢語句
    query = "SELECT * FROM Monthly_Declaration_Report WHERE ReportTitle = '" & ReportTitle & "' AND DataMonthString = '" & DataMonthString & "'"

    ' 建立 Command
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = query
        .CommandType = 1 ' adCmdText (表示這是文字查詢)
    End With

    ' 執行 SQL 並回傳 Recordset
    Set rs = cmd.Execute

    ' 關閉連線
    conn.Close
    Set conn = Nothing
    Set cmd = Nothing

    ' 回傳 Recordset
    Set GetReportData = rs
    Exit Function

ErrHandler:
    MsgBox "發生錯誤: " & Err.Description, vbCritical
    Set GetReportData = Nothing
End Function
    





Public Sub ExportReport()
    Dim pathSheet As Worksheet
    Dim DBsPath As String
    Dim emptyReportPath As String
    Dim reportTitle As String
    Dim dataMonthString As String
    Dim rs As Object
    Dim xlsht As Worksheet
    Dim emptyReportWb As Workbook
    Dim lastRow As Integer
    Dim i As Integer
    Dim fieldCodeDict As Object ' Dictionary 存放 FieldCode 對應的 Excel 欄位
    
    ' 取得控制面板分頁
    Set pathSheet = ThisWorkbook.Sheets("ControlPanel")
    
    ' 取得 Access 資料庫路徑
    DBsPath = pathSheet.Range("DBsPath") & "\" & pathSheet.Range("DBsPathFileName")
    
    ' 轉換資料月份格式
    dataMonthString = pathSheet.Range("DataMonthString")
    
    ' 取得報表路徑
    emptyReportPath = pathSheet.Range("EmptyReportPath")
    
    ' 取得報表標題
    reportTitle = pathSheet.Range("ReportTitle")
    
    ' 查詢資料
    Set rs = GetReportData(DBsPath, reportTitle, dataMonthString)
    
    ' 如果沒有資料則離開
    If rs Is Nothing Or rs.EOF Then
        MsgBox "無法找到符合條件的資料！", vbExclamation
        Exit Sub
    End If
    
    ' 打開空白報表
    Set emptyReportWb = Workbooks.Open(emptyReportPath)
    Set xlsht = emptyReportWb.Sheets(1) ' 假設第一個工作表是要填寫的
    
    ' 設定 Dictionary 來對應 FieldCode 到 Excel 欄位
    Set fieldCodeDict = CreateObject("Scripting.Dictionary")
    
    ' 這裡要根據你的 Excel 欄位來對應 FieldCode
    ' 假設 Excel 欄位 A, B, C, D, E, F 分別對應不同的 FieldCode
    fieldCodeDict.Add "FC001", "B"  ' 假設 FC001 的值填入 B 欄
    fieldCodeDict.Add "FC002", "C"  ' 假設 FC002 的值填入 C 欄
    fieldCodeDict.Add "FC003", "D"  ' 假設 FC003 的值填入 D 欄
    fieldCodeDict.Add "FC004", "E"  ' 假設 FC004 的值填入 E 欄
    fieldCodeDict.Add "FC005", "F"  ' 假設 FC005 的值填入 F 欄
    
    ' 找到最後一列，準備填入新資料
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row + 1
    
    ' 讀取 Recordset 並填入 Excel
    Do While Not rs.EOF
        Dim fieldCode As String
        Dim contentValue As Variant
        Dim targetColumn As String
        
        fieldCode = rs.Fields("FieldCode").Value
        contentValue = rs.Fields("Content").Value
        
        ' 如果 FieldCode 存在於 Dictionary，則填入對應欄位
        If fieldCodeDict.Exists(fieldCode) Then
            targetColumn = fieldCodeDict(fieldCode) ' 取得對應的 Excel 欄位
            xlsht.Range(targetColumn & lastRow).Value = contentValue
        Else
            Debug.Print "找不到 FieldCode 對應的欄位：" & fieldCode
        End If
        
        rs.MoveNext
    Loop
    
    ' 關閉 Recordset
    rs.Close
    Set rs = Nothing
    
    ' 另存為新的
    Dim newFileName As String
    newFileName = "C:\YourPath\NewReport_" & reportTitle & "_" & dataMonthString & ".xlsx"
    emptyReportWb.SaveAs newFileName
    
    ' 關閉新檔案
    emptyReportWb.Close
    
    ' 清除 Dictionary
    Set fieldCodeDict = Nothing
    
    MsgBox "報表已成功匯出並儲存為新檔案！", vbInformation
End Sub




'recordset save to dictionary and for loop to place to specific range

Public Sub ExportReport()
    Dim pathSheet As Worksheet
    Dim DBsPath As String
    Dim emptyReportPath As String
    Dim reportTitle As String
    Dim dataMonthString As String
    Dim rs As Object
    Dim xlsht As Worksheet
    Dim emptyReportWb As Workbook
    Dim lastRow As Integer
    Dim fieldDataDict As Object ' Dictionary 存放 FieldCode → Content
    Dim key As Variant
    
    ' 取得控制面板分頁
    Set pathSheet = ThisWorkbook.Sheets("ControlPanel")
    
    ' 取得 Access 資料庫路徑
    DBsPath = pathSheet.Range("DBsPath") & "\" & pathSheet.Range("DBsPathFileName")
    
    ' 轉換資料月份格式
    dataMonthString = pathSheet.Range("DataMonthString")
    
    ' 取得報表路徑
    emptyReportPath = pathSheet.Range("EmptyReportPath")
    
    ' 取得報表標題
    reportTitle = pathSheet.Range("ReportTitle")
    
    ' 查詢資料
    Set rs = GetReportData(DBsPath, reportTitle, dataMonthString)
    
    ' 如果沒有資料則離開
    If rs Is Nothing Or rs.EOF Then
        MsgBox "無法找到符合條件的資料！", vbExclamation
        Exit Sub
    End If
    
    ' 打開空白報表
    Set emptyReportWb = Workbooks.Open(emptyReportPath)
    Set xlsht = emptyReportWb.Sheets(1) ' 假設第一個工作表是要填寫的
    
    ' 設定 Dictionary 來存放 FieldCode → Content
    Set fieldDataDict = CreateObject("Scripting.Dictionary")
    
    ' 讀取 Recordset，將資料存入 Dictionary
    Do While Not rs.EOF
        Dim fieldCode As String
        Dim contentValue As Variant
        
        fieldCode = rs.Fields("FieldCode").Value
        contentValue = rs.Fields("Content").Value
        
        ' 存入 Dictionary
        fieldDataDict(fieldCode) = contentValue
        
        rs.MoveNext
    Loop
    
    ' 關閉 Recordset
    rs.Close
    Set rs = Nothing
    
    ' 找到最後一列，準備填入新資料
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row + 1
    
    ' 遍歷 Dictionary，根據 FieldCode 來判斷應填入哪個欄位
    For Each key In fieldDataDict.Keys
        Dim targetColumn As String
        
        ' 根據 FieldCode 判斷要填入哪個欄位
        If key = "FC001" Then
            targetColumn = "B"  ' FC001 填入 B 欄
        ElseIf key = "FC002" Then
            targetColumn = "C"  ' FC002 填入 C 欄
        ElseIf key = "FC003" Then
            targetColumn = "D"  ' FC003 填入 D 欄
        ElseIf key = "FC004" Then
            targetColumn = "E"  ' FC004 填入 E 欄
        ElseIf key = "FC005" Then
            targetColumn = "F"  ' FC005 填入 F 欄
        Else
            Debug.Print "找不到 FieldCode 對應的欄位：" & key
            GoTo NextKey ' 若 FieldCode 無對應欄位則跳過
        End If
        
        ' 將對應的資料填入 Excel
        xlsht.Range(targetColumn & lastRow).Value = fieldDataDict(key)
        
NextKey:
    Next key
    
    ' 另存為新的檔案
    Dim newFileName As String
    newFileName = "C:\YourPath\NewReport_" & reportTitle & "_" & dataMonthString & ".xlsx"
    emptyReportWb.SaveAs newFileName
    
    ' 關閉新檔案
    emptyReportWb.Close
    
    ' 清除 Dictionary
    Set fieldDataDict = Nothing
    
    MsgBox "報表已成功匯出並儲存為新檔案！", vbInformation
End Sub




'-----including inputbox

' 以下是一個整合了 InputBox 輸入與格式檢核機制的範例，利用正規表示式驗證使用者輸入必須符合「yyyy/mm」格式（年份4碼，月份01~12），若格式錯誤則會重複提示使用者輸入。

' 請參考以下修改後的完整程式碼：

Option Explicit

Public Sub CNY1()
    Dim xlsht As Worksheet
    Dim controlSheet As Worksheet
    Dim DBsPath As String
    Dim emptyReportPath As String
    Dim dataMonthString As String
    Dim rs As Object
    Dim lastRow As Integer
    Dim sum155930402 As Double
    Dim sum255930402 As Double
    Dim rngs As Range
    Dim rng As Range
    Dim fieldNames As Variant
    Dim i As Integer
    Dim userInput As String
    Dim regex As Object
    
    ' 取得控制面板分頁
    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")
    
    ' 取得 Access 資料庫路徑
    DBsPath = ThisWorkbook.Path & "\" & controlSheet.Range("DBsPathFileName")
    Debug.Print "DBsPath: " & DBsPath
    
    ' 取得報表路徑
    emptyReportPath = ThisWorkbook.Path & "\" & controlSheet.Range("EmptyReportPath")
    Debug.Print "emptyReportPath: " & emptyReportPath
    
    '--- 透過 InputBox 讓使用者輸入資料月份，格式必須為 yyyy/mm
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Pattern = "^\d{4}/(0[1-9]|1[0-2])$"
        .IgnoreCase = True
        .Global = False
    End With
    
    Do
        userInput = InputBox("請輸入資料月份 (格式: yyyy/mm):", "輸入資料月份")
        ' 若使用者按下取消或沒輸入內容，則退出
        If Trim(userInput) = "" Then
            MsgBox "必須輸入資料月份！", vbExclamation, "輸入錯誤"
            Exit Sub
        End If
        If regex.Test(userInput) Then
            Exit Do
        Else
            MsgBox "輸入格式錯誤，請依格式 yyyy/mm 輸入！", vbExclamation, "格式錯誤"
        End If
    Loop
    
    dataMonthString = userInput
    Debug.Print "dataMonthString: " & dataMonthString
    '-------------------------------
    
    ' 設定目標 Excel 分頁（分頁名稱以參數傳入，此例仍用 "CNY1"）
    Set xlsht = ThisWorkbook.Sheets("CNY1")
    xlsht.UsedRange.ClearContents
    
    ' 取得資料庫資料 (參數中 "CNY1_DBU_AC5601" 為資料庫中資料表名稱)
    Set rs = GetAccessData(DBsPath, "CNY1_DBU_AC5601")
    
    ' 如果沒有資料則離開
    If rs Is Nothing Or rs.EOF Then
        MsgBox "無法找到符合條件的資料！", vbExclamation
        Exit Sub
    End If
    
    ' 取得欄位名稱陣列，並寫入 Excel 第 2 列
    fieldNames = GetFieldNamesFromRecordset(rs)
    For i = LBound(fieldNames) To UBound(fieldNames)
        xlsht.Cells(2, i + 1).Value = fieldNames(i)
    Next i
    
    ' 將資料寫入 Excel (從 A3 開始)
    xlsht.Range("A3").CopyFromRecordset rs
    
    ' 關閉 Recordset
    rs.Close
    Set rs = Nothing
    
    ' 計算 sum155930402 與 sum255930402
    sum155930402 = 0
    sum255930402 = 0
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    Set rngs = xlsht.Range("C3:C" & lastRow)
    
    For Each rng In rngs
        ' 判斷 FieldCode (假設 FieldCode 資料存放在 C 欄)
        If rng.Value = "155930402" Then
            ' 假設 Content 值在該列往右兩欄（E 欄）
            sum155930402 = sum155930402 + rng.Offset(0, 2).Value
        ElseIf rng.Value = "255930402" Then
            sum255930402 = sum255930402 + rng.Offset(0, 2).Value
        End If
    Next rng
    
    Debug.Print "sum155930402: " & sum155930402
    Debug.Print "sum255930402: " & sum255930402
    
    ' 寫入合計數值到 H3 & H4
    xlsht.Cells(3, "H").Value = sum155930402
    xlsht.Cells(4, "H").Value = sum255930402
End Sub

' 程式說明
' 	1.	正規表示式檢核：
' 	•	使用 CreateObject("VBScript.RegExp") 建立正規表示式物件。
' 	•	設定 .Pattern = "^\d{4}/(0[1-9]|1[0-2])$"，表示年份必須4位數、斜線，再加上月份（01~12）。
' 	•	進入一個 Do 迴圈，透過 InputBox 讓使用者輸入資料月份。如果輸入為空則退出；若不符合格式則顯示錯誤訊息並重複提示。
' 	2.	資料庫與報表路徑設定：
' 	•	利用 ThisWorkbook.Path 組合存放 Access 資料庫與空白報表的路徑。
' 	3.	讀取資料與輸出：
' 	•	透過 GetAccessData 取得 Recordset，再由 GetFieldNamesFromRecordset 將欄位名稱填入 Excel，最後利用 CopyFromRecordset 將資料寫入工作表中。
' 	4.	計算合計數值：
' 	•	依據欄位 C 中的 FieldCode 值，計算相應資料（位於同一列偏移兩欄）的合計，再將結果寫入 H3 與 H4 儲存格。

' 這樣使用者就能透過視窗輸入符合格式的資料月份，並且在程式中即時檢核輸入是否正確。




'-----including inputbox with function validation

' 下面提供一個將資料月份格式檢核包裝在單獨函數中的範例。此函數會接收使用者輸入的字串，並回傳 Boolean 值，代表該字串是否符合「yyyy/mm」格式（年份4碼、月份01~12）。在主程式中，就可以用這個函數來檢查 InputBox 的輸入是否正確，從而使程式碼更具彈性與模組化。

' 以下是完整的範例程式碼：
    
Option Explicit
    
' 檢查傳入的資料月份格式是否符合 "yyyy/mm"
Function IsValidDataMonth(ByVal userInput As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Pattern = "^\d{4}/(0[1-9]|1[0-2])$"
        .IgnoreCase = True
        .Global = False
    End With
    
    IsValidDataMonth = regex.Test(Trim(userInput))
End Function

Public Sub CNY1()
    Dim xlsht As Worksheet
    Dim controlSheet As Worksheet
    Dim DBsPath As String
    Dim emptyReportPath As String
    Dim dataMonthString As String
    Dim rs As Object
    Dim lastRow As Integer
    Dim sum155930402 As Double
    Dim sum255930402 As Double
    Dim rngs As Range
    Dim rng As Range
    Dim fieldNames As Variant
    Dim i As Integer
    Dim userInput As String
    
    ' 取得控制面板分頁
    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")
    
    ' 取得 Access 資料庫路徑
    DBsPath = ThisWorkbook.Path & "\" & controlSheet.Range("DBsPathFileName")
    Debug.Print "DBsPath: " & DBsPath
    
    ' 取得報表路徑
    emptyReportPath = ThisWorkbook.Path & "\" & controlSheet.Range("EmptyReportPath")
    Debug.Print "emptyReportPath: " & emptyReportPath
    
    ' 透過 InputBox 讓使用者輸入資料月份，格式必須為 yyyy/mm
    Do
        userInput = InputBox("請輸入資料月份 (格式: yyyy/mm):", "輸入資料月份")
        If Trim(userInput) = "" Then
            MsgBox "必須輸入資料月份！", vbExclamation, "輸入錯誤"
            Exit Sub
        End If
        
        If IsValidDataMonth(userInput) Then
            Exit Do
        Else
            MsgBox "輸入格式錯誤，請依格式 yyyy/mm 輸入！", vbExclamation, "格式錯誤"
        End If
    Loop
    
    dataMonthString = userInput
    Debug.Print "dataMonthString: " & dataMonthString
    
    ' 設定目標 Excel 分頁（假設分頁名稱為 "CNY1"）
    Set xlsht = ThisWorkbook.Sheets("CNY1")
    xlsht.UsedRange.ClearContents
    
    ' 取得資料庫資料 (傳入的資料表名稱 "CNY1_DBU_AC5601" 為參數)
    Set rs = GetAccessData(DBsPath, "CNY1_DBU_AC5601")
    
    ' 如果沒有資料則離開
    If rs Is Nothing Or rs.EOF Then
        MsgBox "無法找到符合條件的資料！", vbExclamation
        Exit Sub
    End If
    
    ' 取得欄位名稱陣列，並寫入 Excel 第 2 列
    fieldNames = GetFieldNamesFromRecordset(rs)
    For i = LBound(fieldNames) To UBound(fieldNames)
        xlsht.Cells(2, i + 1).Value = fieldNames(i)
    Next i
    
    ' 將資料寫入 Excel (從 A3 開始)
    xlsht.Range("A3").CopyFromRecordset rs
    
    ' 關閉 Recordset
    rs.Close
    Set rs = Nothing
    
    ' 計算 sum155930402 與 sum255930402
    sum155930402 = 0
    sum255930402 = 0
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    Set rngs = xlsht.Range("C3:C" & lastRow)
    
    For Each rng In rngs
        ' 假設 FieldCode 資料存放在 C 欄
        If rng.Value = "155930402" Then
            ' 假設 Content 值在同一列往右兩欄（E 欄）
            sum155930402 = sum155930402 + rng.Offset(0, 2).Value
        ElseIf rng.Value = "255930402" Then
            sum255930402 = sum255930402 + rng.Offset(0, 2).Value
        End If
    Next rng
    
    Debug.Print "sum155930402: " & sum155930402
    Debug.Print "sum255930402: " & sum255930402
    
    ' 寫入合計數值到 H3 與 H4
    xlsht.Cells(3, "H").Value = sum155930402
    xlsht.Cells(4, "H").Value = sum255930402
End Sub
    
    
    
' ⸻

' 程式說明
'     1.	IsValidDataMonth 函數
'     •	此函數接收一個字串參數，利用正規表示式檢查該字串是否符合「yyyy/mm」格式。
'     •	回傳 True 表示格式正確，否則回傳 False。
'     2.	主程式 CNY1
'     •	在使用 InputBox 取得使用者輸入後，調用 IsValidDataMonth 函數進行檢核。
'     •	若格式不正確則顯示錯誤訊息，並重複要求輸入；若正確則繼續執行後續流程。

' 這樣一來，資料月份的檢核機制就獨立成一個函數，使主程式更為簡潔，同時也提高了程式的彈性與可重用性。