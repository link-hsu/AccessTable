Option Explicit

'=== Global Config Settings ===
Public gDataMonthString As String    ' 使用者輸入的資料月份
Public gDBPath As String               ' 資料庫路徑
Public gReportFolder As String         ' 原始申報報表 Excel 檔所在資料夾
Public gOutputFolder As String         ' 更新後另存新檔的資料夾
Public gReportNames As Variant         ' 報表名稱陣列
Public gReports As Collection          ' Declare Collections that Save all instances of clsReport

'=== 主流程入口 ===
Public Sub Main()
    Dim isInputValid As Boolean
    isInputValid = False
    Do
        gDataMonthString = InputBox("請輸入資料月份 (格式: yyyy/mm):", "輸入資料月份")
        If IsValidDataMonth(gDataMonthString) Then
            isInputValid = True
        ElseIf Trim(gDataMonthString) = "" Then
            MsgBox "請輸入報表資料所屬的年度/月份 (例如: 2024/01)", vbExclamation, "輸入錯誤"
        Else
            MsgBox "格式錯誤，請輸入正確格式 (yyyy/mm)", vbExclamation, "格式錯誤"
        End If
    Loop Until isInputValid
    
    ' 設定其他 config 參數（請根據實際環境調整）
    gDBPath = ThisWorkbook.Path & "\" & Sheets("ControlPanel").Range("DBsPathFileName").Value
    gReportFolder = "C:\申報報表\原始檔\"      ' 調整為實際路徑
    gOutputFolder = "C:\申報報表\Processed\"      ' 調整為實際路徑
    gReportNames = Array("CNY1", "FB2", "FB3", "FM11", "FM13", "AI821", "Table2", "FB5_FB5A", "FM2", "FM10", "F1_F2", "Table41", "AI602", "AI240")
    
    ' (a) 先初始化所有報表，並將初始資料寫入 Access（例如寫入 ReportConfig 資料表）
    Call InitializeReports
    ' (b) 各報表分別進行資料處理（各自邏輯分離）
    Call Process_CNY1
    Call Process_MM4901B
    Call Process_AC5601
    Call Process_AC5602
    ' (c) 最後更新申報 Excel 檔案並另存新檔
    Call UpdateExcelReports
    MsgBox "完成全部流程處理"
End Sub

'=== (a) 初始化所有報表並將初始資料寫入 Access ===
Public Sub InitializeReports()
    Dim rpt As clsReport
    Dim rptName As Variant, key As Variant
    Set gReports = New Collection
    For Each rptName In gReportNames
        Set rpt = New clsReport
        rpt.Init rptName
        gReports.Add rpt, rptName
        ' 將各工作表內每個欄位初始設定寫入 Access（資料表名稱例如 ReportConfig）
        Dim wsPositions As Object
        Dim combinedPositions As Object
        Set combinedPositions = rpt.GetAllFieldPositions ' 合併所有工作表，Key 格式 "wsName|fieldName"
        For Each key In combinedPositions.Keys
            InsertIntoTable gDBPath, "MonthlyDeclarationReport", gDataMonthString, rptName, key, 0, combinedPositions(key)
        Next key
    Next rptName
    MsgBox "報表初始化及初始資料建立完成"
End Sub

'=== (b) 各報表獨立處理邏輯 ===

Public Sub Process_CNY1()
    'Equal Setting
    'Fetch Query Access DB table
    Dim dataArr As Variant

    'Declare worksheet and handle data
    Dim controlSheet As Worksheet
    Dim xlsht As Worksheet

    Dim i As Integer, j As Integer
    Dim lastRow As Integer

    Dim reportTitle As String
    Dim queryTable As String

    'Setting class clsReport
    Dim rpt As clsReport
    '==========
    Set rpt = gReports("CNY1")
    

    '--------------
    'Unique Setting
    '--------------
    Dim rngs As Range
    Dim rng As Range

    Dim fxReceive As Double
    Dim fxPay As Double
    '-------
    '==========
    reportTitle = "CNY1"
    queryTable = "CNY1_DBU_AC5601"

    ' Equal Setting
    ' 取得 Query 資料 (查詢名稱 "CNY1_Query")
    '==========
    ' dataArr = GetAccessDataAsArray(gDBPath, queryTable, gDataMonthString)
    dataArr = GetAccessDataAsArray(gDBPath, queryTable)
    If UBound(dataArr) < 2 Then
        '==========
        MsgBox "CNY1 查詢資料不完整！", vbExclamation
        Exit Sub
    End If
    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    '==========
    xlsht.Range("A:E").ClearContents
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row

    '==========
    'Paste Queyr Table inot Excel
    For j = 0 To UBound(dataArr, 2)
        For i = 0 To UBound(dataArr, 1)
            xlsht.Cells(i + 1, j + 1).Value = dataArr(i, j)
        Next i
    Next j


    '==========
    'Handle Access Table Data
    ' Compute 期收/付遠匯款-換匯遠期
    fxReceive = 0
    fxPay = 0
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    Set rngs = xlsht.Range("C2:C" & lastRow)

    '
    For Each rng In rngs
        If CStr(rng.Value) = "155930402" Then
            fxReceive = fxReceive + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "255930402" Then
            fxPay = fxPay + rng.Offset(0, 2).Value
        End If
    Next rng

    fxReceive = Round(fxReceive / 1000, 0)
    fxPay = Round(fxPay / 1000, 0)
    
    xlsht.Range("CNY1_其他金融資產_淨額").Value = fxReceive
    rpt.SetField "CNY1", "CNY1_其他金融資產_淨額", fxReceive

    xlsht.Range("CNY1_其他").Value = fxReceive
    rpt.SetField "CNY1", "CNY1_其他", fxReceive

    xlsht.Range("CNY1_資產總計").Value = fxReceive
    rpt.SetField "CNY1", "CNY1_資產總計", fxReceive

    xlsht.Range("CNY1_其他金融負債").Value = fxPay
    rpt.SetField "CNY1", "CNY1_其他金融負債", fxPay

    xlsht.Range("CNY1_其他什項金融負債").Value = fxPay
    rpt.SetField "CNY1", "CNY1_其他什項金融負債", fxPay

    xlsht.Range("CNY1_負債總計").Value = fxPay
    rpt.SetField "CNY1", "CNY1_負債總計", fxPay
    
    xlsht.Range("T2:T100").NumberFormat = "#,##,##.00"
    

    ' 驗證並更新資料庫（以 Insert 示範，若記錄已存在則應用 UPDATE）
    If rpt.ValidateFields() Then
        Dim key As Variant
        Dim allValues As Object, allPositions As Object

        Set allValues = rpt.GetAllFieldValues()  ' key 格式 "wsName|fieldName"
        Set allPositions = rpt.GetAllFieldPositions()

        For Each key In allValues.Keys
            UpdateRecord gDBPath, gDataMonthString, rpt.ReportName, key, allPositions(key), CDbl(allValues(key))
        Next key
    End If
End Sub


Public Sub Process_FB2()
    'Equal Setting
    'Fetch Query Access DB table
    Dim dataArr As Variant

    'Declare worksheet and handle data
    Dim controlSheet As Worksheet
    Dim xlsht As Worksheet

    Dim i As Integer, j As Integer
    Dim lastRow As Integer

    Dim reportTitle As String
    Dim queryTable As String

    'Setting class clsReport
    Dim rpt As clsReport
    '==========
    Set rpt = gReports("FB2")
    '==========
    reportTitle = "FB2"
    queryTable = "FB2_OBU_AC4620B"

    ' Equal Setting
    ' 取得 Query 資料 (查詢名稱 "CNY1_Query")
    '==========
    ' dataArr = GetAccessDataAsArray(gDBPath, queryTable, gDataMonthString)
    dataArr = GetAccessDataAsArray(gDBPath, queryTable)
    If UBound(dataArr) < 2 Then
        '==========
        MsgBox "FB2 查詢資料不完整！", vbExclamation
        Exit Sub
    End If
    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    '==========
    xlsht.Range("A:F").ClearContents
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row

    '==========
    'Paste Queyr Table inot Excel
    For j = 0 To UBound(dataArr, 2)
        For i = 0 To UBound(dataArr, 1)
            xlsht.Cells(i + 1, j + 1).Value = dataArr(i, j)
        Next i
    Next j


    '==========
    '--------------
    'Unique Setting
    '--------------
    Dim rngs As Range
    Dim rng As Range

    Dim loanAmount As Double
    Dim loanInterest As Double
    Dim totalAsset As Double

    'Handle Access Table Data
    ' Compute 期收/付遠匯款-換匯遠期
    loanAmount = 0
    loanInterest = 0
    totalAsset = 0
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    Set rngs = xlsht.Range("C2:C" & lastRow)

    '
    For Each rng In rngs
        If CStr(rng.Value) = "115037101" Then
            loanAmount = loanAmount + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "115037115" Then
            loanAmount = loanAmount + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "130152771" Then
            loanInterest = loanInterest + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "130152777" Then
            loanInterest = loanInterest + rng.Offset(0, 2).Value
        End If
    Next rng

    loanAmount = Round(loanAmount / 1000, 0)
    loanInterest = Round(loanInterest / 1000, 0)
    totalAsset = loanAmount + loanInterest
    
    xlsht.Range("FB2_存放及拆借同業").Value = loanAmount
    rpt.SetField "FOA", "FB2_存放及拆借同業", loanAmount

    xlsht.Range("FB2_拆放銀行同業").Value = loanAmount
    rpt.SetField "FOA", "FB2_拆放銀行同業", loanAmount

    xlsht.Range("FB2_應收款項_淨額").Value = loanInterest
    rpt.SetField "FOA", "FB2_應收款項_淨額", loanInterest

    xlsht.Range("FB2_應收利息").Value = loanInterest
    rpt.SetField "FOA", "FB2_應收利息", loanInterest

    xlsht.Range("FB2_資產總計").Value = totalAsset
    rpt.SetField "FOA", "FB2_資產總計", totalAsset
    
    xlsht.Range("T2:T100").NumberFormat = "#,##,##.00"
    

    ' 驗證並更新資料庫（以 Insert 示範，若記錄已存在則應用 UPDATE）
    If rpt.ValidateFields() Then
        Dim key As Variant
        Dim allValues As Object, allPositions As Object

        Set allValues = rpt.GetAllFieldValues()  ' key 格式 "wsName|fieldName"
        Set allPositions = rpt.GetAllFieldPositions()

        For Each key In allValues.Keys
            UpdateRecord gDBPath, gDataMonthString, rpt.ReportName, key, allPositions(key), CDbl(allValues(key))
        Next key
    End If
End Sub



Public Sub Process_FB3()
    'Equal Setting
    'Fetch Query Access DB table
    Dim dataArr_1 As Variant
    Dim dataArr_2 As Variant

    'Declare worksheet and handle data
    Dim controlSheet As Worksheet
    Dim xlsht As Worksheet

    Dim i As Integer, j As Integer
    Dim lastRow As Integer

    Dim reportTitle As String
    Dim queryTable_1 As String
    Dim queryTable_2 As String

    'Setting class clsReport
    Dim rpt As clsReport
    '==========
    Set rpt = gReports("FB3")
    
    '==========
    reportTitle = "FB3"
    queryTable_1 = "FB3_OBU_MM4901B_LIST"
    queryTable_2 = "FB3_OBU_MM4901B_SUM"

    ' Equal Setting
    ' 取得 Query 資料 (查詢名稱 "CNY1_Query")
    '==========
    ' dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1, gDataMonthString)
    ' dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2, gDataMonthString)
    dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1)
    dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2)

    If UBound(dataArr_1) < 2 Or UBound(dataArr_2) < 2 Then
        '==========
        MsgBox "FB3 查詢資料不完整！", vbExclamation
        Exit Sub
    End If
    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    '==========
    xlsht.Range("A:K").ClearContents
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row

    '==========
    'Paste Queyr Table inot Excel
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


    '--------------
    'Unique Setting
    '--------------
    Dim rngs As Range
    Dim rng As Range

    Dim assetTW As Double
    Dim liabilityTW As Double

    '==========
    'Handle Access Table Data
    ' Compute 期收/付遠匯款-換匯遠期
    assetTW = 0
    liabilityTW = 0

    Set rngs = xlsht.Range("J1:K1")

    For Each rng In rngs
        If CStr(rng.Value) = "Sum_MP" Then
            assetTW = assetTW + rng.Offset(1, 0).Value
        ElseIf CStr(rng.Value) = "Sum_MT" Then
            liabilityTW = liabilityTW + rng.Offset(1, 0).Value
        End If
    Next rng

    assetTW = Round(assetTW / 1000, 0)
    liabilityTW = Round(liabilityTW / 1000, 0)
    
    xlsht.Range("FB3_存放及拆借同業_資產面_台灣地區").Value = assetTW
    rpt.SetField "FOA", "FB3_存放及拆借同業_資產面_台灣地區", assetTW

    xlsht.Range("FB3_同業存款及拆放_負債面_台灣地區").Value = liabilityTW
    rpt.SetField "FOA", "FB3_同業存款及拆放_負債面_台灣地區", liabilityTW
    
    xlsht.Range("T2:T100").NumberFormat = "#,##,##.00"
    

    ' 驗證並更新資料庫（以 Insert 示範，若記錄已存在則應用 UPDATE）
    If rpt.ValidateFields() Then
        Dim key As Variant
        Dim allValues As Object, allPositions As Object

        Set allValues = rpt.GetAllFieldValues()  ' key 格式 "wsName|fieldName"
        Set allPositions = rpt.GetAllFieldPositions()

        For Each key In allValues.Keys
            UpdateRecord gDBPath, gDataMonthString, rpt.ReportName, key, allPositions(key), CDbl(allValues(key))
        Next key
    End If
End Sub

Public Sub Process_FB3A()
    'Equal Setting
    'Fetch Query Access DB table
    Dim dataArr As Variant

    'Declare worksheet and handle data
    Dim controlSheet As Worksheet
    Dim xlsht As Worksheet

    Dim i As Integer, j As Integer
    Dim lastRow As Integer

    Dim reportTitle As String
    Dim queryTable As String

    'Setting class clsReport
    Dim rpt As clsReport
    '==========
    Set rpt = gReports("FB3A")

    '==========
    reportTitle = "FB3A"
    queryTable = "FB3A_OBU_MM4901B"

    ' Equal Setting
    ' 取得 Query 資料 (查詢名稱 "CNY1_Query")
    '==========
    ' dataArr = GetAccessDataAsArray(gDBPath, queryTable, gDataMonthString)
    dataArr = GetAccessDataAsArray(gDBPath, queryTable)
    If UBound(dataArr) < 1 Then
        '==========
        MsgBox "FB3A 查詢資料不完整！", vbExclamation
        Exit Sub
    End If
    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    '==========
    xlsht.Range("A:J").ClearContents
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row

    '==========
    'Paste Queyr Table inot Excel
    For j = 0 To UBound(dataArr, 2)
        For i = 0 To UBound(dataArr, 1)
            xlsht.Cells(i + 1, j + 1).Value = dataArr(i, j)
        Next i
    Next j

    '--------------
    'Unique Setting
    '--------------
    '===============================
    Dim BankCode As Variant
    Dim CounterParty As String, Category As String
    Dim Amount As Double

    Dim targetRow As Long
    Dim targetCol As String
    
    ' 設定第二部份記錄的起始列（Row 10）
    targetRow = 10
    ' 逐列處理原始資料（從第二列開始）
    For i = 2 To lastRow
        ' 讀取原始資料欄位值（依照題目定義的欄位順序）
        ' 原始資料欄位：
        ' A: DataID
        ' B: DataMonthString
        ' C: DealDate
        ' D: DealID
        ' E: CounterParty
        ' F: MaturityDate
        ' G: CurrencyType
        ' H: Amount
        ' I: Category
        ' J: BankCode
        
        '銀行代碼
        BankCode = xlsht.Cells(i, "J").Value        
        'CounterParty
        CounterParty = xlsht.Cells(i, "E").Value
        ' 金額
        Amount = Round(xlsht.Cells(i, "H").value / 1000, 0)
        ' 類別 
        Category = xlsht.Cells(i, "I").Value               
        'TWTP_MP / OBU_MP / TWTP_MT / OBU_MT
        
        xlsht.Cells(i, "K").Value = BankCode             ' K：BankCode
        xlsht.Cells(i, "L").Value = CounterParty         ' L：CounterParty
        
        ' 根據 Category 將金額填入對應分類欄位
        Select Case Category
            Case "DBU_MP"
                xlsht.Cells(i, "M").Value = Amount      ' M：DBU_MP
            Case "OBU_MP"
                xlsht.Cells(i, "N").Value = Amount      ' N：OBU_MP
            Case "DBU_MT"
                xlsht.Cells(i, "O").Value = Amount      ' O：DBU_MT
            Case "OBU_MT"
                xlsht.Cells(i, "P").Value = Amount      ' P：OBU_MT
        End Select
        

        ' 二、記錄儲存格位置和數值（輸出位置由 Row 10 開始）
        ' 這邊假設：BankCode 記錄在 C 欄；金額根據 Category 記錄在 E (TWTP_MP) / F (OBU_MP) / G (TWTP_MT) / H (OBU_MT)

        Select Case Category
            Case "DBU_MP"
                targetCol = "E"
            Case "OBU_MP"
                targetCol = "F"
            Case "DBU_MT"
                targetCol = "G"
            Case "OBU_MT"
                targetCol = "H"
        End Select

        xlsht.Cells(i, "Q").Value =  targetCol & CStr(targetRow)
        ' rpt.SetField "FOA", "FB3A_BankCode", "C" & CStr(targetRow), BankCode
        ' rpt.SetField "FOA", "FB3A_Amount", targetCol & CStr(targetRow), Amount

        AddDynamicField "FOA", "FB3A_BankCode", "C" & CStr(targetRow), BankCode
        AddDynamicField "FOA", "FB3A_Amount", targetCol & CStr(targetRow), Amount
        
        InsertIntoTable gDBPath, "MonthlyDeclarationReport", gDataMonthString, "FB3A", "FOA|FB3A_BankCode", BankCode, "C" & CStr(targetRow)
        InsertIntoTable gDBPath, "MonthlyDeclarationReport", gDataMonthString, "FB3A", "FOA|FB3A_Amount", Amount, targetCol & CStr(targetRow)
        
        targetRow = targetRow + 1
    Next i

    xlsht.Range("M2:M100").NumberFormat = "#,##,##.00"
    xlsht.Range("N2:N100").NumberFormat = "#,##,##.00"
    xlsht.Range("O2:O100").NumberFormat = "#,##,##.00"
    xlsht.Range("P2:P100").NumberFormat = "#,##,##.00"

    '==============================
End Sub


'尚無有交易紀錄
Public Sub Process_FM5()
    'Equal Setting
    'Fetch Query Access DB table
    Dim dataArr As Variant

    'Declare worksheet and handle data
    Dim controlSheet As Worksheet
    Dim xlsht As Worksheet

    Dim i As Integer, j As Integer
    Dim lastRow As Integer

    Dim reportTitle As String
    Dim queryTable As String

    'Setting class clsReport
    Dim rpt As clsReport
    '==========
    Set rpt = gReports("FM5")
    

    '--------------
    'Unique Setting
    '--------------
    Dim rngs As Range
    Dim rng As Range

    Dim fxReceive As Double
    Dim fxPay As Double
    '-------
    '==========
    reportTitle = "FM5"
    queryTable = "FM5_OBU_FC9450B"

    ' Equal Setting
    ' 取得 Query 資料 (查詢名稱 "CNY1_Query")
    '==========
    ' dataArr = GetAccessDataAsArray(gDBPath, queryTable, gDataMonthString)
    dataArr = GetAccessDataAsArray(gDBPath, queryTable)
    If UBound(dataArr) < 1 Then
        '==========
        MsgBox "FM5 查詢資料不完整！", vbExclamation
        Exit Sub
    End If
    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    '==========
    xlsht.Range("A:I").ClearContents
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row

    '==========
    'Paste Queyr Table inot Excel
    For j = 0 To UBound(dataArr, 2)
        For i = 0 To UBound(dataArr, 1)
            xlsht.Cells(i + 1, j + 1).Value = dataArr(i, j)
        Next i
    Next j


    '==========
    'Handle Access Table Data
    ' Compute 期收/付遠匯款-換匯遠期
    fxReceive = 0
    fxPay = 0
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    Set rngs = xlsht.Range("C3:C" & lastRow)

    '
    For Each rng In rngs
        If CStr(rng.Value) = "155930402" Then
            fxReceive = fxReceive + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "255930402" Then
            fxPay = fxPay + rng.Offset(0, 2).Value
        End If
    Next rng

    fxReceive = Round(fxReceive / 1000, 0)
    fxPay = Round(fxPay / 1000, 0)
    
    xlsht.Range("CNY1_其他金融資產_淨額").Value = fxReceive
    rpt.SetField "CNY1", "CNY1_其他金融資產_淨額", fxReceive

    xlsht.Range("CNY1_其他").Value = fxReceive
    rpt.SetField "CNY1", "CNY1_其他", fxReceive

    xlsht.Range("CNY1_資產總計").Value = fxReceive
    rpt.SetField "CNY1", "CNY1_資產總計", fxReceive

    xlsht.Range("CNY1_其他金融負債").Value = fxPay
    rpt.SetField "CNY1", "CNY1_其他金融負債", fxPay

    xlsht.Range("CNY1_其他什項金融負債").Value = fxPay
    rpt.SetField "CNY1", "CNY1_其他什項金融負債", fxPay

    xlsht.Range("CNY1_負債總計").Value = fxPay
    rpt.SetField "CNY1", "CNY1_負債總計", fxPay
    
    xlsht.Range("T2:T100").NumberFormat = "#,##,##.00"
    

    ' 驗證並更新資料庫（以 Insert 示範，若記錄已存在則應用 UPDATE）
    If rpt.ValidateFields() Then
        Dim key As Variant
        Dim allValues As Object, allPositions As Object

        Set allValues = rpt.GetAllFieldValues()  ' key 格式 "wsName|fieldName"
        Set allPositions = rpt.GetAllFieldPositions()

        For Each key In allValues.Keys
            UpdateRecord gDBPath, gDataMonthString, rpt.ReportName, key, allPositions(key), CDbl(allValues(key))
        Next key
    End If
End Sub

Public Sub Process_FM11()
    'Equal Setting
    'Fetch Query Access DB table
    Dim dataArr As Variant

    'Declare worksheet and handle data
    Dim controlSheet As Worksheet
    Dim xlsht As Worksheet

    Dim i As Integer, j As Integer
    Dim lastRow As Integer

    Dim reportTitle As String
    Dim queryTable As String

    'Setting class clsReport
    Dim rpt As clsReport
    '==========
    Set rpt = gReports("FM11")
    
    '==========
    reportTitle = "FM11"
    queryTable = "FM11_OBU_AC5411B"

    ' Equal Setting
    ' 取得 Query 資料 (查詢名稱 "CNY1_Query")
    '==========
    ' dataArr = GetAccessDataAsArray(gDBPath, queryTable, gDataMonthString)
    dataArr = GetAccessDataAsArray(gDBPath, queryTable)
    If UBound(dataArr) < 2 Then
        '==========
        MsgBox "FM11 查詢資料不完整！", vbExclamation
        Exit Sub
    End If
    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    '==========
    xlsht.Range("A:E").ClearContents
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row

    '==========
    'Paste Queyr Table inot Excel
    For j = 0 To UBound(dataArr, 2)
        For i = 0 To UBound(dataArr, 1)
            xlsht.Cells(i + 1, j + 1).Value = dataArr(i, j)
        Next i
    Next j

    '--------------
    'Unique Setting
    '--------------
    Dim rngs As Range
    Dim rng As Range

    Dim foreignInterestRevenue As Double
    Dim reversalImpairmentPL As Double
    Dim valuationImpairmentLoss As Double
    Dim domesticInterestRevenue As Double
    '-------

    '==========
    'Handle Access Table Data
    ' Compute 期收/付遠匯款-換匯遠期

    foreignInterestRevenue = 0
    reversalImpairmentPL = 0
    valuationImpairmentLoss = 0
    domesticInterestRevenue = 0
    
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    Set rngs = xlsht.Range("C2:C" & lastRow)

    For Each rng In rngs
        If CStr(rng.Value) = "410331203" Then
            foreignInterestRevenue = foreignInterestRevenue + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "410331211" Then
            foreignInterestRevenue = foreignInterestRevenue + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "410331212" Then
            foreignInterestRevenue = foreignInterestRevenue + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "410331229" Then
            foreignInterestRevenue = foreignInterestRevenue + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "410332203" Then
            foreignInterestRevenue = foreignInterestRevenue + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "410332211" Then
            foreignInterestRevenue = foreignInterestRevenue + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "410332212" Then
            foreignInterestRevenue = foreignInterestRevenue + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "410332229" Then
            foreignInterestRevenue = foreignInterestRevenue + rng.Offset(0, 2).Value



        ElseIf CStr(rng.Value) = "450110105" Then
            reversalImpairmentPL = reversalImpairmentPL + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "450110125" Then
            reversalImpairmentPL = reversalImpairmentPL + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "450110127" Then
            reversalImpairmentPL = reversalImpairmentPL + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "450110143" Then
            reversalImpairmentPL = reversalImpairmentPL + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "450130105" Then
            reversalImpairmentPL = reversalImpairmentPL + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "450130125" Then
            reversalImpairmentPL = reversalImpairmentPL + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "450130147" Then
            reversalImpairmentPL = reversalImpairmentPL + rng.Offset(0, 2).Value

        ElseIf CStr(rng.Value) = "550110105" Then
            valuationImpairmentLoss = valuationImpairmentLoss + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "550110125" Then
            valuationImpairmentLoss = valuationImpairmentLoss + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "550110127" Then
            valuationImpairmentLoss = valuationImpairmentLoss + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "550110143" Then
            valuationImpairmentLoss = valuationImpairmentLoss + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "550130105" Then
            valuationImpairmentLoss = valuationImpairmentLoss + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "550130127" Then
            valuationImpairmentLoss = valuationImpairmentLoss + rng.Offset(0, 2).Value

        ElseIf CStr(rng.Value) = "410997201" Then
            domesticInterestRevenue = domesticInterestRevenue + rng.Offset(0, 2).Value
        End If
    Next rng

    foreignInterestRevenue = Round(foreignInterestRevenue / 1000, 0)
    reversalImpairmentPL = Round(reversalImpairmentPL / 1000, 0)
    valuationImpairmentLoss = Round(valuationImpairmentLoss / 1000, 0)
    domesticInterestRevenue = Round(domesticInterestRevenue / 1000, 0)
    
    xlsht.Range("FM11_一利息股息收入_利息_其他").Value = foreignInterestRevenue
    rpt.SetField "FOA", "FM11_一利息股息收入_利息_其他", foreignInterestRevenue

    xlsht.Range("FM11_五證券投資評價及減損迴轉利益_一年期以上之債權證券").Value = reversalImpairmentPL
    rpt.SetField "FOA", "FM11_五證券投資評價及減損迴轉利益_一年期以上之債權證券", reversalImpairmentPL

    xlsht.Range("FM11_五證券投資評價及減損損失_一年期以上之債權證券").Value = valuationImpairmentLoss
    rpt.SetField "FOA", "FM11_五證券投資評價及減損損失_一年期以上之債權證券", valuationImpairmentLoss

    xlsht.Range("FM11_一利息收入_自中華民國境內其他客戶").Value = domesticInterestRevenue
    rpt.SetField "FOA", "FM11_一利息收入_自中華民國境內其他客戶", domesticInterestRevenue
    
    xlsht.Range("T2:T100").NumberFormat = "#,##,##.00"
    

    ' 驗證並更新資料庫（以 Insert 示範，若記錄已存在則應用 UPDATE）
    If rpt.ValidateFields() Then
        Dim key As Variant
        Dim allValues As Object, allPositions As Object

        Set allValues = rpt.GetAllFieldValues()  ' key 格式 "wsName|fieldName"
        Set allPositions = rpt.GetAllFieldPositions()

        For Each key In allValues.Keys
            UpdateRecord gDBPath, gDataMonthString, rpt.ReportName, key, allPositions(key), CDbl(allValues(key))
        Next key
    End If
End Sub


Public Sub Process_FM13()
    'Equal Setting
    'Fetch Query Access DB table
    Dim dataArr_1 As Variant
    Dim dataArr_2 As Variant

    'Declare worksheet and handle data
    Dim controlSheet As Worksheet
    Dim xlsht As Worksheet

    Dim i As Integer, j As Integer
    Dim lastRow As Integer

    Dim reportTitle As String
    Dim queryTable_1 As String
    Dim queryTable_2 As String

    'Setting class clsReport
    Dim rpt As clsReport
    '==========
    Set rpt = gReports("FM13")
    

    '==========
    reportTitle = "FM13"
    queryTable_1 = "FM13_FXDebtEvaluation_Subtotal_FVandAdjust"
    queryTable_2 = "FM13_FXDebtEvaluation_Subtotal_Impairment"

    ' Equal Setting
    ' 取得 Query 資料 (查詢名稱 "CNY1_Query")
    '==========
    ' dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1, gDataMonthString)
    ' dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2, gDataMonthString)
    dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1)
    dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2)
    If UBound(dataArr_1) < 1 Or UBound(dataArr_2) < 1 Then
        '==========
        MsgBox "FM13 查詢資料不完整！", vbExclamation
        Exit Sub
    End If
    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    '==========
    xlsht.Range("A:E").ClearContents
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row

    '==========
    'Paste Queyr Table inot Excel
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

    '--------------
    'Unique Setting
    '--------------
    Dim rngs As Range
    Dim rng As Range

    Dim bookValue_HongKong As Double
    Dim bookValue_Korea As Double
    Dim bookValue_Thailand As Double
    Dim bookValue_Malaysia As Double
    Dim bookValue_Philippines As Double
    Dim bookValue_Indonesia As Double
    
    Dim valueAdjsut As Double
    Dim accumulateImpairment As Double
    '==========
    'Handle Access Table Data
    ' Compute 期收/付遠匯款-換匯遠期
    bookValue_HongKong = 0
    bookValue_Korea = 0
    bookValue_Thailand = 0
    bookValue_Malaysia = 0
    bookValue_Philippines = 0
    bookValue_Indonesia = 0
    
    valueAdjsut = 0
    accumulateImpairment = 0

    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    Set rngs = xlsht.Range("A2:A" & lastRow)
    For Each rng In rngs
        valueAdjsut = valueAdjsut + rng.Offset(0, 2).Value
        If CStr(rng.Value) = "香港" Then
            bookValue_HongKong = bookValue_HongKong + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "韓國" Then
            bookValue_Korea = bookValue_Korea + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "泰國" Then
            bookValue_Thailand = bookValue_Thailand + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "馬來西亞" Then
            bookValue_Malaysia = bookValue_Malaysia + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "菲律賓" Then
            bookValue_Philippines = bookValue_Philippines + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "印尼" Then
            bookValue_Indonesia = bookValue_Indonesia + rng.Offset(0, 1).Value
        End If
    Next rng

    lastRow = xlsht.Cells(xlsht.Rows.Count, 4).End(xlUp).Row
    Set rngs = xlsht.Range("D2:D" & lastRow)
    For Each rng In rngs
        accumulateImpairment = accumulateImpairment + rng.Offset(0, 2).Value
    Next rng

    bookValue_HongKong = Round(bookValue_HongKong / 1000, 0)
    bookValue_Korea = Round(bookValue_Korea / 1000, 0)
    bookValue_Thailand = Round(bookValue_Thailand / 1000, 0)
    bookValue_Malaysia = Round(bookValue_Malaysia / 1000, 0)
    bookValue_Philippines = Round(bookValue_Philippines / 1000, 0)
    bookValue_Indonesia = Round(bookValue_Indonesia / 1000, 0)
    
    valueAdjsut = Round(ABs(valueAdjsut) / 1000, 0)
    accumulateImpairment = Round(accumulateImpairment / 1000, 0)
    
    xlsht.Range("FM13_OBU_香港_債票券投資").Value = bookValue_HongKong
    rpt.SetField "FOA", "FM13_OBU_香港_債票券投資", bookValue_HongKong

    xlsht.Range("FM13_OBU_韓國_債票券投資").Value = bookValue_Korea
    rpt.SetField "FOA", "FM13_OBU_韓國_債票券投資", bookValue_Korea

    xlsht.Range("FM13_OBU_泰國_債票券投資").Value = bookValue_Thailand
    rpt.SetField "FOA", "FM13_OBU_泰國_債票券投資", bookValue_Thailand

    xlsht.Range("FM13_OBU_馬來西亞_債票券投資").Value = bookValue_Malaysia
    rpt.SetField "FOA", "FM13_OBU_馬來西亞_債票券投資", bookValue_Malaysia

    xlsht.Range("FM13_OBU_菲律賓_債票券投資").Value = bookValue_Philippines
    rpt.SetField "FOA", "FM13_OBU_菲律賓_債票券投資", bookValue_Philippines

    xlsht.Range("FM13_OBU_印尼_債票券投資").Value = bookValue_Indonesia
    rpt.SetField "FOA", "FM13_OBU_印尼_債票券投資", bookValue_Indonesia

    xlsht.Range("FM13_OBU_債票券投資_評價調整").Value = valueAdjsut
    rpt.SetField "FOA", "FM13_OBU_債票券投資_評價調整", valueAdjsut

    xlsht.Range("FM13_OBU_債票券投資_累計減損").Value = accumulateImpairment
    rpt.SetField "FOA", "FM13_OBU_債票券投資_累計減損", accumulateImpairment
    
    xlsht.Range("T2:T100").NumberFormat = "#,##,##.00"
    

    ' 驗證並更新資料庫（以 Insert 示範，若記錄已存在則應用 UPDATE）
    If rpt.ValidateFields() Then
        Dim key As Variant
        Dim allValues As Object, allPositions As Object

        Set allValues = rpt.GetAllFieldValues()  ' key 格式 "wsName|fieldName"
        Set allPositions = rpt.GetAllFieldPositions()

        For Each key In allValues.Keys
            UpdateRecord gDBPath, gDataMonthString, rpt.ReportName, key, allPositions(key), CDbl(allValues(key))
        Next key
    End If
End Sub



Public Sub Process_AI821()
    'Equal Setting
    'Fetch Query Access DB table
    Dim dataArr_1 As Variant
    Dim dataArr_2 As Variant

    'Declare worksheet and handle data
    Dim controlSheet As Worksheet
    Dim xlsht As Worksheet

    Dim i As Integer, j As Integer
    Dim lastRow As Integer

    Dim reportTitle As String
    Dim queryTable_1 As String
    Dim queryTable_2 As String

    'Setting class clsReport
    Dim rpt As clsReport
    '==========
    Set rpt = gReports("AI821")
    
    '==========
    reportTitle = "AI821"
    queryTable_1 = "AI821_OBU_MM4901B_LIST"
    queryTable_2 = "AI821_OBU_MM4901B_SUM"

    ' Equal Setting
    ' 取得 Query 資料 (查詢名稱 "CNY1_Query")
    '==========
    ' dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1, gDataMonthString)
    ' dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2, gDataMonthString)
    dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1)
    dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2)
    If UBound(dataArr_1) < 1 Or UBound(dataArr_2) < 1 Then
        '==========
        MsgBox "AI821 查詢資料不完整！", vbExclamation
        Exit Sub
    End If
    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    '==========
    xlsht.Range("A:K").ClearContents
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row

    '==========
    'Paste Queyr Table inot Excel
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





    '--------------
    'Unique Setting
    '--------------
    Dim rngs As Range
    Dim rng As Range

    Dim domesticBank As Double
    Dim chinaBranchBank As Double
    Dim foreignBranchBank As Double
    Dim chinaBank As Double
    Dim others As Double

    '==========
    'Handle Access Table Data
    ' Compute 期收/付遠匯款-換匯遠期

    domesticBank = 0
    chinaBranchBank = 0
    foreignBranchBank = 0
    chinaBank = 0
    others = 0

    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    Set rngs = xlsht.Range("I2:I" & lastRow)

    For Each rng In rngs
        If CStr(rng.Value) = "本國銀行" Then
            domesticBank = domesticBank + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "陸銀在臺分行" Then
            chinaBranchBank = chinaBranchBank + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "外商銀行在臺分行" Then
            foreignBranchBank = foreignBranchBank + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "大陸地區銀行" Then
            chinaBank = chinaBank + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "其他" Then
            others = others + rng.Offset(0, 2).Value
        End If
    Next rng

    domesticBank = Round(domesticBank / 1000, 0)
    chinaBranchBank = Round(chinaBranchBank / 1000, 0)
    foreignBranchBank = Round(foreignBranchBank / 1000, 0)
    chinaBank = Round(chinaBank / 1000, 0)
    others = Round(others / 1000, 0)
    
    xlsht.Range("AI821_本國銀行").Value = domesticBank
    rpt.SetField "Table1", "AI821_本國銀行", domesticBank

    xlsht.Range("AI821_陸銀在臺分行").Value = chinaBranchBank
    rpt.SetField "Table1", "AI821_陸銀在臺分行", chinaBranchBank

    xlsht.Range("AI821_外商銀行在臺分行").Value = foreignBranchBank
    rpt.SetField "Table1", "AI821_外商銀行在臺分行", foreignBranchBank

    xlsht.Range("AI821_大陸地區銀行").Value = chinaBank
    rpt.SetField "Table1", "AI821_大陸地區銀行", chinaBank

    xlsht.Range("AI821_其他").Value = others
    rpt.SetField "Table1", "AI821_其他", others
    
    xlsht.Range("T2:T100").NumberFormat = "#,##,##.00"
    

    ' 驗證並更新資料庫（以 Insert 示範，若記錄已存在則應用 UPDATE）
    If rpt.ValidateFields() Then
        Dim key As Variant
        Dim allValues As Object, allPositions As Object

        Set allValues = rpt.GetAllFieldValues()  ' key 格式 "wsName|fieldName"
        Set allPositions = rpt.GetAllFieldPositions()

        For Each key In allValues.Keys
            UpdateRecord gDBPath, gDataMonthString, rpt.ReportName, key, allPositions(key), CDbl(allValues(key))
        Next key
    End If
End Sub



Public Sub Process_Table2()
    'Equal Setting
    'Fetch Query Access DB table
    Dim dataArr_1 As Variant
    Dim dataArr_2 As Variant

    'Declare worksheet and handle data
    Dim controlSheet As Worksheet
    Dim xlsht As Worksheet

    Dim i As Integer, j As Integer
    Dim lastRow As Integer

    Dim reportTitle As String
    Dim queryTable_1 As String
    Dim queryTable_2 As String

    'Setting class clsReport
    Dim rpt As clsReport
    '==========
    Set rpt = gReports("Table2")
    
    '==========
    reportTitle = "Table2"
    queryTable_1 = "表2_DBU_AC5602_TWD"
    queryTable_2 = "表2_CloseRate_USDTWD"

    ' Equal Setting
    ' 取得 Query 資料 (查詢名稱 "CNY1_Query")
    '==========
    ' dataArr = GetAccessDataAsArray(gDBPath, queryTable, gDataMonthString)
    dataArr_1 = GetAccessDataAsArray(DBsPath, queryTable_1)
    dataArr_2 = GetAccessDataAsArray(DBsPath, queryTable_2)
    If UBound(dataArr_1) < 1 Or UBound(dataArr_2) < 1  Then
        '==========
        MsgBox "Table2 查詢資料不完整！", vbExclamation
        Exit Sub
    End If
    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    '==========
    xlsht.Range("A:I").ClearContents
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row

    '==========
    'Paste Queyr Table inot Excel
    If Err.Number <> 0 Or LBound(dataArr_1) > UBound(dataArr_1) Then
        MsgBox queryTable_1 & "資料表無資料"
        GoTo ContinueLoop
    Else
        For j = 0 To UBound(dataArr_1, 2)
            For i = 0 To UBound(dataArr_1, 1)
                xlsht.Cells(i + 1, j + 1).Value = dataArr_1(i, j)
            Next i
        Next j
    End If

    If Err.Number <> 0 Or LBound(dataArr_2) > UBound(dataArr_2) Then
        MsgBox queryTable_2 & "資料表無資料"
        GoTo ContinueLoop
    Else
        For j = 0 To UBound(dataArr_2, 2)
            For i = 0 To UBound(dataArr_2, 1)
                xlsht.Cells(i + 1, j + 7).Value = dataArr_2(i, j)
            Next i
        Next j
    End If



    '==========
    'Handle Access Table Data
    ' Compute 期收/付遠匯款-換匯遠期

    '--------------
    'Unique Setting
    '--------------
    Dim rngs As Range
    Dim rng As Range

    Dim marginDeposit_TWD As Double
    Dim marginDeposit_USD As Double
    Dim rateUSDtoTWD As Double

    marginDeposit_TWD = 0
    marginDeposit_USD = 0
    rateUSDtoTWD = 0

    rateUSDtoTWD = xlsht.Range("I2").Value
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    Set rngs = xlsht.Range("C2:C" & lastRow)
    
    
    For Each rng In rngs
        If CStr(rng.Value) = "196017703" Then
            marginDeposit_TWD = marginDeposit_TWD + rng.Offset(0, 3).Value
        End If
    Next rng

    marginDeposit_TWD = Round(marginDeposit_TWD / 1000, 0)
    marginDeposit_USD = Round((marginDeposit_TWD / rateUSDtoTWD) / 1000, 0)
    
    xlsht.Range("Table2_其他").Value = marginDeposit_TWD
    rpt.SetField "FOA", "Table2_其他", marginDeposit_TWD

    xlsht.Range("Table_美元_F1").Value = marginDeposit_USD
    rpt.SetField "FOA", "Table_美元_F1", marginDeposit_USD

    xlsht.Range("Table2_美元_F3").Value = rateUSDtoTWD
    rpt.SetField "FOA", "Table2_美元_F3", rateUSDtoTWD

    xlsht.Range("Table2_美元_F4").Value = marginDeposit_TWD
    rpt.SetField "FOA", "Table2_美元_F4", marginDeposit_TWD
    
    xlsht.Range("T2:T100").NumberFormat = "#,##,##.00"
    

    ' 驗證並更新資料庫（以 Insert 示範，若記錄已存在則應用 UPDATE）
    If rpt.ValidateFields() Then
        Dim key As Variant
        Dim allValues As Object, allPositions As Object

        Set allValues = rpt.GetAllFieldValues()  ' key 格式 "wsName|fieldName"
        Set allPositions = rpt.GetAllFieldPositions()

        For Each key In allValues.Keys
            UpdateRecord gDBPath, gDataMonthString, rpt.ReportName, key, allPositions(key), CDbl(allValues(key))
        Next key
    End If
ContinueLoop:

End Sub




Public Sub Process_FB5_FB5A()
    'Equal Setting
    'Fetch Query Access DB table
    Dim dataArr As Variant

    'Declare worksheet and handle data
    Dim controlSheet As Worksheet
    Dim xlsht As Worksheet

    Dim i As Integer, j As Integer
    Dim lastRow As Integer

    Dim reportTitle As String
    Dim queryTable As String

    'Setting class clsReport
    Dim rpt As clsReport
    '==========
    Set rpt = gReports("FB5_FB5A")
    
    '==========
    reportTitle = "FB5_FB5A"
    queryTable = "FB5_FB5A_DL6320"

    ' Equal Setting
    ' 取得 Query 資料 (查詢名稱 "CNY1_Query")
    '==========
    ' dataArr = GetAccessDataAsArray(gDBPath, queryTable, gDataMonthString)
    dataArr = GetAccessDataAsArray(gDBPath, queryTable)
    If UBound(dataArr) < 1 Then
        '==========
        MsgBox "FB5_FB5A 查詢資料不完整！", vbExclamation
        Exit Sub
    End If
    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    '==========
    xlsht.Range("A:G").ClearContents
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row

    '==========
    'Paste Queyr Table inot Excel
    For j = 0 To UBound(dataArr, 2)
        For i = 0 To UBound(dataArr, 1)
            xlsht.Cells(i + 1, j + 1).Value = dataArr(i, j)
        Next i
    Next j



    '--------------
    'Unique Setting
    '--------------
    Dim rngs As Range
    Dim rng As Range

    Dim SpotToDBU_CNY As Double

    '==========
    'Handle Access Table Data
    ' Compute 期收/付遠匯款-換匯遠期
    SpotToDBU_CNY = 0

    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    Set rngs = xlsht.Range("D1:D" & lastRow)

    For Each rng In rngs
        If CStr(rng.Value) = "CNY" Then
            SpotToDBU_CNY = SpotToDBU_CNY + rng.Offset(0, 1).Value
        End If
    Next rng

    SpotToDBU_CNY = Round(SpotToDBU_CNY / 1000, 0)
    
    xlsht.Range("FB5_外匯交易_即期外匯_DBU").Value = SpotToDBU_CNY
    rpt.SetField "FOA", "FB5_外匯交易_即期外匯_DBU", SpotToDBU_CNY
    
    xlsht.Range("T2:T100").NumberFormat = "#,##,##.00"
    

    ' 驗證並更新資料庫（以 Insert 示範，若記錄已存在則應用 UPDATE）
    If rpt.ValidateFields() Then
        Dim key As Variant
        Dim allValues As Object, allPositions As Object

        Set allValues = rpt.GetAllFieldValues()  ' key 格式 "wsName|fieldName"
        Set allPositions = rpt.GetAllFieldPositions()

        For Each key In allValues.Keys
            UpdateRecord gDBPath, gDataMonthString, rpt.ReportName, key, allPositions(key), CDbl(allValues(key))
        Next key
    End If
End Sub




Public Sub Process_FM2()
    'Equal Setting
    'Fetch Query Access DB table
    Dim dataArr_1 As Variant
    Dim dataArr_2 As Variant
    Dim dataArr_3 As Variant

    'Declare worksheet and handle data
    Dim controlSheet As Worksheet
    Dim xlsht As Worksheet

    Dim i As Integer, j As Integer
    Dim lastRow As Integer

    Dim reportTitle As String
    Dim queryTable_1 As String
    Dim queryTable_2 As String
    Dim queryTable_3 As String

    'Setting class clsReport
    Dim rpt As clsReport
    '==========
    Set rpt = gReports("FM2")
    
    '==========
    reportTitle = "FM2"
    queryTable_1 = "FM2_OBU_MM4901B_LIST"
    queryTable_2 = "FM2_OBU_MM4901B_Subtotal"
    queryTable_3 = "FM2_OBU_MM4901B_Subtotal_BankCode"

    ' Equal Setting
    ' 取得 Query 資料 (查詢名稱 "CNY1_Query")
    '==========
    ' dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1, gDataMonthString)
    ' dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2, gDataMonthString)
    ' dataArr_3 = GetAccessDataAsArray(gDBPath, queryTable_3, gDataMonthString)
    dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1)
    dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2)
    dataArr_3 = GetAccessDataAsArray(gDBPath, queryTable_3)
    If UBound(dataArr_1) < 1 Or UBound(dataArr_2) < 1 Or UBound(dataArr_3) < 1 Then
        '==========
        MsgBox "FM2 查詢資料不完整！", vbExclamation
        Exit Sub
    End If
    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    '==========
    xlsht.Range("A:P").ClearContents
    xlsht.Range("Q:W200").ClearContents
    lastRow = xlsht.Cells(xlsht.Rows.Count, "M").End(xlUp).Row

    '==========
    'Paste Queyr Table inot Excel
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
            xlsht.Cells(i + 1, j + 13).Value = dataArr_3(i, j)
        Next i
    Next j



    '--------------
    'Unique Setting
    '--------------
    '===============================

    '*****這邊還要確認欄位資訊
    Dim BankCode As Variant
    Dim CounterParty As String, Category As String
    Dim Amount As Double

    Dim pasteRow As Long
    Dim targetRow As Long
    Dim targetCol As String
    
    ' 設定第二部份記錄的起始列（Row 10）
    pasteRow = 2
    targetRow = 10
    lastRow = xlsht.Cells(xlsht.Rows.Count, "M").End(xlUp).Row
    
    ' 逐列處理原始資料（從第二列開始）
    For i = 2 To lastRow
        ' 讀取原始資料欄位值（依照題目定義的欄位順序）
        ' 原始資料欄位：
        ' A: DataID
        ' B: DataMonthString
        ' C: DealDate
        ' D: DealID
        ' E: CounterParty
        ' F: MaturityDate
        ' G: CurrencyType
        ' H: Amount
        ' I: Category
        ' J: BankCode
        

        If (Not IsEmpty(xlsht.Cells(i, "P").Value)) Then
            '銀行代碼
            BankCode = xlsht.Cells(i, "P").Value        
            'CounterParty
            CounterParty = xlsht.Cells(i, "M").Value
            ' 金額
            Amount = Round(xlsht.Cells(i, "O").value / 1000, 0)
            ' 類別 
            Category = xlsht.Cells(i, "N").Value               
            'TWTP_MP / OBU_MP / TWTP_MT / OBU_MT
            
            xlsht.Cells(pasteRow, "Q").Value = BankCode             ' K：BankCode
            xlsht.Cells(pasteRow, "R").Value = CounterParty         ' L：CounterParty

            ' 根據 Category 將金額填入對應分類欄位
            Select Case Category
                Case "DBU_MP"
                    xlsht.Cells(pasteRow, "S").Value = Amount      ' M：TWTP_MP
                Case "OBU_MP"
                    xlsht.Cells(pasteRow, "T").Value = Amount      ' N：OBU_MP
                Case "DBU_MT"
                    xlsht.Cells(pasteRow, "U").Value = Amount      ' O：TWTP_MT
                Case "OBU_MT"
                    xlsht.Cells(pasteRow, "V").Value = Amount      ' P：OBU_MT
            End Select
        

            ' 二、記錄儲存格位置和數值（輸出位置由 Row 10 開始）
            ' 這邊假設：BankCode 記錄在 C 欄；金額根據 Category 記錄在 E (TWTP_MP) / F (OBU_MP) / G (TWTP_MT) / H (OBU_MT)

            Select Case Category
                Case "DBU_MP"
                    targetCol = "E"
                Case "OBU_MP"
                    targetCol = "F"
                Case "DBU_MT"
                    targetCol = "G"
                Case "OBU_MT"
                    targetCol = "H"
            End Select
            
            xlsht.Cells(pasteRow, "W").Value =  targetCol & CStr(targetRow)
            ' rpt.SetField "FOA", "FM2_BankCode", "C" & CStr(targetRow), BankCode
            ' rpt.SetField "FOA", "FM2_Amount", targetCol & CStr(targetRow), Amount

            AddDynamicField "FOA", "FM2_BankCode", "C" & CStr(targetRow), BankCode
            AddDynamicField "FOA", "FM2_Amount", targetCol & CStr(targetRow), Amount
            
            InsertIntoTable gDBPath, "MonthlyDeclarationReport", gDataMonthString, "FM2", "FOA|FM2_BankCode", BankCode, "C" & CStr(targetRow)
            InsertIntoTable gDBPath, "MonthlyDeclarationReport", gDataMonthString, "FM2", "FOA|FM2_Amount", Amount, targetCol & CStr(targetRow)
            
            pasteRow = pasteRow + 1
            targetRow = targetRow + 1
        End If
    Next i

    xlsht.Range("S2:S100").NumberFormat = "#,##,##.00"
    xlsht.Range("T2:T100").NumberFormat = "#,##,##.00"
    xlsht.Range("U2:U100").NumberFormat = "#,##,##.00"
    xlsht.Range("V2:V100").NumberFormat = "#,##,##.00"
    '==============================
End Sub



Public Sub Process_FM10()
    'Equal Setting
    'Fetch Query Access DB table
    Dim dataArr_1 As Variant
    Dim dataArr_2 As Variant

    'Declare worksheet and handle data
    Dim controlSheet As Worksheet
    Dim xlsht As Worksheet

    Dim i As Integer, j As Integer
    Dim lastRow As Integer

    Dim reportTitle As String
    Dim queryTable_1 As String
    Dim queryTable_2 As String

    'Setting class clsReport
    Dim rpt As clsReport
    '==========
    Set rpt = gReports("FM10")
    
    '==========
    reportTitle = "FM10"
    queryTable_1 = "FM10_OBU_AC4603_LIST"
    queryTable_2 = "FM10_OBU_AC4603_Subtotal"

    ' Equal Setting
    ' 取得 Query 資料 (查詢名稱 "CNY1_Query")
    '==========
    ' dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1, gDataMonthString)
    ' dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2, gDataMonthString)
    dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1)
    dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2)
    If UBound(dataArr_1) < 1 Or UBound(dataArr_2) < 1 Then
        '==========
        MsgBox "FM10 查詢資料不完整！", vbExclamation
        Exit Sub
    End If
    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    '==========
    xlsht.Range("A:H").ClearContents
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row

    '==========
    'Paste Queyr Table inot Excel
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


    '--------------
    'Unique Setting
    '--------------
    Dim rngs As Range
    Dim rng As Range

    Dim FVOCI_VALUE As Double
    Dim FVOCI_ADJUSTMENT As Double
    Dim FVOCI_NET_VALUE As Double
    Dim AC_VALUE As Double
    Dim AC_ADJUSTMENT As Double
    Dim AC_NET_VALUE As Double
    Dim otherFinancialAssets As Double

    FVOCI_VALUE = 0
    FVOCI_ADJUSTMENT = 0
    FVOCI_NET_VALUE = 0
    AC_VALUE = 0
    AC_ADJUSTMENT = 0
    AC_NET_VALUE = 0
    otherFinancialAssets = 0

    lastRow = xlsht.Cells(xlsht.Rows.Count, "G").End(xlUp).Row
    Set rngs = xlsht.Range("G2:G" & lastRow)

    For Each rng In rngs
        If CStr(rng.Value) = "FVOCI_VALUE" Then
            FVOCI_VALUE = FVOCI_VALUE + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVOCI_ADJUSTMENT" Then
            FVOCI_ADJUSTMENT = FVOCI_ADJUSTMENT + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "AC_VALUE" Then
            AC_VALUE = AC_VALUE + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "AC_ADJUSTMENT" Then
            AC_ADJUSTMENT = AC_ADJUSTMENT + rng.Offset(0, 1).Value
        End If
    Next rng

    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    Set rngs = xlsht.Range("C2:C" & lastRow)
    For Each rng In rngs
        If CStr(rng.Value) = "155517201" Then
            otherFinancialAssets = otherFinancialAssets + rng.Offset(0, 3).Value
        End If
    Next rng

    FVOCI_NET_VALUE = FVOCI_VALUE + FVOCI_ADJUSTMENT
    AC_NET_VALUE = AC_VALUE + AC_ADJUSTMENT

    FVOCI_VALUE = Round(FVOCI_VALUE / 1000, 0)
    FVOCI_NET_VALUE = Round(FVOCI_NET_VALUE / 1000, 0)
    AC_VALUE = Round(AC_VALUE / 1000, 0)
    AC_NET_VALUE = Round(AC_NET_VALUE / 1000, 0)
    otherFinancialAssets = Round(otherFinancialAssets / 1000, 0)
 
    xlsht.Range("FM10_FVOCI_總額C").Value = FVOCI_VALUE
    rpt.SetField "FOA", "FM10_FVOCI_總額C", FVOCI_VALUE

    xlsht.Range("FM10_FVOCI_淨額D").Value = FVOCI_NET_VALUE
    rpt.SetField "FOA", "FM10_FVOCI_淨額D", FVOCI_NET_VALUE

    xlsht.Range("FM10_AC_總額E").Value = AC_VALUE
    rpt.SetField "FOA", "FM10_AC_總額E", AC_VALUE

    xlsht.Range("FM10_AC_淨額F").Value = AC_NET_VALUE
    rpt.SetField "FOA", "FM10_AC_淨額F", AC_NET_VALUE

    xlsht.Range("FM10_四其他_境內_總額H").Value = otherFinancialAssets
    rpt.SetField "FOA", "FM10_四其他_境內_總額H", otherFinancialAssets

    xlsht.Range("FM10_四其他_境內_淨額I").Value = otherFinancialAssets
    rpt.SetField "FOA", "FM10_四其他_境內_淨額I", otherFinancialAssets
    
    xlsht.Range("T2:T100").NumberFormat = "#,##,##.00"
    

    ' 驗證並更新資料庫（以 Insert 示範，若記錄已存在則應用 UPDATE）
    If rpt.ValidateFields() Then
        Dim key As Variant
        Dim allValues As Object, allPositions As Object

        Set allValues = rpt.GetAllFieldValues()  ' key 格式 "wsName|fieldName"
        Set allPositions = rpt.GetAllFieldPositions()

        For Each key In allValues.Keys
            UpdateRecord gDBPath, gDataMonthString, rpt.ReportName, key, allPositions(key), CDbl(allValues(key))
        Next key
    End If
End Sub




Public Sub Process_F1_F2()
    'Equal Setting
    'Fetch Query Access DB table
    Dim dataArr_1 As Variant
    Dim dataArr_2 As Variant
    Dim dataArr_3 As Variant
    Dim dataArr_4 As Variant
    Dim dataArr_5 As Variant
    Dim dataArr_6 As Variant

    'Declare worksheet and handle data
    Dim controlSheet As Worksheet
    Dim xlsht As Worksheet

    Dim i As Integer, j As Integer
    Dim lastRow As Integer

    Dim reportTitle As String
    Dim queryTable_1 As String
    Dim queryTable_2 As String
    Dim queryTable_3 As String
    Dim queryTable_4 As String
    Dim queryTable_5 As String
    Dim queryTable_6 As String

    'Setting class clsReport
    Dim rpt As clsReport
    '==========
    Set rpt = gReports("F1_F2")
    

    '--------------
    'Unique Setting
    '--------------
    Dim rngs As Range
    Dim rng As Range

    Dim fxReceive As Double
    Dim fxPay As Double
    '-------
    '==========
    reportTitle = "F1_F2"
    queryTable_1 = "F1_Foreign_DL6850_FS"
    queryTable_2 = "F1_Foreign_DL6850_SS"
    queryTable_3 = "F1_Domestic_DL6850_FS"
    queryTable_4 = "F1_Domestic_DL6850_SS"
    queryTable_5 = "F1_CM2810_LIST"
    queryTable_6 = "F1_CM2810_Subtotal"

    ' Equal Setting
    ' 取得 Query 資料 (查詢名稱 "CNY1_Query")
    '==========
    ' dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1, gDataMonthString)
    ' dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2, gDataMonthString)
    ' dataArr_3 = GetAccessDataAsArray(gDBPath, queryTable_3, gDataMonthString)
    ' dataArr_4 = GetAccessDataAsArray(gDBPath, queryTable_4, gDataMonthString)
    ' dataArr_5 = GetAccessDataAsArray(gDBPath, queryTable_5, gDataMonthString)
    ' dataArr_5 = GetAccessDataAsArray(gDBPath, queryTable_5, gDataMonthString)
    dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1)
    dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2)
    dataArr_3 = GetAccessDataAsArray(gDBPath, queryTable_3)
    dataArr_4 = GetAccessDataAsArray(gDBPath, queryTable_4)
    dataArr_5 = GetAccessDataAsArray(gDBPath, queryTable_5)
    dataArr_6 = GetAccessDataAsArray(gDBPath, queryTable_6)
    If UBound(dataArr_1) < 1 Or UBound(dataArr_2) < 1 Or UBound(dataArr_3) < 1 Or UBound(dataArr_4) < 1 Or UBound(dataArr_5) < 1 Or UBound(dataArr_6) < 1 Then
        '==========
        MsgBox "F1_F2 查詢資料不完整！", vbExclamation
        Exit Sub
    End If
    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    '==========
    xlsht.Range("A:N").ClearContents
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row

    '==========
    'Paste Queyr Table inot Excel
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


    Dim currencies As Variant
    currencies = Array("JPY", "GBP", "CHF", "CAD", "AUD", "NZD", "SGD", "HKD", "ZAR", "SEK", "THB", "RM", "EUR", "CNY", "OTHER")
    
    ' 定義交易名稱，對應到不同資料表
    Dim transactionTypes As Variant
    transactionTypes = Array("F1_與國外金融機構及非金融機構間交易_SPOT", _
                         "F1_與國外金融機構及非金融機構間交易_SWAP", _
                         "F1_與國內金融機構間交易_SPOT", _
                         "F1_與國內金融機構間交易_SWAP", _
                         "F1_與國內顧客間交易_SPOT")
    
    ' 對應每個交易類型在 Excel 中的欄位範圍
    Dim dataRanges As Variant
    dataRanges = Array("A:B", "C:D", "E:F", "G:H", "Q:R") ' 這裡假設 Cur 在前一欄, Value 在後一欄
    
    For i = LBound(transactionTypes) To UBound(transactionTypes)
        ' 建立字典儲存貨幣數值，並初始化為 0
        Dim curDict As Object
        Set curDict = CreateObject("Scripting.Dictionary")
        For j = LBound(currencies) To UBound(currencies)
            curDict.Add currencies(j), 0
        Next j
        
        ' 確定當前交易的資料範圍
        Dim currCol As Integer
        Dim lastRow As Long
        currCol = xlsht.Range(dataRanges(i)).Column  ' 取得起始欄位（Cur欄）
        lastRow = xlsht.Cells(xlsht.Rows.Count, currCol).End(xlUp).Row  ' 找到最後一列

        For j = 2 To lastRow ' 假設第1列是標題，從第2列開始
            Dim curCode As String, curValue As Variant
            curCode = xlsht.Cells(j, currCol).Value ' 貨幣名稱
            curValue = Round(xlsht.Cells(j, currCol + 1).Value / 1000000, 1) ' 貨幣數值 百萬元，四捨五入小數第一位
            
            ' 確保 Value 為數字，且 Cur 是已定義的貨幣
            If IsNumeric(curValue) And curDict.Exists(curCode) Then
                curDict(curCode) = curValue ' 若累加改成 curDict(curCode) = curDict(curCode) + curValue
            End If
        Next j
        
        ' 依照固定貨幣順序填入 Excel 和報表
        For j = LBound(currencies) To UBound(currencies)
            Dim fieldName As String, valueNum As Variant
            fieldName = transactionTypes(i) & "_" & currencies(j) ' 產生範圍名稱
            valueNum = curDict(currencies(j))
        


            ' 設定 Excel 的 Range 值
            xlsht.Range(fieldName).Value = valueNum
            
            ' 設定報表欄位
            rpt.SetField "F1", fieldName, valueNum
        Next j
    Next i
    
    xlsht.Range("AB2:AB100").NumberFormat = "#,##,##.00"
    xlsht.Range("AF2:AF100").NumberFormat = "#,##,##.00"
    xlsht.Range("AJ2:AJ100").NumberFormat = "#,##,##.00"
    xlsht.Range("AN2:AN100").NumberFormat = "#,##,##.00"
    xlsht.Range("AR2:AR100").NumberFormat = "#,##,##.00"
    

    ' 驗證並更新資料庫（以 Insert 示範，若記錄已存在則應用 UPDATE）
    If rpt.ValidateFields() Then
        Dim key As Variant
        Dim allValues As Object, allPositions As Object

        Set allValues = rpt.GetAllFieldValues()  ' key 格式 "wsName|fieldName"
        Set allPositions = rpt.GetAllFieldPositions()

        For Each key In allValues.Keys
            UpdateRecord gDBPath, gDataMonthString, rpt.ReportName, key, allPositions(key), CDbl(allValues(key))
        Next key
    End If
End Sub




Public Sub Process_Table41()
    'Equal Setting
    'Fetch Query Access DB table
    Dim dataArr_1 As Variant
    Dim dataArr_2 As Variant

    'Declare worksheet and handle data
    Dim controlSheet As Worksheet
    Dim xlsht As Worksheet

    Dim i As Integer, j As Integer
    Dim lastRow As Integer

    Dim reportTitle As String
    Dim queryTable_1 As String
    Dim queryTable_2 As String

    'Setting class clsReport
    Dim rpt As clsReport
    '==========
    Set rpt = gReports("Table41")
    
    '==========
    reportTitle = "Table41"
    queryTable_1 = "表41_DBU_DL9360_LIST"
    queryTable_2 = "表41_DBU_DL9360_Subtotal"

    ' Equal Setting
    ' 取得 Query 資料 (查詢名稱 "CNY1_Query")
    '==========
    ' dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1, gDataMonthString)
    ' dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2, gDataMonthString)
    dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1)
    dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2)
    If UBound(dataArr_1) < 1 Or UBound(dataArr_2) < 1 Then
        '==========
        MsgBox "Table41 查詢資料不完整！", vbExclamation
        Exit Sub
    End If
    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    '==========
    xlsht.Range("A:E").ClearContents
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row

    '==========
    'Paste Queyr Table inot Excel
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



    '--------------
    'Unique Setting
    '--------------

    Dim derivativeGain As Double
    Dim derivativeLoss As Double

    derivativeGain = 0
    derivativeLoss = 0
    
    If xlsht.Cells(1, "I").Value = "SumProfit" Then
        If NOT IsEmpty(xlsht.Cells(2, "I").Value) Then
            derivativeGain = xlsht.Cells(2, "I").Value
        Else
            MsgBox "Error: No Data for Derivative Profit"
            Exit Sub
        End If
    Else
        MsgBox "Error: No Data for Derivative Profit/Loss"
        Exit Sub
    End If

    If xlsht.Cells(1, "J").Value = "SumLoss" Then
        If NOT IsEmpty(xlsht.Cells(2, "J").Value) Then
            derivativeLoss = xlsht.Cells(2, "J").Value
        Else
            MsgBox "Error: No Data for Derivative Loss"
            Exit Sub
        End If
    Else
        MsgBox "Error: No Data for Derivative Profit/Loss"
        Exit Sub
    End If

    derivativeGain = Round(derivativeGain / 1000, 0)
    derivativeLoss = Round(derivativeLoss / 1000, 0)
    
    xlsht.Range("Table41_四衍生工具處分利益").Value = derivativeGain
    rpt.SetField "FOA", "Table41_四衍生工具處分利益", derivativeGain

    xlsht.Range("Table41_四衍生工具處分損失").Value = derivativeLoss
    rpt.SetField "FOA", "Table41_四衍生工具處分損失", derivativeLoss
    
    xlsht.Range("T2:T100").NumberFormat = "#,##,##.00"
    

    ' 驗證並更新資料庫（以 Insert 示範，若記錄已存在則應用 UPDATE）
    If rpt.ValidateFields() Then
        Dim key As Variant
        Dim allValues As Object, allPositions As Object

        Set allValues = rpt.GetAllFieldValues()  ' key 格式 "wsName|fieldName"
        Set allPositions = rpt.GetAllFieldPositions()

        For Each key In allValues.Keys
            UpdateRecord gDBPath, gDataMonthString, rpt.ReportName, key, allPositions(key), CDbl(allValues(key))
        Next key
    End If
End Sub

Public Sub Process_AI602()
    'Equal Setting
    'Fetch Query Access DB table
    Dim dataArr_1 As Variant
    Dim dataArr_2 As Variant
    Dim dataArr_3 As Variant

    'Declare worksheet and handle data
    Dim controlSheet As Worksheet
    Dim xlsht As Worksheet

    Dim i As Integer, j As Integer
    Dim lastRow As Integer

    Dim reportTitle As String
    Dim queryTable_1 As String
    Dim queryTable_2 As String
    Dim queryTable_3 As String

    'Setting class clsReport
    Dim rpt As clsReport
    '==========
    Set rpt = gReports("AI602")
    
    '==========
    reportTitle = "AI602"
    queryTable_1 = "AI602_Impairment_USD"
    queryTable_2 = "AI602_GroupedAC5601"
    queryTable_3 = "AI602_Subtotal"

    ' Equal Setting
    ' 取得 Query 資料 (查詢名稱 "CNY1_Query")
    '==========
    ' dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1, gDataMonthString)
    ' dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2, gDataMonthString)
    ' dataArr_3 = GetAccessDataAsArray(gDBPath, queryTable_3, gDataMonthString)
    dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1)
    dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2)
    dataArr_3 = GetAccessDataAsArray(gDBPath, queryTable_3)
    If UBound(dataArr_1) < 1 Or UBound(dataArr_2) < 1 UBound(dataArr_3) < 1 Then
        '==========
        MsgBox "CNY1 查詢資料不完整！", vbExclamation
        Exit Sub
    End If
    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    '==========
    xlsht.Range("A:K").ClearContents
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row

    '==========
    'Paste Queyr Table inot Excel
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


    '--------------
    'Unique Setting
    '--------------
    Dim rngs As Range
    Dim rng As Range

    Dim FVOCI_GovDebt_Cost As Double
    Dim FVOCI_GovDebt_Adjustment As Double
    Dim FVOCI_GovDebt_Impairment As Double
    Dim FVOCI_GovDebt_BookValue As Double
    
    Dim FVOCI_CompanyDebt_Cost As Double
    Dim FVOCI_CompanyDebt_Adjustment As Double
    Dim FVOCI_CompanyDebt_Impairment As Double
    Dim FVOCI_CompanyDebt_BookValue As Double
    
    Dim FVOCI_FinanceDebt_Cost As Double
    Dim FVOCI_FinanceDebt_Adjustment As Double
    Dim FVOCI_FinanceDebt_Impairment As Double
    Dim FVOCI_FinanceDebt_BookValue As Double

    Dim AC_GovDebt_Cost As Double
    Dim AC_GovDebt_Impairment As Double
    Dim AC_GovDebt_BookValue As Double
    
    Dim AC_CompanyDebt_Cost As Double
    Dim AC_CompanyDebt_Impairment As Double
    Dim AC_CompanyDebt_BookValue As Double
    
    Dim AC_FinanceDebt_Cost As Double
    Dim AC_FinanceDebt_Impairment As Double
    Dim AC_FinanceDebt_BookValue As Double

    FVOCI_GovDebt_Cost = 0
    FVOCI_GovDebt_Adjustment = 0
    FVOCI_GovDebt_Impairment = 0
    FVOCI_GovDebt_BookValue = 0

    FVOCI_CompanyDebt_Cost = 0
    FVOCI_CompanyDebt_Adjustment = 0
    FVOCI_CompanyDebt_Impairment = 0
    FVOCI_CompanyDebt_BookValue = 0

    FVOCI_FinanceDebt_Cost = 0
    FVOCI_FinanceDebt_Adjustment = 0
    FVOCI_FinanceDebt_Impairment = 0
    FVOCI_FinanceDebt_BookValue = 0

    AC_GovDebt_Cost = 0
    AC_GovDebt_Impairment = 0
    AC_GovDebt_BookValue = 0

    AC_CompanyDebt_Cost = 0
    AC_CompanyDebt_Impairment = 0
    AC_CompanyDebt_BookValue = 0

    AC_FinanceDebt_Cost = 0
    AC_FinanceDebt_Impairment = 0
    AC_FinanceDebt_BookValue = 0

    lastRow = xlsht.Cells(xlsht.Rows.Count, "J").End(xlUp).Row
    Set rngs = xlsht.Range("J2:J" & lastRow)

    For Each rng In rngs
        If CStr(rng.Value) = "FVOCI_政府公債_外國_投資成本" Then
            FVOCI_GovDebt_Cost = FVOCI_GovDebt_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVOCI_政府公債_外國_評價調整" Then
            FVOCI_GovDebt_Adjustment = FVOCI_GovDebt_Adjustment + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVOCI_政府公債_外國_減損" Then
            FVOCI_GovDebt_Impairment = FVOCI_GovDebt_Impairment + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "AC_政府公債_外國_投資成本" Then
            AC_GovDebt_Cost = AC_GovDebt_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "AC_政府公債_外國_減損" Then
            AC_GovDebt_Impairment = AC_GovDebt_Impairment + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVOCI_公司債_外國_投資成本" Then
            FVOCI_CompanyDebt_Cost = FVOCI_CompanyDebt_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVOCI_公司債_外國_評價調整" Then
            FVOCI_CompanyDebt_Adjustment = FVOCI_CompanyDebt_Adjustment + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVOCI_公司債_外國_減損" Then
            FVOCI_CompanyDebt_Impairment = FVOCI_CompanyDebt_Impairment + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "AC_公司債_外國_投資成本" Then
            AC_CompanyDebt_Cost = AC_CompanyDebt_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "AC_公司債_外國_減損" Then
            AC_CompanyDebt_Impairment = AC_CompanyDebt_Impairment + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVOCI_金融債_外國_投資成本" Then
            FVOCI_FinanceDebt_Cost = FVOCI_FinanceDebt_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVOCI_金融債_外國_評價調整" Then
            FVOCI_FinanceDebt_Adjustment = FVOCI_FinanceDebt_Adjustment + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "FVOCI_金融債_外國_減損" Then
            FVOCI_FinanceDebt_Impairment = FVOCI_FinanceDebt_Impairment + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "AC_金融債_外國_投資成本" Then
            AC_FinanceDebt_Cost = AC_FinanceDebt_Cost + rng.Offset(0, 1).Value
        ElseIf CStr(rng.Value) = "AC_金融債_外國_減損" Then
            AC_FinanceDebt_Impairment = AC_FinanceDebt_Impairment + rng.Offset(0, 1).Value
        End If
    Next rng

    'FVOCI減損數為正數，需要扣除
    FVOCI_GovDebt_BookValue = FVOCI_GovDebt_Cost + FVOCI_GovDebt_Adjustment - FVOCI_GovDebt_Impairment 
    FVOCI_CompanyDebt_BookValue = FVOCI_CompanyDebt_Cost + FVOCI_CompanyDebt_Adjustment - FVOCI_CompanyDebt_Impairment
    FVOCI_FinanceDebt_BookValue = FVOCI_FinanceDebt_Cost + FVOCI_FinanceDebt_Adjustment - FVOCI_FinanceDebt_Impairment
    
    'AC減損數為負數，相加即可
    AC_GovDebt_BookValue = AC_GovDebt_Cost + AC_GovDebt_Impairment
    AC_CompanyDebt_BookValue = AC_CompanyDebt_Cost + AC_CompanyDebt_Impairment
    AC_FinanceDebt_BookValue = AC_FinanceDebt_Cost + AC_FinanceDebt_Impairment

    Dim sum_GovDebt_Cost As Double
    Dim sum_GovDebt_BookValue As Double
    sum_GovDebt_Cost = 0
    sum_GovDebt_BookValue = 0

    Dim sum_CompanyDebt_Cost As Double
    Dim sum_CompanyDebt_BookValue As Double
    sum_CompanyDebt_Cost = 0
    sum_CompanyDebt_BookValue = 0

    Dim sum_FinanceDebt_Cost As Double
    Dim sum_FinanceDebt_BookValue As Double
    sum_FinanceDebt_Cost = 0
    sum_FinanceDebt_BookValue = 0

    FVOCI_GovDebt_Cost = Round(FVOCI_GovDebt_Cost / 1000, 0)
    FVOCI_GovDebt_BookValue = Round(FVOCI_GovDebt_BookValue / 1000, 0)
    AC_GovDebt_Cost = Round(AC_GovDebt_Cost / 1000, 0)
    AC_GovDebt_BookValue = Round(AC_GovDebt_BookValue / 1000, 0)
    sum_GovDebt_Cost = FVOCI_GovDebt_Cost + AC_GovDebt_Cost
    sum_GovDebt_BookValue = FVOCI_GovDebt_BookValue + AC_GovDebt_BookValue

    FVOCI_CompanyDebt_Cost = Round(FVOCI_CompanyDebt_Cost / 1000, 0)
    FVOCI_CompanyDebt_BookValue = Round(FVOCI_CompanyDebt_BookValue / 1000, 0)
    AC_CompanyDebt_Cost = Round(AC_CompanyDebt_Cost / 1000, 0)
    AC_CompanyDebt_BookValue = Round(AC_CompanyDebt_BookValue / 1000, 0)
    sum_CompanyDebt_Cost = FVOCI_CompanyDebt_Cost + AC_CompanyDebt_Cost
    sum_CompanyDebt_BookValue = FVOCI_CompanyDebt_BookValue + AC_CompanyDebt_BookValue

    FVOCI_FinanceDebt_Cost = Round(FVOCI_FinanceDebt_Cost / 1000, 0)
    FVOCI_FinanceDebt_BookValue = Round(FVOCI_FinanceDebt_BookValue / 1000, 0)
    AC_FinanceDebt_Cost = Round(AC_FinanceDebt_Cost / 1000, 0)
    AC_FinanceDebt_BookValue = Round(AC_FinanceDebt_BookValue / 1000, 0)
    sum_FinanceDebt_Cost = FVOCI_FinanceDebt_Cost + AC_FinanceDebt_Cost
    sum_FinanceDebt_BookValue = FVOCI_FinanceDebt_BookValue + AC_FinanceDebt_BookValue
    
    xlsht.Range("AI602_政府公債_投資成本_FVOCI_F2").Value = FVOCI_GovDebt_Cost
    rpt.SetField "Table1", "AI602_政府公債_投資成本_FVOCI_F2", FVOCI_GovDebt_Cost

    xlsht.Range("AI602_政府公債_帳面價值_FVOCI_F2").Value = FVOCI_GovDebt_BookValue
    rpt.SetField "Table1", "AI602_政府公債_帳面價值_FVOCI_F2", FVOCI_GovDebt_BookValue

    xlsht.Range("AI602_政府公債_投資成本_AC_F3").Value = AC_GovDebt_Cost
    rpt.SetField "Table1", "AI602_政府公債_投資成本_AC_F3", AC_GovDebt_Cost

    xlsht.Range("AI602_政府公債_帳面價值_AC_F3").Value = AC_GovDebt_BookValue
    rpt.SetField "Table1", "AI602_政府公債_帳面價值_AC_F3", AC_GovDebt_BookValue

    xlsht.Range("AI602_政府公債_投資成本_合計_F5").Value = sum_GovDebt_Cost
    rpt.SetField "Table1", "AI602_政府公債_投資成本_合計_F5", sum_GovDebt_Cost

    xlsht.Range("AI602_政府公債_帳面價值_合計_F5").Value = sum_GovDebt_BookValue
    rpt.SetField "Table1", "AI602_政府公債_帳面價值_合計_F5", sum_GovDebt_BookValue


    xlsht.Range("AI602_公司債_投資成本_FVOCI_F7").Value = FVOCI_CompanyDebt_Cost
    rpt.SetField "Table1", "AI602_公司債_投資成本_FVOCI_F7", FVOCI_CompanyDebt_Cost

    xlsht.Range("AI602_公司債_帳面價值_FVOCI_F7").Value = FVOCI_CompanyDebt_BookValue
    rpt.SetField "Table1", "AI602_公司債_帳面價值_FVOCI_F7", FVOCI_CompanyDebt_BookValue

    xlsht.Range("AI602_公司債_投資成本_AC_F8").Value = AC_CompanyDebt_Cost
    rpt.SetField "Table1", "AI602_公司債_投資成本_AC_F8", AC_CompanyDebt_Cost

    xlsht.Range("AI602_公司債_帳面價值_AC_F8").Value = AC_CompanyDebt_BookValue
    rpt.SetField "Table1", "AI602_公司債_帳面價值_AC_F8", AC_CompanyDebt_BookValue

    xlsht.Range("AI602_公司債_投資成本_合計_F10").Value = sum_CompanyDebt_Cost
    rpt.SetField "Table1", "AI602_公司債_投資成本_合計_F10", sum_CompanyDebt_Cost

    xlsht.Range("AI602_公司債_帳面價值_合計_F10").Value = sum_CompanyDebt_BookValue
    rpt.SetField "Table1", "AI602_公司債_帳面價值_合計_F10", sum_CompanyDebt_BookValue


    xlsht.Range("AI602_金融債_投資成本_FVOCI_F2").Value = FVOCI_FinanceDebt_Cost
    rpt.SetField "Table2", "AI602_金融債_投資成本_FVOCI_F2", FVOCI_FinanceDebt_Cost

    xlsht.Range("AI602_金融債_帳面價值_FVOCI_F2").Value = FVOCI_FinanceDebt_BookValue
    rpt.SetField "Table2", "AI602_金融債_帳面價值_FVOCI_F2", FVOCI_FinanceDebt_BookValue

    xlsht.Range("AI602_金融債_投資成本_AC_F3").Value = AC_FinanceDebt_Cost
    rpt.SetField "Table2", "AI602_金融債_投資成本_AC_F3", AC_FinanceDebt_Cost

    xlsht.Range("AI602_金融債_帳面價值_AC_F3").Value = AC_FinanceDebt_BookValue
    rpt.SetField "Table2", "AI602_金融債_帳面價值_AC_F3", AC_FinanceDebt_BookValue

    xlsht.Range("AI602_金融債_投資成本_合計_F5").Value = sum_FinanceDebt_Cost
    rpt.SetField "Table2", "AI602_金融債_投資成本_合計_F5", sum_FinanceDebt_Cost

    xlsht.Range("AI602_金融債_帳面價值_合計_F5").Value = sum_FinanceDebt_BookValue
    rpt.SetField "Table2", "AI602_金融債_帳面價值_合計_F5", sum_FinanceDebt_BookValue
    
    xlsht.Range("T2:T100").NumberFormat = "#,##,##.00"
    

    ' 驗證並更新資料庫（以 Insert 示範，若記錄已存在則應用 UPDATE）
    If rpt.ValidateFields() Then
        Dim key As Variant
        Dim allValues As Object, allPositions As Object

        Set allValues = rpt.GetAllFieldValues()  ' key 格式 "wsName|fieldName"
        Set allPositions = rpt.GetAllFieldPositions()

        For Each key In allValues.Keys
            UpdateRecord gDBPath, gDataMonthString, rpt.ReportName, key, allPositions(key), CDbl(allValues(key))
        Next key
    End If
End Sub




Public Sub Process_AI240()
    'Equal Setting
    'Fetch Query Access DB table
    Dim dataArr_1 As Variant
    Dim dataArr_2 As Variant

    'Declare worksheet and handle data
    Dim controlSheet As Worksheet
    Dim xlsht As Worksheet

    Dim i As Integer, j As Integer
    Dim lastRow As Integer

    Dim reportTitle As String
    Dim queryTable_1 As String
    Dim queryTable_2 As String

    'Setting class clsReport
    Dim rpt As clsReport
    '==========
    Set rpt = gReports("AI240")
    
    '==========
    reportTitle = "AI240"
    queryTable_1 = "AI240_DBU_DL6850_LIST"
    queryTable_2 = "AI240_DBU_DL6850_Subtoal"

    ' Equal Setting
    ' 取得 Query 資料 (查詢名稱 "CNY1_Query")
    '==========
    ' dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1, gDataMonthString)
    ' dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2, gDataMonthString)
    dataArr_1 = GetAccessDataAsArray(gDBPath, queryTable_1)
    dataArr_2 = GetAccessDataAsArray(gDBPath, queryTable_2)
    If UBound(dataArr_1) < 1 Or UBound(dataArr_2) < 1 Then
        '==========
        MsgBox "AI240 查詢資料不完整！", vbExclamation
        Exit Sub
    End If
    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    '==========
    xlsht.Range("A:L").ClearContents
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row

    '==========
    'Paste Queyr Table inot Excel
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

    '--------------
    'Unique Setting
    '--------------
    Dim rngs As Range
    Dim rng As Range

    Dim buyAmountTWD_0to10 As Double
    Dim buyAmountTWD_11to30 As Double
    Dim buyAmountTWD_31to90 As Double
    Dim buyAmountTWD_91to180 As Double
    Dim buyAmountTWD_181to365 As Double
    Dim buyAmountTWD_over365 As Double

    Dim sellAmountTWD_0to10 As Double
    Dim sellAmountTWD_11to30 As Double
    Dim sellAmountTWD_31to90 As Double
    Dim sellAmountTWD_91to180 As Double
    Dim sellAmountTWD_181to365 As Double
    Dim sellAmountTWD_over365 As Double
    
    buyAmountTWD_0to10 = 0
    buyAmountTWD_11to30 = 0
    buyAmountTWD_31to90 = 0
    buyAmountTWD_91to180 = 0
    buyAmountTWD_181to365 = 0
    buyAmountTWD_over365 = 0

    sellAmountTWD_0to10 = 0
    sellAmountTWD_11to30 = 0
    sellAmountTWD_31to90 = 0
    sellAmountTWD_91to180 = 0
    sellAmountTWD_181to365 = 0
    sellAmountTWD_over365 = 0

    lastRow = xlsht.Cells(xlsht.Rows.Count, "J").End(xlUp).Row
    Set rngs = xlsht.Range("J2:J" & lastRow)

    For Each rng In rngs
        If CStr(rng.Value) = "基準日後0-10天" Then
            buyAmountTWD_0to10 = buyAmountTWD_0to10 + rng.Offset(0, 1).Value
            sellAmountTWD_0to10 = sellAmountTWD_0to10 + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "基準日後11-30天" Then
            buyAmountTWD_11to30 = buyAmountTWD_11to30 + rng.Offset(0, 1).Value
            sellAmountTWD_11to30 = sellAmountTWD_11to30 + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "基準日後31-90天" Then
            buyAmountTWD_31to90 = buyAmountTWD_31to90 + rng.Offset(0, 1).Value
            sellAmountTWD_31to90 = sellAmountTWD_31to90 + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "基準日後91-180天" Then
            buyAmountTWD_91to180 = buyAmountTWD_91to180 + rng.Offset(0, 1).Value
            sellAmountTWD_91to180 = sellAmountTWD_91to180 + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "基準日後181天-1年" Then
            buyAmountTWD_181to365 = buyAmountTWD_181to365 + rng.Offset(0, 1).Value
            sellAmountTWD_181to365 = sellAmountTWD_181to365 + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "超過基準日後一年" Then
            buyAmountTWD_over365 = buyAmountTWD_over365 + rng.Offset(0, 1).Value
            sellAmountTWD_over365 = sellAmountTWD_over365 + rng.Offset(0, 2).Value
        End If
    Next rng


    xlsht.Range("AI240_其他到期資金流入項目_10天").Value = buyAmountTWD_0to10
    rpt.SetField "工作表1", "AI240_其他到期資金流入項目_10天", buyAmountTWD_0to10

    xlsht.Range("AI240_其他到期資金流入項目_30天").Value = buyAmountTWD_11to30
    rpt.SetField "工作表1", "AI240_其他到期資金流入項目_30天", buyAmountTWD_11to30

    xlsht.Range("AI240_其他到期資金流入項目_90天").Value = buyAmountTWD_31to90
    rpt.SetField "工作表1", "AI240_其他到期資金流入項目_90天", buyAmountTWD_31to90

    xlsht.Range("AI240_其他到期資金流入項目_180天").Value = buyAmountTWD_91to180
    rpt.SetField "工作表1", "AI240_其他到期資金流入項目_180天", buyAmountTWD_91to180

    xlsht.Range("AI240_其他到期資金流入項目_1年").Value = buyAmountTWD_181to365
    rpt.SetField "工作表1", "AI240_其他到期資金流入項目_1年", buyAmountTWD_181to365

    xlsht.Range("AI240_其他到期資金流入項目_1年以上").Value = buyAmountTWD_over365
    rpt.SetField "工作表1", "AI240_其他到期資金流入項目_1年以上", buyAmountTWD_over365
    

    xlsht.Range("AI240_其他到期資金流出項目_10天").Value = sellAmountTWD_0to10
    rpt.SetField "工作表1", "AI240_其他到期資金流出項目_10天", sellAmountTWD_0to10

    xlsht.Range("AI240_其他到期資金流出項目_30天").Value = sellAmountTWD_11to30
    rpt.SetField "工作表1", "AI240_其他到期資金流出項目_30天", sellAmountTWD_11to30

    xlsht.Range("AI240_其他到期資金流出項目_90天").Value = sellAmountTWD_31to90
    rpt.SetField "工作表1", "AI240_其他到期資金流出項目_90天", sellAmountTWD_31to90

    xlsht.Range("AI240_其他到期資金流出項目_180天").Value = sellAmountTWD_91to180
    rpt.SetField "工作表1", "AI240_其他到期資金流出項目_180天", sellAmountTWD_91to180

    xlsht.Range("AI240_其他到期資金流出項目_1年").Value = sellAmountTWD_181to365
    rpt.SetField "工作表1", "AI240_其他到期資金流出項目_1年", sellAmountTWD_181to365

    xlsht.Range("AI240_其他到期資金流出項目_1年以上").Value = sellAmountTWD_over365
    rpt.SetField "工作表1", "AI240_其他到期資金流出項目_1年以上", sellAmountTWD_over365

    xlsht.Range("T2:T100").NumberFormat = "#,##,##.00"
    

    ' 驗證並更新資料庫（以 Insert 示範，若記錄已存在則應用 UPDATE）
    If rpt.ValidateFields() Then
        Dim key As Variant
        Dim allValues As Object, allPositions As Object

        Set allValues = rpt.GetAllFieldValues()  ' key 格式 "wsName|fieldName"
        Set allPositions = rpt.GetAllFieldPositions()

        For Each key In allValues.Keys
            UpdateRecord gDBPath, gDataMonthString, rpt.ReportName, key, allPositions(key), CDbl(allValues(key))
        Next key
    End If
End Sub


'=== (c) 更新申報 Excel 檔案，將各報表物件數值寫入對應儲存格（各工作表），並另存新檔 ===
Public Sub UpdateExcelReports()
    Dim rpt As clsReport
    Dim rptName As Variant
    Dim wb As Workbook
    Dim reportFilePath As String, outputFilePath As String
    For Each rptName In gReportNames
        Set rpt = gReports(rptName)
        ' 開啟原始 Excel 檔（檔名以報表名稱命名）
        reportFilePath = gReportFolder & rptName & ".xlsx"
        Set wb = Workbooks.Open(reportFilePath)
        If wb Is Nothing Then
            MsgBox "無法開啟檔案: " & reportFilePath, vbExclamation
        End If
        ' 由於報表內可能有多個工作表，呼叫 ApplyToWorkbook 讓 clsReport 自行依各工作表更新
        rpt.ApplyToWorkbook wb
        outputFilePath = gOutputFolder & rptName & "_Processed.xlsx"
        wb.SaveAs Filename:=outputFilePath
        wb.Close SaveChanges:=False
        Set wb = Nothing   ' Release Workbook Object
    Next rptName
    MsgBox "所有 Excel 申報報表已更新並另存！"
End Sub
