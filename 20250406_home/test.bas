



    '===============================
    Dim xlsht As Worksheet
    Dim lastRow As Long
    Dim i As Long


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
        Amount = xlsht.Cells(i, "H").value
        ' 類別 
        Category = xlsht.Cells(i, "I").Value               
        'TWTP_MP / OBU_MP / TWTP_MT / OBU_MT
        
        xlsht.Cells(i, "K").Value = BankCode             ' K：BankCode
        xlsht.Cells(i, "L").Value = CounterParty         ' L：CounterParty
        
        ' 根據 Category 將金額填入對應分類欄位
        Select Case Category
            Case "TWTP_MP"
                xlsht.Cells(i, "M").Value = Amount      ' M：TWTP_MP
            Case "OBU_MP"
                xlsht.Cells(i, "N").Value = Amount      ' N：OBU_MP
            Case "TWTP_MT"
                xlsht.Cells(i, "O").Value = Amount      ' O：TWTP_MT
            Case "OBU_MT"
                xlsht.Cells(i, "P").Value = Amount      ' P：OBU_MT
        End Select
        

        ' 二、記錄儲存格位置和數值（輸出位置由 Row 10 開始）
        ' 這邊假設：BankCode 記錄在 C 欄；金額根據 Category 記錄在 E (TWTP_MP) / F (OBU_MP) / G (TWTP_MT) / H (OBU_MT)

        Select Case Category
            Case "TWTP_MP"
                targetCol = "E"
            Case "OBU_MP"
                targetCol = "F"
            Case "TWTP_MT"
                targetCol = "G"
            Case "OBU_MT"
                targetCol = "H"
        End Select
        
        ' rpt.SetField "FOA", "FB3A_BankCode", "C" & CStr(targetRow), BankCode
        ' rpt.SetField "FOA", "FB3A_Amount", targetCol & CStr(targetRow), Amount

        AddDynamicField "FOA", "FB3A_BankCode", "C" & CStr(targetRow), BankCode
        AddDynamicField "FOA", "FB3A_Amount", targetCol & CStr(targetRow), Amount
        
        InsertIntoTable gDBPath, "MonthlyDeclarationReport", gDataMonthString, "FB3A", "FOA|FB3A_BankCode", BankCode, "C" & CStr(targetRow)
        InsertIntoTable gDBPath, "MonthlyDeclarationReport", gDataMonthString, "FB3A", "FOA|FB3A_Amount", Amount, targetCol & CStr(targetRow)
        
        targetRow = targetRow + 1
    Next i






    '==============================