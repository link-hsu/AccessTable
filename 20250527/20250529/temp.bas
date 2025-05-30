
'尚無有交易紀錄
Public Sub Process_AI233()
    '=== Equal Setting ===
    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("AI233")

    Dim reportTitle As String
    reportTitle = rpt.ReportName

    Dim xlsht As Worksheet
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    ' xlsht.Range("A:G").ClearContents
    ' xlsht.Range("T2:T100").ClearContents

    ' xlsht.Cells.ClearContents  ' 或只清指定区

    ' 呼叫通用函式，並取得 importCols
    Dim importCols As Collection
    Set importCols = ImportQueryTables(gDBPath, xlsht, reportTitle, gDataMonthString)

    If importCols Is Nothing Or importCols.Count = 0 Then
        Exit Sub
    End If
    
    '--------------
    'Unique Setting
    '--------------
     Dim RP_GovBond_Cost As Double: RP_GovBond_Cost = 0
     ' - 透過損益按公允價值衡量A1		
     ' - 投資成本		
     ' - 政府公債A		
     Dim FVPL_GovBond_Domestic_Cost As Double: FVPL_GovBond_Domestic_Cost = 0

     ' - 公司債B		
     Dim FVPL_CompanyBond_Domestic_Cost As Double: FVPL_CompanyBond_Domestic_Cost = 0

     ' - 金融債券C市
     ' - 權益證券投資D
     Dim FVPL_Stock_Cost As Double: FVPL_Stock_Cost = 0

     ' - 結構型商品E		
     ' - 證券化商品F		
     ' - 其它投資部位G		
     Dim FVPL_Other_Cost As Double: FVPL_Other_Cost = 0

     ' - 央行可轉讓定期存單H		
     ' - 帳面價值 = 投資成本 + 以下		
     ' - 政府公債A		
     Dim FVPL_GovBond_Domestic_BV As Double: FVPL_GovBond_Domestic_BV = 0

     ' - 公司債B		
     Dim FVPL_CompanyBond_Domestic_BV As Double: FVPL_CompanyBond_Domestic_BV = 0

     ' - 金融債券C
     ' - 權益證券投資D
     Dim FVPL_Stock_BV As Double: FVPL_Stock_BV = 0

     ' - 結構型商品E		
     ' - 證券化商品F		
     ' - 其它投資部位G		
     Dim FVPL_Other_BV As Double: FVPL_Other_BV = 0

     ' - 央行可轉讓定期存單H

     ' - 透過其他綜合損益按公允價值衡量A6		
     ' - 投資成本		
     ' - 政府公債A		
     Dim FVOCI_GovBond_Domestic_Cost As Double: FVOCI_GovBond_Domestic_Cost = 0

     ' - 公司債B		
     Dim FVOCI_CompanyBond_Domestic_Cost As Double: FVOCI_CompanyBond_Domestic_Cost = 0

     ' - 金融債券C		
     ' - 權益證券投資D		
     Dim FVOCI_Stock_Cost As Double: FVOCI_Stock_Cost = 0
               
     ' - 結構型商品E
     ' - 證券化商品F
     ' - 其它投資部位G
     ' - 央行可轉讓定期存單H
     Dim FVOCI_NCD_CentralBank_Cost As Double: FVOCI_NCD_CentralBank_Cost = 0

     ' - 帳面價值
     ' - 政府公債A
     Dim FVOCI_GovBond_Domestic_BV As Double: FVOCI_GovBond_Domestic_BV = 0

     ' - 公司債B
     Dim FVOCI_CompanyBond_Domestic_BV As Double: FVOCI_CompanyBond_Domestic_BV = 0

     ' - 金融債券C		
     ' - 權益證券投資D		
     Dim FVOCI_Stock_BV As Double: FVOCI_Stock_BV = 0

     ' - 結構型商品E
     ' - 證券化商品F
     ' - 其它投資部位G
     ' - 央行可轉讓定期存單H
     Dim FVOCI_NCD_CentralBank_BV As Double: FVOCI_NCD_CentralBank_BV = 0

     ' - 按攤銷後成本衡量A7
     ' - 投資成本
     ' - 政府公債A
     Dim AC_GovBond_Domestic_Cost As Double: AC_GovBond_Domestic_Cost = 0

     ' - 公司債B
     Dim AC_CompanyBond_Domestic_Cost As Double: AC_CompanyBond_Domestic_Cost = 0

     ' - 金融債券C
     ' - 權益證券投資D
     ' - 結構型商品E
     ' - 證券化商品F
     ' - 其它投資部位G
     ' - 央行可轉讓定期存單H
     Dim AC_NCD_CentralBank_Cost As Double: AC_NCD_CentralBank_Cost = 0

     ' - 帳面價值
     ' - 政府公債A
     Dim AC_GovBond_Domestic_BV As Double: AC_GovBond_Domestic_BV = 0

     ' - 公司債B
     Dim AC_CompanyBond_Domestic_BV As Double: AC_CompanyBond_Domestic_BV = 0

     ' - 金融債券C
     ' - 權益證券投資D
     ' - 結構型商品E
     ' - 證券化商品F
     ' - 其它投資部位G
     ' - 央行可轉讓定期存單H
     Dim AC_NCD_CentralBank_BV As Double: AC_NCD_CentralBank_BV = 0


    Dim lastRow1 As Long, lastRow2 As Long
    Dim rngs1 As Range, rngs2 As Range, rng As Range

    ' --- 【修改1.2】不迴圈，直接取用 importCols(1) 和 importCols(2) ---
    If importCols.Count >= 1 Then
        lastRow1 = xlsht.Cells(xlsht.Rows.Count, importCols(1)).End(xlUp).Row
        Set rngs1 = xlsht.Range(xlsht.Cells(2, importCols(1)), xlsht.Cells(lastRow1, importCols(1)))
        
        For Each rng In rngs1
            Select Case CStr(rng.Value)
                Case "AC_CompanyBond_Domestic_Cost"
                    AC_CompanyBond_Domestic_Cost = AC_CompanyBond_Domestic_Cost + rng.Offset(0, 1).Value
                    AC_CompanyBond_Domestic_BV = AC_CompanyBond_Domestic_BV + rng.Offset(0, 1).Value
                Case "AC_CompanyBond_Domestic_ImpairmentLoss"
                    AC_CompanyBond_Domestic_BV = AC_CompanyBond_Domestic_BV + rng.Offset(0, 1).Value
                Case "AC_GovBond_Domestic_Cost"
                    AC_GovBond_Domestic_Cost = AC_GovBond_Domestic_Cost + rng.Offset(0, 1).Value
                    AC_GovBond_Domestic_BV = AC_GovBond_Domestic_BV + rng.Offset(0, 1).Value
                Case "AC_GovBond_Domestic_ImpairmentLoss"
                    AC_GovBond_Domestic_BV = AC_GovBond_Domestic_BV + rng.Offset(0, 1).Value
                Case "AC_NCD_CentralBank_Cost"
                    AC_NCD_CentralBank_Cost = AC_NCD_CentralBank_Cost + rng.Offset(0, 1).Value
                    AC_NCD_CentralBank_BV = AC_NCD_CentralBank_BV + rng.Offset(0, 1).Value
                Case "AC_NCD_CentralBank_ImpairmentLoss"
                    AC_NCD_CentralBank_BV = AC_NCD_CentralBank_BV + rng.Offset(0, 1).Value
                Case "AFS_FinancialBond_Domestic_Cost"

                Case "AFS_FinancialBond_Domestic_ValuationAdjust"
          
                Case "EquityMethod_Other_Cost"

                Case "FVOCI_CompanyBond_Domestic_Cost"
                    FVOCI_CompanyBond_Domestic_Cost = FVOCI_CompanyBond_Domestic_Cost + rng.Offset(0, 1).Value
                    FVOCI_CompanyBond_Domestic_BV = FVOCI_CompanyBond_Domestic_BV + rng.Offset(0, 1).Value
                Case "FVOCI_CompanyBond_Domestic_ValuationAdjust"
                    FVOCI_CompanyBond_Domestic_BV = FVOCI_CompanyBond_Domestic_BV + rng.Offset(0, 1).Value
                Case "FVOCI_CompanyBond_Domestic_ImpairmentAllowance"
                    FVOCI_CompanyBond_Domestic_BV = FVOCI_CompanyBond_Domestic_BV + rng.Offset(0, 1).Value
                Case "FVOCI_GovBond_Domestic_Cost"
                    FVOCI_GovBond_Domestic_Cost = FVOCI_GovBond_Domestic_Cost + rng.Offset(0, 1).Value
                    FVOCI_GovBond_Domestic_BV = FVOCI_GovBond_Domestic_BV + rng.Offset(0, 1).Value
                Case "FVOCI_GovBond_Domestic_ValuationAdjust"
                    FVOCI_GovBond_Domestic_BV = FVOCI_GovBond_Domestic_BV + rng.Offset(0, 1).Value
                Case "FVOCI_GovBond_Domestic_ImpairmentAllowance"
                    FVOCI_GovBond_Domestic_BV = FVOCI_GovBond_Domestic_BV + rng.Offset(0, 1).Value
                Case "FVOCI_NCD_CentralBank_Cost"
                    FVOCI_NCD_CentralBank_Cost = FVOCI_NCD_CentralBank_Cost + rng.Offset(0, 1).Value
                    FVOCI_NCD_CentralBank_BV = FVOCI_NCD_CentralBank_BV + rng.Offset(0, 1).Value
                Case "FVOCI_NCD_CentralBank_ValuationAdjust"
                    FVOCI_NCD_CentralBank_BV = FVOCI_NCD_CentralBank_BV + rng.Offset(0, 1).Value
                Case "FVOCI_NCD_CentralBank_ImpairmentAllowance"
                    FVOCI_NCD_CentralBank_BV = FVOCI_NCD_CentralBank_BV + rng.Offset(0, 1).Value
                Case "FVOCI_Stock_特別股_Cost"

                Case "FVOCI_Stock_特別股_上市_Cost"
                      
                Case "FVOCI_Stock_特別股_上市_ValuationAdjust"

                Case "FVOCI_Stock_普通股_上市_Cost"
                    FVOCI_Stock_Cost = FVOCI_Stock_Cost + rng.Offset(0, 1).Value
                    FVOCI_Stock_BV = FVOCI_Stock_BV + rng.Offset(0, 1).Value
                Case "FVOCI_Stock_普通股_上市_ValuationAdjust"
                    FVOCI_Stock_BV = FVOCI_Stock_BV + rng.Offset(0, 1).Value
                Case "FVOCI_Stock_普通股_上櫃_Cost"
                    FVOCI_Stock_Cost = FVOCI_Stock_Cost + rng.Offset(0, 1).Value
                    FVOCI_Stock_BV = FVOCI_Stock_BV + rng.Offset(0, 1).Value
                Case "FVOCI_Stock_普通股_上櫃_ValuationAdjust"
                    FVOCI_Stock_BV = FVOCI_Stock_BV + rng.Offset(0, 1).Value
                Case "FVOCI_Stock_普通股_興櫃_Cost"

                Case "FVOCI_Stock_普通股_興櫃_ValuationAdjust"

                Case "FVOCI_Equity_Other_Cost"
                    FVOCI_Stock_Cost = FVOCI_Stock_Cost + rng.Offset(0, 1).Value
                    FVOCI_Stock_BV = FVOCI_Stock_BV + rng.Offset(0, 1).Value
                Case "FVOCI_Equity_Other_ValuationAdjust"
                    FVOCI_Stock_BV = FVOCI_Stock_BV + rng.Offset(0, 1).Value
                Case "FVPL_AssetCertificate_Cost"
                    FVPL_Other_Cost = FVPL_Other_Cost + rng.Offset(0, 1).Value
                    FVPL_Other_BV = FVPL_Other_BV + rng.Offset(0, 1).Value
                Case "FVPL_AssetCertificate_ValuationAdjust"
                    FVPL_Other_BV = FVPL_Other_BV + rng.Offset(0, 1).Value
                Case "FVPL_CompanyBond_Domestic_Cost"
                    FVPL_CompanyBond_Domestic_Cost = FVPL_CompanyBond_Domestic_Cost + rng.Offset(0, 1).Value
                    FVPL_CompanyBond_Domestic_BV = FVPL_CompanyBond_Domestic_BV + rng.Offset(0, 1).Value
                Case "FVPL_SWAP_Cost"
                    FVPL_CompanyBond_Domestic_Cost = FVPL_CompanyBond_Domestic_Cost + rng.Offset(0, 1).Value
                    FVPL_CompanyBond_Domestic_BV = FVPL_CompanyBond_Domestic_BV + rng.Offset(0, 1).Value                
                Case "FVPL_CompanyBond_Domestic_ValuationAdjust"
                    FVPL_CompanyBond_Domestic_BV = FVPL_CompanyBond_Domestic_BV + rng.Offset(0, 1).Value
                Case "FVPL_SWAP_ValuationAdjust"
                    FVPL_CompanyBond_Domestic_BV = FVPL_CompanyBond_Domestic_BV + rng.Offset(0, 1).Value
                Case "FVPL_CVASWAP_ValuationAdjust"
                    FVPL_CompanyBond_Domestic_BV = FVPL_CompanyBond_Domestic_BV + rng.Offset(0, 1).Value
                Case "FVPL_CP_Cost"
                    FVPL_Other_Cost = FVPL_Other_Cost + rng.Offset(0, 1).Value
                    FVPL_Other_BV = FVPL_Other_BV + rng.Offset(0, 1).Value
                Case "FVPL_CP_ValuationAdjust"
                    FVPL_Other_BV = FVPL_Other_BV + rng.Offset(0, 1).Value
                Case "FVPL_GovBond_Domestic_Cost"
                    FVPL_GovBond_Domestic_Cost = FVPL_GovBond_Domestic_Cost + rng.Offset(0, 1).Value
                    FVPL_GovBond_Domestic_BV = FVPL_GovBond_Domestic_BV + rng.Offset(0, 1).Value
                Case "FVPL_GovBond_Domestic_ValuationAdjust"
                    FVPL_GovBond_Domestic_BV = FVPL_GovBond_Domestic_BV + rng.Offset(0, 1).Value
                Case "FVPL_Stock_特別股_上市_Cost"
                    FVPL_Stock_Cost = FVPL_Stock_Cost + rng.Offset(0, 1).Value
                    FVPL_Stock_BV = FVPL_Stock_BV + rng.Offset(0, 1).Value
                Case "FVPL_Stock_特別股_上市_ValuationAdjust"
                    FVPL_Stock_BV = FVPL_Stock_BV + rng.Offset(0, 1).Value
                Case "FVPL_Stock_普通股_上市_Cost"
                    FVPL_Stock_Cost = FVPL_Stock_Cost + rng.Offset(0, 1).Value
                    FVPL_Stock_BV = FVPL_Stock_BV + rng.Offset(0, 1).Value
                Case "FVPL_Stock_普通股_上市_ValuationAdjust"
                    FVPL_Stock_BV = FVPL_Stock_BV + rng.Offset(0, 1).Value
                Case "FVPL_Stock_普通股_上櫃_Cost"
                    FVPL_Stock_Cost = FVPL_Stock_Cost + rng.Offset(0, 1).Value
                    FVPL_Stock_BV = FVPL_Stock_BV + rng.Offset(0, 1).Value
                Case "FVPL_Stock_普通股_上櫃_ValuationAdjust"
                    FVPL_Stock_BV = FVPL_Stock_BV + rng.Offset(0, 1).Value
                Case "FVPL_Stock_普通股_興櫃_Cost"

                Case "FVPL_Stock_普通股_興櫃_ValuationAdjust"
                    FVPL_Stock_BV = FVPL_Stock_BV + rng.Offset(0, 1).Value
            End Select
        Next rng
    End If
    

    ' 單位：新臺幣千元


    RP_CP_Cost = Round(RP_CP_Cost / 1000, 0)
    

    xlsht.Range("Table20_0400_四商業本票_民營企業_其他到期日").Value = RP_CP_Cost
    
    ' 資料一開始要先清掉

    ' --- 【修改2】用 FieldValuePositionMap 批次填 rpt ---
    Dim SetFieldMap As Variant
    Dim tgtSheet As String, srcTag As String, srcVal As Variant
    
    SetFieldMap = GetMapData(gDBPath, reportTitle, "FieldValuePositionMap")
    If Not IsNull(SetFieldMap) And IsArray(SetFieldMap) Then
        For iMap = 0 To UBound(SetFieldMap, 1)
            tgtSheet = CStr(SetFieldMap(iMap, 0))
            srcTag    = CStr(SetFieldMap(iMap, 1))
            
            On Error Resume Next
            srcVal = ThisWorkbook.Sheets(tgtSheet).Range(srcTag).Value
            On Error GoTo 0
            
            rpt.SetField tgtSheet, srcTag, srcVal
        Next iMap
    Else
        WriteLog "無法取得 FieldValuePositionMap for " & reportTitle
    End If

    ' 1.Validation filled all value (NO Null value exist)
    ' 2.Update Access DB
    If rpt.ValidateFields() Then
        Dim key As Variant
        Dim allValues As Object, allPositions As Object
        ' key 格式 "wsName|fieldName"
        Set allValues = rpt.GetAllFieldValues()
        Set allPositions = rpt.GetAllFieldPositions()
        For Each key In allValues.Keys
            ' UpdateRecord gDBPath, gDataMonthString, rpt.ReportName, key, allPositions(key), CDbl(allValues(key))
            UpdateRecord gDBPath, gDataMonthString, rpt.ReportName, key, allPositions(key), allValues(key)
        Next key
    End If
    ' 更改分頁顏色為黃色(6)
    xlsht.Tab.ColorIndex = 6
End Sub