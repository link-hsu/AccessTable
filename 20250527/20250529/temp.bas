Public Sub Process_TABLE10()
    '=== Equal Setting ===
    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("TABLE10")

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

    Dim lastRow1 As Long, lastRow2 As Long
    Dim rngs1 As Range, rngs2 As Range, rng As Range    

    ' --- 【修改1.2】不迴圈，直接取用 importCols(1) 和 importCols(2) ---
    If importCols.Count >= 1 Then
        lastRow1 = xlsht.Cells(xlsht.Rows.Count, importCols(1)).End(xlUp).Row
        Set rngs1 = xlsht.Range(xlsht.Cells(2, importCols(1)), xlsht.Cells(lastRow1, importCols(1)))
        
        For Each rng In rngs1
            Select Case CStr(rng.Value)
                Case "AC_CompanyBond_Domestic_Cost"

                Case "AC_CompanyBond_Domestic_ImpairmentLoss"

                Case "AC_GovBond_Domestic_Cost"

                     =  + rng.Offset(0, 1).Value
                Case "AC_GovBond_Domestic_ImpairmentLoss"
                     =  + rng.Offset(0, 1).Value
                Case "AC_NCD_CentralBank_Cost"
                     =  + rng.Offset(0, 1).Value
                Case "AC_NCD_CentralBank_ImpairmentLoss"
                     =  + rng.Offset(0, 1).Value
                Case "AFS_FinancialBond_Domestic_Cost"
                     =  + rng.Offset(0, 1).Value
                Case "AFS_FinancialBond_Domestic_ValuationAdjust"
                     =  + rng.Offset(0, 1).Value
                Case "EquityMethod_Cost"
                     =  + rng.Offset(0, 1).Value
                Case "EquityMethod_ValuationAdjust"
                     =  + rng.Offset(0, 1).Value            
                Case "EquityMethod_Other_Cost"
                     =  + rng.Offset(0, 1).Value
                Case "FVOCI_CompanyBond_Domestic_Cost"

                Case "FVOCI_CompanyBond_Domestic_ValuationAdjust"

                Case "FVOCI_GovBond_Domestic_Cost"
                     =  + rng.Offset(0, 1).Value
                Case "FVOCI_GovBond_Domestic_ValuationAdjust"
                     =  + rng.Offset(0, 1).Value
                Case "FVOCI_NCD_CentralBank_Cost"
                     =  + rng.Offset(0, 1).Value
                Case "FVOCI_NCD_CentralBank_ValuationAdjust"
                     =  + rng.Offset(0, 1).Value
                Case "FVOCI_Stock_特別股_Cost"
                     =  + rng.Offset(0, 1).Value
                Case "FVOCI_Stock_特別股_上市_Cost"
                     =  + rng.Offset(0, 1).Value                        
                Case "FVOCI_Stock_特別股_上市_ValuationAdjust"
                     =  + rng.Offset(0, 1).Value
                Case "FVOCI_Stock_普通股_上市_Cost"
                     =  + rng.Offset(0, 1).Value
                Case "FVOCI_Stock_普通股_上市_ValuationAdjust"
                     =  + rng.Offset(0, 1).Value
                Case "FVOCI_Stock_普通股_上櫃_Cost"
                     =  + rng.Offset(0, 1).Value
                Case "FVOCI_Stock_普通股_上櫃_ValuationAdjust"
                     =  + rng.Offset(0, 1).Value
                Case "FVOCI_Stock_普通股_興櫃_Cost"
                     =  + rng.Offset(0, 1).Value
                Case "FVOCI_Stock_普通股_興櫃_ValuationAdjust"
                     =  + rng.Offset(0, 1).Value
                Case "FVOCI_Equity_Other_Cost"
                     =  + rng.Offset(0, 1).Value
                Case "FVOCI_Equity_Other_ValuationAdjust"
                     =  + rng.Offset(0, 1).Value
                Case "FVPL_AssetCertificate_Cost"
                     =  + rng.Offset(0, 1).Value
                Case "FVPL_AssetCertificate_ValuationAdjust"
                     =  + rng.Offset(0, 1).Value
                Case "FVPL_CompanyBond_Domestic_Cost"

                Case "FVPL_CompanyBond_Domestic_ValuationAdjust"

                Case "FVPL_CP_Cost"
                     =  + rng.Offset(0, 1).Value
                Case "FVPL_CP_ValuationAdjust"
                     =  + rng.Offset(0, 1).Value
                Case "FVPL_GovBond_Domestic_Cost"
                     =  + rng.Offset(0, 1).Value
                Case "FVPL_GovBond_Domestic_ValuationAdjust"
                     =  + rng.Offset(0, 1).Value
                Case "FVPL_Stock_特別股_上市_Cost"
                     =  + rng.Offset(0, 1).Value
                Case "FVPL_Stock_特別股_上市_ValuationAdjust"
                     =  + rng.Offset(0, 1).Value
                Case "FVPL_Stock_普通股_上市_Cost"
                     =  + rng.Offset(0, 1).Value
                Case "FVPL_Stock_普通股_上市_ValuationAdjust"
                     =  + rng.Offset(0, 1).Value
                Case "FVPL_Stock_普通股_上櫃_Cost"
                     =  + rng.Offset(0, 1).Value
                Case "FVPL_Stock_普通股_上櫃_ValuationAdjust"
                     =  + rng.Offset(0, 1).Value
                Case "FVPL_Stock_普通股_興櫃_Cost"
                     =  + rng.Offset(0, 1).Value
                Case "FVPL_Stock_普通股_興櫃_ValuationAdjust"
                     =  + rng.Offset(0, 1).Value
            End Select
        Next rng
    End If    
    
    If importCols.Count >= 2 Then
        lastRow2 = xlsht.Cells(xlsht.Rows.Count, importCols(2)).End(xlUp).Row
        Set rngs2 = xlsht.Range(xlsht.Cells(2, importCols(2)), xlsht.Cells(lastRow2, importCols(2)))
        
        For Each rng In rngs2
            ' 如果第二筆表也有需要累計的 tag，可以在這裡加
            Select Case CStr(rng.Value)
                Case "強制FVPL金融資產-普通公司債(公營)"
                    FVPL_CompanyBond_Public_Domestic_Cost = FVPL_CompanyBond_Public_Domestic_Cost + rng.Offset(0, 1).Value
                Case "強制FVPL金融資產-普通公司債(民營)"
                    FVPL_CompanyBond_Private_Domestic_Cost = FVPL_CompanyBond_Private_Domestic_Cost + rng.Offset(0, 1).Value
                Case "強制FVPL金融資產評價調整-普通公司債(公營)"
                    FVPL_CompanyBond_Public_Domestic_ValuationAdjust = FVPL_CompanyBond_Public_Domestic_ValuationAdjust + rng.Offset(0, 1).Value
                Case "強制FVPL金融資產評價調整-普通公司債(民營)"
                    FVPL_CompanyBond_Private_Domestic_ValuationAdjust = FVPL_CompanyBond_Private_Domestic_ValuationAdjust + rng.Offset(0, 1).Value
                Case "FVOCI債務工具-普通公司債（公營）"
                    FVOCI_CompanyBond_Public_Domestic_Cost = FVOCI_CompanyBond_Public_Domestic_Cost + rng.Offset(0, 1).Value
                Case "FVOCI債務工具-普通公司債（民營）"
                    FVOCI_CompanyBond_Private_Domestic_Cost = FVOCI_CompanyBond_Private_Domestic_Cost + rng.Offset(0, 1).Value
                Case "FVOCI債務工具評價調整-普通公司債（公營)"
                    FVOCI_CompanyBond_Public_Domestic_ValuationAdjust = FVOCI_CompanyBond_Public_Domestic_ValuationAdjust + rng.Offset(0, 1).Value
                Case "FVOCI債務工具評價調整-普通公司債（民營)"
                    FVOCI_CompanyBond_Private_Domestic_ValuationAdjust = FVOCI_CompanyBond_Private_Domestic_ValuationAdjust + rng.Offset(0, 1).Value
                Case "AC債務工具投資-普通公司債(公營)"
                    AC_CompanyBond_Public_Domestic_Cost = AC_CompanyBond_Public_Domestic_Cost + rng.Offset(0, 1).Value
                Case "AC債務工具投資-普通公司債(民營)"
                    AC_CompanyBond_Private_Domestic_Cost = AC_CompanyBond_Private_Domestic_Cost + rng.Offset(0, 1).Value
                Case "累積減損-累積減損-AC債務工具投資-普通公司(公營)"
                    AC_CompanyBond_Public_Domestic_ImpairmentLoss = AC_CompanyBond_Public_Domestic_ImpairmentLoss + rng.Offset(0, 1).Value
                Case "累積減損-AC債務工具投資-普通公司(民營)"
                    AC_CompanyBond_Private_Domestic_ImpairmentLoss = AC_CompanyBond_Private_Domestic_ImpairmentLoss + rng.Offset(0, 1).Value
            End Select
        Next rng
    End If

    ' HANDLE方式
    
    公債原始成本
    GovBond_Domestic_Cost

    FVPL_GovBond_Domestic_Cost + FVOCI_GovBond_Domestic_Cost + AC_GovBond_Domestic_Cost

    公債
    透過損益按公允價值衡量之金融資產2 A
    FVPL_GovBond_Domestic
    FVPL_GovBond_Domestic_Cost + FVPL_GovBond_Domestic_ValuationAdjust

    公債
    透過其他綜合損益按公允價值衡量之金融資產2 B
    FVOCI_GovBond_Domestic
    FVOCI_GovBond_Domestic_Cost + FVOCI_GovBond_Domestic_ValuationAdjust

    公債
    ac
    AC_GovBond_Domestic
    AC_GovBond_Domestic_Cost + AC_GovBond_Domestic_ImpairmentLoss

    2.公司債		
    2.1.公營事業		
        原始取得成本1		
    CompanyBond_Public_Domestic_Cost
    120050121		強制FVPL金融資產-普通公司債(公營)                 
    121110121		FVOCI債務工具-普通公司債（公營）                  
    122010121		AC債務工具投資-普通公司債(公營)
    
    FVPL_CompanyBond_Public_Domestic_Cost + FVOCI_CompanyBond_Public_Domestic_Cost + AC_CompanyBond_Public_Domestic_Cost

            
        透過損益按公允價值衡量之金融資產2 A		
    FVPL_CompanyBond_Public_Domestic
    120050121		強制FVPL金融資產-普通公司債(公營)                 
    120070121		強制FVPL金融資產評價調整-普通公司債(公營)   
    FVPL_CompanyBond_Public_Domestic_Cost + FVPL_CompanyBond_Public_Domestic_ValuationAdjust
    
            
        透過其他綜合損益按公允價值衡量之金融資產2 B		
    FVOCI_CompanyBond_Public_Domestic
    121110121
    121130121
    ' 重新勾稽
    ' --    

    121110121		FVOCI債務工具-普通公司債（公營）                  
    121130121		FVOCI債務工具評價調整-普通公司債（公營)           
            
    FVOCI_CompanyBond_Public_Domestic_Cost + FVOCI_CompanyBond_Public_Domestic_ValuationAdjust


        按攤銷後成本衡量之債務工具投資2 C		
    AC_CompanyBond_Public_Domestic
    122010121		AC債務工具投資-普通公司債(公營)                   
            

    AC_CompanyBond_Public_Domestic_Cost +


    2.2.民營企業-國內公司債		
        原始取得成本1		
    CompanyBond_Private_Domestic_Cost
    120050123		強制FVPL金融資產-普通公司債(民營)                 
    121110123		FVOCI債務工具-普通公司債（民營）                  
    122010123		AC債務工具投資-普通公司債(民營)                   
    
    FVPL_CompanyBond_Private_Domestic_Cost + FVOCI_CompanyBond_Private_Domestic_Cost + AC_CompanyBond_Private_Domestic_Cost


        透過損益按公允價值衡量之金融資產2 A		
    FVPL_CompanyBond_Private_Domestic
    120050123		強制FVPL金融資產-普通公司債(民營)                 
    120070123		強制FVPL金融資產評價調整-普通公司債(民營)         

    FVPL_CompanyBond_Private_Domestic_Cost + FVPL_CompanyBond_Private_Domestic_ValuationAdjust

        透過其他綜合損益按公允價值衡量之金融資產2 B		
    FVOCI_CompanyBond_Private_Domestic
    121110123		FVOCI債務工具-普通公司債（民營）                  
    121130123		FVOCI債務工具評價調整-普通公司債（民營)               

    FVOCI_CompanyBond_Private_Domestic_Cost + FVOCI_CompanyBond_Private_Domestic_ValuationAdjust

        按攤銷後成本衡量之債務工具投資2 C		
    AC_CompanyBond_Private_Domestic
    122010123		AC債務工具投資-普通公司債(民營)                   
    AC_CompanyBond_Private_Domestic_Cost +
    

    3.股票及股權投資-民營企業
    

        原始取得成本1		    
    Stock_Cost
    1200503
    1210103
    15501
    121019901
    150019901
    ' 重新勾稽
    ' --
    1200503
    FVPL_Stock_PreferredStock_Listed_Cost + FVPL_Stock_CommonStock_Listed_Cost + FVPL_Stock_CommonStock_OTC_Cost + FVPL_Stock_CommonStock_Emergin_Cost
    1210103
    FVOCI_Stock_PreferredStock_Cost + FVOCI_Stock_PreferredStock_Listed_Cost + FVOCI_Stock_CommonStock_Listed_Cost + FVOCI_Stock_CommonStock_OTC_Cost + FVOCI_Stock_CommonStock_Emergin_Cost
    ' * 原來公式寫 15501，實際上這是 15001
    15501 
    EquityMethod_Cost +
    121019901
    FVOCI_Equity_Other_Cost +
    150019901
    EquityMethod_Other_Cost +
            
        透過損益按公允價值衡量之金融資產2 A		
    FVPL_Stock
    1200503
    1200703
    ' 重新勾稽
    ' --

    1200503		強制FVPL金融資產-股票    
    FVPL_Stock_PreferredStock_Listed_Cost + FVPL_Stock_CommonStock_Listed_Cost + FVPL_Stock_CommonStock_OTC_Cost + FVPL_Stock_CommonStock_Emergin_Cost                         
    1200703
    FVPL_Stock_CommonStock_Listed_ValuationAdjust + FVPL_Stock_CommonStock_OTC_ValuationAdjust + FVPL_Stock_CommonStock_Emergin_ValuationAdjust + FVPL_Stock_PreferredStock_Listed_ValuationAdjust
    
        透過其他綜合損益按公允價值衡量之金融資產2 B		
    FVOCI_Stock
    1210103
    1210303
    1210199
    1210399
    ' 重新勾稽
    ' --

    1210103		FVOCI權益工具-股票
    FVOCI_Stock_PreferredStock_Cost + FVOCI_Stock_PreferredStock_Listed_Cost + FVOCI_Stock_CommonStock_Listed_Cost + FVOCI_Stock_CommonStock_OTC_Cost + FVOCI_Stock_CommonStock_Emergin_Cost                          
    1210303		FVOCI權益工具評價調整-股票    
    FVOCI_Stock_PreferredStock_Listed_ValuationAdjust + FVOCI_Stock_CommonStock_Listed_ValuationAdjust + FVOCI_Stock_CommonStock_OTC_ValuationAdjust + FVOCI_Stock_CommonStock_Emergin_ValuationAdjust                    
    1210199		FVOCI權益工具-其他        
    FVOCI_Equity_Other_Cost +
    1210399		FVOCI權益工具評價調整-其他                        
    FVOCI_Equity_Other_ValuationAdjust +
            
        按攤銷後成本衡量之債務工具投資2 C		
        		
        採用權益法之投資-淨額2 E		
    EquityMethod_Stock
    
    15001		採用權益法之投資成本 
    EquityMethod_Other_Cost +                              
    15003		加（減）：採用權益法認列之投資權益調整            
    EquityMethod_ValuationAdjust +            
    4.受益憑證-其他		
            
        原始取得成本1		
    AssetCertificate_Cost
    1200505		強制FVPL金融資產-受益憑證              
    
    FVPL_AssetCertificate_Cost +
            
        透過損益按公允價值衡量之金融資產2 A		
    FVPL_AssetCertificate
    1200505		強制FVPL金融資產-受益憑證                         
    1200705		強制FVPL金融資產評價調整-受益憑證                 
    FVPL_AssetCertificate_Cost + FVPL_AssetCertificate_ValuationAdjust

        透過其他綜合損益按公允價值衡量之金融資產2 B		
            
        按攤銷後成本衡量之債務工具投資2 C		
            
    5.新台幣可轉讓定期存單-中央銀行發行		
            
        原始取得成本1		
    NCD_CentralBank_Cost
    121110911		FVOCI債務工具-央行NCD                             
    122010911		AC債務工具投資-央行NCD   
    FVOCI_NCD_CentralBank_Cost + AC_NCD_CentralBank_Cost
            
        透過損益按公允價值衡量之金融資產2 A		
            
        透過其他綜合損益按公允價值衡量之金融資產2 B		
    FVOCI_NCD_CentralBank
    121110911		FVOCI債務工具-央行NCD                             
    121130911		FVOCI債務工具評價調整-央行NCD                     
    FVOCI_NCD_CentralBank_Cost + FVOCI_NCD_CentralBank_ValuationAdjust
            
        按攤銷後成本衡量之債務工具投資2 C		
    AC_NCD_CentralBank
    122010911		AC債務工具投資-央行NCD                            
    122030911		累積減損-AC債務工具投資-央行NCD                   

    AC_NCD_CentralBank_Cost + AC_NCD_CentralBank_ImpairmentLoss
            
    6.商業本票-民營企業		
            
        原始取得成本1		
    CP_Cost
    120050903		強制FVPL金融資產-商業本票                         
    FVPL_CP_Cost + 
    
            
        透過損益按公允價值衡量之金融資產2 A		
    FVPL_CP
    120050903		強制FVPL金融資產-商業本票                         
    120070903		強制FVPL金融資產評價調整-商業本票                 
    FVPL_CP_Cost + FVPL_CP_ValuationAdjust
            
        透過其他綜合損益按公允價值衡量之金融資產2 B		
            
        按攤銷後成本衡量之債務工具投資2 C		
            
    7.國外機構發行-在國外發行-長期債票券6		
            
        原始取得成本1		
    FinancialBond_Domestic_Cost
    140010147		備供出售-金融債券-海外                  
    AFS_FinancialBond_Domestic_Cost +
    
            
        透過損益按公允價值衡量之金融資產2 A		
            
        透過其他綜合損益按公允價值衡量之金融資產2 B		
    AFS_FinancialBond_Domestic
    140010147		備供出售-金融債券-海外                            
    140030147		備供出售評價調整-金融債券-海外                    

    AFS_FinancialBond_Domestic_Cost + AFS_FinancialBond_Domestic_ValuationAdjust
            
        按攤銷後成本衡量之債務工具投資2 C		

    ' END HANDLE
    
    domesticInterestRevenue = Round(domesticInterestRevenue / 1000, 0)
    
    xlsht.Range("FM11_一利息股息收入_利息_其他").Value = foreignInterestRevenue

    xlsht.Range("FM11_三證券投資處分利益_一年期以上之債權證券").Value = gainOnSecuritiesDisposal

    xlsht.Range("FM11_三證券投資處分損失_一年期以上之債權證券").Value = lossOnSecuritiesDisposal

    xlsht.Range("FM11_五證券投資評價及減損迴轉利益_一年期以上之債權證券").Value = reversalImpairmentPL

    xlsht.Range("FM11_五證券投資評價及減損損失_一年期以上之債權證券").Value = valuationImpairmentLoss

    xlsht.Range("FM11_一利息收入_自中華民國境內其他客戶").Value = domesticInterestRevenue
    
    xlsht.Range("T2:T100").NumberFormat = "#,##,##"

    rpt.SetField "FOA", "FM11_一利息股息收入_利息_其他", CStr(foreignInterestRevenue)

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