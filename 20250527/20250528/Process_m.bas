Public Sub Process_TABLE20()
    '=== Equal Setting ===
    Dim rpt As clsReport
    Set rpt = gReports("TABLE20")
    
    Dim reportTitle As String
    reportTitle = rpt.ReportName
    
    Dim xlsht As Worksheet
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    ' 清空舊資料（視需求）
    ' xlsht.Cells.ClearContents
    
    ' --- 【修改1】宣告 Collection 來存放所有匯入表格的起始欄號 ---
    Dim importCols As New Collection
    
    ' 取得 QueryTableMap
    Dim queryMap As Variant
    queryMap = GetMapData(gDBPath, reportTitle, "QueryTableMap")
    If Not IsArray(queryMap) Or UBound(queryMap, 1) < 0 Then
        WriteLog "未在 QueryTableMap 找到 " & reportTitle & " 的任何配置"
        Exit Sub
    End If
    
    Dim iMap As Long
    For iMap = 0 To UBound(queryMap, 1)
        Dim tblName       As String
        Dim startColLetter As String
        Dim dataArr       As Variant
        Dim r             As Long, c As Long
        
        tblName         = queryMap(iMap, 0)
        startColLetter  = queryMap(iMap, 1)
        
        '【修改1】把欄位字母轉成數字並存入 importCols
        Dim startCol As Long
        startCol = xlsht.Range(startColLetter & "1").Column
        importCols.Add startCol
        
        dataArr = GetAccessDataAsArray(gDBPath, tblName, gDataMonthString)
        If Not IsArray(dataArr) Or UBound(dataArr, 1) < 1 Then
            WriteLog "資料有誤: " & reportTitle & " | " & tblName & " 無法取得資料"
            GoTo NextMap
        End If
        
        ' 把資料貼到 Excel
        For r = 0 To UBound(dataArr, 1)
            For c = 0 To UBound(dataArr, 2)
                xlsht.Cells(r + 1, startCol + c).Value = dataArr(r, c)
            Next c
        Next r
NextMap:
    Next iMap
    
    ' 宣告並初始化累計變數
    Dim RP_GovBond_Cost     As Double: RP_GovBond_Cost = 0    '【修改3】
    Dim RP_CompanyBond_Cost As Double: RP_CompanyBond_Cost = 0
    Dim RP_CP_Cost          As Double: RP_CP_Cost = 0
    
    Dim lastRow1 As Long, lastRow2 As Long
    Dim rngs1 As Range, rngs2 As Range, rng As Range
    
    ' --- 【修改1.2】不迴圈，直接取用 importCols(1) 和 importCols(2) ---
    If importCols.Count >= 1 Then
        lastRow1 = xlsht.Cells(xlsht.Rows.Count, importCols(1)).End(xlUp).Row
        Set rngs1 = xlsht.Range(xlsht.Cells(2, importCols(1)), xlsht.Cells(lastRow1, importCols(1)))
        
        For Each rng In rngs1
            Select Case CStr(rng.Value)
                Case "RP_GovBond_Cost"
                    RP_GovBond_Cost = RP_GovBond_Cost + rng.Offset(0, 1).Value
                Case "AC_CompanyBond_Domestic_ImpairmentLoss"
                    RP_CompanyBond_Cost = RP_CompanyBond_Cost + rng.Offset(0, 1).Value
            End Select
        Next rng
    End If
    
    If importCols.Count >= 2 Then
        lastRow2 = xlsht.Cells(xlsht.Rows.Count, importCols(2)).End(xlUp).Row
        Set rngs2 = xlsht.Range(xlsht.Cells(2, importCols(2)), xlsht.Cells(lastRow2, importCols(2)))
        
        For Each rng In rngs2
            ' 如果第二筆表也有需要累計的 tag，可以在這裡加
            Select Case CStr(rng.Value)
                Case "RP_GovBond_Cost"
                    RP_GovBond_Cost = RP_GovBond_Cost + rng.Offset(0, 1).Value
                Case "AC_CompanyBond_Domestic_ImpairmentLoss"
                    RP_CompanyBond_Cost = RP_CompanyBond_Cost + rng.Offset(0, 1).Value
            End Select
        Next rng
    End If
    
    ' 換算並寫回工作表
    RP_GovBond_Cost     = Round(RP_GovBond_Cost / 1000, 0)
    RP_CompanyBond_Cost = Round(RP_CompanyBond_Cost / 1000, 0)
    RP_CP_Cost          = Round(RP_CP_Cost / 1000, 0)
    
    xlsht.Range("Table20_0200_二公債_民營企業_其他到期日").Value = RP_GovBond_Cost
    xlsht.Range("Table20_0300_三公司債_民營企業_其他到期日").Value = RP_CompanyBond_Cost
    xlsht.Range("Table20_0400_四商業本票_民營企業_其他到期日").Value = RP_CP_Cost
    
    ' --- 【修改2】用 FieldValuePositionMap 批次填 rpt ---
    Dim fvMap    As Variant
    Dim iMap2    As Long
    Dim tgtSheet As String, srcTag As String, srcAddr As String, srcVal As Variant
    
    fvMap = GetMapData(gDBPath, reportTitle, "FieldValuePositionMap")
    If Not IsNull(fvMap) And IsArray(fvMap) Then
        For iMap2 = 0 To UBound(fvMap, 1)
            tgtSheet = CStr(fvMap(iMap2, 0))
            srcTag    = CStr(fvMap(iMap2, 1))
            srcAddr   = CStr(fvMap(iMap2, 2))
            
            On Error Resume Next
            srcVal = ThisWorkbook.Sheets(tgtSheet).Range(srcAddr).Value
            On Error GoTo 0
            
            rpt.SetField tgtSheet, srcTag, srcVal
        Next iMap2
    Else
        WriteLog "無法取得 FieldValuePositionMap for " & reportTitle
    End If
    
    ' 驗證並回寫 DB
    If rpt.ValidateFields() Then
        Dim allVals As Object, allPos As Object, key As Variant
        Set allVals = rpt.GetAllFieldValues()
        Set allPos  = rpt.GetAllFieldPositions()
        
        For Each key In allVals.Keys
            UpdateRecord gDBPath, gDataMonthString, rpt.ReportName, key, allPos(key), allVals(key)
        Next key
    End If
    
    ' 標示此分頁為已處理
    xlsht.Tab.ColorIndex = 6
End Sub
