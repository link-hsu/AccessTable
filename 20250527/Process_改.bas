Public Sub Process_FM11()
    '=== Equal Setting ===
    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("FM11")

    Dim reportTitle As String
    reportTitle = rpt.ReportName    

    Dim xlsht As Worksheet
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    ' xlsht.Range("A:G").ClearContents
    ' xlsht.Range("T2:T100").ClearContents

    ' xlsht.Cells.ClearContents  ' 或只清指定区
    
    ' === 【修改】 改为调用通用函数 GetMapData(..., "Query") ===
    Dim queryMap As Variant
    queryMap = GetMapData(gDBPath, reportTitle, "QueryTableMap")

    '【修改】改為正確判斷 Array 且至少有一筆
    If Not IsArray(queryMap) Or UBound(queryMap, 1) < 0 Then
        WriteLog "未在 QueryTableMap 找到 " & reportTitle & " 的任何配置"
        Exit Sub
    End If
    
    Dim iMap As Long
    For iMap = 0 To UBound(queryMap, 1)
        Dim tblName As String
        Dim startColLetter As String
        Dim numCols As Long
        Dim dataArr As Variant
        Dim r As Long, c As Long

        tblName = queryMap(iMap, 0)
        startColLetter = queryMap(iMap, 1)
        numCols = CLng(queryMap(iMap, 2))
        
        dataArr = GetAccessDataAsArray(gDBPath, tblName, gDataMonthString)

        '【修改】正確判斷是否有回傳陣列，且至少要有 header（row 0）跟一筆資料（row 1）
        If Not IsArray(dataArr) Or UBound(dataArr, 1) < 1 Then
            WriteLog "資料有誤: " & reportTitle & " | " & tblName & " 無法取得資料或只有欄位名稱"
            GoTo NextMap
        End If

        ' 取得欄位起始欄號        
        Dim startCol As Long
        startCol = xlsht.Range(startColLetter & "1").Column

        ' 將整個 dataArr（含 header）貼到 Excel，上下：0..UBound(dataArr,1)，左右：0..UBound(dataArr,2)
        '【修改】改用 UBound(..., 1)/(..., 2) 以符合 VB 陣列維度
        For r = 0 To UBound(dataArr, 1)
            For c = 0 To UBound(dataArr, 2)
                xlsht.Cells(r + 1, startCol + c).Value = dataArr(r, c)
            Next c
        Next r

NextMap:
    Next iMap

    '--------------
    'Unique Setting
    '--------------
    Dim rngs As Range
    Dim rng As Range

    Dim foreignInterestRevenue As Double
    Dim gainOnSecuritiesDisposal As Double
    Dim lossOnSecuritiesDisposal As Double
    Dim reversalImpairmentPL As Double
    Dim valuationImpairmentLoss As Double
    Dim domesticInterestRevenue As Double

    foreignInterestRevenue = 0
    gainOnSecuritiesDisposal = 0
    lossOnSecuritiesDisposal = 0
    reversalImpairmentPL = 0
    valuationImpairmentLoss = 0
    domesticInterestRevenue = 0
    
    lastRow = xlsht.Cells(xlsht.Rows.Count, "F").End(xlUp).Row
    Set rngs = xlsht.Range("F2:F" & lastRow)

    For Each rng In rngs
        If CStr(rng.Value) = "InterestRevenue" Then
            foreignInterestRevenue = foreignInterestRevenue + rng.Offset(0, 1).Value
        
        ElseIf CStr(rng.Value) = "GainOnDisposal" Then
            gainOnSecuritiesDisposal = gainOnSecuritiesDisposal + rng.Offset(0, 1).Value

        ElseIf CStr(rng.Value) = "LossOnDisposal" Then
            lossOnSecuritiesDisposal = lossOnSecuritiesDisposal + rng.Offset(0, 1).Value

        ElseIf CStr(rng.Value) = "ValuationProfit" Then
            reversalImpairmentPL = reversalImpairmentPL + rng.Offset(0, 1).Value

        ' FVPL 金融資產評價損失
        ElseIf CStr(rng.Value) = "ValuationLoss" Then
            valuationImpairmentLoss = valuationImpairmentLoss + rng.Offset(0, 1).Value

        ' 拆放證券公司息 OSU
        ElseIf CStr(rng.Value) = "OSU息" Then
            domesticInterestRevenue = domesticInterestRevenue + rng.Offset(0, 1).Value
        End If
    Next rng

    foreignInterestRevenue = Round(foreignInterestRevenue / 1000, 0)
    gainOnSecuritiesDisposal = Round(gainOnSecuritiesDisposal / 1000, 0)
    lossOnSecuritiesDisposal = Round(lossOnSecuritiesDisposal / 1000, 0)
    reversalImpairmentPL = Round(reversalImpairmentPL / 1000, 0)
    valuationImpairmentLoss = Round(valuationImpairmentLoss / 1000, 0)
    domesticInterestRevenue = Round(domesticInterestRevenue / 1000, 0)
    
    xlsht.Range("FM11_一利息股息收入_利息_其他").Value = foreignInterestRevenue
    rpt.SetField "FOA", "FM11_一利息股息收入_利息_其他", CStr(foreignInterestRevenue)

    xlsht.Range("FM11_三證券投資處分利益_一年期以上之債權證券").Value = gainOnSecuritiesDisposal
    rpt.SetField "FOA", "FM11_三證券投資處分利益_一年期以上之債權證券", CStr(gainOnSecuritiesDisposal)

    xlsht.Range("FM11_三證券投資處分損失_一年期以上之債權證券").Value = lossOnSecuritiesDisposal
    rpt.SetField "FOA", "FM11_三證券投資處分損失_一年期以上之債權證券", CStr(lossOnSecuritiesDisposal)

    xlsht.Range("FM11_五證券投資評價及減損迴轉利益_一年期以上之債權證券").Value = reversalImpairmentPL
    rpt.SetField "FOA", "FM11_五證券投資評價及減損迴轉利益_一年期以上之債權證券", CStr(reversalImpairmentPL)

    xlsht.Range("FM11_五證券投資評價及減損損失_一年期以上之債權證券").Value = valuationImpairmentLoss
    rpt.SetField "FOA", "FM11_五證券投資評價及減損損失_一年期以上之債權證券", CStr(valuationImpairmentLoss)

    xlsht.Range("FM11_一利息收入_自中華民國境內其他客戶").Value = domesticInterestRevenue
    rpt.SetField "FOA", "FM11_一利息收入_自中華民國境內其他客戶", CStr(domesticInterestRevenue)
    
    xlsht.Range("T2:T100").NumberFormat = "#,##,##"
    

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


Que:

請仔細確認這兩個function在幹嘛，
Public Function GetAccessDataAsArray(ByVal DBPath As String, _
                                     ByVal QueryName As String, _
                                     Optional ByVal dataMonthString As String = vbNullString) As Variant
    Dim conn As Object, cmd As Object, rs As Object
    Dim dataArr As Variant
    Dim colCount As Integer, rowCount As Integer
    Dim headerArr() As String, i As Integer, j As Integer
    On Error GoTo ErrHandler
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = conn
    cmd.CommandText = QueryName
    cmd.CommandType = 4 ' 儲存查詢
    If dataMonthString <> vbNullString Then
        cmd.Parameters.Append cmd.CreateParameter("DataMonthParam", 200, 1, 255, dataMonthString)
    End If
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3 ' adUseClient
    rs.Open cmd
    If rs Is Nothing Or rs.EOF Then
        WriteLog QueryName & "查詢結果為空，請檢查資料庫與查詢條件。"
        GetAccessDataAsArray = Array()
        Exit Function
    End If
    colCount = rs.Fields.Count
    ReDim headerArr(0 To colCount - 1)
    For i = 0 To colCount - 1
        headerArr(i) = rs.Fields(i).Name
    Next i
    dataArr = rs.GetRows()
    rowCount = UBound(dataArr, 2) + 1
    Dim resultArr() As Variant
    ReDim resultArr(0 To rowCount, 0 To colCount - 1)
    ' 第一列存放欄位名稱
    For i = 0 To colCount - 1
        resultArr(0, i) = headerArr(i)
    Next i
    ' 後續存放資料
    For i = 0 To colCount - 1
        For j = 1 To rowCount
            resultArr(j, i) = dataArr(i, j - 1)
        Next j
    Next i
    rs.Close
    conn.Close
    Set rs = Nothing
    Set cmd = Nothing
    Set conn = Nothing
    GetAccessDataAsArray = resultArr
    Exit Function
ErrHandler:
    MsgBox "發生錯誤: " & Err.Description, vbCritical
    WriteLog "發生錯誤: " & Err.Description
    GetAccessDataAsArray = Array()
End Function

Public Function GetMapData(ByVal DBPath As String, _
                           ByVal reportName As String, _
                           ByVal tableName As String) As Variant
    Dim conn As Object, rs As Object
    Dim sql As String
    Dim results() As Variant
    Dim i As Long
    Dim arr As Variant

    WriteLog "DBPath: " & DBPath
    WriteLog "tableName: " & tableName

    ' 1. 建立 ADO 连接
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath
    
    ' 2. 根据 tableName 组织 SQL
    Select Case LCase(tableName)
        Case "fiedlvaluepositionmap"
            sql = "SELECT fp.TargetSheetName, fp.SourceNameTag, fp.TargetCellAddress " & _
                  "FROM FiedlValuePositionMap AS fp " & _
                  "INNER JOIN Report AS r ON fp.ReportID = r.ReportID " & _
                  "WHERE r.ReportName = '" & reportName & "' " & _
                  "ORDER BY fp.DataId;"
        Case "querytablemap"
            sql = "SELECT qm.QueryTableName, qm.ImportColName, qm.ImportColNumber " & _
                  "FROM QueryTableMap AS qm " & _
                  "INNER JOIN Report AS r ON qm.ReportID = r.ReportID " & _
                  "WHERE r.ReportName = '" & reportName & "' " & _
                  "ORDER BY qm.DataId;"
        Case Else
            GetMapData = Array()
            Exit Function
    End Select

    WriteLog "sql: " & sql
    
    Set rs = conn.Execute(sql)
    If rs.EOF Then
        GetMapData = Array()    ' 没有紀錄
    Else
        arr = rs.GetRows
    
        ' 转置成 arr2(紀錄, 字段)
        Dim recCount As Long, fldCount As Long
        fldCount = UBound(arr, 1) + 1
        recCount = UBound(arr, 2) + 1
        ReDim results(0 To recCount - 1, 0 To fldCount - 1)
    
        Dim r As Long, f As Long
        For f = 0 To fldCount - 1
            For r = 0 To recCount - 1
                results(r, f) = arr(f, r)
            Next r
        Next f
    
        GetMapData = results
    End If
    
    rs.Close: conn.Close
    Set rs = Nothing: Set conn = Nothing
End Function




在確認我下面的Sub有沒有正確使用
Public Sub Process_FM11()
    '=== Equal Setting ===
    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("FM11")

    Dim reportTitle As String
    reportTitle = rpt.ReportName    

    Dim xlsht As Worksheet
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    ' xlsht.Range("A:G").ClearContents
    ' xlsht.Range("T2:T100").ClearContents

    ' xlsht.Cells.ClearContents  ' 或只清指定区
    
    ' === 【修改】 改为调用通用函数 GetMapData(..., "Query") ===
    Dim queryMap As Variant
    queryMap = GetMapData(gDBPath, reportTitle, "QueryTableMap")

    If IsNull(queryMap) Then
        WriteLog "資料有誤: " & reportTitle & " 在QueryTableMap資料表無資料" 
        Exit Sub
    End If     
    
    If Not IsArray(queryMap) Or UBound(queryMap) < 0 Then
        WriteLog "未在 QueryMap 找到 " & reportTitle & " 的任何查询配置"
        Exit Sub
    End If
    
    Dim iMap As Long
    For iMap = 0 To UBound(queryMap, 1)
        Dim tblName As String
        Dim startColLetter As String
        Dim numCols As Long
        Dim dataArr As Variant
        Dim r As Long, c As Long
        tblName = queryMap(iMap, 0)
        startColLetter = queryMap(iMap, 1)
        numCols = CLng(queryMap(iMap, 2))
        
        dataArr = GetAccessDataAsArray(gDBPath, tblName, gDataMonthString)
        If Not IsArray(dataArr) Then
            WriteLog "資料有誤: " & reportTitle & "| " & tblName & " 資料表無資料"
            GoTo NextMap
        End If
        
        Dim startCol As Long
        startCol = xlsht.Range(startColLetter & "1").Column
        
        For r = 0 To UBound(dataArr, 0)
            For c = 0 To numCols - 1
                xlsht.Cells(r + 1, startCol + c).Value = dataArr(r, c)
            Next c
        Next r
NextMap:
    Next iMap

    '--------------
    'Unique Setting
    '--------------

    

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

我需要將dataArray連同header都列到excel上，
請注意Ubound的使用，
以及col header或row資料的位置和使用有沒有正確，
如果有不適當的地方請幫我修改，
給我完整版本，並標示出修改哪邊


Ans:

以下是修正後的 `Process_FM11`，重點修改及說明已用 `'【修改】` 標示在程式碼中：

```vb
Public Sub Process_FM11()
    '=== Equal Setting ===
    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("FM11")
    
    Dim reportTitle As String
    reportTitle = rpt.ReportName    
    
    Dim xlsht As Worksheet
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    ' 先清空工作表指定範圍（視需求自行調整）
    xlsht.Range("A:Z").ClearContents   '【修改】擴大清除範圍

    ' === 以通用函數讀取 QueryTableMap ===
    Dim queryMap As Variant
    queryMap = GetMapData(gDBPath, reportTitle, "QueryTableMap")

    '【修改】改為正確判斷 Array 且至少有一筆
    If Not IsArray(queryMap) Or UBound(queryMap, 1) < 0 Then
        WriteLog "未在 QueryTableMap 找到 " & reportTitle & " 的任何配置"
        Exit Sub
    End If
    
    Dim iMap As Long
    For iMap = 0 To UBound(queryMap, 1)
        Dim tblName As String
        Dim startColLetter As String
        Dim numCols As Long
        Dim dataArr As Variant
        Dim r As Long, c As Long
        
        tblName = queryMap(iMap, 0)
        startColLetter = queryMap(iMap, 1)
        numCols = CLng(queryMap(iMap, 2))
        
        dataArr = GetAccessDataAsArray(gDBPath, tblName, gDataMonthString)
        
        '【修改】正確判斷是否有回傳陣列，且至少要有 header（row 0）跟一筆資料（row 1）
        If Not IsArray(dataArr) Or UBound(dataArr, 1) < 1 Then
            WriteLog "資料有誤: " & reportTitle & " | " & tblName & " 無法取得資料或只有欄位名稱"
            GoTo NextMap
        End If
        
        ' 取得欄位起始欄號
        Dim startCol As Long
        startCol = xlsht.Range(startColLetter & "1").Column
        
        ' 將整個 dataArr（含 header）貼到 Excel，上下：0..UBound(dataArr,1)，左右：0..UBound(dataArr,2)
        '【修改】改用 UBound(..., 1)/(..., 2) 以符合 VB 陣列維度
        For r = 0 To UBound(dataArr, 1)
            For c = 0 To UBound(dataArr, 2)
                xlsht.Cells(r + 1, startCol + c).Value = dataArr(r, c)
            Next c
        Next r
        
NextMap:
    Next iMap

    '--------------  
    ' Unique Setting  
    ' 1. Validation  
    ' 2. Update Access DB  
    If rpt.ValidateFields() Then
        Dim key As Variant
        Dim allValues As Object, allPositions As Object
        Set allValues = rpt.GetAllFieldValues()
        Set allPositions = rpt.GetAllFieldPositions()
        For Each key In allValues.Keys
            UpdateRecord gDBPath, gDataMonthString, rpt.ReportName, _
                         key, allPositions(key), allValues(key)
        Next key
    End If
    
    ' 將工作表標籤設為黃色
    xlsht.Tab.ColorIndex = 6
End Sub
```

---

### 主要修改說明

1. **清除範圍**

   * 原本註解掉的清除程式碼改放寬到 `A:Z`，確保所有可能貼入之欄位都被清空。
2. **判斷 `queryMap`**

   * 移除 `IsNull` 檢查，改以 `Not IsArray(...) Or UBound(...,1)<0` 確認回傳確實是含一筆以上記錄的二維陣列。
3. **判斷 `dataArr`**

   * 檢查不僅要是陣列，且 `UBound(dataArr,1) >= 1`，確保除了 header 外至少有一筆資料。
4. **使用正確的 `UBound` 維度**

   * VB 的二維陣列以第一維（`UBound(...,1)`）為「列」，第二維（`UBound(...,2)`）為「欄」。原程式誤用 `UBound(dataArr,0)`，已統一改為正確的維度號。
5. **資料貼入方式**

   * 內層迴圈直接從 `0` 跑到 `UBound(dataArr,2)`，完整含欄位名稱與所有資料。

以上修改可確保 header 與資料完整貼入，並避免空陣列或維度誤用的錯誤。若有其他需求或問題，再請告知！
