Public Sub Process_TABLE20()
    '=== Equal Setting ===
    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("TABLE20")

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
    
    Dim importCols As New Collection
    Dim iMap As Long
    For iMap = 0 To UBound(queryMap, 1)
        Dim tblName As String
        Dim startColLetter As String
        Dim dataArr As Variant
        Dim r As Long, c As Long

        tblName = queryMap(iMap, 0)
        startColLetter = queryMap(iMap, 1)

        '【修改1】把欄位字母轉成數字並存入 importCols
        Dim startCol As Long
        startCol = xlsht.Range(startColLetter & "1").Column
        importCols.Add startCol
        
        dataArr = GetAccessDataAsArray(gDBPath, tblName, gDataMonthString)

        '【修改】正確判斷是否有回傳陣列，且至少要有 header（row 0）跟一筆資料（row 1）
        If Not IsArray(dataArr) Or UBound(dataArr, 1) < 1 Then
            WriteLog "資料有誤: " & reportTitle & " | " & tblName & " 無法取得資料或只有欄位名稱"
            GoTo NextMap
        End If

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

    Dim RP_GovBond_Cost As Double: RP_GovBond_Cost = 0
    Dim RP_CompanyBond_Cost As Double: RP_CompanyBond_Cost = 0
    Dim RP_CP_Cost As Double: RP_CP_Cost = 0

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
    
    RP_GovBond_Cost = Round(RP_GovBond_Cost / 1000, 0)
    RP_CompanyBond_Cost = Round(RP_CompanyBond_Cost / 1000, 0)
    RP_CP_Cost = Round(RP_CP_Cost / 1000, 0)
    
    xlsht.Range("Table20_0200_二公債_民營企業_其他到期日").Value = RP_GovBond_Cost
    xlsht.Range("Table20_0300_三公司債_民營企業_其他到期日").Value = RP_CompanyBond_Cost
    xlsht.Range("Table20_0400_四商業本票_民營企業_其他到期日").Value = RP_CP_Cost
    
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


    Dim queryMap As Variant
    queryMap = GetMapData(gDBPath, reportTitle, "QueryTableMap")

    '【修改】改為正確判斷 Array 且至少有一筆
    If Not IsArray(queryMap) Or UBound(queryMap, 1) < 0 Then
        WriteLog "未在 QueryTableMap 找到 " & reportTitle & " 的任何配置"
        Exit Sub
    End If
    
    Dim importCols As New Collection
    Dim iMap As Long
    For iMap = 0 To UBound(queryMap, 1)
        Dim tblName As String
        Dim startColLetter As String
        Dim dataArr As Variant
        Dim r As Long, c As Long

        tblName = queryMap(iMap, 0)
        startColLetter = queryMap(iMap, 1)

        '【修改1】把欄位字母轉成數字並存入 importCols
        Dim startCol As Long
        startCol = xlsht.Range(startColLetter & "1").Column
        importCols.Add startCol
        
        dataArr = GetAccessDataAsArray(gDBPath, tblName, gDataMonthString)

        '【修改】正確判斷是否有回傳陣列，且至少要有 header（row 0）跟一筆資料（row 1）
        If Not IsArray(dataArr) Or UBound(dataArr, 1) < 1 Then
            WriteLog "資料有誤: " & reportTitle & " | " & tblName & " 無法取得資料或只有欄位名稱"
            GoTo NextMap
        End If

        ' 將整個 dataArr（含 header）貼到 Excel，上下：0..UBound(dataArr,1)，左右：0..UBound(dataArr,2)
        '【修改】改用 UBound(..., 1)/(..., 2) 以符合 VB 陣列維度
        For r = 0 To UBound(dataArr, 1)
            For c = 0 To UBound(dataArr, 2)
                xlsht.Cells(r + 1, startCol + c).Value = dataArr(r, c)
            Next c
        Next r
NextMap:
    Next iMap


這區塊我在每一張報表都會使用到，如果一直重複貼也是可以，但是有沒有辦法包在一塊，直接去call



' Ans:

你可以把重複的「從 QueryTableMap 取設定→抓取資料→貼到工作表→回傳 importCols」那一大段，抽成一個「通用 Sub/Function」，然後在各張報表的主流程裡只要呼叫它就好。下面提供一個範例，示範如何封裝並回傳 `importCols`，你可以依照自己專案的命名和參數微調：

```vb
' 放在一個共用 Module 裡
Public Function ImportQueryTables( _
    ByVal dbPath As String, _
    ByVal ws As Worksheet, _
    ByVal reportName As String, _
    ByVal dataMonth As String _
) As Collection
    Dim queryMap As Variant
    Dim importCols As New Collection
    Dim iMap As Long
    
    ' 1. 取配置
    queryMap = GetMapData(dbPath, reportName, "QueryTableMap")
    If Not IsArray(queryMap) Or UBound(queryMap, 1) < 0 Then
        WriteLog "未在 QueryTableMap 找到 " & reportName & " 的任何配置"
        Exit Function
    End If
    
    ' 2. 迭代每張子表
    For iMap = 0 To UBound(queryMap, 1)
        Dim tblName As String
        Dim colLetter As String
        Dim startCol As Long
        Dim dataArr As Variant
        Dim r As Long, c As Long
        
        tblName    = CStr(queryMap(iMap, 0))
        colLetter  = CStr(queryMap(iMap, 1))
        ' A→數字
        startCol   = ws.Range(colLetter & "1").Column
        importCols.Add startCol
        
        ' 取資料
        dataArr = GetAccessDataAsArray(dbPath, tblName, dataMonth)
        If Not IsArray(dataArr) Or UBound(dataArr, 1) < 1 Then
            WriteLog "資料有誤: " & reportName & " | " & tblName & " 無法取得資料或只有欄位名稱"
            GoTo NextMap
        End If
        
        ' 貼到工作表
        For r = 0 To UBound(dataArr, 1)
            For c = 0 To UBound(dataArr, 2)
                ws.Cells(r + 1, startCol + c).Value = dataArr(r, c)
            Next c
        Next r
NextMap:
    Next iMap
    
    ' 回傳已記錄的所有起始欄位
    Set ImportQueryTables = importCols
End Function
```

接著你的 `Process_TABLE20` 就只要改成：

```vb
Public Sub Process_TABLE20()
    Dim rpt As clsReport
    Set rpt = gReports("TABLE20")
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(rpt.ReportName)
    
    ' 呼叫通用函式，並取得 importCols
    Dim importCols As Collection
    Set importCols = ImportQueryTables( _
        gDBPath, ws, rpt.ReportName, gDataMonthString _
    )
    
    If importCols Is Nothing Or importCols.Count = 0 Then
        Exit Sub
    End If
    
    ' 下面就可以用 importCols(1)、importCols(2) … 來做你原本的運算邏輯
    ' … （不需重複那一大段 Query → 貼工作表 的程式碼）
    
    ' … 其餘 Unique Setting、FieldValuePositionMap、UpdateRecord 等
End Sub
```

這樣：

1. **單一維護點**：把抓 Query 和貼工作表的邏輯全寫在 `ImportQueryTables`，未來若要改，只改這裡。
2. **清晰流程**：主流程裡一行呼叫完，馬上拿到 `importCols`，後續累計／計算就跟原本一樣。
3. **可重用性高**：其他報表只要改呼叫參數（reportName、工作表物件），即可沿用同一段程式。

你也可以將回傳的 `Collection` 改成陣列或自訂類別，依專案需要微調即可。
