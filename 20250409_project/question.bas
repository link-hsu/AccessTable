我想寫一個Sub可以傳入分頁名稱和起始Row index和Column index和要使用的Function名稱，
可以選擇Call 這兩隻其中一隻Function，
使用
        For j = 0 To UBound(dataArr, 2)
            For i = 0 To UBound(dataArr, 1)
                xlsht.Cells(i + 1, j + 1).Value = dataArr(i, j)
            Next i
        Next j
將獲得的Array依序貼於Excel分頁的特定欄位中，
並請給我使用的範例

Function GetAccountCodeMapFlex( _
    CategoryParam As Variant, _
    Optional GroupFlagParam As Variant, _       '【★】新增：GroupFlagParam
    Optional SubTypeParam As Variant, _
    Optional TypeParam As Variant) As Variant

    Dim dbPath As String
    dbPath = "C:\Your\Path\To\Database.accdb"  ' ← 修改為你的 Access 資料庫路徑

    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")

    Dim sql As String
    Dim whereClauses As Collection
    Set whereClauses = New Collection

    '── 1) 必填：Category
    Dim catClause As String
    catClause = BuildInClauseParam(CategoryParam, "Category")
    If catClause = "" Then
        GetAccountCodeMapFlex = CVErr(xlErrValue)
        Exit Function
    End If
    whereClauses.Add catClause

    '── 2) 可選：GroupFlag
    Dim grpClause As String
    grpClause = BuildInClauseParam(GroupFlagParam, "GroupFlag")   '【★】新增
    If grpClause <> "" Then whereClauses.Add grpClause         '【★】新增

    '── 3) 可選：AssetMeasurementSubType
    Dim subClause As String
    subClause = BuildInClauseParam(SubTypeParam, "AssetMeasurementSubType")
    If subClause <> "" Then whereClauses.Add subClause

    '── 4) 可選：AssetMeasurementType
    Dim typeClause As String
    typeClause = BuildInClauseParam(TypeParam, "AssetMeasurementType")
    If typeClause <> "" Then whereClauses.Add typeClause

    '── 組出 SQL，調整回傳欄位順序：AccountCode, GroupFlag, AccountTitle
    sql = "SELECT AccountCode, GroupFlag, AccountTitle FROM AccountCodeMap WHERE "
    Dim part As Variant
    For Each part In whereClauses
        sql = sql & part & " AND "
    Next
    sql = Left(sql, Len(sql) - 5)  ' 去掉最後的 " AND "

    '── 執行查詢
    On Error GoTo ErrHandler
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"
    rs.Open sql, conn

    '── 處理回傳：第一列放欄位名稱，之後才是資料
    If rs.EOF Then
        ' 若無資料，只回欄位名稱
        Dim noDataArr() As Variant
        ReDim noDataArr(0 To 0, 0 To rs.Fields.Count - 1)
        Dim fi As Integer
        For fi = 0 To rs.Fields.Count - 1
            noDataArr(0, fi) = rs.Fields(fi).Name
        Next
        GetAccountCodeMapFlex = noDataArr

    Else
        Dim rawArr As Variant
        rawArr = rs.GetRows()  ' fields × rows

        Dim nFields As Long, nRows As Long
        nFields = UBound(rawArr, 1) + 1  ' 3 個欄位
        nRows   = UBound(rawArr, 2) + 1

        Dim outArr() As Variant
        ReDim outArr(0 To nRows, 0 To nFields - 1)

        ' 第一列：欄位名稱
        Dim f As Long, r As Long
        For f = 0 To nFields - 1
            outArr(0, f) = rs.Fields(f).Name
        Next

        ' 之後各列：資料
        For r = 0 To nRows - 1
            For f = 0 To nFields - 1
                outArr(r + 1, f) = rawArr(f, r)
            Next
        Next

        GetAccountCodeMapFlex = outArr
    End If

    rs.Close: conn.Close
    Set rs = Nothing: Set conn = Nothing
    Exit Function

ErrHandler:
    GetAccountCodeMapFlex = "Error: " & Err.Description
    On Error GoTo 0
End Function


Function GetSubtotalBalance( _
    DataMonthStringParam As String, _
    GroupByMode As String, _
    CurrencyTypeParam As String, _
    CategoryParam As Variant, _
    Optional GroupFlagParam As Variant, _     '【★】新增 GroupFlag 篩選
    Optional SubTypeParam As Variant, _
    Optional TypeParam As Variant _
) As Variant

    Dim dbPath As String
    dbPath = "C:\Your\Path\To\Database.accdb"  ' ← 修改成你的 Access 檔案路徑

    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    Dim rs   As Object: Set rs   = CreateObject("ADODB.Recordset")

    Dim sql As String
    Dim whereACM As Collection: Set whereACM = New Collection
    Dim subSQL   As String

    '── 1) AccountCodeMap 的 WHERE 篩選 ──────────────────────────────
    '   Category 必填
    Dim catClause As String
    catClause = BuildInClauseParam(CategoryParam, "AccountCodeMap.Category")
    If catClause = "" Then
        GetSubtotalBalance = CVErr(xlErrValue)
        Exit Function
    End If
    whereACM.Add catClause

    '   GroupFlag 可選  【★】
    Dim grpClause As String
    grpClause = BuildInClauseParam(GroupFlagParam, "AccountCodeMap.GroupFlag")
    If grpClause <> "" Then whereACM.Add grpClause

    '   AssetMeasurementSubType 可選
    Dim subClause As String
    subClause = BuildInClauseParam(SubTypeParam, "AccountCodeMap.AssetMeasurementSubType")
    If subClause <> "" Then whereACM.Add subClause

    '   AssetMeasurementType 可選
    Dim typeClause As String
    typeClause = BuildInClauseParam(TypeParam, "AccountCodeMap.AssetMeasurementType")
    If typeClause <> "" Then whereACM.Add typeClause

    '── 2) 內層子查詢 subSQL ───────────────────────────────────────
    subSQL = ""
    subSQL = subSQL & "SELECT OBU_AC4603.AccountCode, OBU_AC4603.NetBalance" & vbCrLf
    subSQL = subSQL & "FROM OBU_AC4603" & vbCrLf
    subSQL = subSQL & "WHERE OBU_AC4603.CurrencyType = '" & Replace(CurrencyTypeParam, "'", "''") & "'" & vbCrLf
    subSQL = subSQL & "  AND OBU_AC4603.DataMonthString = '" & Replace(DataMonthStringParam, "'", "''") & "'"

    '── 3) 根據 GroupByMode 決定 SELECT / GROUP BY 片段 ─────────────【★】
    Dim selExpr   As String
    Dim grpFields As String

    Select Case LCase(GroupByMode)
        Case "category"
            selExpr   = "AccountCodeMap.Category AS MeasurementCategory"
            grpFields = "AccountCodeMap.Category"
        Case "type_category"
            selExpr   = "AccountCodeMap.AssetMeasurementType & '_' & AccountCodeMap.Category AS MeasurementCategory"
            grpFields = "AccountCodeMap.AssetMeasurementType, AccountCodeMap.Category"
        Case "subtype_category"
            selExpr   = "AccountCodeMap.AssetMeasurementSubType & '_' & AccountCodeMap.Category AS MeasurementCategory"
            grpFields = "AccountCodeMap.AssetMeasurementSubType, AccountCodeMap.Category"
        Case Else
            ' 不支援的模式 → 錯誤
            GetSubtotalBalance = CVErr(xlErrValue)
            Exit Function
    End Select

    '── 4) 組合最終 SQL ───────────────────────────────────────────
    sql = ""
    sql = sql & "SELECT " & selExpr & ", Sum(oa.NetBalance) AS SubtotalBalance" & vbCrLf
    sql = sql & "FROM AccountCodeMap" & vbCrLf
    sql = sql & "INNER JOIN (" & vbCrLf & subSQL & vbCrLf & ") AS oa" & vbCrLf
    sql = sql & "  ON AccountCodeMap.AccountCode = oa.AccountCode" & vbCrLf

    '   外層 WHERE
    sql = sql & "WHERE "
    Dim part As Variant
    For Each part In whereACM
        sql = sql & part & " AND "
    Next
    sql = Left(sql, Len(sql) - 5)  ' 移除最後的 " AND "

    sql = sql & vbCrLf & "GROUP BY " & grpFields

    '── 5) 執行並回傳（含 header） ─────────────────────────────────
    On Error GoTo ErrHandler
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"
    rs.Open sql, conn

    ' 如果沒有資料，只回 header
    If rs.EOF Then
        Dim noArr() As Variant
        ReDim noArr(0 To 0, 0 To rs.Fields.Count - 1)
        Dim fi As Integer
        For fi = 0 To rs.Fields.Count - 1
            noArr(0, fi) = rs.Fields(fi).Name
        Next
        GetSubtotalBalance = noArr
    Else
        ' 抓原始 (fields × rows)
        Dim rawArr As Variant
        rawArr = rs.GetRows()

        Dim nF As Long, nR As Long
        nF = UBound(rawArr, 1) + 1
        nR = UBound(rawArr, 2) + 1

        ' 建立 (rows+1) × fields 輸出陣列
        Dim outArr() As Variant
        ReDim outArr(0 To nR, 0 To nF - 1)

        ' 第一列填 header
        For fi = 0 To nF - 1
            outArr(0, fi) = rs.Fields(fi).Name
        Next
        ' 之後填 data
        Dim r As Long
        For r = 0 To nR - 1
            For fi = 0 To nF - 1
                outArr(r + 1, fi) = rawArr(fi, r)
            Next
        Next

        GetSubtotalBalance = outArr
    End If

    rs.Close: conn.Close
    Set rs = Nothing: Set conn = Nothing
    Exit Function

ErrHandler:
    GetSubtotalBalance = "Error: " & Err.Description
    On Error GoTo 0
End Function
