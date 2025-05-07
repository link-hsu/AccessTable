Thought for 4 seconds


以下 VBA 範例示範一支「通吃」兩種輸入方式（Range 或 逗號分隔字串）的函數 `GetAccountCodeMapFlex`，並附上多組使用範例。

---

## 一、VBA 程式碼

```vb
'──────────────────────────────────────────
' 主函數：支援 Range 或 逗號分隔字串
Public Function GetAccountCodeMapFlex(CategoryParam As Variant, _
                                      Optional SubTypeParam As Variant, _
                                      Optional TypeParam As Variant) As Variant
    Dim dbPath As String
    dbPath = "C:\Users\David Hsu\Documents\Database1.accdb"  ' ← 修改為你的 Access 資料庫路徑

    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")

    Dim sql As String
    Dim whereClauses As Collection
    Set whereClauses = New Collection

    ' 1. 必填：Category
    Dim catClause As String
    catClause = BuildInClauseParam(CategoryParam, "Category")
    If catClause = "" Then
        GetAccountCodeMapFlex = CVErr(xlErrValue)
        Exit Function
    End If
    whereClauses.Add catClause

    ' 2. 可選：AssetMeasurementSubType
    Dim subClause As String
    subClause = BuildInClauseParam(SubTypeParam, "AssetMeasurementSubType")
    If subClause <> "" Then whereClauses.Add subClause

    ' 3. 可選：AssetMeasurementType
    Dim typeClause As String
    typeClause = BuildInClauseParam(TypeParam, "AssetMeasurementType")
    If typeClause <> "" Then whereClauses.Add typeClause

    ' 組出 SQL
    sql = "SELECT AccountCode, AccountTitle FROM AccountCodeMap WHERE "
    Dim part As Variant
    For Each part In whereClauses
        sql = sql & part & " AND "
    Next
    sql = Left(sql, Len(sql) - 5)  ' 去掉最後的 " AND "

    ' 執行查詢
    On Error GoTo ErrHandler
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"
    rs.Open sql, conn
    If rs.EOF Then
        GetAccountCodeMapFlex = Array()  ' or "No records"
    Else
        GetAccountCodeMapFlex = rs.GetRows()
    End If

    rs.Close: conn.Close
    Set rs = Nothing: Set conn = Nothing
    Exit Function

ErrHandler:
    GetAccountCodeMapFlex = "Error: " & Err.Description
    On Error GoTo 0
End Function

'──────────────────────────────────────────
' 輔助函數：根據參數類型（Range 或 String）建構 IN 子句
Private Function BuildInClauseParam(paramIn As Variant, fieldName As String) As String
    Dim items As Collection
    Set items = New Collection
    Dim v As Variant

    ' 1) 如果是 Range
    If TypeName(paramIn) = "Range" Then
        For Each v In paramIn.Cells
            If Trim(v.Value & "") <> "" Then items.Add CStr(v.Value)
        Next

    ' 2) 如果是 String
    ElseIf VarType(paramIn) = vbString Then
        Dim arr As Variant
        arr = Split(paramIn, ",")
        For Each v In arr
            v = Trim(v)
            If v <> "" Then items.Add v
        Next
    End If

    ' 組出 IN (..)
    Dim sqlVals As String
    Dim elem As Variant
    For Each elem In items
        sqlVals = sqlVals & "'" & Replace(elem, "'", "''") & "',"
    Next
    If sqlVals <> "" Then
        sqlVals = Left(sqlVals, Len(sqlVals) - 1)
        BuildInClauseParam = fieldName & " IN (" & sqlVals & ")"
    Else
        BuildInClauseParam = ""
    End If
End Function
'──────────────────────────────────────────
```

---

## 二、使用範例

假設你已經把上述程式貼到 VBA 模組中並存檔，以下示範三種呼叫方式。

### 1. 全 Range 方式

工作表上：

* A1\:A2 放 **Category**（例如：外幣債、國內股）
* B1\:B3 放 **AssetMeasurementSubType**（例如：FVPL\_GovBond\_Foreign、FVOCI\_CompanyBond\_Foreign、AC\_GovBond\_Foreign）
* C1\:C2 放 **AssetMeasurementType**（例如：FVPL、FVOCI）

**公式**（輸出至動態陣列支援的 Excel）：

```excel
=GetAccountCodeMapFlex(A1:A2, B1:B3, C1:C2)
```

---

### 2. 全 逗號分隔字串 方式

直接在公式中傳入三個逗號串：

```excel
=GetAccountCodeMapFlex("外幣債,國內股","FVPL_GovBond_Foreign,FVOCI_CompanyBond_Foreign,AC_GovBond_Foreign","FVPL,FVOCI")
```

---

### 3. 混用方式

* **Category** 用 Range（A1\:A2）
* **SubType** 用 逗號字串
* **Type** 不帶（省略）

```excel
=GetAccountCodeMapFlex(
  A1:A2,
  "FVPL_GovBond_Foreign,AC_GovBond_Foreign",
  )
```

> **注意**：如果你只想依 Category 篩選，後面兩個參數都可以完全不填：
>
> ```excel
> =GetAccountCodeMapFlex(A1:A2)
> ```

---

以上程式與範例，讓你可以靈活地以「Range」或「逗號串」兩種方式傳入參數，無須重寫新函數。若有任何疑問，再告訴我！





' ==============================================
Answer2





Option Explicit

'──────────────────────────────────────────
' 主函數：支援 Range 或 逗號分隔字串，且第一列為欄位名稱
Function GetAccountCodeMapFlex( _
    CategoryParam As Variant, _
    Optional SubTypeParam As Variant, _
    Optional TypeParam As Variant) As Variant

    Dim dbPath As String
    dbPath = "C:\Users\David Hsu\Documents\Database1.accdb"  ' ← 修改為你的 Access 資料庫路徑

    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")

    Dim sql As String
    Dim whereClauses As Collection
    Set whereClauses = New Collection

    ' 1) 必填：Category
    Dim catClause As String
    catClause = BuildInClauseParam(CategoryParam, "Category")
    If catClause = "" Then
        GetAccountCodeMapFlex = CVErr(xlErrValue)
        Exit Function
    End If
    whereClauses.Add catClause

    ' 2) 可選：AssetMeasurementSubType
    Dim subClause As String
    subClause = BuildInClauseParam(SubTypeParam, "AssetMeasurementSubType")
    If subClause <> "" Then whereClauses.Add subClause

    ' 3) 可選：AssetMeasurementType
    Dim typeClause As String
    typeClause = BuildInClauseParam(TypeParam, "AssetMeasurementType")
    If typeClause <> "" Then whereClauses.Add typeClause

    ' 組出 SQL
    sql = "SELECT AccountCode, AccountTitle FROM AccountCodeMap WHERE "
    Dim part As Variant
    For Each part In whereClauses
        sql = sql & part & " AND "
    Next
    sql = Left(sql, Len(sql) - 5)  ' 去掉最後的 " AND "

    ' 執行查詢
    On Error GoTo ErrHandler
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"
    rs.Open sql, conn

    ' ──【修改重點】──
    ' 取 rs 資料並在第一列加入欄位名稱
    If rs.EOF Then
        ' 若無資料，僅回傳欄位名稱
        Dim noDataArr() As Variant
        ReDim noDataArr(0 To 0, 0 To rs.Fields.Count - 1)
        Dim fi As Integer
        For fi = 0 To rs.Fields.Count - 1
            noDataArr(0, fi) = rs.Fields(fi).Name
        Next
        GetAccountCodeMapFlex = noDataArr
    Else
        ' 1. 先用 GetRows 抓原始資料 (fields × rows)
        Dim rawArr As Variant
        rawArr = rs.GetRows()

        Dim nFields As Long, nRows As Long
        nFields = UBound(rawArr, 1) + 1
        nRows   = UBound(rawArr, 2) + 1

        ' 2. 重新建一個 (rows+1) × fields 的陣列
        Dim outArr() As Variant
        ReDim outArr(0 To nRows, 0 To nFields - 1)

        ' 3. 第一列填欄位名稱
        Dim f As Long, r As Long
        For f = 0 To nFields - 1
            outArr(0, f) = rs.Fields(f).Name
        Next

        ' 4. 之後每列填資料
        For r = 0 To nRows - 1
            For f = 0 To nFields - 1
                outArr(r + 1, f) = rawArr(f, r)
            Next
        Next

        GetAccountCodeMapFlex = outArr
    End If
    ' ────────────────────

    rs.Close: conn.Close
    Set rs = Nothing: Set conn = Nothing
    Exit Function

ErrHandler:
    GetAccountCodeMapFlex = "Error: " & Err.Description
    On Error GoTo 0
End Function

'──────────────────────────────────────────
' 輔助函數：根據參數類型（Range 或 String）建構 IN(...) 子句
Private Function BuildInClauseParam(paramIn As Variant, fieldName As String) As String
    Dim items As Collection
    Set items = New Collection
    Dim v As Variant

    ' 1) 如果是 Range
    If TypeName(paramIn) = "Range" Then
        For Each v In paramIn.Cells
            If Trim(v.Value & "") <> "" Then items.Add CStr(v.Value)
        Next

    ' 2) 如果是 String
    ElseIf VarType(paramIn) = vbString Then
        Dim arr As Variant
        arr = Split(paramIn, ",")
        For Each v In arr
            v = Trim(v)
            If v <> "" Then items.Add v
        Next
    End If

    ' 組出 IN (..)
    Dim sqlVals As String, elem As Variant
    For Each elem In items
        sqlVals = sqlVals & "'" & Replace(elem, "'", "''") & "',"
    Next
    If sqlVals <> "" Then
        sqlVals = Left(sqlVals, Len(sqlVals) - 1)
        BuildInClauseParam = fieldName & " IN (" & sqlVals & ")"
    Else
        BuildInClauseParam = ""
    End If
End Function
'──────────────────────────────────────────


=GetAccountCodeMapFlex("Cost","FVPL_GovBond_Foreign,FVOCI_CompanyBond_Foreign,AC_GovBond_Foreign","FVPL,FVOCI")
=GetAccountCodeMapFlex(D5,D7:D9,D11:D12)
=GetAccountCodeMapFlex(D5,"FVPL_GovBond_Foreign,AC_GovBond_Foreign")






' ===================================
answer3 (第一個Final Function)

**Modifying GroupFlag functionality**

The user wants to add an optional GroupFlag parameter, which will act as a filter, but also include it as a second column in the output, between AccountCode and AccountTitle. So, I'll need to modify the function to accept this new GroupFlagParam. It should be optional, like SubTypeParam and TypeParam. I'll also need to adjust the SELECT query to return GroupFlag in the desired order and ensure the filtering works as expected.


Thought for a few seconds


下面是完整修改後的 `GetAccountCodeMapFlex`，已加入可選的 `GroupFlagParam`，並把回傳欄位順序改成 `AccountCode, GroupFlag, AccountTitle`。
我在程式碼中以 `【★】` 標註所有新增／調整的地方。

```vb
Option Explicit

'──────────────────────────────────────────
' 主函數：支援 Range 或 逗號分隔字串，
' 第一列為欄位名稱，且新增可選的 GroupFlag 篩選
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

'──────────────────────────────────────────
' 輔助函數：根據參數類型（Range 或 String）建構 IN(...) 子句
Private Function BuildInClauseParam(paramIn As Variant, fieldName As String) As String
    Dim items As Collection
    Set items = New Collection
    Dim v As Variant

    If TypeName(paramIn) = "Range" Then
        For Each v In paramIn.Cells
            If Trim(v.Value & "") <> "" Then items.Add CStr(v.Value)
        Next
    ElseIf VarType(paramIn) = vbString Then
        Dim arr As Variant
        arr = Split(paramIn, ",")
        For Each v In arr
            v = Trim(v)
            If v <> "" Then items.Add v
        Next
    End If

    Dim sqlVals As String, elem As Variant
    For Each elem In items
        sqlVals = sqlVals & "'" & Replace(elem, "'", "''") & "',"
    Next
    If sqlVals <> "" Then
        sqlVals = Left(sqlVals, Len(sqlVals) - 1)
        BuildInClauseParam = fieldName & " IN (" & sqlVals & ")"
    Else
        BuildInClauseParam = ""
    End If
End Function
'──────────────────────────────────────────


' ### 範例呼叫

' 假設：

' * **A1\:A2** 放 Category（必填）
' * **B1\:B2** 放 GroupFlag（可選）
' * **C1\:C3** 放 AssetMeasurementSubType（可選）
' * **D1\:D2** 放 AssetMeasurementType（可選）


' =GetAccountCodeMapFlex(A1:A2, B1:B2, C1:C3, D1:D2)


' 混用字串方式：
' =GetAccountCodeMapFlex("外幣債,國內股","外幣債,國內債","FVPL_GovBond_Foreign,AC_CompanyBond_Foreign","FVPL")


' * 不想篩 GroupFlag，直接把第二個參數留空或省略即可：

'   ```excel
'   =GetAccountCodeMapFlex(A1:A2,,C1:C3,D1:D2)
'   ```
' * 第一列會自動顯示 `AccountCode | GroupFlag | AccountTitle`，後續列才是資料。



Thought for a couple of seconds


是的，因為我們把 `GroupFlagParam`、`SubTypeParam`、`TypeParam` 都定義成 `Optional`，所以在呼叫時可以只帶：

1. **只帶第一個 (Category)**

   ```excel
   =GetAccountCodeMapFlex(A1:A2)
   ```

   – 只會依 Category 篩選。

2. **只帶前兩個 (Category + GroupFlag)**

   ```excel
   =GetAccountCodeMapFlex(A1:A2, B1:B2)
   ```

   – 依 Category 與 GroupFlag 篩選，SubType/Type 全部都會被忽略。

3. **帶前三個 (Category + GroupFlag + SubType)**

   ```excel
   =GetAccountCodeMapFlex(A1:A2, B1:B2, C1:C3)
   ```

   – 依 Category、GroupFlag、AssetMeasurementSubType 篩選，Type 全部被忽略。

4. **帶前一與第三個，跳過第二個 (Category + SubType)**

   ```excel
   =GetAccountCodeMapFlex(A1:A2, , C1:C3)
   ```

   – 中間的 `GroupFlagParam` 留空，直接以 `,,` 跳過。

5. **全部四個都帶 (Category + GroupFlag + SubType + Type)**

   ```excel
   =GetAccountCodeMapFlex(A1:A2, B1:B2, C1:C3, D1:D2)
   ```

> **提示**：
>
> * 如果想用逗號分隔字串，也完全相同：
>
>   ```excel
>   =GetAccountCodeMapFlex("外幣債,國內股","A,B","FVPL_GovBond_Foreign","FVPL")
>   ```
> * 跳過任一可選參數，都只要在對應位置留空、用 `,` 分隔即可。

只要至少帶第一個參數，函數就能正常執行，其他都可以隨你需要帶或不帶。



' ===============================問題

延續上面的架構，我想要製作一個這樣的SQL
其中有幾個參數
1.其中的 OBU_AC4603 SubSql中的表格名稱
2.DataMonthString
3.GROUP BY 可單獨Group By 或是Grougp by為 AssetMeasurementType and Category 或 AssetMeasurementSubType and Category
如果是後者多個的話 MeasurementCategory 為 兩者的concat，否則就是Groupby的那個類別
4.OBU_AC4603.CurrencyType 也需要是可設定的參數
5.還有 where 的 AccountCodeMap.Category 參數

6.另外可選參數則包含上面範例的那幾個參數，請給我一個完整的版本

SELECT 
  AccountCodeMap.AssetMeasurementType & "_" & AccountCodeMap.Category AS MeasurementCategory,
  Sum(oa.NetBalance) AS SubtotalBalance
FROM
    AccountCodeMap
INNER JOIN
    (
        SELECT
            OBU_AC4603.AccountCode,
            OBU_AC4603.NetBalance
        FROM OBU_AC4603
        WHERE
            OBU_AC4603.CurrencyType = "USD"
            AND OBU_AC4603.DataMonthString = "2024/11"
    ) AS oa
ON
    AccountCodeMap.AccountCode = oa.AccountCode
WHERE
    AccountCodeMap.Category IN ("Cost" , "ValuationAdjust", "ImpairmentLoss", 'OSU')
GROUP BY
    AccountCodeMap.AssetMeasurementType, AccountCodeMap.Category;



' ============================ 
Answer(內容有缺失)

**Modifying VBA function**

The user asks to extend a VBA function to create an SQL query supporting various parameters. The new parameters include DataMonthString, grouping options, and CurrencyType for the OBU\_AC4603.CurrencyType. The function should still handle optional parameters like GroupFlag, AssetMeasurementSubType, and AssetMeasurementType. The final version of the function should return a 2D array with a header row containing MeasurementCategory and SubtotalBalance. The adjustments should include updating the function's name and supporting the new grouping logic.


Thought for 5 seconds


以下是完整的 `GetSubtotalBalance` VBA 程式碼，已整合你要的所有參數、可選篩選，以及動態 `GROUP BY` 邏輯。
我用 `【★】` 標註所有新增／調整重點，並在程式最上方列出參數說明。

```vb
Option Explicit

'──────────────────────────────────────────
' 函式：依參數跑出 MeasurementCategory & SubtotalBalance，
'   支援以下參數：
'   ── DataMonthStringParam (String, 必填)
'   ── GroupByMode (String: "Category" / "Type_Category" / "SubType_Category", 必填)
'   ── CurrencyTypeParam (String, 必填)
'   ── CategoryParam (Range 或 String, 必填，用於 AccountCodeMap.Category 篩選)
'   ── GroupFlagParam (Range 或 String, 選填，用於 AccountCodeMap.GroupFlag 篩選) 【★】
'   ── SubTypeParam (Range 或 String, 選填，用於 AccountCodeMap.AssetMeasurementSubType 篩選)
'   ── TypeParam (Range 或 String, 選填，用於 AccountCodeMap.AssetMeasurementType 篩選)
Function GetSubtotalBalance( _
    DataMonthStringParam As String, _                     '【★】必填：OBU_AC4603.DataMonthString
    GroupByMode As String, _                              '【★】必填："Category" / "Type_Category" / "SubType_Category"
    CurrencyTypeParam As String, _                        '【★】必填：OBU_AC4603.CurrencyType
    CategoryParam As Variant, _                           ' 必填：AccountCodeMap.Category
    Optional GroupFlagParam As Variant, _                  ' 選填：AccountCodeMap.GroupFlag 【★】
    Optional SubTypeParam As Variant, _                    ' 選填：AccountCodeMap.AssetMeasurementSubType
    Optional TypeParam As Variant                          ' 選填：AccountCodeMap.AssetMeasurementType
) As Variant

    Dim dbPath As String
    dbPath = "C:\Your\Path\To\Database.accdb"  ' ← 修改成你的 Access 檔案路徑

    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    Dim sql As String
    Dim whereACM As Collection: Set whereACM = New Collection
    Dim subSQL As String

    '── 1) 先做 AccountCodeMap 的 WHERE 篩選
    Dim catClause As String
    catClause = BuildInClauseParam(CategoryParam, "AccountCodeMap.Category")
    If catClause = "" Then
        GetSubtotalBalance = CVErr(xlErrValue)
        Exit Function
    End If
    whereACM.Add catClause

    '── 2) 可選：GroupFlagParam 篩選 【★】
    Dim grpClause As String
    grpClause = BuildInClauseParam(GroupFlagParam, "AccountCodeMap.GroupFlag")
    If grpClause <> "" Then whereACM.Add grpClause

    '── 3) 可選：SubTypeParam 篩選
    Dim subClause As String
    subClause = BuildInClauseParam(SubTypeParam, "AccountCodeMap.AssetMeasurementSubType")
    If subClause <> "" Then whereACM.Add subClause

    '── 4) 可選：TypeParam 篩選
    Dim typeClause As String
    typeClause = BuildInClauseParam(TypeParam, "AccountCodeMap.AssetMeasurementType")
    If typeClause <> "" Then whereACM.Add typeClause

    '── 5) 構造 OBU_AC4603 的子查詢 【★】
    subSQL = "SELECT OBU_AC4603.AccountCode, OBU_AC4603.NetBalance " & _
             "FROM OBU_AC4603 " & _
             "WHERE OBU_AC4603.CurrencyType = '" & Replace(CurrencyTypeParam, "'", "''") & "' " & _
             "  AND OBU_AC4603.DataMonthString = '" & Replace(DataMonthStringParam, "'", "''") & "'"

    '── 6) 根據 GroupByMode 準備 SELECT 和 GROUP BY 片段 【★】
    Dim selExpr As String, grpFields As String
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
            GetSubtotalBalance = CVErr(xlErrValue)
            Exit Function
    End Select

    '── 7) 完整 SQL
    sql = ""
    sql = sql & "SELECT " & selExpr & ", Sum(oa.NetBalance) AS SubtotalBalance " & _
                "FROM AccountCodeMap " & _
                "INNER JOIN (" & subSQL & ") AS oa " & _
                "ON AccountCodeMap.AccountCode = oa.AccountCode " & _
                "WHERE "

    Dim part As Variant
    For Each part In whereACM
        sql = sql & part & " AND "
    Next
    sql = Left(sql, Len(sql) - 5)  ' 去掉最後的 " AND "

    sql = sql & " GROUP BY " & grpFields

    '── 8) 執行並組成回傳陣列
    On Error GoTo ErrHandler
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"
    rs.Open sql, conn

    ' header + data
    If rs.EOF Then
        ' 只有欄位名稱
        Dim noArr() As Variant
        ReDim noArr(0 To 0, 0 To rs.Fields.Count - 1)
        Dim fi As Integer
        For fi = 0 To rs.Fields.Count - 1
            noArr(0, fi) = rs.Fields(fi).Name
        Next
        GetSubtotalBalance = noArr
    Else
        Dim rawArr As Variant
        rawArr = rs.GetRows()      ' fields × rows

        Dim nF As Long, nR As Long
        nF = UBound(rawArr, 1) + 1
        nR = UBound(rawArr, 2) + 1

        Dim outArr() As Variant
        ReDim outArr(0 To nR, 0 To nF - 1)

        ' 第一列：欄位名稱
        For fi = 0 To nF - 1
            outArr(0, fi) = rs.Fields(fi).Name
        Next

        ' 其餘：資料
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

'──────────────────────────────────────────
' 輔助：建構 IN(...) 子句，同支援 Range 或 逗號字串
Private Function BuildInClauseParam(paramIn As Variant, fieldName As String) As String
    Dim items As Collection: Set items = New Collection
    Dim v As Variant

    If TypeName(paramIn) = "Range" Then
        For Each v In paramIn.Cells
            If Trim(v.Value & "") <> "" Then items.Add CStr(v.Value)
        Next
    ElseIf VarType(paramIn) = vbString Then
        Dim arr As Variant
        arr = Split(paramIn, ",")
        For Each v In arr
            v = Trim(v)
            If v <> "" Then items.Add v
        Next
    End If

    Dim sqlVals As String, elem As Variant
    For Each elem In items
        sqlVals = sqlVals & "'" & Replace(elem, "'", "''") & "',"
    Next
    If sqlVals <> "" Then
        sqlVals = Left(sqlVals, Len(sqlVals) - 1)
        BuildInClauseParam = fieldName & " IN (" & sqlVals & ")"
    Else
        BuildInClauseParam = ""
    End If
End Function
'──────────────────────────────────────────
```

---

## 說明

1. **新增參數**

   * `DataMonthStringParam`、`CurrencyTypeParam`、`GroupByMode` 為必填。
   * `GroupFlagParam` 可像其他可選參數一樣，支援 Range 或 逗號字串（用來篩 `AccountCodeMap.GroupFlag`）。
2. **GroupByMode**

   * `"Category"`：只以 `Category` 分組。
   * `"Type_Category"`：以 `AssetMeasurementType` + `Category` 分組，並以 `Type_Category` 結果顯示。
   * `"SubType_Category"`：以 `AssetMeasurementSubType` + `Category` 分組。
3. **MeasurementCategory**

   * 根據 `GroupByMode` 動態產生：如果是複合分組，會 `& "_" &` 串接產生合併欄位。
4. **內層 SubQuery**

   * 一次拿到符合 `CurrencyTypeParam` 與 `DataMonthStringParam` 的 `AccountCode`/`NetBalance`。
5. **回傳格式**

   * 第一列為欄位名稱：`MeasurementCategory`、`SubtotalBalance`
   * 之後每列為每個分組結果。

只要把這支貼到 VBA 模組，設定好 Access 路徑，就可以在 Excel 直接呼叫：

```excel
=GetSubtotalBalance("2024/11", "Type_Category", "USD", "Cost,ValuationAdjust", "外幣債", "FVPL_GovBond_ForeignAC_CompanyBond_Foreign")
```

或用 Range：

```excel
=GetSubtotalBalance(A1, B1, C1, D1:D4, E1:E2, F1:F3, G1:G2)
```

第一列會顯示 `MeasurementCategory | SubtotalBalance`，後面是分組小計。




' ======= 
Answer (架構較不喜歡)

Thought for a few seconds


下面是整合後的完整 VBA 程式碼，已支援：

1. **SubSql 表格名稱** (`SubTableName`)
2. **DataMonthString**
3. **CurrencyType**
4. **外層篩選**：`CategoryParam`（必填）、`SubTypeParam`、`TypeParam`（可選，Range or comma-string）
5. **Group By** 選項 `GroupByOption`＝

   * `"Category"` → 單一以 Category 分組，`MeasurementCategory = Category`
   * `"SubType"` → 以 `AssetMeasurementSubType, Category` 分組，`MeasurementCategory = SubType & "_" & Category`
   * 其它（含 `"Type"` 或空白）→ 以 `AssetMeasurementType, Category` 分組，`MeasurementCategory = Type & "_" & Category`

所有修改與重點處都已在註解中標示 `【★】`。

```vb
Option Explicit

'──────────────────────────────────────────
' BuildBalanceSQL
' 產生符合所有參數的 SQL 字串
'
' 參數：
'   SubTableName    As String   '【★】內層子查詢表名
'   DataMonthString As String   '【★】
'   CurrencyType    As String   '【★】
'   CategoryParam   As Variant  ' 必填 (Range or "A,B,C")
'   SubTypeParam    As Variant  ' 可選 (Range or "X,Y,Z")
'   TypeParam       As Variant  ' 可選 (Range or "M,N,O")
'   GroupByOption   As String   ' “Category” / “SubType” / others→Type (預設)
'
Function BuildBalanceSQL( _
    SubTableName    As String, _
    DataMonthString As String, _
    CurrencyType    As String, _
    CategoryParam   As Variant, _
    Optional SubTypeParam  As Variant, _
    Optional TypeParam     As Variant, _
    Optional GroupByOption As String = "Type") As String

    Dim sb As String
    Dim outerWhere As Collection
    Set outerWhere = New Collection

    '── 1) 必填：Category 篩選
    Dim catClause As String
    catClause = BuildInClauseParam(CategoryParam, "AccountCodeMap.Category")
    If catClause = "" Then
        BuildBalanceSQL = "#ERROR: 必須提供 Category"
        Exit Function
    End If
    outerWhere.Add catClause

    '── 2) 可選：AssetMeasurementSubType
    Dim subClause As String
    subClause = BuildInClauseParam(SubTypeParam, "AccountCodeMap.AssetMeasurementSubType")
    If subClause <> "" Then outerWhere.Add subClause

    '── 3) 可選：AssetMeasurementType
    Dim typeClause As String
    typeClause = BuildInClauseParam(TypeParam, "AccountCodeMap.AssetMeasurementType")
    If typeClause <> "" Then outerWhere.Add typeClause

    '── 4) 決定 SELECT 與 GROUP BY 欄位
    Dim measureExpr As String, groupFields As String
    Select Case LCase(GroupByOption)
        Case "category"   '【★】只用 Category 分組
            measureExpr = "AccountCodeMap.Category"
            groupFields = "AccountCodeMap.Category"
        Case "subtype"    '【★】以 SubType + Category 分組
            measureExpr = "AccountCodeMap.AssetMeasurementSubType & ""_"" & AccountCodeMap.Category"
            groupFields = "AccountCodeMap.AssetMeasurementSubType, AccountCodeMap.Category"
        Case Else         '【★】預設／Type：以 Type + Category 分組
            measureExpr = "AccountCodeMap.AssetMeasurementType & ""_"" & AccountCodeMap.Category"
            groupFields = "AccountCodeMap.AssetMeasurementType, AccountCodeMap.Category"
    End Select

    '── 5) 組內層子查詢 (SubSql)
    Dim subSql As String
    subSql  = "SELECT " & SubTableName & ".AccountCode, " & SubTableName & ".NetBalance" & vbCrLf
    subSql &= "FROM " & SubTableName & vbCrLf
    subSql &= "WHERE " & SubTableName & ".CurrencyType = '" & Replace(CurrencyType, "'", "''") & "'" & vbCrLf
    subSql &= "  AND " & SubTableName & ".DataMonthString = '" & Replace(DataMonthString, "'", "''") & "'"

    '── 6) 組外層 WHERE 子句
    Dim w As Variant, whereSQL As String
    whereSQL = ""
    For Each w In outerWhere
        whereSQL = whereSQL & w & " AND "
    Next
    whereSQL = Left(whereSQL, Len(whereSQL) - 5)  ' 移除最後 " AND "

    '── 7) 最終組合 SQL
    sb  = "SELECT" & vbCrLf
    sb &= "  " & measureExpr & " AS MeasurementCategory," & vbCrLf
    sb &= "  Sum(oa.NetBalance) AS SubtotalBalance" & vbCrLf
    sb &= "FROM AccountCodeMap" & vbCrLf
    sb &= "INNER JOIN" & vbCrLf
    sb &= "  (" & vbCrLf & subSql & vbCrLf & "  ) AS oa" & vbCrLf
    sb &= "ON AccountCodeMap.AccountCode = oa.AccountCode" & vbCrLf
    sb &= "WHERE " & whereSQL & vbCrLf
    sb &= "GROUP BY " & groupFields & ";"

    BuildBalanceSQL = sb

End Function
'──────────────────────────────────────────

'──────────────────────────────────────────
' BuildInClauseParam
' 將 Range 或 comma-string 轉成 「Field IN ('A','B',…)」 的子句
'
Private Function BuildInClauseParam( _
    paramIn  As Variant, _
    fieldName As String) As String

    Dim items As Collection
    Set items = New Collection
    Dim v As Variant

    If Not IsMissing(paramIn) Then
        If TypeName(paramIn) = "Range" Then
            For Each v In paramIn.Cells
                If Trim(v.Value & "") <> "" Then items.Add CStr(v.Value)
            Next
        ElseIf VarType(paramIn) = vbString Then
            For Each v In Split(paramIn, ",")
                v = Trim(v)
                If v <> "" Then items.Add v
            Next
        End If
    End If

    Dim sqlVals As String
    For Each v In items
        sqlVals = sqlVals & "'" & Replace(v, "'", "''") & "',"
    Next
    If sqlVals <> "" Then
        sqlVals = Left(sqlVals, Len(sqlVals) - 1)  ' 去掉最後逗號
        BuildInClauseParam = fieldName & " IN (" & sqlVals & ")"
    Else
        BuildInClauseParam = ""
    End If

End Function
'──────────────────────────────────────────
```

---

### 呼叫範例

```vb
'— 以 Type + Category 分組（預設）
Debug.Print BuildBalanceSQL( _
  "OBU_AC4603", "2024/11", "USD", _
  Range("A1:A2"), , , "Type" _
)

'— 以 SubType + Category 分組
Debug.Print BuildBalanceSQL( _
  "OBU_AC4603", "2024/11", "USD", _
  "Cost,OSU,ImpairmentLoss", _
  "FVPL_GovBond_Foreign,FVOCI_FinancialBond_Foreign", , _
  "SubType" _
)

'— 只以 Category 分組
Debug.Print BuildBalanceSQL( _
  "OBU_AC4603", "2024/11", "USD", _
  "Cost,ValuationAdjust", , , _
  "Category" _
)
```

把 `Debug.Print` 改成你在程式中要執行 `rs.Open sql, conn` 的地方即可。





' ==============================

Answer(架構內容比較好) (第二個Final Function)
Thought for 6 seconds


以下是依照你提供的 `GetSubtotalBalance` 架構，並且整合之前所有需求（`Category`／`GroupFlag`／`SubType`／`Type` 篩選，三種 `GroupByMode`，內層子查詢 `DataMonthString`、`CurrencyType` 參數等），標示出所有關鍵修改處（以 `【★】` 註解）。

```vb
Option Explicit

'──────────────────────────────────────────
' GetSubtotalBalance
' 依據參數動態產生並執行 SQL，回傳結果（含欄位名稱）
'
' 參數：
'   DataMonthStringParam As String      '【★】OBU_AC4603.DataMonthString
'   GroupByMode          As String      '【★】"Category" / "Type_Category" / "SubType_Category"
'   CurrencyTypeParam    As String      '【★】OBU_AC4603.CurrencyType
'   CategoryParam        As Variant     ' 必填：AccountCodeMap.Category
'   Optional GroupFlagParam  As Variant ' 選填：AccountCodeMap.GroupFlag   【★】
'   Optional SubTypeParam     As Variant ' 選填：AccountCodeMap.AssetMeasurementSubType
'   Optional TypeParam        As Variant ' 選填：AccountCodeMap.AssetMeasurementType
'
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

'──────────────────────────────────────────
' BuildInClauseParam
' Range or comma-string → 「FieldName IN ('A','B',…)」
Private Function BuildInClauseParam( _
    paramIn  As Variant, _
    fieldName As String _
) As String
    Dim items As Collection: Set items = New Collection
    Dim v As Variant

    If Not IsMissing(paramIn) Then
        If TypeName(paramIn) = "Range" Then
            For Each v In paramIn.Cells
                If Trim(v.Value & "") <> "" Then items.Add CStr(v.Value)
            Next
        ElseIf VarType(paramIn) = vbString Then
            For Each v In Split(paramIn, ",")
                v = Trim(v)
                If v <> "" Then items.Add v
            Next
        End If
    End If

    Dim sqlVals As String
    For Each v In items
        sqlVals = sqlVals & "'" & Replace(v, "'", "''") & "',"
    Next
    If sqlVals <> "" Then
        sqlVals = Left(sqlVals, Len(sqlVals) - 1)
        BuildInClauseParam = fieldName & " IN (" & sqlVals & ")"
    Else
        BuildInClauseParam = ""
    End If
End Function
'──────────────────────────────────────────
```

**主要標示**

* `GroupFlagParam` 的篩選條件加在 `whereACM` 集合中【★】
* 三種 `GroupByMode`（`Category`／`Type_Category`／`SubType_Category`）分支決定 `SELECT` 和 `GROUP BY` 片段【★】
* 內層子查詢固定使用 `DataMonthStringParam`、`CurrencyTypeParam`【★】
* 最終回傳二維陣列，第一列為欄位名稱，其後為資料。

貼到你的 VBA 模組後，即可在 Excel 中以：

```vb
=GetSubtotalBalance("2024/11", "Type_Category", "USD", A1:A2, B1:B2, C1:C3, D1:D2)
```

或以逗號串混用方式呼叫。




https://chatgpt.com/share/6818d233-7f40-800b-8191-ee0057c39353
