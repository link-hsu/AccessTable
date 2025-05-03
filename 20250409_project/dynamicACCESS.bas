' =========================
我有一個這樣的資料表 AccountCodeMap
DataID	GroupFlag	AccountCode	AccountTitle	AssetMeasurementSubType	AssetMeasurementType	Category	DeclarationReport	Note	Category
	1	120050105	強制FVPL金融資產-公債-中央政府(外國)	FVPL_GovBond_Foreign	FVPL	Cost	AI602	公債_中央政府	Cost
	1	120050107	強制FVPL金融資產-公債-地方政府(外國)	FVPL_GovBond_Foreign	FVPL	Cost	AI602	公債_地方政府	ValuationAdjust
	1	120050125	強制FVPL金融資產-普通公司債(公營)(外國)	FVPL_CompanyBond_Foreign	FVPL	Cost		公司債_公營	ImpairmentLoss
	1	120050127	強制FVPL金融資產-普通公司債(民營)(外國)	FVPL_CompanyBond_Foreign	FVPL	Cost		公司債_民營	InterestRevenue
	1	120050147	強制FVPL金融資產-金融債(外國)	FVPL_FinancialBond_Foreign	FVPL	Cost			ValuationProfit
	1	121110105	FVOCI債務工具-公債-中央政府(外國)	FVOCI_GovBond_Foreign	FVOCI	Cost			ValuationLoss
	1	121110125	FVOCI債務工具-普通公司債(公營)(外國)	FVOCI_CompanyBond_Foreign	FVOCI	Cost		公司債_公營	GainOnDisposal
	1	121110127	FVOCI債務工具-普通公司債(民營)(外國)	FVOCI_CompanyBond_Foreign	FVOCI	Cost		公司債_民營	LossOnDisposal
	1	121110147	FVOCI債務工具-金融債券-海外	FVOCI_FinancialBond_Foreign	FVOCI	Cost			
	1	122010105	AC債務工具投資-公債-中央政府(外國)	AC_GovBond_Foreign	AC	Cost			



我同時已經寫好了幾個microsoft access sql語法
PARAMETERS DataMonthParam TEXT;
SELECT
    AccountCodeMap.AssetMeasurementSubType AS MeasureType,
    SUM(oa.NetBalance) AS SubNetBalance
FROM
    AccountCodeMap
INNER JOIN
    (
        SELECT
            OBU_AC5601.AccountCode,
            OBU_AC5601.NetBalance
        FROM
            OBU_AC5601
        WHERE
            OBU_AC5601.CurrencyType = 'USD'
            AND OBU_AC5601.DataMonthString = [DataMonthParam]
    ) AS oa
ON
    AccountCodeMap.AccountCode = OBU_AC5411B.AccountCode
WHERE
    AccountCodeMap.Category IN ('Cost', 'ValuationAdjust')
GROUP BY
    AccountCodeMap.AssetMeasurementSubType



PARAMETERS DataMonthParam TEXT;
SELECT
    AccountCodeMap.AssetMeasurementType,
    SUM(OBU_AC4620B.NetBalance) As SubtotalBalance
FROM AccountCodeMap
INNER JOIN
    OBU_AC4620B
ON
    AccountCodeMap.AccountCode = OBU_AC4620B.AccountCode
WHERE
    AccountCodeMap.Category IN ('Cost' , 'ValuationAdjust')
    AND OBU_AC4620B.DataMonthString = [DataMonthParam]
    AND OBU_AC4620B.CurrencyType = "CNY"
GROUP BY
    AccountCodeMap.AssetMeasurementType;



PARAMETERS DataMonthParam TEXT;
SELECT
    OBU_AC4603.DataID,
    OBU_AC4603.DataMonthString,
    AccountCodeMap.AccountCode, 
    AccountCodeMap.AccountTitle, 
    OBU_AC4603.CurrencyType,
    OBU_AC4603.NetBalance
FROM 
    AccountCodeMap
INNER JOIN 
    OBU_AC4603
ON
    AccountCodeMap.AccountCode = OBU_AC4603.AccountCode
WHERE
    AccountCodeMap.GroupFlag IN (1, 2)
    AND AccountCodeMap.Category IN ("Cost" , "ValuationAdjust", "ImpairmentLoss", "otherFinancialAssets")
    AND OBU_AC4603.CurrencyType = "USD" 
    AND OBU_AC4603.DataMonthString = [DataMonthParam];

PARAMETERS DataMonthParam TEXT;
SELECT
    oa.DataID,
    oa.DataMonthString,
    AccountCodeMap.AccountCode, 
    AccountCodeMap.AccountTitle, 
    oa.MonthAmount
FROM
    AccountCodeMap
INNER JOIN
    (
        SELECT
            OBU_AC5411B.DataID,
            OBU_AC5411B.AccountCode,
            OBU_AC5411B.DataMonthString,
            OBU_AC5411B.MonthAmount
        FROM
            OBU_AC5411B
        WHERE
            OBU_AC5411B.DataMonthString = '2024/11'
    ) AS oa
ON
    AccountCodeMap.AccountCode = oa.AccountCode
WHERE
    AccountCodeMap.Category IN ("InterestRevenue" , "GainOnDisposal", "LossOnDisposal", "Interest", "ValuationProfit", "ValuationLoss");




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


dataArr = GetAccessDataAsArray(gDBPath, queryTable, gDataMonthString)


If Err.Number <> 0 Or LBound(dataArr) > UBound(dataArr) Then
    MsgBox reportTitle & ": " & queryTable & " 資料表無資料"
    WriteLog reportTitle & ": " & queryTable & " 資料表無資料"
Else
    For j = 0 To UBound(dataArr, 2)
        For i = 0 To UBound(dataArr, 1)
            xlsht.Cells(i + 1, j + 1).Value = dataArr(i, j)
        Next i
    Next j
End If

以上是相關的素材，GetAccessDataAsArray會從access資料庫取得資料，並回傳一個2d的表，在透過上述貼入excel中，

我想寫一個函數可以傳入像上述SQL範例那樣，透過AccountCodeMap去串聯其他資料表中的數據，其中AccountCodeMap是一個我已經整理好分類的資料表，透過join其他表可以篩選我想要的那些Code，
所以我可以選擇的參數應該包含要貼入excel分頁名稱、要貼入的row and column，AccountCode要篩選的項目，DataMonthString，其他請你在參考上面我給你的資料去設計要那些參數，
我想要可以直接call這個sub或function，就可以在excel中某個分頁的欄位從access撈出想要的數據





' ====================


好的，以下是你需要的整套 VBA 程式碼與使用範例，**不使用分頁**，全部整理在一段中。這樣你可以直接複製貼上到 VBA 編輯器中使用：

---

### 🔧 主功能：自動組 SQL，撈 Access 資料，貼到 Excel

```vb
' 執行 SQL 並回傳含欄位的 2D 陣列
Function ExecuteSQLToArray(DBPath As String, SQL As String) As Variant
    Dim conn As Object, rs As Object
    Dim dataArr As Variant, headerArr() As String
    Dim rowCount As Long, colCount As Long
    Dim i As Long, j As Long

    On Error GoTo ErrHandler

    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath

    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3
    rs.Open SQL, conn

    If rs.EOF Then
        ExecuteSQLToArray = Array()
        GoTo Cleanup
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

    ' 欄位名稱
    For i = 0 To colCount - 1
        resultArr(0, i) = headerArr(i)
    Next i

    ' 資料內容
    For i = 0 To colCount - 1
        For j = 1 To rowCount
            resultArr(j, i) = dataArr(i, j - 1)
        Next j
    Next i

    ExecuteSQLToArray = resultArr

Cleanup:
    If Not rs Is Nothing Then rs.Close
    If Not conn Is Nothing Then conn.Close
    Set rs = Nothing
    Set conn = Nothing
    Exit Function

ErrHandler:
    MsgBox "錯誤：" & Err.Description, vbCritical
    ExecuteSQLToArray = Array()
    Resume Cleanup
End Function

' 主功能：從 AccountCodeMap 對應其他表，自動查詢並貼資料到指定工作表位置
Sub GetMappedDataFromAccess( _
    DBPath As String, _
    DataMonth As String, _
    CurrencyType As String, _
    CategoryList As Variant, _
    SourceTable As String, _
    ValueField As String, _
    GroupByField As String, _
    TargetSheetName As String, _
    StartRow As Long, _
    StartCol As Long)

    Dim ws As Worksheet
    Dim catStr As String, sql As String
    Dim dataArr As Variant
    Dim i As Long, j As Long

    ' 合成 WHERE 的 Category IN ('a','b')
    catStr = "('" & Join(CategoryList, "','") & "')"

    ' 動態 SQL 組合
    If GroupByField <> "" Then
        sql = "SELECT AccountCodeMap." & GroupByField & " AS GroupKey, " & _
              "SUM(" & SourceTable & "." & ValueField & ") AS TotalValue " & _
              "FROM AccountCodeMap INNER JOIN " & SourceTable & " " & _
              "ON AccountCodeMap.AccountCode = " & SourceTable & ".AccountCode " & _
              "WHERE AccountCodeMap.Category IN " & catStr & " " & _
              "AND " & SourceTable & ".CurrencyType = '" & CurrencyType & "' " & _
              "AND " & SourceTable & ".DataMonthString = '" & DataMonth & "' " & _
              "GROUP BY AccountCodeMap." & GroupByField
    Else
        sql = "SELECT " & SourceTable & ".DataID, " & SourceTable & ".DataMonthString, " & _
              "AccountCodeMap.AccountCode, AccountCodeMap.AccountTitle, " & _
              SourceTable & ".CurrencyType, " & SourceTable & "." & ValueField & " " & _
              "FROM AccountCodeMap INNER JOIN " & SourceTable & " " & _
              "ON AccountCodeMap.AccountCode = " & SourceTable & ".AccountCode " & _
              "WHERE AccountCodeMap.Category IN " & catStr & " " & _
              "AND " & SourceTable & ".CurrencyType = '" & CurrencyType & "' " & _
              "AND " & SourceTable & ".DataMonthString = '" & DataMonth & "'"
    End If

    ' 撈資料
    dataArr = ExecuteSQLToArray(DBPath, sql)

    If LBound(dataArr) > UBound(dataArr) Then
        MsgBox "查詢結果為空！", vbInformation
        Exit Sub
    End If

    ' 輸出到 Excel
    Set ws = ThisWorkbook.Sheets(TargetSheetName)
    For i = 0 To UBound(dataArr, 1)
        For j = 0 To UBound(dataArr, 2)
            ws.Cells(StartRow + i, StartCol + j).Value = dataArr(i, j)
        Next j
    Next i

End Sub
```

---

### 🧪 使用範例：只要改參數即可應用不同情境

```vb
Sub 測試_彙總_FVPL_USD()
    Dim cats As Variant
    cats = Array("Cost", "ValuationAdjust")
    Call GetMappedDataFromAccess( _
        "C:\你的路徑\資料庫.accdb", _
        "2024/11", _
        "USD", _
        cats, _
        "OBU_AC5601", _
        "NetBalance", _
        "AssetMeasurementSubType", _
        "Sheet1", _
        1, 1)
End Sub

Sub 測試_明細_利息收入()
    Dim cats As Variant
    cats = Array("InterestRevenue", "GainOnDisposal", "LossOnDisposal", "ValuationProfit", "ValuationLoss")
    Call GetMappedDataFromAccess( _
        "C:\你的路徑\資料庫.accdb", _
        "2024/11", _
        "USD", _
        cats, _
        "OBU_AC5411B", _
        "MonthAmount", _
        "", _
        "Sheet1", _
        20, 1)
End Sub
```

---

這樣一來，你就可以針對不同資料表／類型，快速指定參數呼叫，不需要再為每個查詢重複寫 SQL 或 Excel 貼資料的邏輯。

需要我幫你改寫一個實際的查詢範例嗎？




下面是一個完整的、可直接貼進 VBA 編輯器、針對「2024/11 月份，USD、Cost / ValuationAdjust 類別，依 AssetMeasurementSubType 彙總 OBU_AC5601.NetBalance」的實際查詢範例。

```vb
'——— 實際要執行的 SQL（把下面這段貼到 Immediate Window 用 Debug.Print 就能看到） ———
'SELECT 
'    AccountCodeMap.AssetMeasurementSubType AS GroupKey, 
'    SUM(OBU_AC5601.NetBalance) AS TotalValue
'FROM 
'    AccountCodeMap 
'INNER JOIN 
'    OBU_AC5601 
'ON 
'    AccountCodeMap.AccountCode = OBU_AC5601.AccountCode
'WHERE 
'    AccountCodeMap.Category IN ('Cost','ValuationAdjust')
'    AND OBU_AC5601.CurrencyType = 'USD'
'    AND OBU_AC5601.DataMonthString = '2024/11'
'GROUP BY 
'    AccountCodeMap.AssetMeasurementSubType;

'——— 對應的 VBA Sub 呼叫 ——  
Sub Run_USD_CostValuation_BySubType()
    Dim DBPath       As String
    Dim DataMonth    As String
    Dim cats         As Variant
    
    ' 1. Access 檔案路徑
    DBPath    = "C:\Your\Path\Finance.accdb"     
    ' 2. 查詢年月
    DataMonth = "2024/11"                       
    ' 3. Category 篩選清單
    cats = Array("Cost", "ValuationAdjust")    
    
    ' 4. 呼叫通用子程式：
    '    (1) SourceTable = "OBU_AC5601"
    '    (2) ValueField  = "NetBalance"
    '    (3) GroupByField= "AssetMeasurementSubType"
    '    (4) 結果貼到 Sheet1，從 A1 開始
    Call GetMappedDataFromAccess( _
        DBPath, _                       ' Access 檔案
        DataMonth, _                    ' DataMonthString
        "USD", _                        ' CurrencyType
        cats, _                         ' CategoryList
        "OBU_AC5601", _                 ' SourceTable
        "NetBalance", _                 ' Value 欄位
        "AssetMeasurementSubType", _    ' Group by Map 欄位
        "Sheet1", _                     ' 貼到的工作表
        1, 1)                           ' 起始 Row, Col
End Sub
```

— 呼叫後，Sheet1 的 A1 會開始貼出：

| GroupKey             | TotalValue   |
|----------------------|--------------|
| FVPL_GovBond_Foreign | 123,456,789  |
| …                    | …            |

（以上金額為範例）




https://chatgpt.com/share/68138b9d-51e4-8010-be2c-2001dbe48f81



https://chatgpt.com/share/681598d8-7184-8010-b50a-0d6d55ca7dc7
