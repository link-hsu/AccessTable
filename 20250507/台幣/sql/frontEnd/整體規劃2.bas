Ans:

好的，我们先从第 1 点着手，分批深入讨论。

---

## 1. 用 Access 的 PositionMap 取代 `rptFieldInfo`

### 核心思路

1. **新增一个通用的函数** `GetPositionMapData`

   * 负责从 Access 的 `PositionMap` 表中，根据 `ReportName` 读取所有对应的资料：

     * `TargetSheetName`
     * `SourceNameTag`
     * `TargetCellAddress`
   * 回传一个二维 `Variant` 陣列或 `Collection`，格式与 `rptFieldInfo(reportName)` 一致，方便后续共用。

2. **修改 `clsReport.Init`**

   * 移除原先 hard-code 的 `rptFieldInfo` 定义。
   * 在初始化之初调用 `GetPositionMapData reportName` 得到所有映射，然后对可用记录依序：

     ```vb
     For Each mapping In mappings
         AddWorksheetFields mapping.TargetSheetName, Array( _
             Array(mapping.SourceNameTag, mapping.TargetCellAddress, Null) _
         )
     Next
     ```
   * 后续写入 Access DB、加上动态字段等流程都无需更动。

### 示例代码

```vb
' === 模块：modDataAccess ===
' 通用：从 PositionMap 取出指定报表的所有映射
Public Function GetPositionMapData( _
        ByVal DBPath As String, _
        ByVal reportName As String _
    ) As Variant
    Dim conn As Object, rs As Object
    Dim sql As String
    Dim results() As Variant
    Dim i As Long
    
    ' 1. 建立连接
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath
    
    ' 2. SQL: 先 INNER JOIN Report 找到 ReportID，再取 PositionMap
    sql = "SELECT pm.TargetSheetName, pm.SourceNameTag, pm.TargetCellAddress " & _
          "FROM PositionMap AS pm " & _
          "INNER JOIN Report AS r ON pm.ReportID = r.ReportID " & _
          "WHERE r.ReportName = '" & reportName & "' " & _
          "ORDER BY pm.DataId;"
    
    Set rs = conn.Execute(sql)
    If rs.EOF Then
        GetPositionMapData = Array()  ' 没有记录
    Else
        ' 3. 把结果装进二维数组：每行一笔 mapping
        rs.MoveLast: rs.MoveFirst
        ReDim results(0 To rs.RecordCount - 1, 0 To 2)
        i = 0
        Do Until rs.EOF
            results(i, 0) = rs.Fields("TargetSheetName").Value
            results(i, 1) = rs.Fields("SourceNameTag").Value
            results(i, 2) = rs.Fields("TargetCellAddress").Value
            i = i + 1
            rs.MoveNext
        Loop
        GetPositionMapData = results
    End If
    
    rs.Close: conn.Close
    Set rs = Nothing: Set conn = Nothing
End Function
```

```vb
' === 类别：clsReport 中的 Init 调整 ===
Public Sub Init(ByVal reportName As String, _
                ByVal dataMonthStringROC As String, _
                ByVal dataMonthStringROC_NUM As String, _
                ByVal dataMonthStringROC_F1F2 As String)
    
    Dim mappings As Variant
    Dim i As Long
    Dim sheetName As String, nameTag As String, addr As String
    
    clsReportName = reportName
    Set clsWorksheets = CreateObject("Scripting.Dictionary")
    
    ' —— 1. 从 Access 读出该报表所有映射定义 —— 
    mappings = GetPositionMapData(gDBPath, reportName)
    
    If IsArray(mappings) And UBound(mappings) >= 0 Then
        For i = 0 To UBound(mappings, 1)
            sheetName = mappings(i, 0)
            nameTag   = mappings(i, 1)
            addr      = mappings(i, 2)
            
            ' 跳过空定义
            If Len(Trim(nameTag)) > 0 And Len(Trim(addr)) > 0 Then
                ' 每笔只包含单一 cell，可自行扩展支持多 cell
                AddWorksheetFields sheetName, Array( _
                    Array(nameTag, addr, Null) _
                )
            End If
        Next i
    Else
        WriteLog "Init 未在 PositionMap 找到报表：" & reportName
    End If
    
    ' —— 2. 保持原有的动态字段逻辑不变 —— 
    Select Case reportName
        Case "TABLE10"
            AddDynamicField reportName, "TABLE10_申報時間", "D2", dataMonthStringROC
        ' … 其余 Case 不变
    End Select
End Sub
```

### 说明

* `GetPositionMapData` 足够通用，后续若要改取 `QueryMap`、或支持多目标 cell，只要改此函数即能复用。
* `Init` 中只把 PositionMap 的资料带入 `AddWorksheetFields`，其它流程如写 Access、`ValidateFields`、`ApplyToWorkbook` 均可保留。

---

以上是第 1 部分的解决方案，若理解 OK 或有细节要讨论，我们再进到第 2 部分！请告诉我继续哪个问题。



Que:

請你看一下，我原本是這樣用的，

你的用法有和我一樣嗎？

            AddWorksheetFields "CNY1", Array( _
                Array("CNY1_申報時間", "C2", dataMonthStringROC), _
                Array("CNY1_其他金融資產_淨額", "G98", Null), _
                Array("CNY1_其他", "G100", Null), _
                Array("CNY1_資產總計", "G116", Null), _
                Array("CNY1_其他金融負債", "G170", Null), _
                Array("CNY1_其他什項金融負債", "G172", Null), _
                Array("CNY1_負債總計", "G184", Null) )


請你自行檢討以後重新修改完整版本，並告訴我修改哪邊，清楚標示出來

Ans:

非常好，感謝你提供原本的用法。你使用 `AddWorksheetFields` 的方式是一次傳入：

```vb
AddWorksheetFields "工作表名稱", Array( _
    Array("NameTag1", "儲存格位置1", 預設值1), _
    Array("NameTag2", "儲存格位置2", 預設值2), _
    ...
)
```

而我上面寫的版本是多次呼叫 `AddWorksheetFields`，每次只有一組欄位：

```vb
AddWorksheetFields sheetName, Array( _
    Array(nameTag, addr, Null) _
)
```

### ✅ 相容性檢查與調整

其實你的一次傳多組是完全正確且更有效率的方式，我會根據這個做法修改，讓我們的程式邏輯可以**從 PositionMap 抓多筆資料**，然後**一次傳入 AddWorksheetFields**，這樣才和你現有的架構一致。

---

## 🔧 修正後的完整版本

### 📍**修改 1：`GetPositionMapData` 函數**（**保持不變**）

這部分我們上面提供的程式可以直接使用，回傳一個 `Variant` 的 2D 陣列，格式為：

```vb
results(i, 0) = TargetSheetName  
results(i, 1) = SourceNameTag  
results(i, 2) = TargetCellAddress  
```

這部分不用修改。

---

### 📍**修改 2：clsReport.Init 函數中「載入 PositionMap 設定」的邏輯**

改成將**同一工作表的所有欄位先蒐集起來，最後再用一次 `AddWorksheetFields` 傳入。**

#### ✅ 修正後程式：

```vb
Public Sub Init(ByVal reportName As String, _
                ByVal dataMonthStringROC As String, _
                ByVal dataMonthStringROC_NUM As String, _
                ByVal dataMonthStringROC_F1F2 As String)
    
    Dim mappings As Variant
    Dim i As Long
    Dim sheetName As String, nameTag As String, addr As String
    
    clsReportName = reportName
    Set clsWorksheets = CreateObject("Scripting.Dictionary")

    ' === 從 Access 讀取 PositionMap 資料 ===
    mappings = GetPositionMapData(gDBPath, reportName)

    Dim wsFields As Object
    Set wsFields = CreateObject("Scripting.Dictionary")

    If IsArray(mappings) And UBound(mappings) >= 0 Then
        For i = 0 To UBound(mappings, 1)
            sheetName = mappings(i, 0)
            nameTag = mappings(i, 1)
            addr = mappings(i, 2)

            If Len(Trim(nameTag)) > 0 And Len(Trim(addr)) > 0 Then
                If Not wsFields.exists(sheetName) Then
                    wsFields.Add sheetName, Array()
                End If

                ' 先取出目前該 sheet 的陣列
                Dim tmpList As Variant
                tmpList = wsFields(sheetName)

                ' 將新欄位加入陣列
                ReDim Preserve tmpList(0 To UBound(tmpList) + 1)
                tmpList(UBound(tmpList)) = Array(nameTag, addr, Null)

                ' 寫回 dictionary
                wsFields(sheetName) = tmpList
            End If
        Next i

        ' 最後，依每個工作表將欄位群一次性加到報表設定中
        Dim key As Variant
        For Each key In wsFields.Keys
            AddWorksheetFields key, wsFields(key)
        Next key

    Else
        WriteLog "Init 未在 PositionMap 找到報表：" & reportName
    End If

    ' === 原有的動態欄位設定邏輯保留 ===
    Select Case reportName
        Case "TABLE10"
            AddDynamicField reportName, "TABLE10_申報時間", "D2", dataMonthStringROC
        Case "TABLE11"
            AddDynamicField reportName, "TABLE11_申報時間", "D2", dataMonthStringROC
        ' ... 其他報表
    End Select
End Sub
```

---

## ✅ 這次修改項目總整理

| 區塊                   | 修改內容                                                  | 位置                           |
| -------------------- | ----------------------------------------------------- | ---------------------------- |
| `GetPositionMapData` | 保留不動，回傳 `2D Array` 結構                                 | `modDataAccess`              |
| `clsReport.Init`     | ✅ 改成先蒐集所有欄位後，一次性呼叫 `AddWorksheetFields`<br>✅ 保留動態欄位設定 | `clsReport.cls` 中的 `Init` 函數 |

---

接下來如果你確認這段 OK，我們可以處理第 2 部分：「Process\_FM11 中如何從 QueryMap 撈取查詢名稱與 Excel 放置欄位」。

請告訴我是否繼續 👍


Ans:

下面提供两段关键代码，一个是通用的 `GetQueryMapData` 函数，用来从 `QueryMap` 读出报表对应的所有查询配置；另一个是改写后的 `Process_FM11`，完全取代硬编码的 `queryTable_1`、`queryTable_2`，改用从 `QueryMap` 摸到的每一笔配置来批次下载和贴到 Excel。

---

### 1. 新增：`GetQueryMapData`（放在 `modDataAccess` 模块中，不要动你的 `GetAccessDataAsArray`）

```vb
'=== 从 QueryMap 取出指定报表的所有查询配置 ===
Public Function GetQueryMapData( _
        ByVal DBPath As String, _
        ByVal reportName As String _
    ) As Variant
    Dim conn As Object, rs As Object
    Dim sql As String, results() As Variant
    Dim i As Long
    
    ' 1. 建立 ADO 连接
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath
    
    ' 2. 用 INNER JOIN 取 ReportID，再拿 QueryMap
    sql = "SELECT qm.QueryTableName, " & _
          "qm.ImportColName, qm.ImportColNumber " & _
          "FROM QueryMap AS qm " & _
          "INNER JOIN Report AS r ON qm.ReportID = r.ReportID " & _
          "WHERE r.ReportName='" & reportName & "' " & _
          "ORDER BY qm.DataId;"
    
    Set rs = conn.Execute(sql)
    If rs.EOF Then
        GetQueryMapData = Array()    ' 没有任何配置
    Else
        ' 把结果装二维 Array：(0)=QueryTableName, (1)=ImportColName, (2)=ImportColNumber
        rs.MoveLast: rs.MoveFirst
        ReDim results(0 To rs.RecordCount - 1, 0 To 2)
        i = 0
        Do Until rs.EOF
            results(i, 0) = rs.Fields("QueryTableName").Value
            results(i, 1) = rs.Fields("ImportColName").Value
            results(i, 2) = rs.Fields("ImportColNumber").Value
            i = i + 1
            rs.MoveNext
        Loop
        GetQueryMapData = results
    End If
    
    rs.Close: conn.Close
    Set rs = Nothing: Set conn = Nothing
End Function
```

---

### 2. 修改：`Process_FM11`——改用 `GetQueryMapData` 动态跑所有查询

```vb
Public Sub Process_FM11()
    Dim rpt As clsReport
    Set rpt = gReports("FM11")
    
    Dim reportTitle As String
    reportTitle = rpt.ReportName
    
    Dim xlsht As Worksheet
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    ' === 先清空整张表（或局部你可改） ===
    xlsht.Cells.ClearContents
    
    ' === 动态取 QueryMap 配置 ===
    Dim qmap As Variant
    qmap = GetQueryMapData(gDBPath, reportTitle)
    
    If Not IsArray(qmap) Or UBound(qmap) < 0 Then
        MsgBox "未在 QueryMap 找到 FM11 的任何查询配置！", vbExclamation
        Exit Sub
    End If
    
    Dim iMap As Long, tblName As String
    Dim startColLetter As String, startCol As Long, numCols As Long
    Dim dataArr As Variant
    Dim r As Long, c As Long
    
    For iMap = 0 To UBound(qmap, 1)
        tblName        = qmap(iMap, 0)
        startColLetter = qmap(iMap, 1)
        numCols        = CLng(qmap(iMap, 2))
        
        ' 1) 取资料
        dataArr = GetAccessDataAsArray(gDBPath, tblName, gDataMonthString)
        If Not IsArray(dataArr) Then
            MsgBox "查询无资料: " & tblName, vbExclamation
            WriteLog "FM11 无资料: " & tblName
            GoTo NextMap
        End If
        
        ' 2) 找出起始列 (第 1 列用 row 1)
        startCol = xlsht.Range(startColLetter & "1").Column
        
        ' 3) 将 dataArr 贴进 Excel
        ' dataArr 的维度： (0 To rowCount, 0 To colCount-1)
        For r = 0 To UBound(dataArr, 0)
            For c = 0 To numCols - 1
                xlsht.Cells(r + 1, startCol + c).Value = dataArr(r, c)
            Next c
        Next r
        
NextMap:
    Next iMap
    
    ' === 后续你原来的“Unique Setting”不动 ===
    ' …(略)…
    
    ' 最后 Validation + 更新 DB
    If rpt.ValidateFields() Then
        Dim key As Variant
        Dim allValues As Object, allPositions As Object
        Set allValues    = rpt.GetAllFieldValues()
        Set allPositions = rpt.GetAllFieldPositions()
        For Each key In allValues.Keys
            UpdateRecord gDBPath, gDataMonthString, rpt.ReportName, _
                         key, allPositions(key), allValues(key)
        Next key
    End If
    
    ' 分页标色
    xlsht.Tab.ColorIndex = 6
End Sub
```

---

#### ✏️ 修改要点

1. **移除**：

   ```vb
   queryTable_1 = "FM11_OBU_AC5411B"
   queryTable_2 = "FM11_OBU_AC5411B_Subtotal"
   dataArr_1  = GetAccessDataAsArray(...)
   dataArr_2  = GetAccessDataAsArray(...)
   ```
2. **加入**：

   ```vb
   qmap = GetQueryMapData(gDBPath, reportTitle)
   For iMap = 0 To UBound(qmap,1)
       tblName = qmap(iMap,0)
       startColLetter = qmap(iMap,1)
       numCols = CLng(qmap(iMap,2))
       dataArr = GetAccessDataAsArray(...)
       ' …贴值逻辑…
   Next iMap
   ```
3. 使用 `xlsht.Range(startColLetter & "1").Column` 自动算出起始列编号，便于把同一 Query 的多栏位连续贴上。

这样一来，**新增或调整 `QueryMap` 中的几笔记录**就能立刻生效，无需改程式。下一步，我们可以再讨论你的第 4 部分：**如何在 Excel 端做下拉＋多层级的动态维护界面**，并把修改写回 Access + 更新名称定义。请告诉我是否现在继续！


Ans:

下面示範一種以「UserForm + Excel 工作表驗證」為主的做法，讓您可以在 Excel 端：

1. **選擇報表** → 2. **選擇欄位 (NameTag)** → 3. **查看／編輯其對應的設定** → 4. **按一下「儲存」** → 同步更新 Access 資料表 & Excel 裡的 Named Range。

---

## 一、建立維護介面

### 1. 在 ControlSheet 上設置下拉驗證

在工作簿中新增一張「Config」工作表，A1\:D1 各放：

| A1     | B1      | C1          | D1      |
| ------ | ------- | ----------- | ------- |
| Report | NameTag | CellAddress | \[修改後值] |

1. **A2**：資料驗證 → 來源：`=ReportList`
2. **B2**：資料驗證 → 來源動態依 A2 而定（見下）
3. **C2**：直接顯示目前的 CellAddress（用公式或 VBA 一拉就寫入）
4. **D2**：使用者可在此輸入新的 CellAddress

### 2. 定義動態 Named Range

* `ReportList`：來自 `PositionMap` 取得的所有不重複 ReportName
* `NameTagList_[Report]`：對應每個 Report 在 `PositionMap` 下的所有 `SourceNameTag`

可先寫一支小程式在開檔時跑一次，將這些清單寫到隱藏工作表，並定義成 Named Range。

---

## 二、UserForm 設計

1. **ComboBox\_Report**：載入所有 ReportName
2. **ComboBox\_NameTag**：當 Report 改選時，重新 Load 該報表的所有 NameTag
3. **TextBox\_CellAddress**：顯示目前設定的 CellAddress，可讓使用者修改
4. **CommandButton\_Save**：按下後執行「將修改寫回 Access & 更新 Excel 名稱」

---

## 三、核心 VBA 代碼結構

以下程式碼提供核心邏輯，您可將它放在 `ThisWorkbook` 或模組中（依專案架構調整）。

### 1. 啟動表單、初始化清單

```vb
' 開啟維護表單
Sub ShowConfigForm()
    Dim frm As New frmConfig
    frm.InitializeForm  ' 自訂初始化
    frm.Show
End Sub
```

```vb
' frmConfig.InitializeForm
Public Sub InitializeForm()
    ' 1. 載入所有 ReportName
    Dim reports As Variant
    reports = GetDistinctReports(gDBPath)  ' 下方示範
    
    Me.ComboBox_Report.List = reports
    
    ' 2. If 之前曾選擇，預設選第一個
    If Me.ComboBox_Report.ListCount > 0 Then
        Me.ComboBox_Report.ListIndex = 0
        LoadNameTags Me.ComboBox_Report.Value
    End If
End Sub
```

```vb
' frmConfig.ComboBox_Report_Change
Private Sub ComboBox_Report_Change()
    LoadNameTags Me.ComboBox_Report.Value
End Sub

' 依 ReportName 載入 NameTag
Private Sub LoadNameTags(ByVal reportName As String)
    Dim q As Variant
    q = GetPositionMapData(gDBPath, reportName)  ' 前面已定義
    
    Dim arr() As String, i As Long
    ReDim arr(0 To UBound(q, 1))
    For i = 0 To UBound(q, 1)
        arr(i) = q(i, 1)  ' SourceNameTag
    Next
    
    Me.ComboBox_NameTag.List = arr
    If UBound(arr) >= 0 Then
        Me.ComboBox_NameTag.ListIndex = 0
        Me.TextBox_CellAddress.Value = GetCellAddress(gDBPath, reportName, arr(0))
    End If
End Sub

' 取得單一欄位的 CellAddress
Private Function GetCellAddress(dbPath As String, rpt As String, nameTag As String) As String
    Dim q As Variant
    q = GetPositionMapData(dbPath, rpt)
    Dim i As Long
    For i = 0 To UBound(q, 1)
        If q(i, 1) = nameTag Then
            GetCellAddress = q(i, 2)
            Exit Function
        End If
    Next
    GetCellAddress = ""
End Function

' frmConfig.ComboBox_NameTag_Change
Private Sub ComboBox_NameTag_Change()
    Me.TextBox_CellAddress.Value = GetCellAddress( _
        gDBPath, Me.ComboBox_Report.Value, Me.ComboBox_NameTag.Value)
End Sub
```

### 2. 按「儲存」：同步更新 Access + Excel

```vb
' frmConfig.CommandButton_Save_Click
Private Sub CommandButton_Save_Click()
    Dim rpt As String, nt As String, newAddr As String
    rpt = Me.ComboBox_Report.Value
    nt  = Me.ComboBox_NameTag.Value
    newAddr = Me.TextBox_CellAddress.Value
    
    ' 1. 更新 Access：PositionMap
    Call UpdatePositionMapEntry(gDBPath, rpt, nt, newAddr)
    
    ' 2. 更新 Excel Named Range
    Call UpdateExcelNameTag(rpt, nt, newAddr)
    
    MsgBox "已儲存設定並同步", vbInformation
End Sub
```

```vb
' 更新 PositionMap 資料表中的指定欄位
Public Sub UpdatePositionMapEntry( _
        ByVal DBPath As String, _
        ByVal reportName As String, _
        ByVal nameTag As String, _
        ByVal newAddr As String)
    Dim conn As Object, cmd As Object
    Dim sql As String
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath
    
    sql = "UPDATE PositionMap " & _
          "SET TargetCellAddress = ? " & _
          "WHERE ReportID = (SELECT ReportID FROM Report WHERE ReportName=?) " & _
          "  AND SourceNameTag = ?;"
    
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = sql
        .Parameters.Append .CreateParameter("pAddr", 200, 1, 255, newAddr)
        .Parameters.Append .CreateParameter("pRpt", 200, 1, 50, reportName)
        .Parameters.Append .CreateParameter("pNT", 200, 1, 255, nameTag)
        .Execute
    End With
    
    conn.Close
    Set cmd = Nothing: Set conn = Nothing
End Sub

' 更新 Excel 中對應的 Name 定義
Public Sub UpdateExcelNameTag( _
        ByVal reportName As String, _
        ByVal nameTag As String, _
        ByVal newAddr As String)
    Dim nm As Name
    Dim fullName As String
    
    ' 假設您在 Excel 裡的名稱是 ReportName & "_" & NameTag
    fullName = reportName & "_" & nameTag
    
    On Error Resume Next
    Set nm = ThisWorkbook.Names(fullName)
    On Error GoTo 0
    
    If Not nm Is Nothing Then
        nm.RefersTo = "=" & Replace(newAddr, "$", "")  ' 或者 "='" & Sheet & "'!" & newAddr
    Else
        ' 如果不存在，就新增
        ThisWorkbook.Names.Add Name:=fullName, RefersTo:="=" & newAddr
    End If
End Sub
```

---

## 四、Excel 端「公式下拉」整合（可選）

若您不想用 UserForm，也可以直接在「Config」表用 **Data Validation** 做下拉清單，並用一個「儲存」按鈕跑上面 `UpdatePositionMapEntry` 與 `UpdateExcelNameTag` 的程式。

---

### 小結

1. **UserForm**：最清爽的多層級選單＋即時顯示編輯
2. **Data Validation**：Excel 內建下拉＋按鈕觸發儲存邏輯
3. **同步**：`UpdatePositionMapEntry` (Access) + `UpdateExcelNameTag` (Excel)

這樣能同時滿足——在 Excel 端檢視、編輯欄位設定，又能一鍵推回 Access，且更新 Excel 的 Named Range，使後續用 `clsReport.Init` 或 Process 皆生效。

如果上述有任何細節要再調整，或進一步針對 Drop-down 工作表範例演示，請隨時告訴我！

Que:



' === 模块：modDataAccess ===
' 通用：从 PositionMap 取出指定报表的所有映射
Public Function GetPositionMapData( _
        ByVal DBPath As String, _
        ByVal reportName As String _
    ) As Variant
    Dim conn As Object, rs As Object
    Dim sql As String
    Dim results() As Variant
    Dim i As Long
    
    ' 1. 建立连接
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath
    
    ' 2. SQL: 先 INNER JOIN Report 找到 ReportID，再取 PositionMap
    sql = "SELECT pm.TargetSheetName, pm.SourceNameTag, pm.TargetCellAddress " & _
          "FROM PositionMap AS pm " & _
          "INNER JOIN Report AS r ON pm.ReportID = r.ReportID " & _
          "WHERE r.ReportName = '" & reportName & "' " & _
          "ORDER BY pm.DataId;"
    
    Set rs = conn.Execute(sql)
    If rs.EOF Then
        GetPositionMapData = Array()  ' 没有记录
    Else
        ' 3. 把结果装进二维数组：每行一笔 mapping
        rs.MoveLast: rs.MoveFirst
        ReDim results(0 To rs.RecordCount - 1, 0 To 2)
        i = 0
        Do Until rs.EOF
            results(i, 0) = rs.Fields("TargetSheetName").Value
            results(i, 1) = rs.Fields("SourceNameTag").Value
            results(i, 2) = rs.Fields("TargetCellAddress").Value
            i = i + 1
            rs.MoveNext
        Loop
        GetPositionMapData = results
    End If
    
    rs.Close: conn.Close
    Set rs = Nothing: Set conn = Nothing
End Function


'=== 从 QueryMap 取出指定报表的所有查询配置 ===
Public Function GetQueryMapData( _
        ByVal DBPath As String, _
        ByVal reportName As String _
    ) As Variant
    Dim conn As Object, rs As Object
    Dim sql As String, results() As Variant
    Dim i As Long
    
    ' 1. 建立 ADO 连接
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath
    
    ' 2. 用 INNER JOIN 取 ReportID，再拿 QueryMap
    sql = "SELECT qm.QueryTableName, " & _
          "qm.ImportColName, qm.ImportColNumber " & _
          "FROM QueryMap AS qm " & _
          "INNER JOIN Report AS r ON qm.ReportID = r.ReportID " & _
          "WHERE r.ReportName='" & reportName & "' " & _
          "ORDER BY qm.DataId;"
    
    Set rs = conn.Execute(sql)
    If rs.EOF Then
        GetQueryMapData = Array()    ' 没有任何配置
    Else
        ' 把结果装二维 Array：(0)=QueryTableName, (1)=ImportColName, (2)=ImportColNumber
        rs.MoveLast: rs.MoveFirst
        ReDim results(0 To rs.RecordCount - 1, 0 To 2)
        i = 0
        Do Until rs.EOF
            results(i, 0) = rs.Fields("QueryTableName").Value
            results(i, 1) = rs.Fields("ImportColName").Value
            results(i, 2) = rs.Fields("ImportColNumber").Value
            i = i + 1
            rs.MoveNext
        Loop
        GetQueryMapData = results
    End If
    
    rs.Close: conn.Close
    Set rs = Nothing: Set conn = Nothing
End Function


我看你的這兩個函數還蠻像的，有沒有辦法把這函數併為一個，
簡化為通用函數，
讓
Public Sub Init(ByVal reportName As String, _
                ByVal dataMonthStringROC As String, _
                ByVal dataMonthStringROC_NUM As String, _
                ByVal dataMonthStringROC_F1F2 As String)
和
Public Sub Process_FM11()
共同call一隻function就好，請把完整代碼提供給我，請記得Init要以你給我的最新一個版本為基礎去修改，修改的地方請詳細標示出來

Ans:

下面是一支合并版本的通用取得 Mapping 资料函数 `GetMapData`，它可分别获取 `PositionMap` 或 `QueryMap`，并回传一个二维 `Variant` Array。之后你只要在 `Init` 或 `Process_*` 中指定第三个参数 `"Position"` 或 `"Query"` 即可。

---

```vb
'=== 模块：modDataAccess ===

'—— 通用：从任一 Mapping 表取指定报表的所有记录 —— 
' mapType: "Position" 或 "Query"
'   Position ⇒ 取 PositionMap.TargetSheetName, SourceNameTag, TargetCellAddress
'   Query    ⇒ 取 QueryMap.QueryTableName, ImportColName, ImportColNumber
Public Function GetMapData( _
        ByVal DBPath As String, _
        ByVal reportName As String, _
        ByVal mapType As String _
    ) As Variant

    Dim conn As Object, rs As Object
    Dim sql As String, results() As Variant
    Dim i As Long
    
    ' 1. 建立 ADO 连接
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath
    
    ' 2. 根据 mapType 组织 SQL
    Select Case LCase(mapType)
        Case "position"
            sql = "SELECT pm.TargetSheetName, pm.SourceNameTag, pm.TargetCellAddress " & _
                  "FROM PositionMap AS pm " & _
                  "INNER JOIN Report AS r ON pm.ReportID = r.ReportID " & _
                  "WHERE r.ReportName = '" & reportName & "' " & _
                  "ORDER BY pm.DataId;"
        Case "query"
            sql = "SELECT qm.QueryTableName, qm.ImportColName, qm.ImportColNumber " & _
                  "FROM QueryMap AS qm " & _
                  "INNER JOIN Report AS r ON qm.ReportID = r.ReportID " & _
                  "WHERE r.ReportName = '" & reportName & "' " & _
                  "ORDER BY qm.DataId;"
        Case Else
            GetMapData = Array()
            Exit Function
    End Select
    
    Set rs = conn.Execute(sql)
    If rs.EOF Then
        GetMapData = Array()    ' 没有记录
    Else
        ' 把结果装二维 Array：0..RecordCount-1, 0..2
        rs.MoveLast: rs.MoveFirst
        ReDim results(0 To rs.RecordCount - 1, 0 To 2)
        i = 0
        Do Until rs.EOF
            results(i, 0) = rs.Fields(0).Value
            results(i, 1) = rs.Fields(1).Value
            results(i, 2) = rs.Fields(2).Value
            i = i + 1
            rs.MoveNext
        Loop
        GetMapData = results
    End If
    
    rs.Close: conn.Close
    Set rs = Nothing: Set conn = Nothing
End Function
```

---

### 修改后 `clsReport.Init` （基于你最新版本，改用 `GetMapData`）

```vb
Public Sub Init(ByVal reportName As String, _
                ByVal dataMonthStringROC As String, _
                ByVal dataMonthStringROC_NUM As String, _
                ByVal dataMonthStringROC_F1F2 As String)
    
    clsReportName = reportName
    Set clsWorksheets = CreateObject("Scripting.Dictionary")

    ' === 【修改】 改为调用通用函数 GetMapData(..., "Position") ===
    Dim mappings As Variant
    mappings = GetMapData(gDBPath, reportName, "Position")
    
    Dim wsFields As Object
    Set wsFields = CreateObject("Scripting.Dictionary")

    If IsArray(mappings) And UBound(mappings) >= 0 Then
        Dim i As Long
        For i = 0 To UBound(mappings, 1)
            Dim sheetName As String, nameTag As String, addr As String
            sheetName = mappings(i, 0)
            nameTag   = mappings(i, 1)
            addr      = mappings(i, 2)
            
            If Len(Trim(nameTag)) > 0 And Len(Trim(addr)) > 0 Then
                If Not wsFields.Exists(sheetName) Then
                    wsFields.Add sheetName, Array()
                End If
                Dim tmpList As Variant
                tmpList = wsFields(sheetName)
                ReDim Preserve tmpList(0 To UBound(tmpList) + 1)
                tmpList(UBound(tmpList)) = Array(nameTag, addr, Null)
                wsFields(sheetName) = tmpList
            End If
        Next i

        Dim key As Variant
        For Each key In wsFields.Keys
            AddWorksheetFields key, wsFields(key)
        Next key
    Else
        WriteLog "Init 未在 PositionMap 找到报表：" & reportName
    End If

    ' === 保留原有动态字段逻辑（不变） ===
    Select Case reportName
        Case "TABLE10"
            AddDynamicField reportName, "TABLE10_申報時間", "D2", dataMonthStringROC
        Case "TABLE15A"
            AddDynamicField reportName, "TABLE15A_申報時間", "D2", dataMonthStringROC
        ' … 其他 Case 保留不变 …
    End Select
End Sub
```

---

### 修改后 `Process_FM11` （直接调用 `GetMapData(..., "Query")`）

```vb
Public Sub Process_FM11()
    Dim rpt As clsReport
    Set rpt = gReports("FM11")
    
    Dim reportTitle As String
    reportTitle = rpt.ReportName
    
    Dim xlsht As Worksheet
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    xlsht.Cells.ClearContents  ' 或只清指定区
    
    ' === 【修改】 改为调用通用函数 GetMapData(..., "Query") ===
    Dim qmap As Variant
    qmap = GetMapData(gDBPath, reportTitle, "Query")
    
    If Not IsArray(qmap) Or UBound(qmap) < 0 Then
        MsgBox "未在 QueryMap 找到 " & reportTitle & " 的任何查询配置！", vbExclamation
        Exit Sub
    End If
    
    Dim iMap As Long
    For iMap = 0 To UBound(qmap, 1)
        Dim tblName As String
        Dim startColLetter As String
        Dim numCols As Long
        Dim dataArr As Variant
        Dim r As Long, c As Long
        tblName        = qmap(iMap, 0)
        startColLetter = qmap(iMap, 1)
        numCols        = CLng(qmap(iMap, 2))
        
        dataArr = GetAccessDataAsArray(gDBPath, tblName, gDataMonthString)
        If Not IsArray(dataArr) Then GoTo NextMap
        
        Dim startCol As Long
        startCol = xlsht.Range(startColLetter & "1").Column
        
        For r = 0 To UBound(dataArr, 0)
            For c = 0 To numCols - 1
                xlsht.Cells(r + 1, startCol + c).Value = dataArr(r, c)
            Next c
        Next r
NextMap:
    Next iMap
    
    ' … 后续“Unique Setting”、“ValidateFields”、“UpdateRecord”等逻辑不变 …
    
    If rpt.ValidateFields() Then
        Dim key As Variant
        Dim allValues As Object, allPositions As Object
        Set allValues    = rpt.GetAllFieldValues()
        Set allPositions = rpt.GetAllFieldPositions()
        For Each key In allValues.Keys
            UpdateRecord gDBPath, gDataMonthString, rpt.ReportName, _
                         key, allPositions(key), allValues(key)
        Next key
    End If
    
    xlsht.Tab.ColorIndex = 6
End Sub
```

---

## 小结

* **新增** 一支通用函数 `GetMapData(DBPath, reportName, mapType)`，替代原本两个几乎相同的函数。
* **Init** 改用 `GetMapData(..., "Position")`，并保持你的批次 `AddWorksheetFields` 逻辑不变。
* **Process\_FM11** 改用 `GetMapData(..., "Query")` 动态取得查询设定。

这样后续如果你要再新增其他类型的 Map，只需在同一函数中加一个 `Case` 即可，整个架构更为精简、一致。



Que:

我希望下拉驗證可以單純只能查看特定Report包含哪些NameTag，以及對應的儲存格和分頁．
只有進入UserForm才能進行資料更新的動作，
請將下拉驗證和UserForm各自進行的程序詳細定義和說明清楚，
任何細節都要說明，
並且要把兩者間的互動過程仔細考慮到，
另外再更新的時候不只要更新資料庫的資料，
還要針對所更新的ReportName(即分頁名稱）、SheetName、儲存格名稱 和 對應的NameTag在Excel中進行更新，請詳細說明整個過程和各個細節

Ans:

下面的方案分两大块：

1. **工作表端“只读”下拉验证**——让使用者在「Config」页、无需 VBA Form 即可查看某报表所有 NameTag、对应分页与单元格，却不能修改。
2. **UserForm 编辑／更新流程**——只有进入专用表单才可修改，再一键把变动写回 Access、同步更新 Excel 中的 Named Range（含 ReportName、SheetName、NameTag、CellAddress）。

---

## 一、工作表端「下拉验证」：只供查看

### 1. 新增「Config」工作表

在你的活跃工作簿里新增一个名为 **Config** 的工作表，并按下列方式布局：

| 单元格 | 含义          | 备注                      |
| --- | ----------- | ----------------------- |
| A1  | Report      | 报表名称（下拉）                |
| B1  | NameTag     | 该报表所有 SourceNameTag（下拉） |
| C1  | SheetName   | 对应的目标分頁（只读，公式拉取）        |
| D1  | CellAddress | 对应的儲存格地址（只读，公式拉取）       |

### 2. 准备动态清单数据源

为了让 A2、B2 的下拉能对应 Access 里最新的 PositionMap，我们需要：

1. **隐藏页**（例如 “\_Lists”）：存放

   * **ReportList**：所有不重复的 ReportName
   * **NameTagList\_<Report>**：每个 Report 对应的 SourceNameTag 列表
2. **开档时**用宏或 `Workbook_Open` 事件把这两张清单从 Access 拉下来，写进“\_Lists”页，并建立 Named Range。

示例伪代码（放在 `ThisWorkbook`）：

```vb
Private Sub Workbook_Open()
    Call RefreshReportList
    Call RefreshNameTagLists
End Sub

Sub RefreshReportList()
    Dim arrRpt As Variant
    arrRpt = GetDistinctReports(gDBPath)  ' 从 Access 的 Report 表取所有 ReportName
    ' 把 arrRpt 写入 "_Lists"!A2:A...
    ' 再定义 Named Range "ReportList" 指向那段区域
End Sub

Sub RefreshNameTagLists()
    Dim rpt As Variant, i As Long
    Dim mappings As Variant
    For i = LBound(arrRpt) To UBound(arrRpt)
        rpt = arrRpt(i)
        mappings = GetPositionMapData(gDBPath, rpt)  ' 从 PositionMap 拉这个报表的所有(…)映射
        ' 把 mappings(:,1) — 即 SourceNameTag — 写到 "_Lists"!B2:B...
        ' 再定义 Named Range "NameTagList_" & rpt
    Next
End Sub
```

### 3. 给 A2、B2 设下拉验证

* **A2**：数据验证 → 允选区域 → `=ReportList`
* **B2**：数据验证 → 允选区域 → `=INDIRECT("NameTagList_" & A2)`

这样，用户在 A2 选报表，B2 就自动只列出该报表所有 NameTag；但 C2、D2 只是用公式显示详情，不能改。

### 4. C2、D2 的显示公式

* **C2**（SheetName）：

  ```excel
  =IFERROR(
    INDEX(_Lists!SheetCol, MATCH(1, (_Lists!RptCol=A2)*(_Lists!NameTagCol=B2), 0)),
    ""
  )
  ```
* **D2**（CellAddress）：

  ```excel
  =IFERROR(
    INDEX(_Lists!AddrCol, MATCH(1, (_Lists!RptCol=A2)*(_Lists!NameTagCol=B2), 0)),
    ""
  )
  ```

务必把 `_Lists` 页中对应的列定义成 Named Ranges（例如 `RptCol`, `NameTagCol`, `SheetCol`, `AddrCol`），并以 Ctrl+Shift+Enter 确认数组运算。

> **效果**：
>
> * 选 A2 报表 → B2 下拉只列出该报表 NameTag → C2、D2 即时显示对应 SheetName、CellAddress。
> * 所有单元都设为读锁（或保护工作表），以免误改。

---

## 二、UserForm 编辑／更新流程

### 1. UserForm 元件布局

在 VBAEditor 里新增一个 **UserForm**（例如命名 `frmConfig`），控件如下：

| 控件名称                  | 类型       | 说明                               |
| --------------------- | -------- | -------------------------------- |
| ComboBox\_Report      | ComboBox | 报表下拉，列出所有 ReportName             |
| ComboBox\_NameTag     | ComboBox | NameTag 下拉，依 ComboBox\_Report 切换 |
| TextBox\_SheetName    | TextBox  | 仅读，显示对应 SheetName                |
| TextBox\_CellAddress  | TextBox  | 可编辑，显示／修改对应 CellAddress          |
| CommandButton\_Save   | Button   | 「儲存變更」，写回 DB & 更新 Excel 名称定义     |
| CommandButton\_Cancel | Button   | 关闭 Form，不做任何储存                   |

### 2. Form 初始化（`UserForm_Initialize` 或自定义 `InitializeForm`）

```vb
Private Sub UserForm_Initialize()
    Dim reports As Variant
    reports = GetDistinctReports(gDBPath)  ' 从 Access Report 表拉 ReportName
    
    Me.ComboBox_Report.List = reports
    If Me.ComboBox_Report.ListCount > 0 Then
        Me.ComboBox_Report.ListIndex = 0
        LoadNameTags Me.ComboBox_Report.Value
    End If
End Sub
```

### 3. 切换报表时载入 NameTag（`ComboBox_Report_Change`）

```vb
Private Sub ComboBox_Report_Change()
    LoadNameTags Me.ComboBox_Report.Value
End Sub

Private Sub LoadNameTags(ByVal rptName As String)
    Dim mappings As Variant, i As Long
    mappings = GetPositionMapData(gDBPath, rptName)
    
    Dim arrTags() As String
    If IsArray(mappings) Then
        ReDim arrTags(0 To UBound(mappings,1))
        For i = 0 To UBound(mappings,1)
            arrTags(i) = mappings(i,1)  ' SourceNameTag
        Next
        Me.ComboBox_NameTag.List = arrTags
        If UBound(arrTags) >= 0 Then
            Me.ComboBox_NameTag.ListIndex = 0
            UpdateDetailFields rptName, arrTags(0)
        End If
    Else
        Me.ComboBox_NameTag.Clear
        Me.TextBox_SheetName.Value = ""
        Me.TextBox_CellAddress.Value = ""
    End If
End Sub
```

### 4. 切换 NameTag 即时显示详情（`ComboBox_NameTag_Change`）

```vb
Private Sub ComboBox_NameTag_Change()
    UpdateDetailFields Me.ComboBox_Report.Value, Me.ComboBox_NameTag.Value
End Sub

Private Sub UpdateDetailFields(ByVal rpt As String, ByVal nt As String)
    Dim mappings As Variant, i As Long
    mappings = GetPositionMapData(gDBPath, rpt)
    For i = 0 To UBound(mappings,1)
        If mappings(i,1) = nt Then
            Me.TextBox_SheetName.Value   = mappings(i,0)
            Me.TextBox_CellAddress.Value = mappings(i,2)
            Exit For
        End If
    Next
End Sub
```

### 5. 按「儲存變更」：同步更新 Access + Excel（`CommandButton_Save_Click`）

```vb
Private Sub CommandButton_Save_Click()
    Dim rpt As String, nt As String, newAddr As String, sht As String
    rpt     = Me.ComboBox_Report.Value
    nt      = Me.ComboBox_NameTag.Value
    sht     = Me.TextBox_SheetName.Value
    newAddr = Me.TextBox_CellAddress.Value
    
    If rpt = "" Or nt = "" Or sht = "" Or newAddr = "" Then
        MsgBox "所有欄位皆須選擇／填寫完整", vbExclamation
        Exit Sub
    End If
    
    ' 1) 更新 Access
    UpdatePositionMapEntry gDBPath, rpt, nt, newAddr, sht
    
    ' 2) 更新 Excel Named Range & Worksheet 中的 Name 定义
    UpdateExcelNameTag rpt, nt, sht, newAddr
    
    MsgBox "設定已儲存並同步更新！", vbInformation
    Me.Hide
End Sub
```

#### 5.1 更新 Access：`UpdatePositionMapEntry`

```vb
Public Sub UpdatePositionMapEntry( _
        ByVal DBPath     As String, _
        ByVal reportName As String, _
        ByVal nameTag    As String, _
        ByVal newAddr    As String, _
        ByVal newSheet   As String)
    Dim conn As Object, cmd As Object, sql As String
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath
    
    sql = "UPDATE PositionMap SET " & _
          "TargetCellAddress=?, TargetSheetName=? " & _
          "WHERE ReportID=(SELECT ReportID FROM Report WHERE ReportName=?) " & _
          "  AND SourceNameTag=?;"
    
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText     = sql
        .Parameters.Append .CreateParameter("pAddr", 200, 1, 255, newAddr)
        .Parameters.Append .CreateParameter("pSheet", 200, 1, 255, newSheet)
        .Parameters.Append .CreateParameter("pRpt", 200, 1, 50, reportName)
        .Parameters.Append .CreateParameter("pNT", 200, 1, 255, nameTag)
        .Execute
    End With
    
    conn.Close
End Sub
```

#### 5.2 更新 Excel：`UpdateExcelNameTag`

```vb
Public Sub UpdateExcelNameTag( _
        ByVal reportName   As String, _
        ByVal nameTag      As String, _
        ByVal sheetName    As String, _
        ByVal cellAddress  As String)
    Dim nm As Name, fullName As String
    ' 假设 Name 定义格式：ReportName & "_" & NameTag
    fullName = reportName & "_" & nameTag
    
    ' 1) 更新该名称指向的新分頁与单元格
    On Error Resume Next
    Set nm = ThisWorkbook.Names(fullName)
    On Error GoTo 0
    
    Dim refersToStr As String
    refersToStr = "= '" & sheetName & "'!" & cellAddress
    
    If Not nm Is Nothing Then
        nm.RefersTo = refersToStr
    Else
        ThisWorkbook.Names.Add Name:=fullName, RefersTo:=refersToStr
    End If
    
    ' 2) （可选）更新对应 Worksheet 上的单元格名称
    '    如果你在 sheet 中也建立了同名命名区域，可考虑同步。
End Sub
```

---

## 三、使用者操作／互动流程

1. **日常查看**

   * 切到 **Config** 工作表
   * A2 选报表 → B2 选 NameTag → C2、D2 即时显示对应的分頁与儲存格
   * 无法直接修改，保证只读安全。

2. **修改配置**

   * 按「Ribbon 上的自定义按钮」或直接在 VBA 中呼叫 `ShowConfigForm`，弹出 `frmConfig`
   * 在 Form 里：

     1. 选报表 → NameTag 下拉自动加载
     2. TextBox\_SheetName 只读显示
     3. TextBox\_CellAddress 可修改
     4. 按「儲存變更」 → 自动完成：

        * Update Access `PositionMap` 的 `TargetSheetName` 与 `TargetCellAddress`
        * Update Excel 中 Named Range（`ReportName_NameTag`）指向新 sheet & cell
   * Form 关闭后：

     * 若需要立即更新 Config 页的下拉与显示，可再呼叫一次 `RefreshNameTagLists` 与刷新公式。

---

### 注意事项

* **保护 Config 工作表**：锁定 C2、D2，避免意外手动编辑。
* **Named Range 命名规范**：确保 `ReportName_NameTag` 无特殊字符，或自行清理空格／标点。
* **同步刷新**：每次修改后，若要让“Config”表里看到最新的下拉选项（例如改了 NameTag），需执行一次 `RefreshNameTagLists`。可在 `CommandButton_Save_Click` 结尾加上。
* **Error Handling**：在生产环境建议加上错误提示、事务回滚等，确保数据库／Excel 始终一致。

---

这样就完整串起了：

* **只读查看**：用「Config」表内的下拉＋公式实时展现 PositionMap；
* **编辑 & 更新**：通过 `frmConfig` 一条龙写回 Access、同步更新 Excel 名称定义。

如有任何细节需深入或范例演示，请随时告诉我！

Ans:

还没给出具体的保护代码，下面补上完整的 VBA 流程，帮你在打开档案或更新完下拉清单后，一键锁定 “Config” 工作表中 C、D 两栏并启用保护，只留 A、B 可编辑。

---

### 1. 在 `ThisWorkbook` 模块中加入保护逻辑

```vb
Private Sub Workbook_Open()
    ' 先刷新清单（已有）
    Call RefreshReportList
    Call RefreshNameTagLists
    ' 再保护 Config 表
    Call ProtectConfigSheet
End Sub
```

---

### 2. 新增 `ProtectConfigSheet` 子程序

把下面这段放在标准模块（例如 `modConfig`）里：

```vb
' 锁定 Config 工作表上的 C:D，保护工作表但让 VBA 代码可操作（UserInterfaceOnly）
Public Sub ProtectConfigSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Config")
    
    ' 1. 先关闭保护，避免重复报错
    On Error Resume Next
    ws.Unprotect Password:="1234"   ' 请自行替换或移除密码
    On Error GoTo 0
    
    ' 2. 默认：先把全表所有单元格设为可编辑（Unlocked）
    ws.Cells.Locked = False
    
    ' 3. 锁定 C、D 两列（或你想保护的范围）
    ws.Columns("C:D").Locked = True
    
    ' 4. 重新启用保护，仅禁止用户手动修改锁定区
    ws.Protect Password:="1234", _
               DrawingObjects:=True, _
               Contents:=True, _
               Scenarios:=True, _
               UserInterfaceOnly:=True
    ' UserInterfaceOnly:=True 让 VBA 宏仍可通过代码写入 C、D
End Sub
```

* **步骤说明**

  1. `Unprotect`：先解锁已有保护（若未设密码可省略密码参数）。
  2. `ws.Cells.Locked = False`：取消所有单元格的锁定状态，为后续精确锁定做准备。
  3. `ws.Columns("C:D").Locked = True`：把 C、D 列（包括标题行）设为锁定。
  4. `ws.Protect … UserInterfaceOnly:=True`：开启保护，让使用者不能手动改动 C、D，但 VBA 程序仍可写入。

---

### 3. 在「刷新下拉清单」后也调用一次

若你每次更新完 NameTag 清单希望及时把“Config”表也重新锁定，可在 `RefreshNameTagLists` 的最后，加入：

```vb
    Call ProtectConfigSheet
```

这样无论是开档还是手动重刷，都能确保 C、D 列始终上锁。

---

#### 小贴士

* 若要锁定 **单元格范围**（如 C2\:D100 而非整列），可以用：

  ```vb
  ws.Range("C2:D100").Locked = True
  ```
* 若不想用密码，则把 `Password:="1234"` 删掉即可。
* `UserInterfaceOnly:=True` 非常关键，否则保护后连宏都不能写入锁定区——开档事件中必须每次都呼叫一次，因 Excel 重启会重置此属性。

这样就能保证：在 **Config** 表里，A、B 两列依然自由下拉选择，C、D 列无论公式或后端刷值，都无法被用户误删、误改。
