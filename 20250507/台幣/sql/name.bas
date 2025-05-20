Answer1.


下面是一支通用的 VBA 子程序（Sub），它會自動把工作表上以「第一欄做列名、第一列做欄名」的儲存格，依規則命名為：

```
Table10_<列名稱>_<欄名稱>
```

你只要把它貼到任何一個標準模組（Module）裡，就可以在不同工作表、不同範圍下重複使用。

---

```vb
Option Explicit

'---------------------------------------------------------
' 將工作表中以「第一欄做列標題、第一列做欄標題」的資料區
' 按 Table10_<列名稱>_<欄名稱> 命名每一個儲存格
'
' 參數：
'   tblName   - 前綴（如 "Table10"）
'   headerRow - 欄標題所在列 (預設 1)
'   headerCol - 列標題所在欄 (預設 1)
'---------------------------------------------------------
Public Sub NameTableCells( _
        Optional ByVal tblName   As String = "Table10", _
        Optional ByVal headerRow As Long   = 1, _
        Optional ByVal headerCol As Long   = 1)
    
    Dim ws        As Worksheet
    Dim lastRow   As Long, lastCol As Long
    Dim r         As Long, c       As Long
    Dim rowName   As String, colName As String
    Dim nm        As String
    
    Set ws = Application.Caller.Worksheet  ' 宏從工作表上執行時自動指向
        
    ' 找出使用範圍的最底列和最右欄
    With ws
        lastRow = .Cells(.Rows.Count, headerCol).End(xlUp).Row
        lastCol = .Cells(headerRow, .Columns.Count).End(xlToLeft).Column
    End With
    
    Application.ScreenUpdating = False
    On Error Resume Next
    
    ' 逐一命名：從 (headerRow+1, headerCol+1) 到 (lastRow, lastCol)
    For r = headerRow + 1 To lastRow
        rowName = ws.Cells(r, headerCol).Value
        If Len(Trim(rowName)) = 0 Then GoTo NextRow
        ' 清理不合法字元：只留字母數字底線
        rowName = Application.WorksheetFunction.Clean(rowName)
        rowName = Replace(rowName, " ", "_")
        
        For c = headerCol + 1 To lastCol
            colName = ws.Cells(headerRow, c).Value
            If Len(Trim(colName)) = 0 Then GoTo NextCell
            colName = Application.WorksheetFunction.Clean(colName)
            colName = Replace(colName, " ", "_")
            
            nm = tblName & "_" & rowName & "_" & colName
            ' 如果同名已存在，先刪除舊的
            ThisWorkbook.Names(nm).Delete
            
            ' 新增名稱
            ThisWorkbook.Names.Add _
                Name:=nm, _
                RefersTo:="=" & ws.Name & "!" & ws.Cells(r, c).Address(False, False)
            
NextCell:
        Next c
NextRow:
    Next r
    
    Application.ScreenUpdating = True
    On Error GoTo 0
    
    MsgBox "命名完成！共處理 " & (lastRow - headerRow) * (lastCol - headerCol) & " 個儲存格", vbInformation
End Sub
```

---

## 使用步驟

1. 按 **Alt+F11** 開啟 VBA 編輯器。
2. 在左側專案樹上，右鍵點你的活頁簿名 → **Insert** → **Module**。
3. 把上面的程式碼貼進去，存檔並關閉編輯器。
4. 回到 Excel，在要命名的工作表上按 **Alt+F8**，選擇 `NameTableCells` → **Run**。

   * 若你的資料列標題不在第 1 欄、欄標題不在第 1 列，可在執行時按「編輯參數」（Edit）填入自訂的 `headerRow`、`headerCol`。
   * `tblName` 參數預設為 `"Table10"`，也可以改成別的前綴。

執行完畢後，每一個資料儲存格就會有對應的名稱。之後在公式或 VBA 裡都能直接用：

```excel
=SUM(Table10_列名_欄名, …)
```

這支程式寫成 Sub，也很容易移植到其他檔案重複使用。若還有進階需求（例如範圍參數化、掃描多個工作表⋯⋯），再跟我說！





Answer2.



下面這支簡化版的 VBA 巨集示範了「以第 x 列做標題，把每一個標題下方（第 x+1 列）對應的儲存格，命名成 `前綴_標題文字`」的做法：

```vb
Option Explicit

' ---------------------------------------------------
' 將第 headerRow 列的欄標題，命名第 headerRow+1 列的儲存格
'
' 參數：
'   prefix    - 名稱前綴（例如 "Table10"）
'   headerRow - 標題所在列 (例如 5 表示第 5 列是你要取標題的列)
'   startCol  - 從哪一欄開始 (預設 1)
' ---------------------------------------------------
Public Sub NameCellsBelowHeader( _
        Optional ByVal prefix    As String = "Table10", _
        Optional ByVal headerRow As Long   = 1, _
        Optional ByVal startCol  As Long   = 1)
    
    Dim ws       As Worksheet
    Dim lastCol  As Long
    Dim c        As Long
    Dim hdr      As String
    Dim nm       As String
    
    Set ws = Application.Caller.Worksheet
    
    ' 找到第 headerRow 列最後一個有值的欄號
    With ws
        lastCol = .Cells(headerRow, .Columns.Count).End(xlToLeft).Column
    End With
    
    On Error Resume Next   ' 如果同名已存在就跳過
    For c = startCol To lastCol
        hdr = Trim(ws.Cells(headerRow, c).Value)
        If hdr <> "" Then
            ' 把空格都換成底線，並去除不可見字元
            hdr = Application.WorksheetFunction.Clean(hdr)
            hdr = Replace(hdr, " ", "_")
            
            nm = prefix & "_" & hdr
            ' 刪掉同名（若存在）
            ThisWorkbook.Names(nm).Delete
            ' 新增名稱：指向 row=headerRow+1, col=c 的儲存格
            ThisWorkbook.Names.Add _
                Name:=nm, _
                RefersTo:="=" & ws.Name & "!" & _
                          ws.Cells(headerRow + 1, c).Address(False, False)
        End If
    Next c
    On Error GoTo 0
    
    MsgBox "已為第 " & headerRow & " 列下方儲存格命名：" & _
           (lastCol - startCol + 1) & " 個", vbInformation

End Sub
```

---

### 怎麼用

1. **貼到 VBA 模組**：

   * 按 Alt+F11 → 插入 Module → 貼上程式。

2. **執行**：

   * 回到 Excel 按 Alt+F8 選 `NameCellsBelowHeader` → **執行**
   * 執行時可以按「編輯參數」（Edit）改：

     * `prefix`：你想要的前綴字串（預設 `"Table10"`）
     * `headerRow`：標題所在列（例如你要第 5 列當標題就填 5）
     * `startCol`：從哪一欄開始（預設從第 1 欄）

它就會把第 `headerRow` 列的每個標題下方（第 `headerRow+1` 列）儲存格，以 `prefix_標題文字` 的方式加上名稱。

---

**範例**

* 假設第 3 列有欄標題：`產品`、`數量`、`價格`
* 你執行：

  ```vb
  NameCellsBelowHeader prefix:="Table10", headerRow:=3
  ```
* 那麼第 4 列的 A、B、C 欄就會被命名成

  * `Table10_產品` → 指向 A4
  * `Table10_數量` → 指向 B4
  * `Table10_價格` → 指向 C4

有問題或要再客製化再跟我說！



Answer3

你可以用以下幾種方式來呼叫這兩支 Sub（`NameTableCells` 與 `NameCellsBelowHeader`）：

---

## 一、使用 Alt+F8 執行（最簡單）

1. 在 Excel 按 **Alt + F8**
2. 在宏清單中選取「NameTableCells」或「NameCellsBelowHeader」
3. 按 **執行**（Run）
4. 如果宏有 Optional 參數，會跳出「編輯參數」視窗，你可以填：

   * `tblName` 或 `prefix`：前綴字串
   * `headerRow`：標題列
   * `headerCol` 或 `startCol`：標題欄或起始欄

---

## 二、在另一支 Sub 裡以程式碼呼叫

你可以寫一支「呼叫器」Macro，直接在裡面帶入你要的參數：

```vb
Sub Test_Call_NameTableCells()
    ' 範例：用 Table10_ 前綴，第 1 列／第 1 欄當標題
    Call NameTableCells(tblName:="Table10", headerRow:=1, headerCol:=1)
End Sub

Sub Test_Call_NameCellsBelowHeader()
    ' 範例：用 MyTbl_ 前綴，第 5 列當標題，從第 2 欄開始
    Call NameCellsBelowHeader(prefix:="MyTbl", headerRow:=5, startCol:=2)
End Sub
```

1. 同樣貼到 Module 裡
2. 按 Alt+F8 執行 `Test_Call_…`

---

## 三、指派到工作表按鈕或捷徑鍵

1. **插入按鈕**

   * 開啟「開發人員」索引標籤 → **插入** → 選「表單按鈕(Form Control)」
   * 在工作表畫一個按鈕，跳出指派巨集對話框時選 `NameTableCells` 或 `NameCellsBelowHeader`
   * 完成後點按鈕就會執行

2. **設定捷徑鍵**

   * 按 **Alt + F8** → 選取某一宏 → 點「選項…」
   * 指定一個 Ctrl+<字母> 做為快速鍵

---

### 小結

* **互動式**：Alt+F8 → 選擇 → 執行／參數
* **程式化**：在另一個 Sub 裡 `Call NameTableCells(...)`
* **UI 觸發**：按鈕或快捷鍵

這三種方式都可以靈活地在不同檔案或情境下，快速呼叫並帶入你的命名邏輯。
