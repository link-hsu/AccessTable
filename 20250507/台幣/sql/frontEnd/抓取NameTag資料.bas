Q:
我在Excel中的儲存格有儲存許多NameTag， 請問我要怎麼使用VBA遍歷擷取所有儲存NameTag資訊， 貼到其中一個分頁，以資料表的方式呈現， 第一個欄位名稱為儲存的NameTag名稱， 第二個欄位名稱為分頁名稱， 第三個欄位名稱為儲存格的Range 例如: H8

A:
下面是一段範例 VBA 程式，示範如何將活頁簿中所有的 NameTag（Named Range）依序擷取，並彙整到一個名為「NameTagList」的工作表中，格式如下：

| NameTag 名稱 | 分頁名稱 | 儲存格範圍 |
| ---------- | ---- | ----- |

請將這段程式貼到 VBA IDE 的一個新模組中（Alt + F11 → 插入 → 模組），然後執行 `ExtractNameTags` 即可。

```vb
Option Explicit

Sub ExtractNameTags()
    Dim wb As Workbook
    Dim nm As Name
    Dim rng As Range
    Dim outWs As Worksheet
    Dim nextRow As Long
    
    Set wb = ThisWorkbook
    
    '── 建立或重新建立匯出用工作表 ──
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets("NameTagList").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set outWs = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    outWs.Name = "NameTagList"
    
    ' 設定標題列
    With outWs
        .Range("A1").Value = "NameTag 名稱"
        .Range("B1").Value = "分頁名稱"
        .Range("C1").Value = "儲存格範圍"
        .Rows(1).Font.Bold = True
    End With
    
    nextRow = 2
    
    '── 逐一處理所有 Name 物件 ──
    For Each nm In wb.Names
        On Error Resume Next
        Set rng = nm.RefersToRange   ' 嘗試取回對應的 Range
        If Err.Number = 0 Then
            ' 成功取得 Range 才輸出
            With outWs
                .Cells(nextRow, 1).Value = nm.Name                       ' 名稱
                .Cells(nextRow, 2).Value = rng.Parent.Name              ' 所在工作表
                .Cells(nextRow, 3).Value = rng.Address(RowAbsolute:=False, ColumnAbsolute:=False)  ' 範圍 (例如 H8)
            End With
            nextRow = nextRow + 1
        End If
        Err.Clear
    Next nm
    
    ' 調整欄寬
    outWs.Columns("A:C").AutoFit
    
    MsgBox "已完成擷取，共找到 " & nextRow - 2 & " 個 NameTag。", vbInformation
End Sub
```

---

### 程式說明

1. **刪除既有工作表並重新建立**
   為避免多次執行造成重複，程式一開始會嘗試刪除名為 `NameTagList` 的工作表，再重新建立一個新的。

2. **標題列設定**
   在新工作表的第一列，設定好三個欄位標題並加粗，以符合「NameTag 名稱／分頁名稱／儲存格範圍」的格式需求。

3. **遍歷所有 Named Range**
   `For Each nm In wb.Names` 會抓出所有名稱，透過 `nm.RefersToRange` 取得對應的 Range 物件（如果該 NameTag 指向常數或外部參照就會出錯，程式會自動跳過）。

4. **資料寫入**

   * `nm.Name`：NameTag 本身的名稱。
   * `rng.Parent.Name`：該 NameTag 所在的工作表名稱。
   * `rng.Address(False, False)`：不固定列、欄，純粹「H8」這種格式的位址。

5. **美化與提示**

   * 執行完畢後自動調整欄寬（`AutoFit`），並以訊息提示共找到幾個 NameTag。

---

這段程式只要貼到 VBA 模組，按下執行，就能快速把所有 NameTag 的資訊整理成一張清單，方便後續管理或匯出。若有客製化需求（例如只列出特定前綴的 NameTag，或是輸出至 CSV），再追加條件即可。祝開發順利！
