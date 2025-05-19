下面這段範例 VBA 程式示範，如何用迴圈依照你指定的規則，快速為範圍內的每一個儲存格建立名稱（Name）。假設你要為工作表 `Sheet1` 上從 A2 到 A10 的每個儲存格，依序命名為 `Item1`、`Item2`……`Item9`：

```vb
Sub NameCellsInLoop()
    Dim ws    As Worksheet
    Dim rng   As Range
    Dim cell  As Range
    Dim idx   As Long
    Dim baseName As String
    
    ' 1. 指定工作表和範圍
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    Set rng = ws.Range("A2:A10")
    
    ' 2. 命名前綴
    baseName = "Item"
    
    ' 3. 清除既有同名範圍（可選）
    On Error Resume Next
      For idx = 1 To rng.Count
        ThisWorkbook.Names(baseName & idx).Delete
      Next idx
    On Error GoTo 0
    
    ' 4. 逐一為每個 cell 建立 Name
    idx = 1
    For Each cell In rng
        ThisWorkbook.Names.Add _
          Name:=baseName & idx, _
          RefersTo:= "=" & ws.Name & "!" & cell.Address(False, False)
        idx = idx + 1
    Next cell
    
    MsgBox "已為範圍 " & rng.Address & " 命名為 " & baseName & "1～" & baseName & (idx - 1), vbInformation
End Sub
```

---

### 重點說明

1. **清除既有名稱**
   如果你事先已經用過同樣的 `Item1`、`Item2` … 名稱，可以先刪除避免衝突（步驟 3）。

2. **`ThisWorkbook.Names.Add`**

   * `Name:=` 參數指定名稱字串。
   * `RefersTo:=` 參數一定要以等號 `=` 開頭，後面接工作表名稱和儲存格地址。

3. **規則自由定義**

   * 只要把 `baseName = "Item"` 改成 `"CP2_"`、`"Dept"`… 或把迴圈走過的是欄、列、區塊都可以。
   * 例如：若想把 B2\:F2 五個標題依序命名為 Header\_B、Header\_C…，可改成：

     ```vb
     Set rng = ws.Range("B2:F2")
     baseName = "Header_"
     For Each cell In rng
       ThisWorkbook.Names.Add Name:=baseName & cell.Column, RefersTo:="=" & ws.Name & "!" & cell.Address(False, False)
     Next cell
     ```

4. **使用方式**

   * 打開 VBA 編輯器，插入模組，貼上程式，按 F5 執行即可。
   * 執行後在「名稱管理員」(Name Manager) 中就能看到 `Item1`～`Item9`。

這樣就能夠以最簡潔的迴圈邏輯，為成百上千個儲存格建立有規律的命名！
