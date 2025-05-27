' 將SQL印出來丟進去SQL跑看看資料有沒有抓到



在 VBA 裡 ，當你把一個值為 `Null` 的變數，丟到需要「字串」輸入的函數（像是 `Trim`、`Len`、`InStr`…）裡，就會跳出「 **無法正確使用 Null** 」的錯誤訊息。

在你貼的 `clsReport.Init` 裡，最可能出問題的，就是這一段：

```vb
sheetName = addressMap(i, 0)
nameTag   = addressMap(i, 1)
addr      = addressMap(i, 2)

If Len(Trim(nameTag)) > 0 And Len(Trim(addr)) > 0 Then
    ’… 
End If
```

其中 `addressMap = GetMapData(...)` 取回的陣列裡，如果某一筆 row 的第 1 欄（`nameTag`）或第 2 欄（`addr`）本身就是 `Null`，你一執行到 `Trim(nameTag)` 或 `Len(Trim(addr))`，就會馬上跳「Invalid use of Null」。

---

## 可能的原因

1. **Access 資料庫裡的對應表 (field\_value) 有欄位留空**
   你用 `GetMapData` 把〈工作表名稱／欄位代號〉抓出來，如果有人在 Access 那邊留空 (`Null`)，VBA 就拿不出字串來做 `Trim`。

2. **`GetMapData` 在找不到任何紀錄時直接回傳 `Null` 而不是「空陣列」**
   如果完全沒有任何 mapping，`addressMap` 不是一個空陣列，而是 `Null`，第一個 `IsArray(addressMap)` 雖然過不了，不過在之後對 `addressMap(i,…)` 的取用就會惹錯。

---

## 解法

### 1. 先過濾掉 `Null` 值

在取用前，先把可能的 `Null` 強制轉成空字串：

```vb
sheetName = addressMap(i, 0) & ""  
nameTag   = addressMap(i, 1) & ""  
addr      = addressMap(i, 2) & ""
```

這樣即使資料庫裡那個欄位是 `Null`，`& ""` 也會自動把它變成空字串 `""`，就不會再丟到 `Trim` 裡出錯。

完整範例：

```vb
For i = 0 To UBound(addressMap, 1)
    sheetName = addressMap(i, 0) & ""
    nameTag   = addressMap(i, 1) & ""
    addr      = addressMap(i, 2) & ""
    
    If Len(Trim(nameTag)) > 0 And Len(Trim(addr)) > 0 Then
        ’ …
    End If
Next i
```

---

### 2. 檢查 `addressMap` 是否為 `Null`

在 `Init` 一開始就加一行保護程式，若 `GetMapData` 回傳 `Null`，直接跳出：

```vb
addressMap = GetMapData(gDBPath, reportName, "field_value")
If IsNull(addressMap) Then
    WriteLog "GetMapData 回傳 Null，無法初始化報表：" & reportName
    Exit Sub
End If
```

---

### 3. 確保 Access 表裡的對應資料完整

最根本的方式，是到你那支 `GetMapData` 所讀的 Access table（或 query）裡，檢查：

* **`sheetName`、`nameTag`、`addr` 這三個欄位都不可以留空。**
* 如果有一筆空白，先用 Access 裡的 UPDATE 把它改成某個預設值，或是乾脆刪掉那筆 mapping。

---

### 小結

「Invalid use of Null」通常就是你在對一個 `Null` 執行字串函數。
最簡單的防禦式寫法，就是拿到每個可能會是 `Null` 的變數時，立刻做：

```vb
myVar = myVar & ""
```

強制轉成 `""`，就再也不會跑到 `Trim(Null)` 或 `Len(Null)` 的坑裡了。

把上面的修法加到 `clsReport.Init` 裡，應該就能解掉你那行 `rpt.Init …` 的錯。
