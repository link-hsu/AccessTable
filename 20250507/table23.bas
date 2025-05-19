初級市場買入
        - 1-30天
=IF(SUMIF(承銷交易!AF:AF,30,承銷交易!S:S)=0,0,SUMIF(承銷交易!AF:AF,30,承銷交易!AE:AE)/SUMIF(承銷交易!AF:AF,30,承銷交易!S:S))
        - 31-90天
=IF(SUMIF(承銷交易!AF:AF,90,承銷交易!S:S)=0,0,SUMIF(承銷交易!AF:AF,90,承銷交易!AE:AE)/SUMIF(承銷交易!AF:AF,90,承銷交易!S:S))
        - 91-180天
=IF(SUMIF(承銷交易!AF:AF,180,承銷交易!S:S)=0,0,SUMIF(承銷交易!AF:AF,180,承銷交易!AE:AE)/SUMIF(承銷交易!AF:AF,180,承銷交易!S:S))
        - 181-270天
=IF(SUMIF(承銷交易!AF:AF,270,承銷交易!S:S)=0,0,SUMIF(承銷交易!AF:AF,270,承銷交易!AE:AE)/SUMIF(承銷交易!AF:AF,270,承銷交易!S:S))
        - 271-365天
=IF(SUMIF(承銷交易!AF:AF,365,承銷交易!S:S)=0,0,SUMIF(承銷交易!AF:AF,365,承銷交易!AE:AE)/SUMIF(承銷交易!AF:AF,365,承銷交易!S:S))


承銷交易AF欄位
=IF(V2<=30,30,IF(V2<=90,90,IF(V2<=180,180,IF(V2<=270,270,365))))

AE欄位
=S2*U2

V欄位 天數
S欄位 面額
U欄位 成交利率



我會access 資料表中抓資料到Excel中，
然後我要寫迴圈從該資料表中抓資料，需要抓取的資料如上面的Excel函數內容，
請幫我寫vba




下面這隻 VBA 巨集示範如何 **不** 用 `SUMIF`，而是透過迴圈一次跑過整張「承銷交易」表，計算「初級市場買入」在不同天數區間的加權平均成交利率（即 AE總和/面額總和），並把結果寫回你指定的儲存格。

```vb
Sub CalcPrimaryMarketYield()
    Dim ws       As Worksheet
    Dim lastRow  As Long
    Dim i        As Long, j As Long
    Dim days     As Long
    Dim faceVal  As Double, rateVal As Double
    Dim buckets  As Variant
    Dim faceSum () As Double, prodSum() As Double
    Dim yieldRes() As Double
    Dim outputCells As Variant
    
    '── 1. 設定工作表＆找最後一列 ─────────────────────
    Set ws = ThisWorkbook.Worksheets("承銷交易")
    lastRow = ws.Cells(ws.Rows.Count, "V").End(xlUp).Row  ' V 欄是「天數」
    
    '── 2. 定義天數上限和對應輸出儲存格 ───────────────
    '    5 個區間：1–30, 31–90, 91–180, 181–270, 271–365
    buckets = Array(30, 90, 180, 270, 365)
    '    輸出到哪幾個儲存格？請改成你實際要放結果的位址
    outputCells = Array("E2", "E3", "E4", "E5", "E6")
    
    '── 3. 建立累加器陣列 ───────────────────────────────
    ReDim faceSum(0 To UBound(buckets))    ' 面額加總
    ReDim prodSum(0 To UBound(buckets))    ' 面額×利率加總
    ReDim yieldRes(0 To UBound(buckets))   ' 最後要輸出的比值
    
    '── 4. 主迴圈：掃描每一筆交易 ─────────────────────
    For i = 2 To lastRow
        days    = ws.Cells(i, "V").Value
        faceVal = ws.Cells(i, "S").Value      ' S 欄：面額
        rateVal = ws.Cells(i, "U").Value      ' U 欄：成交利率
        
        ' 找到屬於哪個 bucket
        For j = 0 To UBound(buckets)
            If days <= buckets(j) Then
                ' 累加面額與面額×利率
                faceSum(j) = faceSum(j) + faceVal
                prodSum(j) = prodSum(j) + faceVal * rateVal
                Exit For
            End If
        Next j
    Next i
    
    '── 5. 計算加權平均利率（或 0）並寫回儲存格 ───────
    For j = 0 To UBound(buckets)
        If faceSum(j) = 0 Then
            yieldRes(j) = 0
        Else
            yieldRes(j) = prodSum(j) / faceSum(j)
        End If
        ws.Range(outputCells(j)).Value = yieldRes(j)
    Next j
    
    MsgBox "初級市場買入利率已計算完畢！", vbInformation
End Sub
```

---

### 程式說明

1. **掃描整張表一次就能完成所有區間的累加**
   我們定義了一個 `buckets` 陣列，裡面放各天數上限（30、90、180、270、365）。
   每掃描到一筆交易，就用這五個上限去判斷它落在哪個區間，然後把面額 (`S` 欄) 和「面額×利率」累加到對應的 bucket。

2. **三組陣列管理數值**

   * `faceSum(j)`：第 j 個區間的面額總和
   * `prodSum(j)`：第 j 個區間的面額×利率總和
   * 最後再算 `prodSum(j)/faceSum(j)`，就是你原本公式的分子／分母。

3. **完全不靠 `SUMIF`**
   迴圈內全部以 VBA 變數累加，速度上若資料筆數不超過數萬筆，其實效能相當足。

4. **輸出位置可自行調整**
   範例預設把 1–30 天的結果寫到 `E2`，31–90 到 `E3`，…，271–365 到 `E6`。
   如果你在 Excel 上想放在別的地方，只要把程式開頭的

   ```vb
   outputCells = Array("E2","E3","E4","E5","E6")
   ```

   改成你要的儲存格陣列即可。

只要把這段貼進 VBA 編輯器並執行，就可以一次跑出所有區間的初級市場買入加權平均利率！
